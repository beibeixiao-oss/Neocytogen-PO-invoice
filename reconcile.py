"""
reconcile.py — PDF invoice × Procurement Tracking List 对账主程序

    python reconcile.py <tracking_list.xlsx> <invoice_pdf_folder> [输出.xlsx]

输出四个 sheet，表头统一（12 列模板 + Source + Note）：
    matched                  两边都有且金额一致
    discrepancy              两边都有但金额对不上
    PDF to Excel - not match 有发票、Excel 里查无此单
    Excel to PDF - not match Excel 说已开票、但没找到 PDF
"""

__version__ = "2026-09-11.3"

import os
import re
import sys
import glob

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment
from openpyxl.utils import get_column_letter

from excel_loader import load_tracking_list, normalize_invoice_no, GST_RATE
from invoice_extractor import extract_invoice

TEMPLATE_COLS = [
    "Invoice Number", "PO Number", "Supplier", "Invoice date", "Due Date",
    "Description", "Unit no", "Currency", "Unit Price",
    "Amount excl. GST", "GST", "Amount incl. GST",
]
OUT_COLS = TEMPLATE_COLS + ["Source", "Note"]
TOLERANCE = 0.05          # 金额容差：行级 GST 各家舍入方式不同，只在发票合计层面卡


def pdf_rows(inv, source, note=""):
    """一张发票 -> 若干行（每个 line item 一行），字段对齐模板"""
    rows = []
    for it in inv["line_items"]:
        rows.append({
            "Invoice Number": inv.get("invoice_no"),
            "PO Number": inv.get("po_no"),
            "Supplier": inv.get("supplier"),
            "Invoice date": inv.get("invoice_date"),
            "Due Date": inv.get("due_date"),
            "Description": it["description"],
            "Unit no": it["qty"],
            "Currency": inv.get("currency"),
            "Unit Price": it["unit_price"],
            "Amount excl. GST": it["amount_excl"],
            "GST": it["gst"],
            "Amount incl. GST": it["amount_incl"],
            "Source": source,
            "Note": note,
        })
    if not rows and (inv.get("subtotal") or {}).get("amount_incl") is not None:
        s = inv["subtotal"]
        rows.append({c: None for c in OUT_COLS} | {
            "Invoice Number": inv.get("invoice_no"), "PO Number": inv.get("po_no"),
            "Supplier": inv.get("supplier"), "Invoice date": inv.get("invoice_date"),
            "Due Date": inv.get("due_date"), "Description": None,
            "Currency": inv.get("currency"),
            "Amount excl. GST": s.get("amount_excl"), "GST": s.get("gst"),
            "Amount incl. GST": s.get("amount_incl"),
            "Source": source,
            "Note": "；".join(x for x in [note, "仅取到发票合计，无明细行"] if x),
        })
        return rows
    if not rows:      # 没抽到 line item 也要留痕，否则这张发票会凭空消失
        rows.append({c: None for c in OUT_COLS} | {
            "Invoice Number": inv.get("invoice_no"), "PO Number": inv.get("po_no"),
            "Supplier": inv.get("supplier"), "Source": source,
            "Note": (note + " / " if note else "") + "未抽到明细行",
        })
    return rows


def merged_rows(inv, grp, note=""):
    """两边都有的发票：明细取台账（Item Description 是人写的，比从 PDF 抠出来的规范），
    发票号 / PO / 供应商 / 日期 / 币种取 PDF（发票是付款依据）。"""
    rows = []
    for _, r in grp.iterrows():
        rows.append({
            "Invoice Number": inv.get("invoice_no"),
            "PO Number": inv.get("po_no") or r["PO Number"],
            "Supplier": inv.get("supplier") or r["Supplier"],
            "Invoice date": inv.get("invoice_date") or r["Invoice date"],
            "Due Date": inv.get("due_date") or r["Due Date"],
            "Description": r["Description"],
            "Unit no": r["Unit no"],
            "Currency": inv.get("currency"),
            "Unit Price": r["Unit Price"],
            "Amount excl. GST": r["Amount excl. GST"],
            "GST": r["GST"],
            "Amount incl. GST": r["Amount incl. GST"],
            "Source": "Both",
            "Note": note,
        })
    return rows


def excel_rows(df, note=""):
    out = df[TEMPLATE_COLS].copy()
    out["Source"] = "Excel"
    out["Note"] = note
    return out.to_dict("records")


def find_candidates(expected, inv, limit=5):
    """发票在台账里找不到时，列出最可能的候选行供人工判断。

    只在「同一供应商」内找 —— BioLabs 与 BioBasic 名字相近却是两家公司，
    跨供应商排序会让人逐条核对，比不给建议还费时间。
    注意：这只是给人看的参考，程序绝不据此自动匹配（发票号必须精确相等）。
    """
    from rapidfuzz import fuzz
    from excel_loader import normalize_supplier

    sup = normalize_supplier(inv.get("supplier") or "")
    date = inv.get("invoice_date")
    total = (inv.get("subtotal") or {}).get("amount_incl")
    if not sup:
        return []

    rows = []
    for key, grp in expected.groupby("invoice_key"):
        row = grp.iloc[0]
        if not row["supplier_key"]:
            continue
        sim = max(fuzz.ratio(sup, row["supplier_key"]),
                  fuzz.partial_ratio(sup, row["supplier_key"]))
        if sim < 85:                       # BioLab vs BioBasic 恰好卡在 80，阈值必须高于它
            continue
        amt = round(grp["Amount incl. GST"].sum(), 2)
        diff = None if total is None else round(amt - total, 2)
        gap = None
        if date is not None and pd.notna(row["Invoice date"]):
            gap = abs((pd.Timestamp(date) - pd.Timestamp(row["Invoice date"])).days)
        # 金额一致最有说服力（多半只是发票号登错），其次是日期接近
        score = (0 if diff is None else max(0, 60 - min(abs(diff), 60))) \
                + (0 if gap is None else max(0, 30 - gap)) + sim * 0.1
        rows.append({
            "Excel Invoice Number": row["Invoice Number"],
            "Supplier": row["Supplier"],
            "Invoice date": row["Invoice date"],
            "Amount incl. GST": amt,
            "金额差": diff,
            "日期差(天)": gap,
            "_s": score,
        })
    rows.sort(key=lambda r: r["_s"], reverse=True)
    for r in rows:
        r.pop("_s")
    return rows[:limit]


def flag_suspicious(invoices, expected):
    """自动挑出「值抽到了但可能抽错」的发票，按可疑程度排序。

    抽不到会报错，抽错了不会 —— 后者才危险。这里用同供应商的历史分布做基准：
    金额差一两个数量级、发票号格式与同行不一致、日期落在合理区间之外，都要人看一眼。
    """
    import statistics
    from rapidfuzz import fuzz
    from excel_loader import normalize_supplier

    # 按供应商归集台账里的金额与发票号格式，作为比较基准
    by_sup = {}
    for _, r in expected.iterrows():
        k = r["supplier_key"]
        if not k:
            continue
        by_sup.setdefault(k, {"amts": [], "shapes": set(), "numbers": set()})
        if pd.notna(r["Amount incl. GST"]):
            by_sup[k]["amts"].append(float(r["Amount incl. GST"]))
        by_sup[k]["shapes"].add(_shape(r["Invoice Number"]))
        n = r["Invoice Number"]
        if n is not None and not (isinstance(n, float) and pd.isna(n)):
            n = str(int(n)) if isinstance(n, float) and n.is_integer() else str(n).strip()
            by_sup[k]["numbers"].add(n)

    out = []
    for inv in invoices.values():
        reasons = []
        sup = normalize_supplier(inv.get("supplier") or "")
        ref = None
        for k, v in by_sup.items():
            if sup and max(fuzz.ratio(sup, k), fuzz.partial_ratio(sup, k)) >= 85:
                ref = v
                break

        total = (inv.get("subtotal") or {}).get("amount_incl")
        if total is None:
            reasons.append("未取到发票合计")
        elif ref and len(ref["amts"]) >= 3:
            med = statistics.median(ref["amts"])
            if med > 0 and (total < med / 10 or total > med * 10):
                reasons.append(f"金额 {total} 与该供应商中位数 {round(med, 2)} 相差十倍以上")

        shape = _shape(inv.get("invoice_no"))
        # Proforma 的号码由供应商另一套规则生成（如 PI- 对 INV-），格式不同属正常
        if ref and ref["shapes"] and shape not in ref["shapes"] and not inv.get("is_proforma"):
            msg = f"发票号格式 {shape} 与该供应商台账中的 {'/'.join(sorted(ref['shapes'])[:3])} 不一致"
            fix = suggest_repair(inv.get("invoice_no"), ref.get("numbers", set()))
            if fix:
                msg += f"；{fix}"
            reasons.append(msg)

        d = inv.get("invoice_date")
        if d is None:
            reasons.append("未取到发票日期")

        s = inv.get("subtotal") or {}
        e, g, i = s.get("amount_excl"), s.get("gst"), s.get("amount_incl")
        if None not in (e, g, i) and abs(e + g - i) > 0.05:
            reasons.append("税前+税额 ≠ 税后")
        if i and e and i > 0 and not (0 <= (i - e) / i < 0.15):
            reasons.append(f"税率异常：{round((i - e) / i * 100, 1)}%")

        if inv.get("ai"):
            reasons.append("由视觉模型识别，数字需人工复核")
        if inv.get("ocr"):
            reasons.append("扫描件经 OCR 识别，数字需人工复核")
        # 按用户要求：Proforma 跟正常 Tax Invoice 一样处理，不再仅因为是 Proforma
        # 就单独进「待核查」——它本来就正常参与对账（见 README「关于 matched 与
        # 待核查」），这里只是不再额外拿这一条刷疑点数。真正该被挑出来复核的
        # 还是金额等式、税率、OCR/AI 来源这些信号，跟是不是 Proforma 无关。
        if not inv.get("supplier"):
            reasons.append("未取到供应商")

        if reasons:
            out.append({
                "发票号": inv.get("invoice_no"), "供应商": inv.get("supplier"),
                "发票日期": inv.get("invoice_date"), "币种": inv.get("currency"),
                "合计": total, "疑点数": len(reasons),
                "需要核查的原因": "；".join(reasons), "文件": inv.get("source_file"),
            })
    out.sort(key=lambda r: r["疑点数"], reverse=True)
    return out


# OCR 最常混淆的字符对。只在「疑似识别错误」时用来试探，绝不做全文替换。
_CONFUSIONS = [("O", "0"), ("I", "1"), ("L", "1"), ("S", "5"),
               ("B", "8"), ("Z", "2"), ("G", "6"), ("Q", "0"), ("D", "0")]
_CONFUSION_SET = {p for a, b in _CONFUSIONS for p in ((a, b), (b, a))}


def suggest_repair(raw, candidates, limit=3):
    """发票号疑似抽错时，在同供应商的台账号码里找出可能的正确值。

    两种情形：
      1) OCR 字符混淆 —— LBHO00 里的 O 其实是数字 0
      2) 被截断 —— INV-26 其实是 INV-26-03283

    逐位比对而非整串替换：整串替换会把 INV 里的 I 也换成 1，反而破坏号码。
    只有长度相同、且所有不同的位置都落在混淆字符对里，才算候选。

    永远只是建议。程序不会据此自动匹配，发票号必须完全相等才算 matched。
    """
    if not raw or not candidates:
        return None
    s = str(raw).strip()

    # 情形 1：逐位比对
    hits = []
    for c in candidates:
        if c == s or len(c) != len(s):
            continue
        diff = [(a, b) for a, b in zip(s.upper(), c.upper()) if a != b]
        if diff and len(diff) <= 3 and all(pair in _CONFUSION_SET for pair in diff):
            hits.append(c)
    if hits:
        return "疑似 OCR 字符误识，台账中有 " + " / ".join(sorted(hits)[:limit])

    # 情形 2：抽到的值是某个台账号码的前缀
    pre = sorted(c for c in candidates if c != s and c.startswith(s) and len(s) >= 4)
    if pre:
        tail = " / ".join(pre[:limit]) + ("…" if len(pre) > limit else "")
        return f"疑似被截断，台账中有 {tail}"

    return None


def _shape(v):
    """把发票号抽象成格式：97809367 -> 99999999 ； INV-0386 -> AAA-9999"""
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return ""
    s = str(v).strip()
    if s.endswith(".0"):
        s = s[:-2]
    return re.sub(r"[A-Za-z]", "A", re.sub(r"\d", "9", s))


def reconcile(xlsx_path, pdf_source):
    """pdf_source 可以是文件夹路径，也可以是 pdf 路径列表"""
    expected, pending = load_tracking_list(xlsx_path)

    if isinstance(pdf_source, (str, os.PathLike)):
        pdf_files = sorted(glob.glob(os.path.join(pdf_source, "*.pdf")))
    else:
        pdf_files = list(pdf_source)

    invoices, failed = {}, []
    all_parsed = []
    for f in pdf_files:
        inv = extract_invoice(f)
        all_parsed.append(inv)
        key = normalize_invoice_no(inv.get("invoice_no"))
        if not key:
            failed.append(inv)
            continue
        invoices[key] = inv

    superseded = []
    po_has_tax = {
        normalize_invoice_no(i.get("po_no"))
        for i in all_parsed
        if not i.get("is_proforma") and i.get("po_no")
    }
    for key in list(invoices):
        inv = invoices[key]
        po = normalize_invoice_no(inv.get("po_no"))
        if inv.get("is_proforma") and po and po in po_has_tax:
            superseded.append({
                "发票号": inv.get("invoice_no"), "PO": inv.get("po_no"),
                "供应商": inv.get("supplier"),
                "合计": (inv.get("subtotal") or {}).get("amount_incl"),
                "忽略原因": "同一 PO 已有 Tax Invoice，按规则只算 Tax Invoice",
                "文件": inv.get("source_file"),
            })
            del invoices[key]

    matched, discrepancy, pdf_only, excel_only = [], [], [], []

    # 两级匹配。
    # 台账里 PO 号 59 个、发票号只有 54 个且格式各异，而每张 PDF 的文件名都带 PO，
    # 所以 PO 才是两边都完整规范的键。发票号仍优先——它更精确、能区分同一 PO 的多张发票；
    # 发票号对不上时退回 PO，并在 Note 里注明用了哪一级，保持可追溯。
    by_inv, by_po = {}, {}
    for _, r in expected.iterrows():
        if r["invoice_key"]:
            by_inv.setdefault(r["invoice_key"], []).append(r)
        if r["po_key"]:
            by_po.setdefault(r["po_key"], []).append(r)

    # OCR 混淆折叠索引。Tesseract 在发票号上最常见的错认是 O/0、I/l/1、S/5、B/8，
    # 实测 INV260123-LBH001 被读成 INV260123-LBHOO1 —— 格式合法、校验能过，
    # 所以 fix_ocr_digits 不触发，精确匹配又对不上，白白掉进「查无此单」。
    # 折叠后若在台账中唯一命中才认，命中多条则宁可不认，交人工。
    def _fold(k):
        return (str(k or "").upper().replace("O", "0").replace("I", "1")
                .replace("L", "1").replace("S", "5").replace("B", "8"))

    folded = {}
    for k in by_inv:
        folded.setdefault(_fold(k), []).append(k)

    used_inv, used_po = set(), set()
    # 同一 PO 可能对应多张发票（分批开票），先数一下，多张时不能整组认领
    po_pdf_count = {}
    for inv in list(invoices.values()) + failed:
        k = normalize_invoice_no(inv.get("po_no"))
        if k:
            po_pdf_count[k] = po_pdf_count.get(k, 0) + 1

    # 没抽到发票号的（多为扫描件）也一并参与：它们的文件名 PO 依然可靠
    candidates = list(invoices.items()) + [("", i) for i in failed]
    failed = []

    # ---- 分批开票：同一 PO 多张发票，先加总再比 ----
    # 逐张比一定不符（每张只是总额的一部分），会把本来正确的账报成异常。
    # 先按 PO 把这些发票的合计加起来与台账整组比：对得上 -> 整组 matched；
    # 对不上 -> 退回逐张处理，让人看清楚是哪一张出了问题。
    batch_ok = {}          # po_key -> 该 PO 下所有发票（已确认加总相符）
    _by_po_pdf = {}
    for key, inv in candidates:
        if key and (key in by_inv
                    or (inv.get("ocr") and len(folded.get(_fold(key), [])) == 1)):
            continue       # 发票号能直接对上的，不走 PO 聚合
        pk = normalize_invoice_no(inv.get("po_no"))
        if pk and pk in by_po:
            _by_po_pdf.setdefault(pk, []).append((key, inv))

    for pk, group in _by_po_pdf.items():
        if len(group) < 2:
            continue
        totals = [(i.get("subtotal") or {}).get("amount_incl") for _, i in group]
        if any(t is None for t in totals):
            continue       # 有一张没抽到合计，加总没有意义
        x_incl = round(pd.DataFrame(by_po[pk])["Amount incl. GST"].sum(), 2)
        if abs(round(sum(totals), 2) - x_incl) <= TOLERANCE:
            batch_ok[pk] = group

    for pk, group in batch_ok.items():
        used_po.add(pk)
        gdf = pd.DataFrame(by_po[pk])
        n = len(group)
        for _, inv in group:
            t = (inv.get("subtotal") or {}).get("amount_incl")
            matched.extend(merged_rows(
                inv, gdf,
                f"按 PO 号匹配：该 PO 共 {n} 张发票分批开具，"
                f"合计 {round(sum((i.get('subtotal') or {}).get('amount_incl') for _, i in group), 2)}"
                f" 与台账一致（本张 {t}）"))
    _batched = {id(i) for g in batch_ok.values() for _, i in g}
    candidates = [(k, i) for k, i in candidates if id(i) not in _batched]

    # 同一 PO 下多张发票都没能识别出发票号、加总也对不上台账（上面「分批开票」
    # 那步失败）：以前这种情况会让每一张都去跟台账里没认领的全部行比总和，
    # 报出来的差额是「台账整组之和 - 这一张」，既没意义还会连累明明对得上的那张
    # 也被判成 mismatch（真实案例：同一 PO 两张洗衣发票，其中一张金额和台账某一
    # 行完全一致，只因为另一张对不上就被一起报错，Note 里的台账合计跟这张发票
    # 实际对应的那行对不上号，容易让人怀疑是不是台账录错了）。
    # 张数和台账未认领行数对得上时，改成按金额就近一对一配对：每张发票配它金额
    # 最接近的那一行台账，再各自判断是否在容差内——这样报出来的差额才是这一张
    # 真正对应哪一行、差多少。张数对不上（更常见是行数与发票数不等）时不做这个
    # 配对，退回原来的整组比较，避免瞎猜配错。
    _paired = set()
    for pk, group in _by_po_pdf.items():
        if pk in batch_ok or len(group) < 2:
            continue
        rows = by_po.get(pk, [])
        if len(rows) != len(group):
            continue
        if any((i.get("subtotal") or {}).get("amount_incl") is None for _, i in group):
            continue
        remaining_rows = list(rows)
        for _key, inv in sorted(group, key=lambda kv: (kv[1].get("subtotal") or {}).get("amount_incl")):
            p_incl = (inv.get("subtotal") or {}).get("amount_incl")
            # 注意：rows 里是 pandas Series，list.remove() 靠 == 比较会在多值场景下
            # 抛 "truth value of a Series is ambiguous"，改成按下标 pop，避免这个坑。
            best_idx = min(range(len(remaining_rows)),
                           key=lambda idx: abs(remaining_rows[idx]["Amount incl. GST"] - p_incl))
            best = remaining_rows.pop(best_idx)
            gdf = pd.DataFrame([best])
            x_incl = best["Amount incl. GST"]
            tag = (f"按金额就近配对（该 PO 下 {len(group)} 张发票均未能识别出发票号，"
                   f"按合计金额分别配对台账对应行，而非整组合计比较）")
            if abs(x_incl - p_incl) <= TOLERANCE:
                matched.extend(merged_rows(inv, gdf, tag))
            else:
                diff = round(p_incl - x_incl, 2)
                discrepancy.extend(merged_rows(
                    inv, gdf, "；".join([tag, f"金额不符：台账 {x_incl} vs 发票 {p_incl}（差 {diff}）"])))
            _paired.add(id(inv))
        used_po.add(pk)
    candidates = [(k, i) for k, i in candidates if id(i) not in _paired]

    # 预扫一遍：先确定哪些台账行会被发票号直接认领。
    # 必须放在主循环之前 —— 否则「已认领」取决于遍历顺序，
    # 同一 PO 的两张发票谁先被处理，结果就不一样。
    claimed_inv = set()
    for _k, _i in candidates:
        if _k and _k in by_inv:
            claimed_inv.add(_k)
        elif _k and _i.get("ocr") and len(folded.get(_fold(_k), [])) == 1:
            claimed_inv.add(folded[_fold(_k)][0])

    for key, inv in candidates:
        grp, level = None, None
        if key and key in by_inv:
            grp, level = by_inv[key], "发票号"
            used_inv.add(key)
        elif key and inv.get("ocr") and len(folded.get(_fold(key), [])) == 1:
            real = folded[_fold(key)][0]
            grp, level = by_inv[real], f"发票号（OCR 字符纠正：{key} → {real}，请人工确认）"
            used_inv.add(real)
            # 之前这里只用 real 去台账里取匹配的行，抽取结果本身（inv["invoice_no"]）
            # 从没改过，导致 outcome.xlsx 和「待核查」里显示的仍是 OCR 读错的原始号码
            # （比如 LBHOO1），配对成功了但号码本身还是错的，容易被当成没修好。
            # inv 和 invoices 字典引用的是同一个对象，这里改了后面 flag_suspicious
            # 用到的也是修正后的号码，格式比对不会再误报「与台账不一致」。
            n = by_inv[real][0]["Invoice Number"]
            fixed_no = (str(int(n)) if isinstance(n, float) and n.is_integer() else str(n).strip()) if n is not None else None
            if fixed_no and normalize_invoice_no(fixed_no) == real:
                old_no = inv.get("invoice_no")
                if fixed_no != old_no:
                    inv["invoice_no"] = fixed_no
                    inv.setdefault("warnings", []).append(
                        f"发票号经 OCR 字符纠正：{old_no} → {fixed_no}（依据台账唯一匹配），请人工确认")
        else:
            pk = normalize_invoice_no(inv.get("po_no"))
            if pk in by_po:
                # 同一 PO 下的其他发票可能已按发票号认领了对应台账行。
                # 若仍拿台账整组来比，剩下这张必然对不上（差额正好是别人那份），
                # 报出来的 "台账 577.7 vs 发票 957.52" 毫无参考价值。
                # 所以只跟尚未被认领的行比。
                rest = [r for r in by_po[pk]
                        if not r["invoice_key"] or r["invoice_key"] not in claimed_inv]
                if rest:
                    grp = rest
                    level = ("PO 号" if len(rest) == len(by_po[pk])
                             else "PO 号（已扣除同 PO 下按发票号认领的行）")
                else:
                    grp, level = by_po[pk], "PO 号（该 PO 台账行已全部被认领，请人工确认）"
                used_po.add(pk)

        if grp is None:
            po = str(inv.get("po_no") or "")
            hint = ("PO 为 2025 年，请核对 2025 年度台账"
                    if re.match(r"PONCG2025", po, re.I) else "台账中查无此 PO 与发票号")
            pdf_only.extend(pdf_rows(inv, "PDF", hint))
            continue

        gdf = pd.DataFrame(grp)
        x_incl = round(gdf["Amount incl. GST"].sum(), 2)
        p_incl = (inv.get("subtotal") or {}).get("amount_incl")
        tag = "" if level == "发票号" else f"按 {level} 匹配"
        if not key and tag:
            tag += "（发票号未能识别）"

        if p_incl is None:
            discrepancy.extend(merged_rows(inv, gdf, "；".join(x for x in [tag, "未抽到发票合计，无法比对金额"] if x)))
        elif abs(x_incl - p_incl) <= TOLERANCE:
            matched.extend(merged_rows(inv, gdf, tag))
        else:
            diff = round(p_incl - x_incl, 2)
            note = f"金额不符：台账 {x_incl} vs 发票 {p_incl}（差 {diff}）"
            discrepancy.extend(merged_rows(inv, gdf, "；".join(x for x in [tag, note] if x)))

    # 台账有、但没找到对应 PDF 的
    from excel_loader import STATUS_INVOICE_EXPECTED
    still_pending = []
    for _, r in expected.iterrows():
        if r["invoice_key"] in used_inv or r["po_key"] in used_po:
            continue
        if r["_status"] not in STATUS_INVOICE_EXPECTED:
            still_pending.append(r)          # 货未到，本就不该有发票
        else:
            excel_only.extend(excel_rows(pd.DataFrame([r]), "台账已记录，但未找到对应发票 PDF"))

    for inv in failed:
        pdf_only.extend(pdf_rows(inv, "PDF", "未能抽出发票号：" + "；".join(inv["warnings"])))

    sheets = {
        "matched": matched,
        "discrepancy": discrepancy,
        "PDF to Excel - not match": pdf_only,
        "Excel to PDF - not match": excel_only,
    }
    context = {"expected": expected, "invoices": invoices,
               "superseded": superseded,
               "suspicious": flag_suspicious(invoices, expected),
               "unmatched_pdf": [i for k, i in invoices.items()
                                 if k not in used_inv
                                 and normalize_invoice_no(i.get("po_no")) not in used_po],
               "failed": failed, "n_pdf": len(pdf_files)}
    if still_pending:
        pending = pd.concat([pending, pd.DataFrame(still_pending)], ignore_index=True)
    return sheets, pending, context


def write_output(sheets, pending, path, suspicious=None, superseded=None):
    wb = Workbook()
    wb.remove(wb.active)
    order = ["matched", "discrepancy", "PDF to Excel - not match", "Excel to PDF - not match"]
    for name in order:
        ws = wb.create_sheet(name[:31])
        ws.append(OUT_COLS)
        for c in ws[1]:
            c.font = Font(name="Arial", bold=True)
            c.alignment = Alignment(vertical="center")
        for r in sheets[name]:
            ws.append([r.get(c) for c in OUT_COLS])
        for col in range(1, len(OUT_COLS) + 1):
            L = get_column_letter(col)
            for cell in ws[L]:
                if cell.row > 1:
                    cell.font = Font(name="Arial")
            ws.column_dimensions[L].width = 34 if OUT_COLS[col-1] in ("Description", "Note") else 17
        ws.freeze_panes = "A2"

    ws = wb.create_sheet("已忽略的 Proforma")
    cols2 = ["发票号", "PO", "供应商", "合计", "忽略原因", "文件"]
    ws.append(cols2)
    for c in ws[1]:
        c.font = Font(name="Arial", bold=True)
    for r in (superseded or []):
        ws.append([r.get(c) for c in cols2])
    for n, L in enumerate("ABCDEF"):
        ws.column_dimensions[L].width = 46 if L in "EF" else 18
    ws.freeze_panes = "A2"

    ws = wb.create_sheet("待核查")
    cols = ["发票号", "供应商", "发票日期", "币种", "合计", "疑点数", "需要核查的原因", "文件"]
    ws.append(cols)
    for c in ws[1]:
        c.font = Font(name="Arial", bold=True)
    for r in (suspicious or []):
        ws.append([r.get(c) for c in cols])
    for n, L in enumerate("ABCDEFGH"):
        ws.column_dimensions[L].width = 52 if L in "GH" else 16
    ws.freeze_panes = "A2"

    ws = wb.create_sheet("pending (no invoice yet)")
    ws.append(["Invoice Number", "PO Number", "Supplier", "Description", "Unit no",
               "Amount incl. GST", "Status", "Excel row"])
    for c in ws[1]:
        c.font = Font(name="Arial", bold=True)
    for _, r in pending.iterrows():
        ws.append([r["Invoice Number"], r["PO Number"], r["Supplier"], r["Description"],
                   r["Unit no"], r["Amount incl. GST"], r["_status"], r["_row"]])
    for col in range(1, 9):
        ws.column_dimensions[get_column_letter(col)].width = 30 if col == 4 else 17
    ws.freeze_panes = "A2"

    wb.save(path)


if __name__ == "__main__":
    xlsx, pdf_dir = sys.argv[1], sys.argv[2]
    out = sys.argv[3] if len(sys.argv) > 3 else "Neocytogen - outcome.xlsx"
    sheets, pending, ctx = reconcile(xlsx, pdf_dir)
    write_output(sheets, pending, out, ctx["suspicious"], ctx["superseded"])
    for k, v in sheets.items():
        print(f"{k:28s} {len(v):4d} 行")
    print(f"{'已忽略 Proforma':25s} {len(ctx['superseded']):4d} 张")
    print(f"{'待核查':29s} {len(ctx['suspicious']):4d} 张")
    print(f"{'pending (未开票)':26s} {len(pending):4d} 行")
    print("已写出:", out)
