"""
app.py — 采购发票对账工具（界面版）

    streamlit run app.py

左边传 Procurement Tracking List，右边传发票 PDF（可多选）。
传一张 = 单张核对，传一整个月 = 批量对账，走的是同一套逻辑。

交互式编辑（2026-09-11.5 起）：
    对账结果第一次算出来后存进 st.session_state，后面点"移到 matched"
    "已核实"这些按钮触发的都是页面重跑（Streamlit 的机制），如果每次重跑
    都重新调 reconcile()，编辑过的东西全部作废。所以本文件下半部分全部
    对 st.session_state["reconciled"] 这份可变状态操作，只有再点一次
    "开始对账" 才会用新上传的文件重新生成、扔掉旧的编辑。这些编辑只在
    本次会话里有效——下载 outcome.xlsx 时会带上，但刷新页面/重新上传
    同一批文件不会记得之前做过的操作（按用户要求，暂不做跨会话持久化）。
"""

__version__ = "2026-09-11.5"

import os
import tempfile
from io import BytesIO

import pandas as pd
import streamlit as st

# Streamlit Community Cloud 没有 .env 文件，key 存在 App -> Settings -> Secrets 里。
# ai_extract.py 只认 os.environ / 本地 .env，这里把 Secrets 桥接成环境变量，
# 本地开发（用 .env）和云端部署（用 st.secrets）走的是同一份代码不用改。
if "ANTHROPIC_API_KEY" not in os.environ:
    try:
        os.environ["ANTHROPIC_API_KEY"] = st.secrets["ANTHROPIC_API_KEY"]
    except (KeyError, FileNotFoundError):
        pass

from reconcile import reconcile, find_candidates, write_output, OUT_COLS

NUMERIC_COLS = ("Unit no", "Unit Price", "Amount excl. GST", "GST", "Amount incl. GST")
DATE_COLS = ("Invoice date", "Due Date")


def _safe_df(rows):
    """安全转成表格：object 列一律转字符串，避免 Arrow 类型冲突。
    rows 可以是 list[dict] 也可以是现成的 DataFrame。"""
    df = pd.DataFrame(rows).copy()
    for c in df.columns:
        if df[c].dtype == "object":
            df[c] = df[c].map(lambda v: "" if v is None or (isinstance(v, float) and pd.isna(v))
                              else (str(int(v)) if isinstance(v, float) and v.is_integer() else str(v)))
    return df


def show(df, **kw):
    st.dataframe(_safe_df(df), width="stretch", hide_index=True, **kw)


def _to_out_rows(rows):
    """把任意来源的行（sheets 里的 dict，或 pending DataFrame 转出来的 dict）
    统一成「只含 OUT_COLS」的干净 dict list，缺的字段补 None。这样 matched /
    discrepancy / pdf_only / excel_only / pending 五份数据格式完全一致，
    下面的编辑、分组、移动逻辑才能共用同一套代码，不用为 pending 的列名
    差异（原本没有 Source/Note）单独写一套。"""
    return [{c: r.get(c) for c in OUT_COLS} for r in rows]


def _ensure_out_cols(r):
    """确保 dict 至少包含 OUT_COLS 的每一个字段（缺的补 None），但不动其余
    多出来的字段。pending 从 excel_loader 出来时压根没有 Source/Note 这两列——
    直接转 dict 会缺这两个键，渲染时选 OUT_COLS 会直接 KeyError（这个坑是写
    自动化测试时才测出来的：本地随手点一下不一定会踩到，得真有 pending 记录
    才会炸）。同时要保留 _status/_row，write_output() 写 pending sheet 要用。"""
    out = dict(r)
    for c in OUT_COLS:
        out.setdefault(c, None)
    return out


def _coerce_row(r):
    """st.data_editor 编辑完吐出来的是字符串（_safe_df 把所有列都转成了字符串），
    金额/数量/日期这几列要转回真正的数值/日期类型再放进 matched——否则导出的
    outcome.xlsx 和 Xero 清单里这些格子会变成文本，不能被当数字/日期用，
    Excel 里没法求和、Xero 那边大概率也认不出。"""
    out = dict(r)
    for c in NUMERIC_COLS:
        v = out.get(c)
        if v in (None, ""):
            out[c] = None
        else:
            try:
                out[c] = float(str(v).replace(",", "").strip())
            except (TypeError, ValueError):
                pass
    for c in DATE_COLS:
        v = out.get(c)
        if v in (None, ""):
            out[c] = None
        else:
            ts = pd.to_datetime(v, errors="coerce")
            out[c] = ts.date() if pd.notna(ts) else None
    return out


def _group_contiguous(rows, keys=("Invoice Number", "PO Number", "Note")):
    """把列表里连续且 (keys) 相同的行分成一组——同一张发票的多个明细行是
    merged_rows()/pdf_rows() 一次性 extend 进去的，在列表里天然是连续的一段，
    按这个分组就能把"编辑/移动"作用在整张发票上，而不是拆散成单独的行。"""
    groups = []
    for row in rows:
        k = tuple(row.get(x) for x in keys)
        if groups and groups[-1]["key"] == k:
            groups[-1]["rows"].append(row)
        else:
            groups.append({"key": k, "rows": [row]})
    return [g["rows"] for g in groups]


def render_movable_cases(groups, state, source_key, title_fn, note_fn, key_prefix,
                          button_label="✅ 确认并移到 matched"):
    """discrepancy / 台账查无此单 / 台账有单缺发票 / 未开票 四个 tab 共用的渲染逻辑：
    每一组（通常是同一张发票的若干明细行，或台账里的一行）放进一个可编辑表格，
    配一个按钮，点了就把（可能已编辑过的）内容整组移进 matched，同时从原来
    的列表里删掉、打上"这是人工手动处理的"说明，再 st.rerun() 刷新页面。"""
    if not groups:
        st.write("无")
        return
    for gi, grp in enumerate(groups):
        with st.expander(title_fn(grp)):
            edited = st.data_editor(_safe_df(grp)[OUT_COLS], key=f"{key_prefix}_edit_{gi}",
                                     num_rows="fixed", width="stretch")
            if st.button(button_label, key=f"{key_prefix}_move_{gi}"):
                note = note_fn(grp)
                new_rows = [_coerce_row(r) for r in edited.to_dict("records")]
                for r in new_rows:
                    r["Note"] = note
                state["matched"].extend(new_rows)
                ids = {id(r) for r in grp}
                state[source_key] = [r for r in state[source_key] if id(r) not in ids]
                st.rerun()


def _xero_export(rows):
    """把当前 matched 里的行导出成一份清单，跟 outcome.xlsx 的 matched sheet
    是同一套列（OUT_COLS）。这不是 Xero 官方导入模板——真要对上 Xero 的
    Bills 导入格式还需要 AccountCode/TaxType 这些映射信息，等确认了 Xero
    那边具体要什么字段再调整，现在先给一份通用、可核对的清单。"""
    buf = BytesIO()
    df = pd.DataFrame(rows)[OUT_COLS] if rows else pd.DataFrame(columns=OUT_COLS)
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        df.to_excel(w, index=False, sheet_name="To Import to Xero")
    buf.seek(0)
    return buf


st.set_page_config(page_title="采购发票对账", layout="wide")
st.title("采购发票对账")
st.caption("Excel 直读，PDF 抽取，按发票号精确匹配")

# 各模块版本必须一致。文件没换全是最常见的故障，放在最显眼处
import excel_loader as _el, invoice_extractor as _ie, reconcile as _rc
_vers = {"app": __version__, "reconcile": getattr(_rc, "__version__", "?"),
         "invoice_extractor": getattr(_ie, "__version__", "?"),
         "excel_loader": getattr(_el, "__version__", "?")}
try:
    import ocr as _ocr
    _vers["ocr"] = getattr(_ocr, "__version__", "?")
    _ocr_ok, _ocr_path = _ocr.available(), _ocr.where()
except ImportError:
    _vers["ocr"] = "未安装"
    _ocr_ok, _ocr_path = False, None

if len(set(_vers.values())) > 1:
    st.error("文件版本不一致，请把所有 .py 重新覆盖一遍： "
             + "；".join(f"{k} {v}" for k, v in _vers.items()))
else:
    st.caption(f"版本 {__version__} · OCR "
               + (f"就绪（{_ocr_path}）" if _ocr_ok else "不可用"))

uploads = st.file_uploader(
    "把台账 Excel 和发票 PDF 一起拖进来（可多选，顺序无所谓）",
    type=["xlsx", "xlsm", "pdf"], accept_multiple_files=True,
)

xlsx_ups = [f for f in (uploads or []) if f.name.lower().endswith((".xlsx", ".xlsm"))]
pdf_ups = [f for f in (uploads or []) if f.name.lower().endswith(".pdf")]

c1, c2 = st.columns(2)
c1.metric("台账 Excel", len(xlsx_ups))
c2.metric("发票 PDF", len(pdf_ups))

if not uploads:
    st.info("台账和发票一起拖进来即可。PDF 可以只传一张做单笔核对，也可以整个月一起传。")
    st.stop()

if not xlsx_ups:
    st.warning("还缺台账 Excel（.xlsx）。")
    st.stop()
if not pdf_ups:
    st.warning("还缺发票 PDF。")
    st.stop()

if len(xlsx_ups) > 1:
    names = [f.name for f in xlsx_ups]
    pick = st.selectbox("传了多份 Excel，用哪一份作为台账？", names)
    xlsx_file = next(f for f in xlsx_ups if f.name == pick)
else:
    xlsx_file = xlsx_ups[0]
pdf_files = pdf_ups

run_clicked = st.button("开始对账", type="primary")

if run_clicked:
    # 上传的是内存对象，落到临时目录后复用命令行版的同一套逻辑
    with tempfile.TemporaryDirectory() as tmp:
        xlsx_path = os.path.join(tmp, "tracking.xlsx")
        with open(xlsx_path, "wb") as f:
            f.write(xlsx_file.getbuffer())

        pdf_paths = []
        for up in pdf_files:
            p = os.path.join(tmp, up.name)
            with open(p, "wb") as f:
                f.write(up.getbuffer())
            pdf_paths.append(p)

        with st.spinner(f"正在处理 {len(pdf_paths)} 份发票…"):
            sheets, pending, ctx = reconcile(xlsx_path, pdf_paths)

        # 兼容旧版 reconcile.py：缺字段时降级，不让整个页面崩掉
        ctx.setdefault("suspicious", [])
        ctx.setdefault("superseded", [])
        ctx.setdefault("unmatched_pdf", [])
        ctx.setdefault("failed", [])

        # pending 比 matched/discrepancy 那几个多留了 _status/_row 两列——
        # write_output() 写「pending (no invoice yet)」那个 sheet 要用到，
        # 不能像其他几份那样直接压成只剩 OUT_COLS，否则导出会报 KeyError。
        # 显示/编辑的时候（render_movable_cases 里）只会挑 OUT_COLS 出来用，
        # 这两列不会出现在编辑框里，纯粹是留着给导出用的。
        pending_rows = [_ensure_out_cols(r) for r in
                        pending.drop(columns=["invoice_key", "po_key", "supplier_key"],
                                     errors="ignore").to_dict("records")]

        # 这里点一次"开始对账"就是全新算一遍：之前编辑/移动/打勾的状态全部
        # 扔掉重建——按用户要求，这些操作只在本次会话里有效，不用跨批次记住。
        st.session_state["reconciled"] = {
            "matched": _to_out_rows(sheets["matched"]),
            "discrepancy": _to_out_rows(sheets["discrepancy"]),
            "pdf_only": _to_out_rows(sheets["PDF to Excel - not match"]),
            "excel_only": _to_out_rows(sheets["Excel to PDF - not match"]),
            "pending": pending_rows,
            "suspicious": list(ctx["suspicious"]),
            "ctx": ctx,
        }

if "reconciled" not in st.session_state:
    st.stop()

state = st.session_state["reconciled"]
ctx = state["ctx"]
n_inv = len(ctx["invoices"])

# 下载内容每次都从当前（可能已被编辑/移动过的）state 现算，这样下载的
# outcome.xlsx 才会带上页面上做过的所有手动操作，不是对账刚跑完那一刻的快照。
_sheets_now = {
    "matched": state["matched"],
    "discrepancy": state["discrepancy"],
    "PDF to Excel - not match": state["pdf_only"],
    "Excel to PDF - not match": state["excel_only"],
}
_pending_now = pd.DataFrame(state["pending"]) if state["pending"] else \
    pd.DataFrame(columns=list(OUT_COLS) + ["_status", "_row"])
_buf = BytesIO()
write_output(_sheets_now, _pending_now, _buf, state["suspicious"], ctx.get("superseded"))
_buf.seek(0)

m1, m2, m3, m4, m5 = st.columns(5)
m1.metric("读入发票", n_inv)
m2.metric("匹配成功", len({r["Invoice Number"] for r in state["matched"]}))
m3.metric("金额不符", len(_group_contiguous(state["discrepancy"])))
m4.metric("台账查无此单", len(_group_contiguous(state["pdf_only"])))
m5.metric("台账有单缺发票", len(state["excel_only"]))

st.download_button("下载 outcome.xlsx", _buf,
                   file_name="Neocytogen - outcome.xlsx",
                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# 待核查：程序自己算出哪些结果可疑并排序，避免人工逐张点开核对。
# 每条配一个"已核实"按钮，点了就从这个清单里拿掉——只是人工复核的勾选记录，
# 不影响这张发票在 matched/discrepancy 里的状态，纯粹用来盯"还剩几张没看"。
if ctx["superseded"]:
    with st.expander(f"已忽略的 Proforma（{len(ctx['superseded'])} 张）"):
        st.caption("同一 PO 已有 Tax Invoice，按规则只算 Tax Invoice。列在此处仅供查证。")
        show(ctx["superseded"])

susp = state["suspicious"]
if susp:
    st.subheader(f"待核查 · {len(susp)} 张（共 {n_inv} 张）")
    st.caption("按疑点数量排序。核实过的点右边「已核实」拿掉，不影响 matched/discrepancy 的判定。")
    for i, row in enumerate(list(susp)):
        sc1, sc2 = st.columns([9, 1])
        with sc1:
            st.write(f"**{row.get('发票号') or '(号未识别)'}** · {row.get('供应商')} · "
                     f"合计 {row.get('合计')} · 疑点数 {row.get('疑点数')}  \n"
                     f"{row.get('需要核查的原因')}  \n"
                     f"*{row.get('文件')}*")
        with sc2:
            if st.button("✓ 已核实", key=f"susp_done_{i}"):
                state["suspicious"] = [r for r in state["suspicious"] if r is not row]
                st.rerun()
        st.divider()
else:
    st.success(f"{n_inv} 张发票全部通过自动校验，无需人工核查。")

# 抽取告警：字段没抽到、行加总对不上等，先让人知道哪些结果不可全信
warned = [i for i in ctx["invoices"].values() if i["warnings"]]
if warned or ctx["failed"]:
    with st.expander(f"提取告警（{len(warned) + len(ctx['failed'])} 份）", expanded=False):
        for inv in warned + ctx["failed"]:
            st.write(f"**{inv.get('invoice_no') or inv['source_file']}** — "
                     + "；".join(inv["warnings"]))

tab_xero, tab_matched, tab_disc, tab_pdf_only, tab_excel_only, tab_pending = st.tabs(
    ["📤 To Import to Xero", "matched", "discrepancy", "台账查无此单", "台账有单缺发票", "未开票"]
)

with tab_xero:
    st.caption("matched 的发票都会自动出现在这里——包括原本就 matched 的，以及从其他 "
               "tab 手动确认移过来的。这份清单可以直接下载去核对再导入 Xero。")
    if state["matched"]:
        show(pd.DataFrame(state["matched"])[OUT_COLS])
    else:
        st.write("暂无 matched 发票。")
    st.download_button("下载 Xero 导入清单", _xero_export(state["matched"]),
                       file_name="To Import to Xero.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                       key="xero_download")

with tab_matched:
    if state["matched"]:
        show(pd.DataFrame(state["matched"])[OUT_COLS])
    else:
        st.write("无")

with tab_disc:
    st.caption("金额跟台账对不上的发票。可以在下面直接改内容（比如台账本身有误、或者要按发票"
               "实收金额为准），改完点按钮就整张移到 matched（会同时出现在 To Import to Xero 里）。")
    render_movable_cases(
        _group_contiguous(state["discrepancy"]), state, "discrepancy",
        title_fn=lambda grp: (f"发票 {grp[0].get('Invoice Number') or '(号未识别)'} · "
                               f"{grp[0].get('Supplier')} · {(grp[0].get('Note') or '')[:60]}"),
        note_fn=lambda grp: f"人工核对后手动确认匹配（原自动比对结果：{grp[0].get('Note') or ''}）",
        key_prefix="disc",
    )

with tab_pdf_only:
    st.caption("有发票 PDF，但台账里查无此单（发票号/PO 都对不上）。核实清楚后可以手动把它配到 "
               "matched——移过去会带一条说明，写清楚这不是程序自动匹配的。")
    render_movable_cases(
        _group_contiguous(state["pdf_only"]), state, "pdf_only",
        title_fn=lambda grp: (f"发票 {grp[0].get('Invoice Number') or grp[0].get('PO Number') or '(未识别)'} · "
                               f"{grp[0].get('Supplier')}"),
        note_fn=lambda grp: f"⚠ 未自动匹配到台账——人工手动添加至 matched（原提示：{grp[0].get('Note') or ''}）",
        key_prefix="pdfonly", button_label="➕ 手动加入 matched",
    )

with tab_excel_only:
    st.caption("台账已记录、但没找到对应发票 PDF。如果发票已经拿到了、只是这次没传或者没识别出来，"
               "可以直接在这条上手动确认——移过去会带一条说明，写清楚这不是程序自动匹配的。")
    render_movable_cases(
        [[r] for r in state["excel_only"]], state, "excel_only",
        title_fn=lambda grp: f"{grp[0].get('Invoice Number') or grp[0].get('PO Number')} · {grp[0].get('Supplier')}",
        note_fn=lambda grp: f"⚠ 台账记录未找到对应发票 PDF——人工手动添加至 matched（原提示：{grp[0].get('Note') or ''}）",
        key_prefix="exlonly", button_label="➕ 手动加入 matched",
    )

with tab_pending:
    st.caption("货未到或未开票，不算 mismatch。如果其实已经开票入账了，只是状态还没更新，"
               "也可以在这条上手动确认——移过去会带一条说明，写清楚这不是程序自动匹配的。")
    render_movable_cases(
        [[r] for r in state["pending"]], state, "pending",
        title_fn=lambda grp: f"{grp[0].get('Invoice Number') or grp[0].get('PO Number')} · {grp[0].get('Supplier')}",
        note_fn=lambda grp: "⚠ 原状态为未开票/货未到——人工手动添加至 matched",
        key_prefix="pending", button_label="➕ 手动加入 matched",
    )

# 台账里找不到时，列出最接近的几条供人工判断
if ctx["unmatched_pdf"]:
    st.divider()
    st.subheader("人工核对建议")
    st.caption("以下仅为参考排序，程序不会据此自动匹配 —— 发票号必须完全相等才算 matched")
    for inv in ctx["unmatched_pdf"]:
        total = (inv.get("subtotal") or {}).get("amount_incl")
        with st.expander(f"发票 {inv.get('invoice_no')} · {inv.get('supplier')} · {total}"):
            cands = find_candidates(ctx["expected"], inv)
            if cands:
                show(cands)
            else:
                st.write("台账里没有相近的记录。")
            for w in inv["warnings"]:
                st.warning(w)
