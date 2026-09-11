"""
excel_loader.py — Neocytogen 采购对账系统 · Excel 侧加载与标准化

职责：把 Procurement Tracking List 读成一张干净的、可与 PDF invoice 匹配的表。
输出列名已对齐 "Neocytogen - outcome.xlsx" 模板。

    from excel_loader import load_tracking_list
    df, pending = load_tracking_list("Neocytogen Procurement Tracking List (2026).xlsx")
"""

__version__ = "2026-09-11.2"

import re
import pandas as pd

SHEET_MAIN = "Expense Tracker"
SHEET_VENDOR = "Vendor lists"
GST_RATE = 0.09

# 只有这些状态才应该已经有发票。其余（Ordered / Ordering / Pending Delivery）
# 是货未到、发票尚未产生 —— 不能算进 "Excel to PDF - not match"，否则全是假警报。
STATUS_INVOICE_EXPECTED = {
    "Order Complete",
    "Delivered",
    "Partial Delivered",
    "Pending Payment",
}


def normalize_invoice_no(value) -> str:
    """统一发票号格式，供两边做 join key。

    坑：pandas 把纯数字发票号读成 float（523 -> 523.0），
    跟 PDF 抽出来的文本 "523" 永远匹配不上。
    """
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    if isinstance(value, float) and value.is_integer():
        value = int(value)
    s = str(value).strip().upper()
    s = re.sub(r"\s+", "", s)          # 去掉所有空白
    s = re.sub(r"[^\w\-/]", "", s)     # 去掉 # 、 . 等杂符号，保留 - 和 /
    return s


def normalize_supplier(name) -> str:
    """供应商名标准化：去空格、去公司后缀、转大写。
    'Agilent Technologies Singapore' 与 'Agilent' 会收敛到同一个可比形式的前缀。
    """
    if name is None or (isinstance(name, float) and pd.isna(name)):
        return ""
    s = str(name).upper()
    s = re.sub(r"\b(PTE|LTD|LIMITED|INC|LLC|CO|CORP|SINGAPORE|ASIA)\b", "", s)
    s = re.sub(r"[^A-Z0-9]", "", s)
    return s


def load_vendor_aliases(path: str) -> dict:
    """从 Vendor lists sheet 建 别名 -> 标准供应商名 的字典。
    Brands 列是逗号/中文逗号分隔的品牌名，都指向同一个 Vendor。
    """
    aliases = {}
    try:
        v = pd.read_excel(path, sheet_name=SHEET_VENDOR)
    except Exception:
        return aliases
    for _, row in v.iterrows():
        vendor = row.get("Vendor")
        if not isinstance(vendor, str) or not vendor.strip():
            continue
        vendor = vendor.strip()
        aliases[normalize_supplier(vendor)] = vendor
        brands = row.get("Brands")
        if isinstance(brands, str):
            for b in re.split(r"[,，/]", brands):
                b = b.strip()
                if b:
                    aliases.setdefault(normalize_supplier(b), vendor)
    return aliases


def load_tracking_list(path: str):
    """返回 (df_expected, df_pending)。

    df_expected —— 应该有发票的行，参与对账。
    df_pending  —— 货未到 / 未开票的行，报告里单独说明，不算 mismatch。
    """
    raw = pd.read_excel(path, sheet_name=SHEET_MAIN)

    # 坑 1：文件有 1095 行，真实数据只到第 123 行，其余是空的公式行。
    df = raw[raw["Status"].notna()].copy()

    aliases = load_vendor_aliases(path)

    # 供应商：用 'Company Ordered From'（开票方），不用 'Brand'（品牌方）。
    # 例：Brand=SinoBio 但 Company Ordered From=Afirmus，发票是 Afirmus 开的。
    supplier = df["Company Ordered From"].fillna(df["Brand"])
    df["supplier_raw"] = supplier
    df["Supplier"] = supplier.map(
        lambda x: aliases.get(normalize_supplier(x), x if isinstance(x, str) else "")
    )
    df["supplier_key"] = df["Supplier"].map(normalize_supplier)

    df["invoice_key"] = df["Invoice Number"].map(normalize_invoice_no)
    df["po_key"] = df["PO Number"].map(normalize_invoice_no)

    # 金额三件套。Unit Cost 是税前单价，Total Cost 是税后总额（已验证 115/122 行比值 = 1.09）。
    excl = df["Unit Cost"] * df["Units"]
    incl = df["Total Cost "]          # 注意原表头结尾有一个空格
    df["Amount excl. GST"] = excl.round(2)
    df["Amount incl. GST"] = incl.round(2)
    df["GST"] = (incl - excl).round(2)

    # Excel 里没有 Currency 和 Due Date —— 只能从 PDF 抽。
    # Due Date 兜底：Notes 写 "30 Days from Invoice" 时按 Invoice Date + 30 天推算。
    df["Currency"] = None
    df["Due Date"] = df.apply(_infer_due_date, axis=1)

    out = pd.DataFrame({
        "Invoice Number": df["Invoice Number"],
        "PO Number": df["PO Number"],
        "Supplier": df["Supplier"],
        "Invoice date": df["Invoice Date"],
        "Due Date": df["Due Date"],
        "Description": df["Item Description"],
        "Unit no": df["Units"],
        "Currency": df["Currency"],
        "Unit Price": df["Unit Cost"],
        "Amount excl. GST": df["Amount excl. GST"],
        "GST": df["GST"],
        "Amount incl. GST": df["Amount incl. GST"],
        # 内部用，写出 Excel 前 drop 掉
        "invoice_key": df["invoice_key"],
        "po_key": df["po_key"],
        "supplier_key": df["supplier_key"],
        "_status": df["Status"],
        "_row": df.index + 2,          # 对应 Excel 里的实际行号，方便人工核对
    })

    # 全部有 PO 或发票号的行都参与匹配 —— 发票常常先于状态更新到达，
    # 用状态过滤会把「货未到但发票已开」的单子挡在外面（实测漏掉 3 张）。
    # 状态只用于判断「没找到发票」算不算问题。
    expected = out[(out["invoice_key"] != "") | (out["po_key"] != "")]
    pending = out[~out.index.isin(expected.index)]
    return expected.reset_index(drop=True), pending.reset_index(drop=True)


def _infer_due_date(row):
    terms = row.get("Notes")
    inv_date = row.get("Invoice Date")
    if pd.isna(inv_date) or not isinstance(terms, str):
        return None
    m = re.search(r"(\d+)\s*Days", terms, re.I)
    if m:
        return inv_date + pd.Timedelta(days=int(m.group(1)))
    return None


if __name__ == "__main__":
    import sys
    expected, pending = load_tracking_list(sys.argv[1])
    print(f"参与对账的行: {len(expected)}   不同发票号: {expected['invoice_key'].nunique()}")
    print(f"未开票/货未到的行: {len(pending)}")
    print()
    print(expected.head(8).to_string())
