"""
excel_loader.py — Neocytogen procurement reconciliation · Excel-side loading and normalization

Responsibility: turn the Procurement Tracking List into a clean table that can
be matched against PDF invoices. Output column names are already aligned to
the "Neocytogen - outcome.xlsx" template.

    from excel_loader import load_tracking_list
    df, pending = load_tracking_list("Neocytogen Procurement Tracking List (2026).xlsx")
"""

__version__ = "2026-09-11.10"

import re
import pandas as pd

SHEET_MAIN = "Expense Tracker"
SHEET_VENDOR = "Vendor lists"
GST_RATE = 0.09

# Only these statuses should already have an invoice. The rest (Ordered /
# Ordering / Pending Delivery) mean the goods haven't arrived and no invoice
# exists yet — these must not count toward "Excel to PDF - not match", or
# every one of them would be a false alarm.
STATUS_INVOICE_EXPECTED = {
    "Order Complete",
    "Delivered",
    "Partial Delivered",
    "Pending Payment",
}


def normalize_invoice_no(value) -> str:
    """Normalize invoice-number formatting so both sides can use it as a join key.

    Edge case: pandas reads a purely numeric invoice number as a float
    (523 -> 523.0), which will never match the text "523" extracted from the PDF.
    """
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    if isinstance(value, float) and value.is_integer():
        value = int(value)
    s = str(value).strip().upper()
    s = re.sub(r"\s+", "", s)          # Strip all whitespace
    s = re.sub(r"[^\w\-/]", "", s)     # Strip stray symbols like # and . , keeping - and /
    return s


def normalize_supplier(name) -> str:
    """Normalize a supplier name: strip whitespace, strip company suffixes, uppercase.
    'Agilent Technologies Singapore' and 'Agilent' converge to the same comparable prefix.
    """
    if name is None or (isinstance(name, float) and pd.isna(name)):
        return ""
    s = str(name).upper()
    s = re.sub(r"\b(PTE|LTD|LIMITED|INC|LLC|CO|CORP|SINGAPORE|ASIA)\b", "", s)
    s = re.sub(r"[^A-Z0-9]", "", s)
    return s


def load_vendor_aliases(path: str) -> dict:
    """Build an alias -> canonical supplier name dict from the Vendor lists sheet.
    The Brands column is a comma-separated (including full-width comma) list of
    brand names, all pointing to the same Vendor.
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
    """Returns (df_expected, df_pending).

    df_expected — rows that should have an invoice, taking part in reconciliation.
    df_pending  — rows where goods haven't arrived / no invoice yet, reported separately, not counted as a mismatch.
    """
    raw = pd.read_excel(path, sheet_name=SHEET_MAIN)

    # Edge case 1: the file has 1095 rows, but real data only goes down to row 123 — the rest are empty formula rows.
    df = raw[raw["Status"].notna()].copy()

    aliases = load_vendor_aliases(path)

    # Supplier: use 'Company Ordered From' (who issues the invoice), not
    # 'Brand' (the product brand). Example: Brand=SinoBio but Company Ordered
    # From=Afirmus — the invoice is issued by Afirmus.
    supplier = df["Company Ordered From"].fillna(df["Brand"])
    df["supplier_raw"] = supplier
    df["Supplier"] = supplier.map(
        lambda x: aliases.get(normalize_supplier(x), x if isinstance(x, str) else "")
    )
    df["supplier_key"] = df["Supplier"].map(normalize_supplier)

    df["invoice_key"] = df["Invoice Number"].map(normalize_invoice_no)
    df["po_key"] = df["PO Number"].map(normalize_invoice_no)

    # The amount trio. Unit Cost is the excl.-tax unit price; Total Cost is the
    # incl.-tax total (verified against rows 115/122, ratio = 1.09).
    excl = df["Unit Cost"] * df["Units"]
    incl = df["Total Cost "]          # note: the original header has a trailing space
    df["Amount excl. GST"] = excl.round(2)
    df["Amount incl. GST"] = incl.round(2)
    df["GST"] = (incl - excl).round(2)

    # Excel has no Currency or Due Date column — these can only come from the PDF.
    # Due Date fallback: when Notes says "30 Days from Invoice", derive it as Invoice Date + 30 days.
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
        # Internal use only, dropped before writing to Excel
        "invoice_key": df["invoice_key"],
        "po_key": df["po_key"],
        "supplier_key": df["supplier_key"],
        "_status": df["Status"],
        "_row": df.index + 2,          # Corresponds to the actual Excel row number, for manual cross-checking
    })

    # Every row with a PO or invoice number takes part in matching — invoices
    # often arrive before the status gets updated, so filtering by status
    # would exclude "goods not received yet, but already invoiced" cases
    # (verified to drop 3 in practice). Status is only used to judge whether
    # "no invoice found" actually counts as a problem.
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
    print(f"Rows in reconciliation: {len(expected)}   distinct invoice numbers: {expected['invoice_key'].nunique()}")
    print(f"Not yet invoiced / goods not received: {len(pending)}")
    print()
    print(expected.head(8).to_string())
