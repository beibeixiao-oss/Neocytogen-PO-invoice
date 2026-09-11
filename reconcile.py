"""
reconcile.py — PDF invoice x Procurement Tracking List reconciliation engine

    python reconcile.py <tracking_list.xlsx> <invoice_pdf_folder> [output.xlsx]

Produces four sheets with a shared header (the 12-column template + Source + Note):
    matched                  present on both sides, amounts agree
    discrepancy              present on both sides, amounts don't agree
    PDF to Excel - not match invoice exists, no matching entry in Excel
    Excel to PDF - not match Excel says invoiced, but no matching PDF found

Every row produced from a PDF also carries two extra fields beyond OUT_COLS
(not written to the Excel export, only used by the UI): "source_file" (the
original PDF's filename, so the app can offer to open/preview it) and
"quality_flags" (a list of reasons this invoice's extracted data might not be
fully trustworthy — e.g. it was OCR-read, or its tax rate looks unusual —
independent of whether the amount happened to reconcile). See
quality_flags_by_reasons()/_quality_reasons() below.
"""

__version__ = "2026-09-11.11"

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
TOLERANCE = 0.05          # Amount tolerance: line-level GST rounding differs by vendor, so we only gate on the invoice total


def pdf_rows(inv, source, note="", quality_flags=None):
    """One invoice -> one row per line item, fields aligned to the template."""
    extra = {"source_file": inv.get("source_file"), "quality_flags": quality_flags or None}
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
            **extra,
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
            "Note": "; ".join(x for x in [note, "Only the invoice total was extracted, no line items"] if x),
            **extra,
        })
        return rows
    if not rows:      # Keep a trace even when no line items were extracted, or the invoice would vanish silently
        rows.append({c: None for c in OUT_COLS} | {
            "Invoice Number": inv.get("invoice_no"), "PO Number": inv.get("po_no"),
            "Supplier": inv.get("supplier"), "Source": source,
            "Note": (note + " / " if note else "") + "No line items extracted",
            **extra,
        })
    return rows


def merged_rows(inv, grp, note="", quality_flags=None):
    """A row present on both sides: line-item detail comes from the ledger (Item
    Description is human-written and cleaner than what's scraped from the PDF);
    invoice number / PO / supplier / date / currency come from the PDF (the
    invoice is the basis for payment)."""
    extra = {"source_file": inv.get("source_file"), "quality_flags": quality_flags or None}
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
            **extra,
        })
    return rows


def excel_rows(df, note=""):
    out = df[TEMPLATE_COLS].copy()
    out["Source"] = "Excel"
    out["Note"] = note
    return out.to_dict("records")


def find_candidates(expected, inv, limit=5):
    """When an invoice can't be found in the ledger, list the most likely
    candidate rows for manual review.

    Only searches within the same supplier — BioLabs and BioBasic have similar
    names but are different companies; cross-supplier ranking would just make
    someone check every row, which wastes more time than giving no suggestion.
    Note: this is reference only for a human — the program never auto-matches
    on this basis (the invoice number must match exactly).
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
        if sim < 85:                       # BioLab vs BioBasic lands right at 80, so the threshold must be above that
            continue
        amt = round(grp["Amount incl. GST"].sum(), 2)
        diff = None if total is None else round(amt - total, 2)
        gap = None
        if date is not None and pd.notna(row["Invoice date"]):
            gap = abs((pd.Timestamp(date) - pd.Timestamp(row["Invoice date"])).days)
        # A matching amount is the strongest signal (usually just a mistyped invoice number); a close date is next
        score = (0 if diff is None else max(0, 60 - min(abs(diff), 60))) \
                + (0 if gap is None else max(0, 30 - gap)) + sim * 0.1
        rows.append({
            "Excel Invoice Number": row["Invoice Number"],
            "Supplier": row["Supplier"],
            "Invoice date": row["Invoice date"],
            "Amount incl. GST": amt,
            "Amount diff": diff,
            "Date diff (days)": gap,
            "_s": score,
        })
    rows.sort(key=lambda r: r["_s"], reverse=True)
    for r in rows:
        r.pop("_s")
    return rows[:limit]


def _supplier_baseline(expected):
    """Group the ledger's amounts and invoice-number shapes by supplier, as a
    comparison baseline for _quality_reasons()."""
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
    return by_sup


def _quality_reasons(inv, baseline):
    """Return the list of reasons this invoice's extracted data might not be
    fully trustworthy (OCR/AI source, missing date, unusual tax rate, an
    invoice-number format that doesn't match this supplier's history, etc.).

    This is independent of whether the amount reconciled — a matched invoice
    can still come back with reasons here (e.g. it was OCR-read but happened
    to add up correctly), and a genuine amount mismatch won't show up here at
    all unless one of these other signals also fired. Called fresh at the
    point each output row is produced, so it always reflects the invoice's
    current state (e.g. after an OCR invoice-number correction has already
    been applied — see the OCR-fold branch in reconcile() below).

    A field that fails to extract raises an error elsewhere; a field that
    extracts to the wrong value doesn't — the latter is what this function is
    for. It uses the historical distribution for the same supplier as a
    baseline: an amount off by an order of magnitude, an invoice-number format
    that doesn't match its peers, or a date outside a plausible range are all
    worth a human glance.
    """
    import statistics
    from rapidfuzz import fuzz
    from excel_loader import normalize_supplier

    reasons = []
    sup = normalize_supplier(inv.get("supplier") or "")
    ref = None
    for k, v in baseline.items():
        if sup and max(fuzz.ratio(sup, k), fuzz.partial_ratio(sup, k)) >= 85:
            ref = v
            break

    total = (inv.get("subtotal") or {}).get("amount_incl")
    if total is None:
        reasons.append("Invoice total not extracted")
    elif ref and len(ref["amts"]) >= 3:
        med = statistics.median(ref["amts"])
        if med > 0 and (total < med / 10 or total > med * 10):
            reasons.append(f"Amount {total} is more than 10x off this supplier's median of {round(med, 2)}")

    shape = _shape(inv.get("invoice_no"))
    # A Proforma's number follows the supplier's separate numbering scheme (e.g. PI- vs INV-), so a different shape is expected
    if ref and ref["shapes"] and shape not in ref["shapes"] and not inv.get("is_proforma"):
        msg = f"Invoice number format {shape} doesn't match this supplier's ledger formats ({'/'.join(sorted(ref['shapes'])[:3])})"
        fix = suggest_repair(inv.get("invoice_no"), ref.get("numbers", set()))
        if fix:
            msg += f"; {fix}"
        reasons.append(msg)

    d = inv.get("invoice_date")
    if d is None:
        reasons.append("Invoice date not extracted")

    s = inv.get("subtotal") or {}
    e, g, i = s.get("amount_excl"), s.get("gst"), s.get("amount_incl")
    if None not in (e, g, i) and abs(e + g - i) > 0.05:
        reasons.append("excl. GST + GST != incl. GST")
    if i and e and i > 0 and not (0 <= (i - e) / i < 0.15):
        reasons.append(f"Unusual tax rate: {round((i - e) / i * 100, 1)}%")

    if inv.get("ai"):
        reasons.append("Read by the vision model — figures need manual review")
    if inv.get("ocr"):
        reasons.append("Scanned document read via OCR — figures need manual review")
    # A Proforma is treated the same as a regular Tax Invoice and never flagged
    # just for being a Proforma — the signals that should actually surface a
    # problem are the ones above (amount equation, tax rate, OCR/AI source),
    # none of which are about whether it's a Proforma.
    if not inv.get("supplier"):
        reasons.append("Supplier not extracted")

    return reasons


# The character pairs OCR confuses most often. Only used to probe "possibly misread" cases — never a blanket text-wide replacement.
_CONFUSIONS = [("O", "0"), ("I", "1"), ("L", "1"), ("S", "5"),
               ("B", "8"), ("Z", "2"), ("G", "6"), ("Q", "0"), ("D", "0")]
_CONFUSION_SET = {p for a, b in _CONFUSIONS for p in ((a, b), (b, a))}


def suggest_repair(raw, candidates, limit=3):
    """When an invoice number looks like it was mis-extracted, look for a
    likely correct value among this supplier's ledger numbers.

    Two cases:
      1) OCR character confusion — the O in LBHO00 is actually the digit 0
      2) Truncation — INV-26 is actually INV-26-03283

    Compares character-by-character rather than doing a blanket substring
    replace — a blanket replace would also turn the I in INV into a 1 and
    corrupt the number. A candidate only counts if it's the same length and
    every differing position falls within a known confusion pair.

    Always just a suggestion. The program never auto-matches on this basis —
    the invoice number must match exactly to count as matched.
    """
    if not raw or not candidates:
        return None
    s = str(raw).strip()

    # Case 1: character-by-character comparison
    hits = []
    for c in candidates:
        if c == s or len(c) != len(s):
            continue
        diff = [(a, b) for a, b in zip(s.upper(), c.upper()) if a != b]
        if diff and len(diff) <= 3 and all(pair in _CONFUSION_SET for pair in diff):
            hits.append(c)
    if hits:
        return "Possible OCR character misread — the ledger has " + " / ".join(sorted(hits)[:limit])

    # Case 2: the extracted value is a prefix of some ledger number
    pre = sorted(c for c in candidates if c != s and c.startswith(s) and len(s) >= 4)
    if pre:
        tail = " / ".join(pre[:limit]) + ("..." if len(pre) > limit else "")
        return f"Possibly truncated — the ledger has {tail}"

    return None


def _shape(v):
    """Abstract an invoice number into a shape: 97809367 -> 99999999 ; INV-0386 -> AAA-9999"""
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return ""
    s = str(v).strip()
    if s.endswith(".0"):
        s = s[:-2]
    return re.sub(r"[A-Za-z]", "A", re.sub(r"\d", "9", s))


def reconcile(xlsx_path, pdf_source):
    """pdf_source can be a folder path or a list of PDF paths"""
    expected, pending = load_tracking_list(xlsx_path)
    baseline = _supplier_baseline(expected)

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
                "Invoice No": inv.get("invoice_no"), "PO": inv.get("po_no"),
                "Supplier": inv.get("supplier"),
                "Total": (inv.get("subtotal") or {}).get("amount_incl"),
                "Reason Ignored": "This PO already has a Tax Invoice; by rule only the Tax Invoice is counted",
                "File": inv.get("source_file"),
            })
            del invoices[key]

    matched, discrepancy, pdf_only, excel_only = [], [], [], []

    # Two-level matching.
    # The ledger has 59 PO numbers but only 54 invoice numbers in varying formats,
    # and every PDF's filename carries a PO number, so the PO is the key that's
    # fully populated and normalized on both sides. Invoice number is still tried
    # first — it's more precise and can distinguish multiple invoices on the same
    # PO; when it doesn't match, we fall back to PO and note which level was used
    # in the Note field, so it stays traceable.
    by_inv, by_po = {}, {}
    for _, r in expected.iterrows():
        if r["invoice_key"]:
            by_inv.setdefault(r["invoice_key"], []).append(r)
        if r["po_key"]:
            by_po.setdefault(r["po_key"], []).append(r)

    # OCR-confusion folding index. Tesseract's most common misreads on invoice
    # numbers are O/0, I/l/1, S/5, B/8 — in practice INV260123-LBH001 gets read
    # as INV260123-LBHOO1: the format is still valid and passes validation, so
    # fix_ocr_digits never triggers, and the exact match fails too, so it falls
    # uselessly into "no ledger match". Folded keys are only accepted when they
    # hit exactly one ledger entry; multiple hits are left for a human rather
    # than guessed.
    def _fold(k):
        return (str(k or "").upper().replace("O", "0").replace("I", "1")
                .replace("L", "1").replace("S", "5").replace("B", "8"))

    folded = {}
    for k in by_inv:
        folded.setdefault(_fold(k), []).append(k)

    used_inv, used_po = set(), set()
    # The same PO can correspond to multiple invoices (invoiced in batches) — count them first, since a group of more than one can't be claimed as a whole
    po_pdf_count = {}
    for inv in list(invoices.values()) + failed:
        k = normalize_invoice_no(inv.get("po_no"))
        if k:
            po_pdf_count[k] = po_pdf_count.get(k, 0) + 1

    # Invoices with no extracted invoice number (mostly scans) still take part: their filename PO is still reliable
    candidates = list(invoices.items()) + [("", i) for i in failed]
    failed = []

    # ---- Batch invoicing: multiple invoices on the same PO — sum first, then compare ----
    # Comparing each one individually is guaranteed to mismatch (each is only
    # part of the total), which would flag an otherwise-correct account as an
    # error. Instead, sum these invoices per PO and compare against the whole
    # ledger group: if it matches, the whole group is matched; if not, fall
    # back to per-invoice handling so a person can see exactly which one is off.
    batch_ok = {}          # po_key -> all invoices under that PO (confirmed the sum matches)
    _by_po_pdf = {}
    for key, inv in candidates:
        if key and (key in by_inv
                    or (inv.get("ocr") and len(folded.get(_fold(key), [])) == 1)):
            continue       # Invoices whose number matches directly skip PO aggregation
        pk = normalize_invoice_no(inv.get("po_no"))
        if pk and pk in by_po:
            _by_po_pdf.setdefault(pk, []).append((key, inv))

    for pk, group in _by_po_pdf.items():
        if len(group) < 2:
            continue
        totals = [(i.get("subtotal") or {}).get("amount_incl") for _, i in group]
        if any(t is None for t in totals):
            continue       # If one invoice's total wasn't extracted, the sum is meaningless
        x_incl = round(pd.DataFrame(by_po[pk])["Amount incl. GST"].sum(), 2)
        if abs(round(sum(totals), 2) - x_incl) <= TOLERANCE:
            batch_ok[pk] = group

    for pk, group in batch_ok.items():
        used_po.add(pk)
        gdf = pd.DataFrame(by_po[pk])
        n = len(group)
        for key, inv in group:
            t = (inv.get("subtotal") or {}).get("amount_incl")
            matched.extend(merged_rows(
                inv, gdf,
                f"Matched by PO: this PO has {n} invoices issued in batches, "
                f"total {round(sum((i.get('subtotal') or {}).get('amount_incl') for _, i in group), 2)}"
                f" agrees with the ledger (this invoice: {t})",
                quality_flags=_quality_reasons(inv, baseline)))
    _batched = {id(i) for g in batch_ok.values() for _, i in g}
    candidates = [(k, i) for k, i in candidates if id(i) not in _batched]

    # Multiple invoices under the same PO where none of them could be matched
    # by invoice number, and the sum didn't match the ledger either (the
    # "batch invoicing" step above failed): previously this would compare
    # every single one against the sum of all unclaimed ledger rows, which
    # produced a meaningless diff ("ledger group total minus this one
    # invoice") and dragged an otherwise-correct invoice into mismatch too
    # (real case: two laundry invoices on the same PO, one of which matched a
    # ledger row exactly, but got reported as a mismatch anyway just because
    # the other one didn't match — the ledger total in the Note didn't
    # correspond to the row this invoice actually matched, which made it look
    # like the ledger itself might be wrong).
    # When the invoice count matches the number of unclaimed ledger rows,
    # pair them by nearest amount instead: each invoice is paired with the
    # ledger row closest to it in amount, and each pair is then judged
    # against the tolerance independently — so any reported diff is actually
    # about the specific row that invoice corresponds to, and by how much.
    # When the counts don't match (more commonly rows != invoices), skip this
    # pairing and fall back to the original whole-group comparison, to avoid
    # guessing a wrong pairing.
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
            # Note: rows are pandas Series — list.remove()'s == comparison
            # raises "truth value of a Series is ambiguous" when there are
            # multiple candidates, so we pop by index instead to sidestep that.
            best_idx = min(range(len(remaining_rows)),
                           key=lambda idx: abs(remaining_rows[idx]["Amount incl. GST"] - p_incl))
            best = remaining_rows.pop(best_idx)
            gdf = pd.DataFrame([best])
            x_incl = best["Amount incl. GST"]
            tag = (f"Paired by nearest amount (all {len(group)} invoices under this PO had no "
                   f"recognizable invoice number, so each was paired to its corresponding ledger "
                   f"row by total amount rather than compared against the whole-group sum)")
            qf = _quality_reasons(inv, baseline)
            if abs(x_incl - p_incl) <= TOLERANCE:
                matched.extend(merged_rows(inv, gdf, tag, quality_flags=qf))
            else:
                diff = round(p_incl - x_incl, 2)
                discrepancy.extend(merged_rows(
                    inv, gdf, "; ".join([tag, f"Amount mismatch: ledger {x_incl} vs invoice {p_incl} (diff {diff})"]),
                    quality_flags=qf))
            _paired.add(id(inv))
        used_po.add(pk)
    candidates = [(k, i) for k, i in candidates if id(i) not in _paired]

    # Pre-scan: figure out up front which ledger rows will be directly claimed
    # by invoice number. This must happen before the main loop — otherwise
    # "already claimed" would depend on iteration order, and whichever of two
    # invoices on the same PO gets processed first would change the outcome.
    claimed_inv = set()
    for _k, _i in candidates:
        if _k and _k in by_inv:
            claimed_inv.add(_k)
        elif _k and _i.get("ocr") and len(folded.get(_fold(_k), [])) == 1:
            claimed_inv.add(folded[_fold(_k)][0])

    for key, inv in candidates:
        grp, level = None, None
        if key and key in by_inv:
            grp, level = by_inv[key], "Invoice No"
            used_inv.add(key)
        elif key and inv.get("ocr") and len(folded.get(_fold(key), [])) == 1:
            real = folded[_fold(key)][0]
            grp, level = by_inv[real], f"Invoice No (OCR character correction: {key} -> {real}, please confirm manually)"
            used_inv.add(real)
            # Previously this only used `real` to look up the matching ledger
            # row — the extraction result itself (inv["invoice_no"]) was never
            # updated, so outcome.xlsx kept showing the original OCR-misread
            # number (e.g. LBHOO1): the pairing succeeded but the displayed
            # number was still wrong, which looked unfixed. `inv` and the
            # `invoices` dict reference the same object, so updating it here
            # means the quality-flag check further below (which runs after
            # this) also sees the corrected number, and the format check
            # won't falsely flag "doesn't match the ledger" anymore.
            n = by_inv[real][0]["Invoice Number"]
            fixed_no = (str(int(n)) if isinstance(n, float) and n.is_integer() else str(n).strip()) if n is not None else None
            if fixed_no and normalize_invoice_no(fixed_no) == real:
                old_no = inv.get("invoice_no")
                if fixed_no != old_no:
                    inv["invoice_no"] = fixed_no
                    inv.setdefault("warnings", []).append(
                        f"Invoice number corrected via OCR: {old_no} -> {fixed_no} (based on a unique ledger match), please confirm manually")
        else:
            pk = normalize_invoice_no(inv.get("po_no"))
            if pk in by_po:
                # Other invoices under the same PO may already have claimed
                # their matching ledger rows by invoice number. If we still
                # compared against the whole ledger group, this one would
                # necessarily mismatch (the diff would just be someone else's
                # share) — reporting "ledger 577.7 vs invoice 957.52" would be
                # meaningless. So only compare against rows that aren't
                # claimed yet.
                rest = [r for r in by_po[pk]
                        if not r["invoice_key"] or r["invoice_key"] not in claimed_inv]
                if rest:
                    grp = rest
                    level = ("PO" if len(rest) == len(by_po[pk])
                             else "PO (excluding rows already claimed by invoice number under the same PO)")
                else:
                    grp, level = by_po[pk], "PO (every ledger row under this PO is already claimed, please confirm manually)"
                used_po.add(pk)

        # Computed after any OCR invoice-number correction above, so it
        # reflects the invoice's final, corrected state.
        qf = _quality_reasons(inv, baseline)

        if grp is None:
            po = str(inv.get("po_no") or "")
            hint = ("PO is from 2025 — please check the 2025 ledger"
                    if re.match(r"PONCG2025", po, re.I) else "No matching PO or invoice number in the ledger")
            pdf_only.extend(pdf_rows(inv, "PDF", hint, quality_flags=qf))
            continue

        gdf = pd.DataFrame(grp)
        x_incl = round(gdf["Amount incl. GST"].sum(), 2)
        p_incl = (inv.get("subtotal") or {}).get("amount_incl")
        tag = "" if level == "Invoice No" else f"Matched by {level}"
        if not key and tag:
            tag += " (invoice number not recognized)"

        if p_incl is None:
            discrepancy.extend(merged_rows(inv, gdf, "; ".join(x for x in [tag, "Invoice total not extracted, can't compare amounts"] if x), quality_flags=qf))
        elif abs(x_incl - p_incl) <= TOLERANCE:
            matched.extend(merged_rows(inv, gdf, tag, quality_flags=qf))
        else:
            diff = round(p_incl - x_incl, 2)
            note = f"Amount mismatch: ledger {x_incl} vs invoice {p_incl} (diff {diff})"
            discrepancy.extend(merged_rows(inv, gdf, "; ".join(x for x in [tag, note] if x), quality_flags=qf))

    # In the ledger, but no matching PDF was found
    from excel_loader import STATUS_INVOICE_EXPECTED
    still_pending = []
    for _, r in expected.iterrows():
        if r["invoice_key"] in used_inv or r["po_key"] in used_po:
            continue
        if r["_status"] not in STATUS_INVOICE_EXPECTED:
            still_pending.append(r)          # Goods not received yet, so there shouldn't be an invoice yet
        else:
            excel_only.extend(excel_rows(pd.DataFrame([r]), "Recorded in the ledger, but no matching invoice PDF was found"))

    for inv in failed:
        pdf_only.extend(pdf_rows(inv, "PDF", "Could not extract an invoice number: " + "; ".join(inv["warnings"]),
                                  quality_flags=_quality_reasons(inv, baseline)))

    sheets = {
        "matched": matched,
        "discrepancy": discrepancy,
        "PDF to Excel - not match": pdf_only,
        "Excel to PDF - not match": excel_only,
    }
    context = {"expected": expected, "invoices": invoices,
               "superseded": superseded,
               "unmatched_pdf": [i for k, i in invoices.items()
                                 if k not in used_inv
                                 and normalize_invoice_no(i.get("po_no")) not in used_po],
               "failed": failed, "n_pdf": len(pdf_files)}
    if still_pending:
        pending = pd.concat([pending, pd.DataFrame(still_pending)], ignore_index=True)
    return sheets, pending, context


def write_output(sheets, pending, path, superseded=None):
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

    ws = wb.create_sheet("Ignored Proformas")
    cols2 = ["Invoice No", "PO", "Supplier", "Total", "Reason Ignored", "File"]
    ws.append(cols2)
    for c in ws[1]:
        c.font = Font(name="Arial", bold=True)
    for r in (superseded or []):
        ws.append([r.get(c) for c in cols2])
    for n, L in enumerate("ABCDEF"):
        ws.column_dimensions[L].width = 46 if L in "EF" else 18
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
    write_output(sheets, pending, out, ctx["superseded"])
    for k, v in sheets.items():
        print(f"{k:28s} {len(v):4d} rows")
    print(f"{'Ignored Proformas':25s} {len(ctx['superseded']):4d}")
    print(f"{'pending (not invoiced)':26s} {len(pending):4d} rows")
    print("Written to:", out)
