# Purchase Invoice Reconciliation Tool

Ledger Excel x invoice PDFs, matched exactly by invoice number, producing outcome.xlsx.

## Install

    pip install -r requirements.txt

## Usage

UI version (recommended for non-technical colleagues):

    streamlit run app.py

Command-line batch:

    python reconcile.py "Neocytogen Procurement Tracking List (2026).xlsx" "./2026 Jan/" "outcome.xlsx"

## Files

| File | Purpose |
|---|---|
| `excel_loader.py` | Ledger Excel cleanup and normalization |
| `invoice_extractor.py` | PDF invoice field extraction |
| `reconcile.py` | Matching logic + Excel output |
| `app.py` | Streamlit UI (interface layer only — swap this out when moving to Zoho) |

## Key design decisions

- **Invoice numbers are matched exactly, never fuzzily.** The same supplier's
  invoice numbers can differ by just a few characters (97809367 /
  97809395) — fuzzy matching is guaranteed to pair the wrong ones. Similarity
  scores are only used to suggest candidates for a human to review.
- **Amounts are only validated at the invoice-total level**, with a tolerance
  of 0.05 — line-level GST rounding varies by supplier.
- **Excel is read directly, never through a model.** Numbers must be read deterministically.
- **Rows with status Ordered / Pending don't count as a mismatch** — they go into their own pending sheet.
- **Orders spanning a year boundary**: when the PO number is from the prior year (e.g. PONCG2025xxxxx), the Note flags it for checking against the prior year's ledger.
- `reconcile()` returns a plain data structure, and `write_output()` is a
  separate output layer. Moving to Zoho only means swapping the output and UI
  layers — the matching logic doesn't need to change.
