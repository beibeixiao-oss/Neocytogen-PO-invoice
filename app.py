"""
app.py — Purchase Invoice Reconciliation Tool (UI)

    streamlit run app.py

Upload the Procurement Tracking List on one side and the invoice PDF(s) on
the other (multiple files allowed). Upload one PDF for a single check, or a
whole month's worth for a batch reconciliation — both paths use the same
logic.

Interactive editing (since 2026-09-11.5):
    Once a reconciliation run has been computed, the results are stored in
    st.session_state. Buttons like "Move to matched" or "Verified" all work by
    triggering a page rerun (that's how Streamlit works) — if every rerun
    called reconcile() again from scratch, any edits made on screen would be
    discarded. So the lower half of this file operates entirely on the mutable
    state in st.session_state["reconciled"]; only clicking "Run Reconciliation"
    again regenerates it from the newly uploaded files and drops prior edits.
    These edits are only valid for the current session — they are included
    when you download outcome.xlsx, but refreshing the page or re-uploading
    the same files will not remember what you did before (per the user's
    request, there is no cross-session persistence yet).

Data-quality flagging, in place of a standalone review section (since
2026-09-11.9):
    There used to be a separate "Needs Review" section listing invoices whose
    extracted data looked untrustworthy (OCR/AI-read, missing date, unusual
    tax rate, etc.), on top of the discrepancy / No Ledger Match tabs. Per the
    user's request that section is gone; instead every row carries this signal
    inline, wherever it already appears:
      - matched / Xero import (flat tables, can't have per-row buttons):
        flagged rows are highlighted (a warm peach, see HIGHLIGHT below) via a
        pandas Styler, and a small picker tool below the table
        (styled_table + render_pdf_viewer) lets you select one of the flagged
        invoices and open its PDF inline.
      - discrepancy / No Ledger Match / No Invoice Found (expander cards):
        each flagged group shows its reasons via st.warning(...) plus a
        "View invoice PDF" button that toggles an inline preview — all inside
        the existing per-invoice expander (render_movable_cases).
    The underlying reasons and the PDF itself come from reconcile.py's
    "quality_flags"/"source_file" fields (see that module's docstring) and
    from pdf_bytes captured into session_state at upload time.
"""

__version__ = "2026-09-11.11"

import base64
import os
import tempfile
from io import BytesIO

import pandas as pd
import streamlit as st

# Streamlit Community Cloud has no .env file — the key lives in
# App -> Settings -> Secrets. ai_extract.py only looks at os.environ / a local
# .env, so we bridge Secrets into an environment variable here. That way local
# development (.env) and the cloud deployment (st.secrets) run the exact same
# code with no changes needed.
if "ANTHROPIC_API_KEY" not in os.environ:
    try:
        os.environ["ANTHROPIC_API_KEY"] = st.secrets["ANTHROPIC_API_KEY"]
    except (KeyError, FileNotFoundError):
        pass

from reconcile import reconcile, find_candidates, write_output, OUT_COLS

NUMERIC_COLS = ("Unit no", "Unit Price", "Amount excl. GST", "GST", "Amount incl. GST")
DATE_COLS = ("Invoice date", "Due Date")
# Warm peach, not yellow: the rest of the app (see .streamlit/config.toml) is
# themed in a fresh sage/mint palette, so a flagged row needs a color that
# still reads as "needs a second look" by contrast rather than blending in.
HIGHLIGHT = "background-color: #FBE3C8"


def _safe_df(rows):
    """Convert to a table safely: every object column is cast to string to
    avoid Arrow type conflicts. rows can be a list[dict] or an existing
    DataFrame."""
    df = pd.DataFrame(rows).copy()
    for c in df.columns:
        if df[c].dtype == "object":
            df[c] = df[c].map(lambda v: "" if v is None or (isinstance(v, float) and pd.isna(v))
                              else (str(int(v)) if isinstance(v, float) and v.is_integer() else str(v)))
    return df


def show(df, **kw):
    st.dataframe(_safe_df(df), width="stretch", hide_index=True, **kw)


def styled_table(rows, **kw):
    """Like show(), but highlights (a warm peach — see HIGHLIGHT) rows whose invoice carries a
    quality flag (OCR/AI-read, unusual tax rate, missing date, etc.) — used
    for matched and the Xero import list. Those are flat tables and can't
    have a per-row "view PDF" button the way the expander-based tabs do;
    render_pdf_viewer() below is the click-to-open counterpart for these two.
    Expects the raw row dicts (with the extra "quality_flags" field, not yet
    trimmed to OUT_COLS) — the extra field itself is never displayed."""
    if not rows:
        st.write("None")
        return
    df = _safe_df(rows)
    cols = [c for c in OUT_COLS if c in df.columns]
    disp = df[cols]
    flagged = df["quality_flags"].map(bool) if "quality_flags" in df.columns else None
    if flagged is None or not flagged.any():
        st.dataframe(disp, width="stretch", hide_index=True, **kw)
        return
    styler = (disp.style
              .apply(lambda row: [HIGHLIGHT if flagged.loc[row.name] else "" for _ in row], axis=1)
              .hide(axis="index"))
    st.dataframe(styler, width="stretch", **kw)


def _render_pdf_inline(data, height=600):
    b64 = base64.b64encode(data).decode()
    st.markdown(
        f'<iframe src="data:application/pdf;base64,{b64}" width="100%" height="{height}" '
        f'style="border:1px solid #ddd" type="application/pdf"></iframe>',
        unsafe_allow_html=True,
    )


def _flagged_pdf_options(rows, pdf_bytes):
    """One entry per invoice (line items collapsed) among `rows` that both
    carries a quality flag and has its original PDF available in pdf_bytes."""
    opts, seen = [], set()
    for grp in _group_contiguous(rows):
        r = grp[0]
        qf, src = r.get("quality_flags"), r.get("source_file")
        if not qf or not src or src not in pdf_bytes:
            continue
        label = f"{r.get('Invoice Number') or '(number not recognized)'} · {r.get('Supplier')} · {src}"
        if label in seen:
            continue
        seen.add(label)
        opts.append((label, qf, src))
    return opts


def render_pdf_viewer(rows, state, key_prefix):
    """Small picker tool for the flat matched/Xero tables: pick a flagged
    invoice from the dropdown, see why it was flagged, and open its PDF
    inline — the click-to-open counterpart to styled_table()'s highlighting,
    for tabs where a per-row button isn't possible."""
    opts = _flagged_pdf_options(rows, state.get("pdf_bytes") or {})
    if not opts:
        return
    st.caption("Rows highlighted above have a data-quality flag (OCR/AI read, unusual tax "
               "rate, missing date, etc.) — pick one below to check the original PDF.")
    labels = [o[0] for o in opts]
    pick = st.selectbox("Flagged invoice", labels, key=f"{key_prefix}_pick")
    _, qf, src = next(o for o in opts if o[0] == pick)
    for reason in qf:
        st.warning(reason)
    toggle_key = f"{key_prefix}_pdfopen_{src}"
    if st.button("📄 Open invoice PDF", key=f"{key_prefix}_openbtn_{src}"):
        st.session_state[toggle_key] = not st.session_state.get(toggle_key, False)
    if st.session_state.get(toggle_key):
        _render_pdf_inline(state["pdf_bytes"][src])


def _to_out_rows(rows):
    """Normalize rows from any source (dicts from the sheets, or dicts
    converted from the pending DataFrame) into a clean dict list containing
    every OUT_COLS field (filling any missing one with None), plus the extra
    "quality_flags"/"source_file" fields when present. This keeps matched /
    discrepancy / pdf_only / excel_only / pending in exactly the same shape,
    so the editing, grouping, and move logic below can share one code path
    instead of a separate one for pending's column differences (it originally
    had no Source/Note column). The extra fields are what let the UI
    highlight a row and offer to open its PDF (see styled_table /
    render_pdf_viewer / render_movable_cases) — they're dropped again before
    anything is written to outcome.xlsx."""
    out = []
    for r in rows:
        d = {c: r.get(c) for c in OUT_COLS}
        for extra in ("quality_flags", "source_file"):
            if extra in r:
                d[extra] = r[extra]
        out.append(d)
    return out


def _ensure_out_cols(r):
    """Make sure a dict contains every field in OUT_COLS (filling any missing
    one with None), without touching any extra fields it already has. Rows
    coming out of excel_loader's pending DataFrame simply don't have a
    Source/Note column — converting straight to dict leaves those keys
    missing, and rendering with OUT_COLS would raise a KeyError (this was
    only caught by writing automated tests — casual manual clicking doesn't
    necessarily hit it; it only blows up once there's an actual pending
    record). Also keeps _status/_row, which write_output() needs for the
    pending sheet."""
    out = dict(r)
    for c in OUT_COLS:
        out.setdefault(c, None)
    return out


def _coerce_row(r):
    """st.data_editor returns edited values as strings (_safe_df cast every
    column to string) — the amount/quantity/date columns need to be converted
    back to real numeric/date types before going into matched. Otherwise
    those cells end up as text in the exported outcome.xlsx and the Xero
    list, which can't be summed in Excel and likely won't be recognized by
    Xero either."""
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
    """Group consecutive rows that share the same (keys) into one group —
    the multiple line items of one invoice are extended into the list
    together by merged_rows()/pdf_rows(), so they're naturally contiguous.
    Grouping this way lets "edit/move" act on the whole invoice at once
    instead of splitting it into separate rows."""
    groups = []
    for row in rows:
        k = tuple(row.get(x) for x in keys)
        if groups and groups[-1]["key"] == k:
            groups[-1]["rows"].append(row)
        else:
            groups.append({"key": k, "rows": [row]})
    return [g["rows"] for g in groups]


def render_movable_cases(groups, state, source_key, title_fn, note_fn, key_prefix,
                          button_label="✅ Confirm & move to matched"):
    """Shared rendering logic for the discrepancy / No Ledger Match /
    No Invoice Found / Not Yet Invoiced tabs: each group (usually the line
    items of one invoice, or one ledger row) is shown in an editable table
    with a button; clicking it moves the (possibly edited) content into
    matched as a group, removes it from its original list, tags it with a
    note explaining this was handled manually, and calls st.rerun() to
    refresh the page.

    When the group carries a quality flag (see reconcile.py's
    quality_flags), its reasons are shown via st.warning(...) right in the
    expander, and — when the underlying PDF was captured — a plain button
    toggles an inline preview underneath. This intentionally avoids
    st.expander/st.popover for the toggle: Streamlit doesn't allow nesting
    either of those inside another expander, and this whole block already
    renders inside one."""
    if not groups:
        st.write("None")
        return
    pdf_bytes = state.get("pdf_bytes") or {}
    for gi, grp in enumerate(groups):
        with st.expander(title_fn(grp)):
            qf = grp[0].get("quality_flags")
            src = grp[0].get("source_file")
            for reason in (qf or []):
                st.warning(reason)
            if src and src in pdf_bytes:
                toggle_key = f"{key_prefix}_pdfopen_{gi}"
                if st.button("📄 View invoice PDF", key=f"{key_prefix}_pdfbtn_{gi}"):
                    st.session_state[toggle_key] = not st.session_state.get(toggle_key, False)
                if st.session_state.get(toggle_key):
                    _render_pdf_inline(pdf_bytes[src])
            edited = st.data_editor(_safe_df(grp)[OUT_COLS], key=f"{key_prefix}_edit_{gi}",
                                     num_rows="fixed", width="stretch")
            if st.button(button_label, key=f"{key_prefix}_move_{gi}"):
                note = note_fn(grp)
                new_rows = [_coerce_row(r) for r in edited.to_dict("records")]
                for r, orig in zip(new_rows, grp):
                    r["Note"] = note
                    for extra in ("quality_flags", "source_file"):
                        if extra in orig:
                            r[extra] = orig[extra]
                state["matched"].extend(new_rows)
                ids = {id(r) for r in grp}
                state[source_key] = [r for r in state[source_key] if id(r) not in ids]
                st.rerun()


def _xero_date(v):
    """Format a date-ish value (python date, pandas Timestamp, NaT, None, or
    a plain string) as "dd/mm/yyyy" — the format Xero's import expects.
    Returns "" when there's nothing usable, rather than letting a NaT or a
    datetime's "00:00:00" time component leak into the exported cell."""
    if v is None:
        return ""
    if isinstance(v, float) and pd.isna(v):
        return ""
    ts = pd.to_datetime(v, errors="coerce")
    if pd.isna(ts):
        return ""
    return ts.strftime("%d/%m/%Y")


def _xero_export(rows):
    """Export the current matched rows as a downloadable list, using the same
    columns (OUT_COLS) as the matched sheet in outcome.xlsx. This is not an
    official Xero import template — matching Xero's actual Bills import
    format would need extra mappings like AccountCode/TaxType. Once we know
    exactly what fields Xero needs, this can be adjusted; for now it's a
    general-purpose, reviewable list.

    Invoice date / Due Date are written as plain "dd/mm/yyyy" text, per
    Xero's expected import format — not a datetime cell (which Excel would
    otherwise show with a trailing "00:00:00") and never blank via NaT. When
    an invoice's Due Date wasn't extracted, it falls back to that same
    invoice's Invoice date rather than exporting an empty cell."""
    buf = BytesIO()
    df = pd.DataFrame(rows)[OUT_COLS] if rows else pd.DataFrame(columns=OUT_COLS)
    if not df.empty:
        df["Invoice date"] = df["Invoice date"].map(_xero_date)
        df["Due Date"] = df["Due Date"].map(_xero_date)
        df["Due Date"] = df["Due Date"].where(df["Due Date"] != "", df["Invoice date"])
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        df.to_excel(w, index=False, sheet_name="To Import to Xero")
    buf.seek(0)
    return buf


st.set_page_config(page_title="Purchase Invoice Reconciliation", layout="wide")
st.title("Purchase Invoice Reconciliation")
st.caption("Reads Excel directly, extracts PDFs, matches exactly by invoice number")

# All modules must be on the same version. Forgetting to replace every file is
# the most common failure mode, so this check is placed front and center.
import excel_loader as _el, invoice_extractor as _ie, reconcile as _rc
_vers = {"app": __version__, "reconcile": getattr(_rc, "__version__", "?"),
         "invoice_extractor": getattr(_ie, "__version__", "?"),
         "excel_loader": getattr(_el, "__version__", "?")}
try:
    import ocr as _ocr
    _vers["ocr"] = getattr(_ocr, "__version__", "?")
    _ocr_ok, _ocr_path = _ocr.available(), _ocr.where()
except ImportError:
    _vers["ocr"] = "not installed"
    _ocr_ok, _ocr_path = False, None

if len(set(_vers.values())) > 1:
    st.error("File versions don't match — please re-upload all .py files: "
             + "; ".join(f"{k} {v}" for k, v in _vers.items()))
else:
    st.caption(f"Version {__version__} · OCR "
               + (f"ready ({_ocr_path})" if _ocr_ok else "unavailable"))

uploads = st.file_uploader(
    "Drag in the ledger Excel and invoice PDF(s) together (multiple files OK, any order)",
    type=["xlsx", "xlsm", "pdf"], accept_multiple_files=True,
)

xlsx_ups = [f for f in (uploads or []) if f.name.lower().endswith((".xlsx", ".xlsm"))]
pdf_ups = [f for f in (uploads or []) if f.name.lower().endswith(".pdf")]

c1, c2 = st.columns(2)
c1.metric("Ledger Excel", len(xlsx_ups))
c2.metric("Invoice PDFs", len(pdf_ups))

if not uploads:
    st.info("Drag in the ledger and invoices together. You can upload a single PDF for a one-off check, or a whole month's worth for batch reconciliation.")
    st.stop()

if not xlsx_ups:
    st.warning("Still missing the ledger Excel (.xlsx).")
    st.stop()
if not pdf_ups:
    st.warning("Still missing invoice PDFs.")
    st.stop()

if len(xlsx_ups) > 1:
    names = [f.name for f in xlsx_ups]
    pick = st.selectbox("Multiple Excel files were uploaded — which one is the ledger?", names)
    xlsx_file = next(f for f in xlsx_ups if f.name == pick)
else:
    xlsx_file = xlsx_ups[0]
pdf_files = pdf_ups

run_clicked = st.button("Run Reconciliation", type="primary")

if run_clicked:
    # Uploads are in-memory objects; write them to a temp dir so we can reuse
    # the exact same logic as the command-line version.
    with tempfile.TemporaryDirectory() as tmp:
        xlsx_path = os.path.join(tmp, "tracking.xlsx")
        with open(xlsx_path, "wb") as f:
            f.write(xlsx_file.getbuffer())

        # Also keep every PDF's raw bytes in memory (keyed by filename), so
        # the UI can offer to preview them later — the temp dir itself is
        # gone by the time the page re-renders on the next interaction.
        pdf_paths = []
        pdf_bytes_by_name = {}
        for up in pdf_files:
            data = bytes(up.getbuffer())
            pdf_bytes_by_name[up.name] = data
            p = os.path.join(tmp, up.name)
            with open(p, "wb") as f:
                f.write(data)
            pdf_paths.append(p)

        with st.spinner(f"Processing {len(pdf_paths)} invoices…"):
            sheets, pending, ctx = reconcile(xlsx_path, pdf_paths)

        # Backward compatible with older reconcile.py: degrade gracefully on
        # missing fields instead of crashing the whole page.
        ctx.setdefault("superseded", [])
        ctx.setdefault("unmatched_pdf", [])
        ctx.setdefault("failed", [])

        # pending keeps two extra columns, _status/_row, that matched/
        # discrepancy/etc. don't — write_output() needs them for the
        # "pending (no invoice yet)" sheet, so it can't be trimmed down to
        # just OUT_COLS like the others or the export will raise a KeyError.
        # When displaying/editing (in render_movable_cases) only OUT_COLS is
        # picked out, so these two columns never show up in the edit form —
        # they're kept purely for the export.
        pending_rows = [_ensure_out_cols(r) for r in
                        pending.drop(columns=["invoice_key", "po_key", "supplier_key"],
                                     errors="ignore").to_dict("records")]

        # Clicking "Run Reconciliation" always recomputes everything from
        # scratch: any prior edits/moves/ticks are discarded and rebuilt —
        # per the user's request, those actions are only valid for the
        # current session and don't need to persist across runs.
        st.session_state["reconciled"] = {
            "matched": _to_out_rows(sheets["matched"]),
            "discrepancy": _to_out_rows(sheets["discrepancy"]),
            "pdf_only": _to_out_rows(sheets["PDF to Excel - not match"]),
            "excel_only": _to_out_rows(sheets["Excel to PDF - not match"]),
            "pending": pending_rows,
            "pdf_bytes": pdf_bytes_by_name,
            "ctx": ctx,
        }

if "reconciled" not in st.session_state:
    st.stop()

state = st.session_state["reconciled"]
ctx = state["ctx"]
n_inv = len(ctx["invoices"])

# The download is regenerated every time from the current state (which may
# have been edited/moved since the run), so the downloaded outcome.xlsx
# always reflects every manual action taken on screen, not just a snapshot
# from the moment reconciliation finished.
_sheets_now = {
    "matched": state["matched"],
    "discrepancy": state["discrepancy"],
    "PDF to Excel - not match": state["pdf_only"],
    "Excel to PDF - not match": state["excel_only"],
}
_pending_now = pd.DataFrame(state["pending"]) if state["pending"] else \
    pd.DataFrame(columns=list(OUT_COLS) + ["_status", "_row"])
_buf = BytesIO()
write_output(_sheets_now, _pending_now, _buf, ctx.get("superseded"))
_buf.seek(0)

m1, m2, m3, m4, m5 = st.columns(5)
m1.metric("Invoices read", n_inv)
m2.metric("Matched", len({r["Invoice Number"] for r in state["matched"]}))
m3.metric("Amount mismatch", len(_group_contiguous(state["discrepancy"])))
m4.metric("No Ledger Match", len(_group_contiguous(state["pdf_only"])))
m5.metric("No Invoice Found", len(state["excel_only"]))

st.download_button("Download outcome.xlsx", _buf,
                   file_name="Neocytogen - outcome.xlsx",
                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

if ctx["superseded"]:
    with st.expander(f"Ignored Proformas ({len(ctx['superseded'])})"):
        st.caption("This PO already has a Tax Invoice, and by rule only the Tax Invoice is counted. Listed here for reference only.")
        show(ctx["superseded"])

# Extraction warnings: fields that couldn't be extracted, line totals that
# don't add up, etc. — surfaced up front so it's clear which results might
# not be fully trustworthy.
warned = [i for i in ctx["invoices"].values() if i["warnings"]]
if warned or ctx["failed"]:
    with st.expander(f"Extraction warnings ({len(warned) + len(ctx['failed'])})", expanded=False):
        for inv in warned + ctx["failed"]:
            st.write(f"**{inv.get('invoice_no') or inv['source_file']}** — "
                     + "; ".join(inv["warnings"]))

tab_xero, tab_matched, tab_disc, tab_pdf_only, tab_excel_only, tab_pending = st.tabs(
    ["📤 To Import to Xero", "matched", "discrepancy", "No Ledger Match", "No Invoice Found", "Not Yet Invoiced"]
)

with tab_xero:
    st.caption("Every matched invoice shows up here automatically — both the ones that matched "
               "automatically and the ones manually confirmed from other tabs. Download this list "
               "directly to review before importing into Xero. Highlighted rows have a "
               "data-quality flag — see the picker below to check the original PDF.")
    styled_table(state["matched"])
    render_pdf_viewer(state["matched"], state, key_prefix="xero")
    st.download_button("Download Xero import list", _xero_export(state["matched"]),
                       file_name="To Import to Xero.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                       key="xero_download")

with tab_matched:
    st.caption("Highlighted rows were read by OCR/AI or have another data-quality flag "
               "worth a second look — use the picker below to open the original PDF.")
    styled_table(state["matched"])
    render_pdf_viewer(state["matched"], state, key_prefix="matched")

with tab_disc:
    st.caption("Invoices whose amount doesn't match the ledger. You can edit the content directly "
               "below (e.g. if the ledger itself was wrong, or the invoice's actual amount should "
               "win) — click the button after editing to move the whole invoice to matched (it will "
               "also appear in To Import to Xero). A yellow warning under an invoice's title means "
               "its extracted data also has a quality flag — open the PDF there to double check.")
    render_movable_cases(
        _group_contiguous(state["discrepancy"]), state, "discrepancy",
        title_fn=lambda grp: (f"Invoice {grp[0].get('Invoice Number') or '(number not recognized)'} · "
                               f"{grp[0].get('Supplier')} · {(grp[0].get('Note') or '')[:60]}"),
        note_fn=lambda grp: f"Manually confirmed as matched after review (original auto-match result: {grp[0].get('Note') or ''})",
        key_prefix="disc",
    )

with tab_pdf_only:
    st.caption("There's an invoice PDF, but no matching entry in the ledger (neither invoice number "
               "nor PO matches). Once you've confirmed it's correct, you can manually match it to "
               "matched — moving it adds a note making clear this wasn't an automatic match.")
    render_movable_cases(
        _group_contiguous(state["pdf_only"]), state, "pdf_only",
        title_fn=lambda grp: (f"Invoice {grp[0].get('Invoice Number') or grp[0].get('PO Number') or '(not recognized)'} · "
                               f"{grp[0].get('Supplier')}"),
        note_fn=lambda grp: f"⚠ Not auto-matched to the ledger — manually added to matched (original note: {grp[0].get('Note') or ''})",
        key_prefix="pdfonly", button_label="➕ Manually add to matched",
    )

with tab_excel_only:
    st.caption("The ledger has this entry, but no matching invoice PDF was found. If you already "
               "have the invoice and it just wasn't uploaded or recognized this time, you can "
               "confirm it manually here — moving it adds a note making clear this wasn't an "
               "automatic match.")
    render_movable_cases(
        [[r] for r in state["excel_only"]], state, "excel_only",
        title_fn=lambda grp: f"{grp[0].get('Invoice Number') or grp[0].get('PO Number')} · {grp[0].get('Supplier')}",
        note_fn=lambda grp: f"⚠ No matching invoice PDF found for this ledger entry — manually added to matched (original note: {grp[0].get('Note') or ''})",
        key_prefix="exlonly", button_label="➕ Manually add to matched",
    )

with tab_pending:
    st.caption("Goods not yet received or not yet invoiced — not counted as a mismatch. If it's "
               "actually already been invoiced and the status just hasn't been updated yet, you "
               "can confirm it manually here — moving it adds a note making clear this wasn't an "
               "automatic match.")
    render_movable_cases(
        [[r] for r in state["pending"]], state, "pending",
        title_fn=lambda grp: f"{grp[0].get('Invoice Number') or grp[0].get('PO Number')} · {grp[0].get('Supplier')}",
        note_fn=lambda grp: "⚠ Originally marked not yet invoiced / goods not received — manually added to matched",
        key_prefix="pending", button_label="➕ Manually add to matched",
    )

# When nothing in the ledger matches, list the closest candidates for manual review.
if ctx["unmatched_pdf"]:
    st.divider()
    st.subheader("Suggestions for Manual Review")
    st.caption("These are ranked for reference only — the program never auto-matches on this basis; "
               "the invoice number must match exactly to count as matched.")
    for inv in ctx["unmatched_pdf"]:
        total = (inv.get("subtotal") or {}).get("amount_incl")
        with st.expander(f"Invoice {inv.get('invoice_no')} · {inv.get('supplier')} · {total}"):
            cands = find_candidates(ctx["expected"], inv)
            if cands:
                show(cands)
            else:
                st.write("No similar record found in the ledger.")
            for w in inv["warnings"]:
                st.warning(w)
