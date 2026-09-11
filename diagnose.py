"""
diagnose.py — single-PDF troubleshooting tool

    python diagnose.py "path\\some_scanned_file.pdf"

Checks, in order: how much text pdfplumber can read -> whether OCR is
available -> how much text OCR reads -> which fields end up extracted. It's
obvious at a glance where the chain breaks.
"""

import sys
import os


def main(path):
    print("=" * 60)
    print("File:", os.path.basename(path))
    print("Exists:", os.path.exists(path))
    if not os.path.exists(path):
        return

    # 1. Native text
    try:
        import pdfplumber
        with pdfplumber.open(path) as pdf:
            native = "\n".join((p.extract_text() or "") for p in pdf.pages)
        print(f"\n[1] pdfplumber native text: {len(native)} characters  ({len(pdf.pages) if False else ''})")
        if native.strip():
            print("    -> Has text, won't go through OCR. First 120 chars:")
            print("   ", repr(native[:120]))
    except Exception as e:
        native = ""
        print("\n[1] pdfplumber failed:", e)

    # 2. OCR environment
    try:
        import ocr
        ok = ocr.available()
        print(f"\n[2] OCR available: {ok}   path: {ocr.where()}")
    except ImportError as e:
        print("\n[2] Could not import ocr:", e)
        return
    except Exception as e:
        print("\n[2] OCR detection error:", e)
        return

    # 3. Actual recognition
    if not native.strip() and ok:
        print("\n[3] Running OCR, please wait...")
        try:
            plain, layout = ocr.ocr_pdf(path)
            print(f"    OCR plain text: {len(plain)} characters / layout text: {len(layout)} characters")
            if plain.strip():
                print("    First 400 chars:")
                print("   ", repr(plain[:400]))
            else:
                print("    !! Not a single character was recognized — the image may be too blurry, skewed, or purely graphical")
        except Exception as e:
            print("    OCR error:", type(e).__name__, e)
    elif native.strip():
        print("\n[3] Skipping OCR (already has text)")

    # 4. Full extraction result
    print("\n[4] Extraction result:")
    try:
        from invoice_extractor import extract_invoice
        r = extract_invoice(path)
        for k in ["invoice_no", "invoice_date", "po_no", "due_date",
                  "supplier", "currency", "is_proforma", "ocr"]:
            print(f"    {k:14} {r.get(k)}")
        print(f"    subtotal       {r.get('subtotal')}")
        print(f"    line items     {len(r.get('line_items') or [])}")
        print(f"    warnings       {r.get('warnings')}")
        if "ocr" not in r:
            print("\n    !! The result has no 'ocr' field — invoice_extractor.py is an old version, please update")
    except Exception as e:
        import traceback
        traceback.print_exc()


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print('Usage: python diagnose.py "some_scanned_file.pdf"')
    else:
        main(sys.argv[1])
