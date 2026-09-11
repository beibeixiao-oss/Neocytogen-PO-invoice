"""
ocr.py — text recognition for scanned PDF invoices

Only kicks in when pdfplumber can't extract any text. Uses PyMuPDF to render
pages as high-resolution images, then hands them to Tesseract for recognition.

Dependencies:
    pip install pytesseract pymupdf pillow
    Tesseract itself must also be installed (Windows: https://github.com/UB-Mannheim/tesseract/wiki)

Note: OCR text is far less accurate than native PDF text, and digits are
especially error-prone (0/O, 1/l/I, 5/S, 8/B). So OCR results always go
through the amount-equation validation; anything that fails goes into "needs
review" — it's never trusted outright.
"""

import os
import re
import shutil

__version__ = "2026-09-11.7"

DPI = 300          # Recognition quality drops noticeably below 300; higher gives little benefit and is much slower
_TESS_OK = None


def _find_tesseract():
    """Locate tesseract.exe.

    On Windows it's often not on PATH after install, and the Tesseract-OCR
    folder in the Start menu is just a shortcut, not the install directory.
    So this tries, in order: PATH -> registry -> common install paths -> a
    broad search of common locations.
    """
    import pytesseract

    if shutil.which("tesseract"):
        return True

    cands = []

    # The UB Mannheim installer writes the install directory to the registry
    if os.name == "nt":
        try:
            import winreg
            for root, key in [
                (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Tesseract-OCR"),
                (winreg.HKEY_CURRENT_USER, r"SOFTWARE\Tesseract-OCR"),
                (winreg.HKEY_LOCAL_MACHINE,
                 r"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\Tesseract-OCR"),
            ]:
                for name in ("InstallDir", "Path", "InstallLocation"):
                    try:
                        with winreg.OpenKey(root, key) as k:
                            v, _ = winreg.QueryValueEx(k, name)
                            if v:
                                cands.append(os.path.join(v, "tesseract.exe"))
                    except OSError:
                        continue
        except ImportError:
            pass

    cands += [
        r"C:\Program Files\Tesseract-OCR\tesseract.exe",
        r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe",
        os.path.expanduser(r"~\AppData\Local\Programs\Tesseract-OCR\tesseract.exe"),
        os.path.expanduser(r"~\AppData\Local\Tesseract-OCR\tesseract.exe"),
        "/opt/homebrew/bin/tesseract",
        "/usr/local/bin/tesseract",
        "/usr/bin/tesseract",
    ]

    # Still not found — search one level deep under common root directories
    if os.name == "nt":
        import glob
        for root in [r"C:\Program Files", r"C:\Program Files (x86)",
                     os.path.expanduser(r"~\AppData\Local\Programs"),
                     os.path.expanduser(r"~\AppData\Local")]:
            cands += glob.glob(os.path.join(root, "*esseract*", "tesseract.exe"))

    for c in cands:
        if c and os.path.exists(c):
            pytesseract.pytesseract.tesseract_cmd = c
            return True
    return False


def where():
    """Returns the tesseract path actually in use, to make troubleshooting easier."""
    import pytesseract
    return (shutil.which("tesseract")
            or getattr(pytesseract.pytesseract, "tesseract_cmd", None))


def available():
    """Whether OCR is available. When it isn't, callers should skip it rather than error out."""
    global _TESS_OK
    if _TESS_OK is None:
        try:
            import pytesseract, fitz          # noqa: F401
            _TESS_OK = _find_tesseract()
        except ImportError:
            _TESS_OK = False
    return _TESS_OK


def ocr_pdf(path, max_pages=4):
    """Returns (plain text, layout text), matching pdfplumber's two modes so
    downstream parsing logic doesn't need to change at all. Returns ("", "")
    on failure."""
    if not available():
        return "", ""
    import fitz
    import pytesseract
    from PIL import Image

    plain, layout = [], []
    try:
        doc = fitz.open(path)
    except Exception:
        return "", ""

    for page in doc[:max_pages]:
        try:
            pix = page.get_pixmap(dpi=DPI)
            img = Image.frombytes("RGB", (pix.width, pix.height), pix.samples)
            if img.mode != "L":
                img = img.convert("L")          # Grayscale — more stable for receipt/invoice recognition
            # psm 6: treat the whole page as one uniform text block, better suited to invoice layouts than the default mode
            plain.append(pytesseract.image_to_string(img, config="--psm 6"))
            # Keep character positions, for the "label above, value below" column-alignment logic to use
            layout.append(_layout_from_data(
                pytesseract.image_to_data(img, config="--psm 6",
                                          output_type=pytesseract.Output.DICT)))
        except Exception:
            continue
    doc.close()
    return "\n".join(plain), "\n".join(layout)


def _layout_from_data(data, char_w=9):
    """Reconstruct Tesseract's per-word coordinates into fixed-width layout
    text, so _label_below's column alignment also works on scanned documents."""
    rows = {}
    n = len(data.get("text", []))
    for i in range(n):
        word = (data["text"][i] or "").strip()
        if not word:
            continue
        try:
            if int(data["conf"][i]) < 30:      # Drop words with too-low confidence outright
                continue
        except (ValueError, TypeError):
            pass
        line_key = (data["block_num"][i], data["par_num"][i], data["line_num"][i])
        rows.setdefault(line_key, []).append((data["left"][i] // char_w, word))

    out = []
    for key in sorted(rows):
        line = ""
        for col, word in sorted(rows[key]):
            if col > len(line):
                line += " " * (col - len(line))
            line += word + " "
        out.append(line.rstrip())
    return "\n".join(out)


def fix_ocr_digits(s):
    """Only corrects common confusions on fields that are expected to be
    numeric — never a blanket text-wide replace, which would turn the O in a
    company name into a 0 and create a new error."""
    if s is None:
        return None
    return (str(s).replace("O", "0").replace("o", "0")
            .replace("l", "1").replace("I", "1")
            .replace("S", "5").replace("B", "8"))
