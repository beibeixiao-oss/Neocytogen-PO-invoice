__version__ = "2026-09-11.7"

"""
ai_extract.py — reads invoices with a vision model (only for the ones that regex and OCR can't handle)

Call strategy: native-text PDFs go through regex — fast, accurate, and free;
only scanned documents, or ones where regex couldn't extract an invoice
number, get sent to the model. Out of 60 invoices, usually fewer than 10
actually get sent.

The model only "reads" — it doesn't "judge": it emits fields according to a
fixed schema, and matching, amount validation, and de-duplication are still
handled by deterministic code. Financial reconciliation can't let a model draw conclusions.

The key lives in a .env file in the same directory:
    ANTHROPIC_API_KEY=sk-ant-xxxx
Don't commit .env, and don't distribute it with the scripts.
"""

import base64
import json
import os
import re

MODEL = "claude-sonnet-5"      # switch to claude-haiku-4-5-20251001 for lower cost
MAX_PAGES = 3                  # the invoice body is usually on the first few pages; later pages are mostly terms and remittance info
DPI = 200

_client = None

SCHEMA_PROMPT = """You are reading a supplier invoice for a reconciliation system.

Return ONLY a JSON object, no prose, no markdown fences, with exactly these keys:

{
  "invoice_no": string or null,      // the supplier's invoice number, exactly as printed
  "invoice_date": string or null,    // ISO format YYYY-MM-DD
  "due_date": string or null,        // ISO format YYYY-MM-DD
  "po_no": string or null,           // buyer's purchase order, usually starts with PONCG
  "supplier": string or null,        // the company ISSUING the invoice
  "currency": string or null,        // ISO code: SGD, USD, EUR...
  "amount_excl_gst": number or null, // invoice total BEFORE tax
  "gst": number or null,             // tax amount only, NOT the tax rate
  "amount_incl_gst": number or null, // invoice total AFTER tax
  "is_proforma": true or false,      // true only if the document says Proforma
  "line_items": [
    {"description": string, "qty": number or null, "unit_price": number or null,
     "amount_excl_gst": number or null, "amount_incl_gst": number or null}
  ]
}

Rules:
- The buyer is Neocytogen Therapeutics. NEVER return Neocytogen as the supplier.
- gst is an amount, not a percentage. If the invoice shows "Tax Amount 9.000 % 11.71",
  gst is 11.71, not 9.
- amount_excl_gst + gst must equal amount_incl_gst. If the invoice has no tax, gst is 0.
- Use null when a field is genuinely not printed. Do not guess, do not infer from
  other fields, do not invent an invoice number.
- Copy the invoice number character for character, including letters and hyphens.
"""


def _load_env():
    """Read the .env file in the same directory, with the simplest possible parsing and no extra dependencies."""
    if os.environ.get("ANTHROPIC_API_KEY"):
        return os.environ["ANTHROPIC_API_KEY"]
    for d in (os.path.dirname(os.path.abspath(__file__)), os.getcwd()):
        f = os.path.join(d, ".env")
        if os.path.exists(f):
            for line in open(f, encoding="utf-8"):
                line = line.strip()
                if line.startswith("ANTHROPIC_API_KEY"):
                    v = line.split("=", 1)[-1].strip().strip('"').strip("'")
                    if v:
                        os.environ["ANTHROPIC_API_KEY"] = v
                        return v
    return None


def available():
    """Only available when there's a key and the SDK is installed. When unavailable, callers should skip it rather than error out."""
    try:
        import anthropic          # noqa: F401
    except ImportError:
        return False
    return bool(_load_env())


def _pages_as_png(path, max_pages=MAX_PAGES):
    import fitz
    out = []
    doc = fitz.open(path)
    for page in doc[:max_pages]:
        pix = page.get_pixmap(dpi=DPI)
        out.append(pix.tobytes("png"))
    doc.close()
    return out


def _parse_json(text):
    """The model occasionally wraps the output in ``` fences or adds a lead-in sentence — this tolerates both."""
    t = text.strip()
    t = re.sub(r"^```(?:json)?\s*|\s*```$", "", t)
    try:
        return json.loads(t)
    except json.JSONDecodeError:
        m = re.search(r"\{.*\}", t, re.S)
        if m:
            try:
                return json.loads(m.group(0))
            except json.JSONDecodeError:
                pass
    return None


def extract(path, filename_hint=""):
    """Returns a field dict shaped the same way as invoice_extractor's; returns None on failure."""
    global _client
    if not available():
        return None
    import anthropic

    if _client is None:
        _client = anthropic.Anthropic(api_key=_load_env())

    try:
        images = _pages_as_png(path)
    except Exception:
        return None
    if not images:
        return None

    content = [{"type": "image",
                "source": {"type": "base64", "media_type": "image/png",
                           "data": base64.b64encode(p).decode()}}
               for p in images]
    hint = f"\n\nThe file is named: {filename_hint}" if filename_hint else ""
    content.append({"type": "text", "text": SCHEMA_PROMPT + hint})

    try:
        resp = _client.messages.create(
            model=MODEL, max_tokens=2000,
            messages=[{"role": "user", "content": content}],
        )
    except Exception as e:
        return {"_error": f"{type(e).__name__}: {e}"}

    text = "".join(b.text for b in resp.content if getattr(b, "type", "") == "text")
    data = _parse_json(text)
    if not data:
        return {"_error": "Model did not return valid JSON"}

    # The model might mistake the tax rate for the tax amount, or the three amounts might not agree — the equation catches that
    e, g, i = (data.get("amount_excl_gst"), data.get("gst"), data.get("amount_incl_gst"))
    if None not in (e, g, i) and abs(e + g - i) > 0.05:
        data["gst"] = round(i - e, 2)
        data["_amount_fixed"] = True
    if g is None and None not in (e, i):
        data["gst"] = round(i - e, 2)

    return data
