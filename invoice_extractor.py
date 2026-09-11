"""
invoice_extractor.py — extracts structured fields from an invoice PDF

Design choices:
  * Prefer pdfplumber's table extraction for line items (works best on native-text PDFs)
  * Header fields use "label regex" patterns — each field lists multiple label
    aliases to cover different supplier templates
  * The filename itself is structured (PO - date - type - supplier) and is used
    as a second, cross-checking signal
  * A field that can't be extracted is always left as None — never guessed
"""

import os
import re
from datetime import datetime

import pdfplumber

__version__ = "2026-09-11.7"

# Multiple label spellings for the same field, tried in order
# The buyer is Neocytogen — any supplier result that extracts to this is wrong
BUYER_HINTS = ["NEOCYTOGEN"]

# An invoice number is never one of these (a mis-captured label or body word)
INVOICE_NO_STOPWORDS = {
    "ON", "IN", "DO", "DATE", "GST", "NO", "TO", "OF", "THE", "INVOICE",
    "TAX", "BILL", "SHIP", "PO", "TOTAL", "AMOUNT", "PAGE", "REF",
}

LABEL_PATTERNS = {
    "invoice_no": [
        r"(?:Tax\s*)?Invoice\s*(?:NO|No\.?|Number|#)[ \t]*[:.#]?[ \t]*([A-Za-z0-9][\w\-/]{2,})",
        r"(?:TAX\s+)?INVOICE\s*\n\s*No\.?[ \t]*[:.#]?[ \t]*([A-Za-z0-9][\w\-/]{2,})",
        r"(?:Tax\s*)?Invoice\s*(?:NO|No\.?|Number|#)[ \t]*[:.#]?[ \t]*\n[ \t]*([A-Za-z0-9][\w\-/]{2,})",
        r"\bInv(?:oice)?\s*(?:No\.?|#)[ \t]*([A-Za-z0-9][\w\-/]{2,})",
        r"\b(?:Document|Doc)\s*(?:No\.?|Number)[ \t]*[:.#]?[ \t]*([A-Za-z0-9][\w\-/]{2,})",
        r"\bP\.?I\.?\s*(?:NO|No)\.?[ \t]*[:.#]?[ \t]*([A-Za-z0-9][\w\-/]{2,})",
        # Edge case (Roylab): this invoice's layout uses the document title itself
        # as the label — "Tax Invoice INV/2026/0117" — the invoice number follows
        # the title directly, with no "No./Number/#" word in between, so none of
        # the rules above match. We don't treat any word after "Tax Invoice" as
        # the invoice number (many invoices have a company name or address right
        # after the title instead) — so the following word must itself look like
        # an invoice number: starts with letters, has a - or / in the middle,
        # followed by digits, which rules out ordinary words and place names.
        r"\bTax\s+Invoice\s+([A-Za-z]{2,6}[\-/][\dA-Za-z\-/]{3,})\b",
    ],
    "invoice_date": [
        # Lonza: the label "Our invoice" and "<number> dated <date>" are split
        # across a line break, and earlier in the body there's a standalone line
        # "dated 27-Jan-2026" (which is actually the order date). So "Our invoice"
        # must be anchored across lines, and tried before the generic Date rule.
        r"Our\s*invoice\b[\s\S]{0,80}?dated\s*(\d{1,2}[-\s]*[A-Za-z]{3,9}[-\s]*\d{4})",
        r"Invoice\s*Date\s*[:.]?\s*([\d]{1,4}[./\-][\d]{1,2}[./\-][\d]{1,4})",
        r"(?<!Due)(?<!Due )\bDate[dD]?\s*[:.]\s*([\d]{1,4}[./\-][\d]{1,2}[./\-][\d]{1,4})",
        r"Invoice\s*Date[ \t]*[:.]?\s*\n?\s*(\d{1,2}\s*[A-Za-z]{3,9}\s*\d{4})",
        # If "dated" is preceded by Your order / Our order, that's the order date, not the invoice date.
        # Lonza's header reads "Your order PONCG... dated 27-Jan-2026" while the actual invoice date is 05-Feb-2026.
        r"(?<!Due)(?<!Due )(?<!order )(?<!Order )\bDate[dD]?\s*[:.]?\s*(\d{1,2}\s*[A-Za-z]{3,9}\s*\d{4})",
        r"Date\s*of\s*Invoice\s*[:.]?\s*([\d]{2,4}[./\-][\d]{1,2}[./\-][\d]{1,4})",
    ],
    "po_no": [
        r"(?<![A-Za-z0-9])(PONCG\d{9,})(?!\d)",     # Neocytogen's own PO format, the most reliable
        r"P\.?O\.?\s*(?:NO|No|Number)\s*[:.]?\s*([\w\-/]+)",
        r"Your\s*Ref\.?\s*[:.]?\s*([\w\-/]+)",
        r"Purchase\s*Order\s*[:.]?\s*\n?\s*([\w\-/]+)",
    ],
    "due_date": [
        r"Due\s*Date\s*[:.]?\s*([\d]{2,4}[./\-][\d]{1,2}[./\-][\d]{1,4})",
    ],
}

CURRENCY_HINTS = [
    # USD must be checked first: US$ contains S$, otherwise a USD invoice gets misread as SGD (amounts off by 30%+)
    (r"\bUSD\b|US\$", "USD"),
    (r"\bSGD\b|(?<![A-Za-z])S\$", "SGD"),
    (r"\bEUR\b|€", "EUR"),   (r"\bJPY\b|¥", "JPY"),
    (r"\bGBP\b|£", "GBP"),
]

QTY_HEADS = {"qty", "quantity", "units", "unit no", "no of units"}
DESC_HEADS = {"description", "item description", "item", "particulars", "product"}


def _money(raw):
    """'S$1,234.56' -> 1234.56 ; empty/dash -> None"""
    if raw is None:
        return None
    s = re.sub(r"[^\d.\-]", "", str(raw))
    if s in ("", "-", "."):
        return None
    try:
        return round(float(s), 2)
    except ValueError:
        return None


def _date(raw):
    if not raw:
        return None
    s = str(raw).strip()
    m = re.fullmatch(r"(\d{1,2})\s*([A-Za-z]{3,9})\s*(\d{4})", s)
    if m:
        s = f"{m.group(1)} {m.group(2)[:3]} {m.group(3)}"
    if not re.search(r"[A-Za-z]", s):           # Numeric-only dates keep the separator style as-is
        s = s.replace(".", "-").replace("/", "-")
    for fmt in ("%Y-%m-%d", "%d-%m-%Y", "%m-%d-%Y", "%d-%b-%Y",
                "%d %b %Y", "%d %B %Y", "%Y-%m-%d %H:%M:%S"):
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            continue
    return None


def normalize_cmp(s):
    return re.sub(r"[^a-z0-9]", "", str(s or "").lower())


def valid_invoice_no(v):
    """An invoice number must actually look like one: contains a digit, is long
    enough, and isn't a mis-captured label word. This one check alone filters
    out junk values like on / Date / GST / in / DO."""
    if not v:
        return False
    s = str(v).strip()
    if s.upper() in INVOICE_NO_STOPWORDS:
        return False
    if not re.search(r"\d", s):          # No digits at all -> definitely not an invoice number
        return False
    if len(re.sub(r"\W", "", s)) < 4:
        return False
    if re.fullmatch(r"(19|20)\d{2}", s):  # A bare year
        return False
    # A mis-captured address (things like 114AJlnJurongKechil)
    if re.search(r"\b(Jln|Jalan|Road|Rd|Street|St|Ave|Avenue|Blk|Block|Singapore|Lorong)\b", s, re.I):
        return False
    if re.search(r"(Jln|Jalan|Road|Street|Avenue|Singapore|Lorong)", s, re.I):
        return False
    return True


def clean_supplier(name):
    """Strip out labels, tax IDs, and the buyer's name that end up mixed into the header."""
    if not name:
        return None
    s = re.sub(r"\b(TAX\s+)?INVOICE\b", " ", str(name), flags=re.I)
    s = re.sub(r"\b(Bill|Ship|Sold)\s*To\s*:?", " ", s, flags=re.I)
    s = re.sub(r"\bGST\s*Reg(?:istration)?\s*(?:No|Number)?\s*[:.]?\s*[\w\-]+", " ", s, flags=re.I)
    s = re.sub(r"\b(UEN|Customer\s*Code|Invoice\s*Date)\b.*", " ", s, flags=re.I)
    s = re.sub(r"\s{2,}", " ", s).strip(" ,.:;-")
    if not s or any(b in s.upper() for b in BUYER_HINTS):
        return None                       # Extracting to the buyer means it's wrong — better to leave it blank
    if len(s) < 3 or not re.search(r"[A-Za-z]", s):
        return None
    # If all that's left is a company suffix (LTD / PTE LTD / INC), the header wasn't captured in full
    if re.fullmatch(r"(PTE|LTD|LIMITED|INC|LLC|CO|CORP|BHD|SDN|GMBH|[.\s&,]+)+", s, re.I):
        return None
    return s


# Filename parentheses often contain more than just the supplier name: "item 1
# out of 3", "8 bottles MEM", "Tax Invoice No. MCIN060317". These aren't
# company names.
NOTE_PAT = re.compile(
    r"(out\s*of|\bitem\b|\bpc?s\b|\bpieces?\b|\bbottles?\b|\bboxes?\b|\bvials?\b"
    r"|\bfinal\b|\btax\s*invoice\b|\bproforma\b|\binvoice\s*no\b|^no\.|^\d)", re.I)


def _is_note(s):
    s = s.strip()
    if not s or NOTE_PAT.search(s):
        return True
    letters = sum(c.isalpha() for c in s)
    return letters < 3 or letters < len(s) * 0.4        # Too high a proportion of digits/symbols


def parse_filename(path):
    """'PONCG202512185 - 2026-01-14 Tax Invoice (Genscript).pdf'
       -> {'po_no', 'invoice_date', 'supplier'}"""
    stem = os.path.splitext(os.path.basename(path))[0]
    out = {}
    # Can't use \b here: underscore is a word character, so in
    # "PONCG202601008_-_..." there's no word boundary between the digits and
    # the underscore, and \b never matches there — the filename PO fallback
    # was broken for a long time because of this.
    m = re.search(r"(?<![A-Za-z0-9])(PONCG\d{9,})(?!\d)", stem)
    if m:
        out["po_no"] = m.group(1)
    m = re.search(r"(\d{4}-\d{2}-\d{2})", stem)
    if m:
        out["invoice_date"] = _date(m.group(1))
    m = re.search(r"\(([^)]+)\)\s*$", stem)
    if m:
        inside = m.group(1).strip()
        # The parentheses might read "Agilent, item 1 out of 3" (supplier first)
        # or "Tax Invoice No. MCIN060317, AsiaMedEnviro" (supplier last), so this
        # has to be judged by content, not position.
        parts = [x.strip() for x in inside.split(",") if x.strip()]
        cands = [x for x in parts if not _is_note(x)]
        out["supplier"] = (cands[-1] if cands else (parts[0] if parts else inside))
        mi = re.search(r"Invoice\s*(?:No\.?|Number|#)\s*([A-Za-z0-9][\w\-/]{3,})", inside, re.I)
        if mi and valid_invoice_no(mi.group(1)):
            out["invoice_no_hint"] = mi.group(1).strip()
    # Filenames that end like "Tax Invoice Final 7010416250"
    if "invoice_no_hint" not in out:
        mt = re.search(r"Invoice\s+(?:Final\s+)?([A-Z0-9][\w\-]{5,})\s*$", stem, re.I)
        if mt and valid_invoice_no(mt.group(1)):
            out["invoice_no_hint"] = mt.group(1).strip()
    out["is_proforma"] = bool(re.search(r"proforma|pro\s*forma", stem, re.I))
    return out


def _label_below(layout_text, label_pat, max_lines=3):
    """A layout where the label is on one line and the value is directly below
    it in the same column (common in Xero/Zoho-style templates). Other fields
    (an address, etc.) often sit to the right on the same row, so this has to
    align by column position rather than just grabbing the whole line."""
    lines = layout_text.splitlines()
    for idx, line in enumerate(lines):
        m = re.search(label_pat, line, re.I)
        if not m:
            continue
        col = m.start()
        for nxt in lines[idx + 1: idx + 1 + max_lines]:
            for tok in re.finditer(r"\S+", nxt):
                if abs(tok.start() - col) <= 3:
                    return tok.group(0).strip(" :,")
    return None


def _table_kv(tables):
    """Flatten a "header row / value row" horizontal table into {label: value}.
    GenScript's PO NO / DUE DATE / TERMS block is this kind of structure — the line-based regexes can't find it."""
    kv = {}
    pairs = []
    for tbl in tables:
        rows = [r for r in tbl if r]
        for a in range(0, len(rows) - 1, 2):      # Pair up rows (0,1) (2,3) ...
            pairs.append((rows[a], rows[a + 1]))
    for head, val in pairs:
        if len(head) != len(val) or len(head) < 2:
            continue
        for h, v in zip(head, val):
            if h and v:
                key = re.sub(r"\s+", " ", str(h)).lower()
                key = re.sub(r"[^a-z ]", "", key).strip()
                kv.setdefault(key, str(v).strip())
    return kv


def _detect_currency(*texts):
    """Currency must be judged from the totals line. The body often lists both
    SGD and USD bank account details, so matching anywhere in the text would
    misjudge an SGD invoice as USD."""
    ctx = [
        r"(?:Total|Balance|Grand\s*Total|Amount\s*Due)\s*(SGD|USD|EUR|GBP|JPY|RMB|CNY)\b",
        r"(?:Total|Balance|Amount)[^\n]{0,30}?\b(SGD|USD|EUR|GBP|JPY|RMB|CNY)\b[^\n]{0,12}[\d,]+\.\d{2}",
        r"Amount\s*\((SGD|USD|EUR|GBP|JPY)\)",
        r"\b(SGD|USD|EUR|GBP|JPY)\b[^\n]{0,10}[\d,]+\.\d{2}\s*$",
    ]
    for text in texts:
        if not text:
            continue
        for pat in ctx:
            m = re.search(pat, text, re.I | re.M)
            if m:
                return m.group(1).upper()
    for text in texts:                      # Fallback: currency symbols, USD checked first (US$ contains S$)
        if not text:
            continue
        for pat, code in CURRENCY_HINTS:
            if re.search(pat, text):
                return code
    return None


# Payment-terms paragraphs are full of company names with "Pte Ltd" (payee
# account name, bank name), but that's not the header. In practice this
# mis-extracted things like "1. Cheque crossed and made payable to...",
# "Banker : United Overseas Bank", "All cheque shall be made payable to...".
SUPPLIER_NOISE = re.compile(
    r"(cheque|payable|remit|bank|banker|account|swift|paynow|giro|beneficiary"
    r"|payment|please|kindly|refer to|made out|branch\s*code|reg(?:istration)?\s*no"
    r"|property of|finance charge|terms|www\.|@"
    r"|defect|liability|sold shall|limited to the|overdue|crossed)", re.I)

# The header is often split across two lines: "Agilent Technologies" /
# "Singapore (Sales) Pte Ltd. 199904761K". The suffixed line is the second
# one, and taken alone it would produce "Singapore (Sales) Pte Ltd." with no
# brand name at all.
_SUFFIX_ONLY = re.compile(
    r"^\s*(?:[A-Z][a-z]+\s+)?\(?(?:Singapore|Asia|SG)?\)?\s*"
    r"\(?(?:Sales|Trading|Distribution)?\)?\s*(?:Pte|Sdn|Co)\b", re.I)


def _guess_supplier(text):
    """The header company name. Takes the first line in the first several lines
    of the body that has a company suffix, and trims off any label block to the
    right on the same line (like 'Invoice NO:97809367')."""
    lines = text.splitlines()[:60]
    for idx, line in enumerate(lines):
        if SUPPLIER_NOISE.search(line):
            continue
        line = re.split(r"\s{2,}|\s(?=(?:Invoice|Order|Customer|Quote|GST)\s*(?:NO|No|Date))",
                        line.strip())[0].strip()
        if any(b in line.upper() for b in BUYER_HINTS):
            continue                                   # The buyer — skip
        if re.search(r"\b(PTE|LTD|LIMITED|INC|LLC|GMBH|CORP|BHD|SDN)\b\.?", line, re.I) \
                or re.search(r"(Pte\.?\s*Ltd|Co\.,\s*Ltd)", line, re.I):
            # If all that's left is region + suffix (missing the brand name), pull the previous line's brand name back in
            if _SUFFIX_ONLY.match(line) and idx > 0:
                prev = lines[idx - 1].strip()
                if (prev and not SUPPLIER_NOISE.search(prev)
                        and re.search(r"[A-Za-z]{3}", prev)
                        and len(prev) < 60
                        and not any(b in prev.upper() for b in BUYER_HINTS)):
                    return f"{prev} {line}".strip()
            return line
    return None


# Each vendor's own wording for "excl. tax / tax / incl. tax", ordered by reliability
TOTAL_PATTERNS = {
    "amount_excl": [
        r"Total\s*Amount\s*payable\s*excluding\s*GST[^\d\-]{0,12}([\d,]+\.\d{2})",
        r"Net\s*(?:Total|Amount)[^\d\-]{0,14}([\d,]+\.\d{2})",
        r"Subtotal[^\n]{0,60}?([\d,]+\.\d{2})\s*$",
        r"Subtotal[^\d\-]{0,20}([\d,]+\.\d{2})",
        r"^\s*Total\s*[:：][^\d\-]{0,12}([\d,]+\.\d{2})",
        # Edge case (Roylab, a common label in Odoo-generated invoice templates):
        # "Untaxed Amount" is the excl.-tax subtotal. None of the labels above
        # (Net Amount / Subtotal) cover it, so it was never recognized before —
        # leaving the excl.-tax amount permanently None, which triggered the
        # zero-tax-rate fallback further down ("excl. = incl., GST = 0") and
        # misclassified this invoice as zero-tax when it was actually 9% GST.
        r"Untaxed\s*Amount[^\d\-]{0,14}([\d,]+\.\d{2})",
    ],
    "gst": [
        r"Add\s*GST\s*\(?\s*[\d.]+\s*\)?\s*%[^\d\-]{0,14}([\d,]+\.\d{2})",
        r"Add\s*[\d.]+\s*%?\s*GST[^\d\-]{0,12}([\d,]+\.\d{2})",
        r"Tax\s*Amount\s*(?:[\d.]+\s*%)?[^\d\-]{0,12}([\d,]+\.\d{2})",
        r"Total\s*GST\s*(?:Amount)?\s*(?:[\d.]+\s*%)?[^\d\-]{0,12}([\d,]+\.\d{2})",
        # Edge case (Roylab): the tax line reads "TAX 9% 47.16", without the word
        # "GST" at all, so none of the rules above match. There are also lines
        # like "Sales Tax S$ 247.50" among the line items, but that's a per-line
        # tax amount with no percentage right after it, so it's not caught by
        # this rule — this one specifically requires a percentage number right
        # after "TAX" ("TAX 9%"), which is the fixed wording of the bottom
        # summary line.
        r"\bTAX\s*[\d.]+\s*%[^\d\-]{0,14}([\d,]+\.\d{2})",
        # Edge case (found while debugging today's 5 invoices — matches the
        # "unusual tax rate: 91.7%" needs-review record from Lonza): this is the
        # loosest rule in the whole table, requiring only that the word "GST"
        # appear somewhere. Lonza prints both "TOTAL EXCL. GST 1,176.00" and
        # "GST 9% 105.84", with the former appearing first — without an
        # exclusion, this rule would grab "EXCL. GST 1,176.00" first, treating
        # the excl.-tax subtotal as the tax amount; since the incl.-tax total
        # was fixed, the equation then backed out an excl.-tax amount of
        # 105.84 — a tax rate that displays as 91.7%, exactly the needs-review
        # record seen in the ledger. After adding the exclusion, the amount
        # equation check fills in GST correctly instead.
        r"(?<!EXCL\. )(?<!EXCL )(?<!EXCLUDING )"
        r"GST\s*(?:@|Amount)?\s*(?:\(?\s*[\d.]+\s*\)?\s*%)?[^\d\-]{0,12}([\d,]+\.\d{2})",
    ],
    "amount_incl": [
        # Order is priority. The rules at the top are unambiguous "amount payable" labels.
        r"Balance\s*Payable[^\d\-]{0,20}([\d,]+\.\d{2})",
        r"Amount\s*Due[^\d\-]{0,12}([\d,]+\.\d{2})",
        r"Grand\s*Total[^\d\-]{0,12}([\d,]+\.\d{2})",
        r"TOTAL\s*AMOUNT\s*(?:in\s*)?(?:[A-Z]{3})?[^\d\-]{0,12}([\d,]+\.\d{2})",
        r"SAY\s*TOTAL\s*[:：]?[^\d\-]{0,16}([\d,]+\.\d{2})",
        r"Balance\s*\([A-Z]{3}\)[^\d\-]{0,12}([\d,]+\.\d{2})",
        # Edge case: the 'Sub' exclusion can't be dropped. "SubTotal SGD 416.00"
        # contains "Total SGD 416.00" — without (?<!Sub) this treats the
        # excl.-tax subtotal as the incl.-tax total (hit Linde/Aik Moh both).
        # Same idea for 'Untaxed' (Roylab).
        # The currency may or may not be in parentheses: "TOTAL (SGD) 305.20", "TOTAL SGD 453.44", "TOTAL: SGD 1,234.00"
        r"(?<!Sub)(?<!SUB)(?<!Sub )(?<!SUB )(?<!Untaxed )\bTOTAL\b[^\d\n]{0,10}?\(?\s*(?:SGD|USD|EUR|GBP|JPY)\s*\)?[^\d\-]{0,14}([\d,]+\.\d{2})",
        # Edge case: in some native-text PDFs like Acoerela's, the text
        # pdfplumber extracts has no space between words (not a scan — this
        # particular PDF's character spacing is too tight for pdfplumber's
        # word-break heuristic to detect a space), so "TOTAL SGD" reads as the
        # run-together "TOTALSGD". \bTOTAL\b requires a word boundary right
        # after TOTAL, but the following S is also a word character, so there's
        # no boundary there and every \bTOTAL\b-anchored rule above fails —
        # this invoice's total wasn't extracted at all before. Added a rule
        # specifically for "TOTAL glued directly to a currency code", without
        # touching the global word-break threshold (too risky — it would affect every other field's extraction too).
        r"\bTOTAL(?:SGD|USD|EUR|GBP|JPY)\b[^\d\-]{0,14}([\d,]+\.\d{2})",
        # Edge case: Lonza prints both "TOTAL EXCL. GST 1,176.00" and "TOTAL
        # 1,281.84" — a generic ^TOTAL rule would hit the former first. Any
        # label containing EXCL/BEFORE is never the incl.-tax amount.
        # Edge case: Genomax prints "Total Discount 0.00" — a generic ^TOTAL
        # rule would grab 0.00. Labels containing DISCOUNT/EXCL/BEFORE/PAID/UNITS are never the amount payable.
        # Edge case (Roylab): this line-anchored rule should have matched
        # "Total S$ 571.06" just fine — the currency symbol S$ falls within the
        # allowed [^\d\n\-]{0,40} character set. What actually blocked it was
        # the trailing \s*$: OCR read one extra stray period after the amount
        # ("571.06 ."), and a period isn't whitespace, so \s*$ never matched
        # the end of line and the whole rule failed — none of Roylab's three
        # line items and no invoice total were extracted at all. Loosened the
        # end-of-line condition to allow a short trailing run of periods/dashes (this kind of OCR noise) after the amount.
        r"^\s*(?<!Sub)(?<!Sub )TOTAL\b(?![^\d\n]{0,20}(?:EXCL|BEFORE|EXCLUDING|DISCOUNT|PAID|UNITS|QTY))[^\d\n\-]{0,40}([\d,]+\.\d{2})[\s.\-]{0,5}$",
        # Fallback: OCR sometimes reads two side-by-side columns as a single
        # line, so what follows "Total" isn't a number but text from the other
        # column, and by the time the real amount shows up it's no longer at
        # the start of the line and has no currency before it. In practice a
        # scanned Vazyme invoice read as "...Singapore Branch Total 279.04"
        # (the bank-details column and the amount column got stitched
        # together). Placed last, reusing the same exclusion words as above to
        # avoid reintroducing the Lonza/Genomax edge cases; doesn't require
        # start-of-line or a currency, just a number immediately after Total.
        # The (?<!Sub-) exclusion here was found while testing the Sigma sample:
        # Sigma prints "Items Sub-Total 551.70", where Sub and Total are joined
        # by a hyphen rather than nothing or a space — the original exclusion
        # only guarded against "Subtotal"/"Sub Total", and "Sub-Total" slips
        # through the \bTotal\b word boundary (a hyphen is a non-word character,
        # so there's still a boundary right before Total), which mistook a
        # line-item subtotal for the invoice total.
        r"(?<!Sub)(?<!SUB)(?<!Sub )(?<!SUB )(?<!Sub-)(?<!SUB-)(?<!Untaxed )\bTotal\b"
        r"(?![^\d\n]{0,20}(?:EXCL|BEFORE|EXCLUDING|DISCOUNT|PAID|UNITS|QTY))"
        r"[ \t:]{0,10}([\d,]+\.\d{2})(?!\s*%)",
    ],
}


def _totals_from_text(*texts):
    """When the line-item table can't be extracted, find the invoice total
    directly from the body text. Reconciliation only needs the invoice number
    plus the total — line items only affect how detailed the output is."""
    out = {}
    for field, pats in TOTAL_PATTERNS.items():
        for text in texts:
            if not text:
                continue
            for pat in pats:
                m = re.search(pat, text, re.I | re.M)
                if m:
                    out[field] = _money(m.group(1))
                    break
            if out.get(field) is not None:
                break
    # Zero-rated / overseas invoice: no GST line, excl. and incl. amounts are equal
    if out.get("gst") is None and out.get("amount_excl") is not None \
            and out.get("amount_incl") is not None \
            and abs(out["amount_excl"] - out["amount_incl"]) <= 0.01:
        out["gst"] = 0.0
    if out.get("amount_excl") is None and out.get("amount_incl") is not None \
            and out.get("gst") is None:
        out["amount_excl"] = out["amount_incl"]
        out["gst"] = 0.0

    # excl. + tax = incl. If this doesn't hold, one of the numbers was mis-extracted (often the tax rate got grabbed instead)
    e, g, i = out.get("amount_excl"), out.get("gst"), out.get("amount_incl")
    if None not in (e, g, i) and abs(e + g - i) > 0.05:
        out["gst"] = round(i - e, 2)     # excl./incl. are usually printed unambiguously; the tax amount is the one most likely mis-captured

    # A total of zero or negative always means the wrong line was captured
    # (usually a label like "Total Discount 0.00"). Better to mark it as "not
    # extracted" and let the later fallback handle it than let a fake number
    # into reconciliation.
    for k in ("amount_incl", "amount_excl"):
        v = out.get(k)
        if v is not None and v <= 0:
            out[k] = None
            out.pop("gst", None)

    # If one of the three is missing, it can be derived from the other two
    e, g, i = out.get("amount_excl"), out.get("gst"), out.get("amount_incl")
    if i is None and e is not None and g is not None:
        out["amount_incl"] = round(e + g, 2)
    if e is None and i is not None and g is not None:
        out["amount_excl"] = round(i - g, 2)
    if g is None and i is not None and e is not None:
        out["gst"] = round(i - e, 2)
    return out


MIN_CHARS_PER_PAGE = 40


def _needs_ocr(pages_text):
    """Decide whether OCR is needed.

    Edge case: can't use `not full_text.strip()`. A scanned page is often read
    by pdfplumber as 1-3 stray characters (page borders or scan noise
    misread as glyphs), which is non-empty after strip(), so the OCR branch
    gets skipped and every field on the whole invoice ends up None with no
    indication why. In practice two Roylab pages read as 1 and 3 characters
    respectively — judging by "average characters per page" is the reliable way.
    """
    if not pages_text:
        return True
    total = sum(len((t or "").strip()) for t in pages_text)
    return total < MIN_CHARS_PER_PAGE * len(pages_text)


def _from_ai(data):
    """Map ai_extract's schema onto this module's field structure. Pure field
    relabeling — no judgment calls here; amount validation and matching are
    still handled by deterministic code."""
    if not data or data.get("_error"):
        return None
    header = {
        "invoice_no": data.get("invoice_no"),
        "invoice_date": _date(data.get("invoice_date")),
        "due_date": _date(data.get("due_date")),
        "po_no": data.get("po_no"),
        "supplier": clean_supplier(data.get("supplier")),
        "currency": data.get("currency"),
        "is_proforma": bool(data.get("is_proforma")),
        "terms": None,
    }
    subtotal = {
        "amount_excl": data.get("amount_excl_gst"),
        "gst": data.get("gst"),
        "amount_incl": data.get("amount_incl_gst"),
    }
    items = []
    for it in (data.get("line_items") or []):
        if not isinstance(it, dict) or not it.get("description"):
            continue
        items.append({
            "description": str(it["description"]).strip(),
            "qty": it.get("qty"),
            "unit_price": it.get("unit_price"),
            "amount_excl": it.get("amount_excl_gst"),
            "gst": None,
            "amount_incl": it.get("amount_incl_gst"),
        })
    return header, items, subtotal


def _find_item_table(page):
    """Pick out the line-item table: the one whose header row has both a "quantity"-type word and a "description"-type word"""
    for tbl in page.extract_tables():
        if len(tbl) < 2 or not tbl[0]:
            continue
        heads = [re.sub(r"\s+", " ", (c or "")).strip().lower() for c in tbl[0]]
        joined = " | ".join(heads)
        if any(q in joined for q in QTY_HEADS) and any(d in joined for d in DESC_HEADS):
            return heads, tbl[1:]
    return None, None


def _col(heads, keywords):
    for i, h in enumerate(heads):
        if any(k in h for k in keywords):
            return i
    return None


def extract_invoice(path):
    """Returns {header fields..., 'line_items': [...], 'warnings': [...]}"""
    warnings = []
    with pdfplumber.open(path) as pdf:
        pages_text = [p.extract_text() or "" for p in pdf.pages]
        full_text = "\n".join(pages_text)
        # Layout-mode text preserves column spacing — many invoices have labels/values in the same row but different columns, which normal mode would run together
        layout_text = "\n".join((p.extract_text(layout=True) or "") for p in pdf.pages)
        all_tables, heads, rows = [], None, None
        for p in pdf.pages:
            all_tables.extend(p.extract_tables())
            if rows is None:
                h, r = _find_item_table(p)
                if r:
                    heads, rows = h, r
    kv = _table_kv(all_tables)

    ocr_used = False
    is_scanned = _needs_ocr(pages_text)
    if is_scanned:
        try:
            import ocr
        except ImportError:
            ocr = None
        if ocr and ocr.available():
            o_plain, o_layout = ocr.ocr_pdf(path)
            if o_plain.strip():
                # Tesseract sometimes reads a printed hyphen "-" as a tilde "~"
                # in some italic/condensed fonts — commonly seen with
                # "INV-26-03285" being read as "INV-26~03285". The invoice-number
                # regex charset only allows \w - /, so it truncates at the tilde,
                # and the result looks fine even though it's cut short
                # (valid_invoice_no still passes, since the truncated "INV-26" is
                # itself a valid-looking shape). A tilde essentially never
                # appears in real invoice text, so one sandwiched between
                # digits/letters is treated as a misread hyphen — normalizing
                # this at the text level before regex matching is less likely to
                # cause side effects than loosening every individual regex's charset.
                o_plain = re.sub(r"(?<=[0-9A-Za-z])~(?=[0-9A-Za-z])", "-", o_plain)
                o_layout = re.sub(r"(?<=[0-9A-Za-z])~(?=[0-9A-Za-z])", "-", o_layout)
                full_text, layout_text = o_plain, o_layout
                ocr_used = True
                warnings.append("Scanned document, content read via OCR — figures may be inaccurate, please review manually")
            else:
                warnings.append("Scanned document, OCR could not recognize any text")
        else:
            warnings.append("PDF has no readable text (scanned) and OCR is not installed on this machine")

    header = {}
    for field, patterns in LABEL_PATTERNS.items():
        val = None
        if field == "invoice_date":          # Label-above-value layout — try column alignment first
            c = _label_below(layout_text, r"Invoice\s*Date\b")
            if c and _date(c):
                val = c
        for text in (full_text, layout_text):        # Plain text first, layout text as fallback
            for pat in patterns:
                for m in re.finditer(pat, text, re.I):
                    cand = m.group(1).strip()
                    if field != "invoice_no" or valid_invoice_no(cand):
                        val = cand
                        break
                if val:
                    break
            if val:
                break
        if field in ("invoice_date", "invoice_no") and not val:
            pass
        if not val:
            below = {
                "invoice_no": r"Invoice\s*(?:Number|No\.?|NO)\b",
                "invoice_date": r"Invoice\s*Date\b",
                "po_no": r"(?:Reference|Purchase\s*Order|Your\s*Ref)\b",
                "due_date": r"Due\s*Date\b",
            }.get(field)
            if below:
                cand = _label_below(layout_text, below)
                if cand and (field != "invoice_no" or valid_invoice_no(cand)):
                    val = cand
        header[field] = val

    # Table key-value pairs fill in fields the regexes missed (label and value on different lines)
    KV_MAP = {"invoice_no": ["invoice no", "invoice number"],
              "invoice_date": ["invoice date"],
              "po_no": ["po no", "po number", "purchase order"],
              "due_date": ["due date", "payment due"]}
    for field, keys in KV_MAP.items():
        if header.get(field) in (None, ""):
            for k in keys:
                if kv.get(k):
                    header[field] = kv[k]
                    break

    header["invoice_date"] = _date(header["invoice_date"])
    header["due_date"] = _date(header["due_date"])
    header["terms"] = kv.get("terms")

    header["currency"] = _detect_currency(full_text, layout_text)

    # Supplier: the first line of the body is usually the header; the name in the filename's parentheses is a fallback and cross-check
    header["supplier"] = clean_supplier(_guess_supplier(full_text))

    fn = parse_filename(path)
    header["is_proforma"] = fn.get("is_proforma", False)
    header["supplier_from_filename"] = fn.get("supplier")

    if header.get("po_no") in (None, "") and fn.get("po_no"):
        header["po_no"] = fn["po_no"]
    elif fn.get("po_no") and str(header.get("po_no")) != str(fn["po_no"]):
        # If the PO extracted from the body isn't in PONCG format but the
        # filename's is, the body one is very likely not actually
        # Neocytogen's own PO — just some other number in the document that
        # happens to also be labeled "Reference No. / PO No". In practice,
        # UPS's Import Tax Invoice has no real PO field in the body; the
        # generic po_no fallback regex grabbed the freight table's
        # "Reference No." (the carrier's own waybill reference,
        # 5282293306) instead. PONCG format is already treated in this code
        # as the most reliable signal — when the filename has that format and
        # the body doesn't, trust the filename.
        body_po = str(header.get("po_no") or "")
        fn_po = str(fn["po_no"])
        if not re.fullmatch(r"PONCG\d{9,}", body_po, re.I) and re.fullmatch(r"PONCG\d{9,}", fn_po, re.I):
            warnings.append(f"Body PO number {header['po_no']} is not in PONCG format and is likely the wrong field — switched to {fn_po} from the filename")
            header["po_no"] = fn_po
        else:
            warnings.append(f"PO number mismatch: body {header['po_no']} vs filename {fn['po_no']}")

    if not valid_invoice_no(header.get("invoice_no")) and fn.get("invoice_no_hint"):
        header["invoice_no"] = fn["invoice_no_hint"]
        warnings.append("Invoice number taken from the filename")

    # Prefer the short name in the filename's parentheses for the supplier —
    # the ledger's "Company Ordered From" is also a short name, while the body
    # header is often the full legal name mixed with an address/tax ID, which is harder to match against.
    if fn.get("supplier"):
        if header.get("supplier") and normalize_cmp(header["supplier"]) != normalize_cmp(fn["supplier"]):
            header["supplier_from_body"] = header["supplier"]
        header["supplier"] = fn["supplier"]

    # The date only falls back to the filename when it can't be extracted from
    # the body, and always with a warning — filename dates sometimes have typos
    # (e.g. PONCG202511160 written as 2025-01-03 when it's actually 2026).
    if header.get("invoice_date") is None and fn.get("invoice_date"):
        header["invoice_date"] = fn["invoice_date"]
        warnings.append("Invoice date taken from the filename — filename dates often have typos, please confirm manually")

    if header.get("is_proforma"):
        warnings.append("This is a Proforma invoice, not a basis for payment")

    # ---- line items ----
    items, subtotal = [], {}
    if rows:
        i_desc = _col(heads, DESC_HEADS)
        i_qty = _col(heads, QTY_HEADS)
        i_price = _col(heads, ["unit price", "unit cost", "rate"])
        i_excl = _col(heads, ["excluding gst", "extended", "amount excl", "net"])
        i_gst = _col(heads, ["gst rate", "gst amount", "add gst", "tax"])
        i_incl = _col(heads, ["including gst", "total amount", "amount incl"])

        for r in rows:
            cell0 = re.sub(r"\s+", " ", str(r[0] or "")).strip().lower()
            desc = re.sub(r"\s+", " ", str(r[i_desc] or "")).strip() if i_desc is not None else ""
            # A totals row: the first cell says subtotal/total and there's no description
            if re.search(r"subtotal|^total\b|grand total", cell0) and not desc:
                subtotal = {
                    "amount_excl": _money(r[i_excl]) if i_excl is not None else None,
                    "gst": _money(r[i_gst]) if i_gst is not None else None,
                    "amount_incl": _money(r[i_incl]) if i_incl is not None else None,
                }
                continue
            if not desc:
                continue
            items.append({
                "description": desc,
                "qty": _money(r[i_qty]) if i_qty is not None else None,
                "unit_price": _money(r[i_price]) if i_price is not None else None,
                "amount_excl": _money(r[i_excl]) if i_excl is not None else None,
                "gst": _money(r[i_gst]) if i_gst is not None else None,
                "amount_incl": _money(r[i_incl]) if i_incl is not None else None,
            })
    else:
        warnings.append("No line-item table found")

    if not subtotal or subtotal.get("amount_incl") is None:
        fallback = _totals_from_text(full_text, layout_text)
        if fallback.get("amount_incl") is not None:
            subtotal = {**fallback, **{k: v for k, v in (subtotal or {}).items() if v is not None}}
            if not items:
                warnings.append("No line items extracted — reconciled using the invoice total instead")

    # One more fallback: TOTAL_PATTERNS all require the label and the number on
    # the same line, but "label above, value directly below in the same
    # column" is another common invoice layout (especially visible on
    # system-generated invoices like Sigma's: one line prints "... TOTAL
    # CURRENCY", with the value on the next line in the same column, "...
    # 601.35 SGD"). Reuses _label_below's column-alignment logic to find the
    # number directly under the TOTAL label.
    if (subtotal or {}).get("amount_incl") is None:
        tok = _label_below(layout_text, r"\bTOTAL\b")
        val = _money(tok)
        if val is not None and val > 0:
            subtotal = dict(subtotal or {})
            subtotal["amount_incl"] = val
            warnings.append("Invoice total taken from a \"value directly below the TOTAL label\" layout, please confirm manually")

    # Validation: line-item sum vs invoice total (line-level GST rounding varies by vendor, so this is only gated at the total level, tolerance 0.05)
    priced = [i for i in items if i.get("amount_incl") is not None]
    if items and not priced:
        warnings.append("Line items had no extractable amount — reconciled using the invoice total instead")
        items = []
    elif priced and len(priced) == len(items) and subtotal.get("amount_incl") is not None:
        s = round(sum(i["amount_incl"] for i in priced), 2)
        if abs(s - subtotal["amount_incl"]) > 0.05:
            warnings.append(f"Line-item sum {s} doesn't match the invoice total {subtotal['amount_incl']}")

    # OCR commonly confuses O/0, l/1, S/5, B/8 in invoice numbers. Only tries a
    # correction when validation fails, and only touches the invoice number
    # field — a blanket text replace would turn the O in a company name into a 0.
    if ocr_used and not valid_invoice_no(header.get("invoice_no")) and header.get("invoice_no"):
        try:
            import ocr as _ocr_mod
            fixed = _ocr_mod.fix_ocr_digits(header["invoice_no"])
        except Exception:
            fixed = None
        if fixed and valid_invoice_no(fixed):
            header["invoice_no"] = fixed
            warnings.append(f"Invoice number corrected via OCR character correction (originally read as {header['invoice_no']}), please confirm manually")

    # ---- vision-model fallback ----
    # Triggers when: the document is scanned, or there's text but either the
    # invoice number or the total couldn't be extracted. The original code
    # only went to OCR when full_text was empty and never called ai_extract at
    # all — so scanned invoices failed across the board. Here the model only
    # "reads" — whatever it reads still goes through the same amount-equation
    # validation.
    need_ai = (is_scanned
               or not valid_invoice_no(header.get("invoice_no"))
               or subtotal.get("amount_incl") is None)
    if need_ai:
        try:
            import ai_extract
        except ImportError:
            ai_extract = None
        if ai_extract and ai_extract.available():
            data = ai_extract.extract(path, filename_hint=os.path.basename(path))
            if data and data.get("_error"):
                warnings.append(f"Vision model call failed: {data['_error']}")
            else:
                mapped = _from_ai(data)
                if mapped:
                    ai_header, ai_items, ai_sub = mapped
                    used = []
                    for k, v in ai_header.items():
                        if v not in (None, "") and header.get(k) in (None, "", False):
                            header[k] = v
                            used.append(k)
                    # Invoice number: what the model read takes priority over a "taken from filename" guess
                    if valid_invoice_no(ai_header.get("invoice_no")):
                        if normalize_cmp(ai_header["invoice_no"]) != normalize_cmp(header.get("invoice_no")):
                            if not valid_invoice_no(header.get("invoice_no")):
                                header["invoice_no"] = ai_header["invoice_no"]
                                used.append("invoice_no")
                            else:
                                warnings.append(
                                    f"Invoice number uncertain: regex read {header['invoice_no']}, "
                                    f"model read {ai_header['invoice_no']}")
                    # If the date was a "taken from filename" guess, prefer the model's actual printed date
                    if (ai_header.get("invoice_date")
                            and any("taken from the filename" in w for w in warnings)
                            and ai_header["invoice_date"] != header.get("invoice_date")):
                        header["invoice_date"] = ai_header["invoice_date"]
                        warnings = [w for w in warnings if "taken from the filename" not in w]
                        used.append("invoice_date")
                    if subtotal.get("amount_incl") is None and ai_sub.get("amount_incl") is not None:
                        subtotal = ai_sub
                        used.append("invoice total")
                    if not items and ai_items:
                        items = ai_items
                    if used:
                        warnings.append("The following fields were read by the vision model, please review manually: " + ", ".join(used))
        elif is_scanned:
            warnings.append("Scanned document and the vision model is unavailable (missing ANTHROPIC_API_KEY or anthropic isn't installed)")

    # Amounts read back from the model/OCR go through the same equation
    # validation too — no free pass just because the source is different.
    #
    # Which side gets corrected depends on the source:
    #   Native PDF — all three numbers are read deterministically, so the total
    #                is authoritative; back out the tax amount from it.
    #   OCR        — the other way around. The line items (63.20 / 5.69) have
    #                fewer digits and simpler glyphs, so they're less likely to
    #                be misread than the total; the total is often the
    #                largest, most stylized line on the page, which makes it
    #                the most error-prone. In practice a Lip Laundry receipt's
    #                total of 68.90 was misread as 68.50, while 63.20 + 5.69 =
    #                68.89, only a one-cent rounding difference from the
    #                ledger. Trusting the total unconditionally would have
    #                invented a 0.39 gap and reported it as a mismatch.
    e, g, i = subtotal.get("amount_excl"), subtotal.get("gst"), subtotal.get("amount_incl")
    if None not in (e, g, i) and abs(e + g - i) > 0.05:
        if ocr_used:
            subtotal["amount_incl"] = round(e + g, 2)
            warnings.append(
                f"OCR amount equation doesn't hold ({e} + {g} != {i}) — used the line-item sum {round(e + g, 2)} as the total instead"
                f" (originally read as {i}), please confirm manually")
        else:
            subtotal["gst"] = round(i - e, 2)
            warnings.append(f"Amount equation doesn't hold ({e} + {g} != {i}) — recalculated the tax amount as incl. minus excl.")

    if not valid_invoice_no(header.get("invoice_no")):
        header["invoice_no"] = None
        warnings.append("Invoice number not extracted — can't take part in matching")
    if not header.get("supplier"):
        warnings.append("Supplier not extracted")

    return {**header, "ocr": ocr_used, "line_items": items, "subtotal": subtotal,
            "source_file": os.path.basename(path), "warnings": warnings}


if __name__ == "__main__":
    import sys, json
    r = extract_invoice(sys.argv[1])
    print(json.dumps(r, indent=2, ensure_ascii=False, default=str))
