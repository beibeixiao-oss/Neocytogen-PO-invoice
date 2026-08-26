__version__ = "2026-08-21.14"

"""
ai_extract.py — 用视觉模型读发票（仅用于正则和 OCR 都拿不下的那批）

调用策略：原生 PDF 走正则，又快又准且不花钱；只有扫描件、或正则抽不到发票号的
才发给模型。60 张发票里通常只有不到 10 张真正发出去。

模型只负责「读」，不负责「判断」：它按固定 schema 吐字段，匹配、金额校验、
去重仍由确定性代码完成。财务对账不能让模型下结论。

Key 放同目录 .env：
    ANTHROPIC_API_KEY=sk-ant-xxxx
.env 不要提交、不要随脚本分发。
"""

import base64
import json
import os
import re

MODEL = "claude-sonnet-5"      # 想更省可换 claude-haiku-4-5-20251001
MAX_PAGES = 3                  # 发票正文通常在前几页，后面多是条款和汇款信息
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
    """读同目录 .env。用最朴素的解析，不引入额外依赖。"""
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
    """有 key 且 SDK 装好了才算可用。不可用时上层应跳过而非报错。"""
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
    """模型偶尔会裹 ``` 或加一句前言，这里做容错。"""
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
    """返回与 invoice_extractor 同构的字段字典；失败返回 None。"""
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
        return {"_error": "模型返回的不是合法 JSON"}

    # 模型可能把税率当税额，或三个金额对不上 —— 用等式兜一道
    e, g, i = (data.get("amount_excl_gst"), data.get("gst"), data.get("amount_incl_gst"))
    if None not in (e, g, i) and abs(e + g - i) > 0.05:
        data["gst"] = round(i - e, 2)
        data["_amount_fixed"] = True
    if g is None and None not in (e, i):
        data["gst"] = round(i - e, 2)

    return data
