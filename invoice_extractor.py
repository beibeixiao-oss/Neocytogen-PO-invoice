"""
invoice_extractor.py — 从 invoice PDF 抽取结构化字段

设计取舍：
  * 优先用 pdfplumber 的 table 抽 line item（原生文字 PDF 效果最好）
  * header 字段用「标签正则」，同一字段列多个别名，兼容不同供应商模板
  * 文件名本身是结构化的（PO - 日期 - 类型 - 供应商），作为交叉验证的第二信号
  * 抽不到的字段一律留 None，绝不猜
"""

import os
import re
from datetime import datetime

import pdfplumber

__version__ = "2026-09-11.5"

# 同一字段的多种标签写法，按顺序尝试
# 买方是 Neocytogen —— 任何抽成这个的供应商结果都是错的
BUYER_HINTS = ["NEOCYTOGEN"]

# 发票号绝不会是这些（标签或正文词被误抓）
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
        # 坑（Roylab）：这份发票的版式是文档标题本身就是标签——"Tax Invoice
        # INV/2026/0117"，标题后面直接跟发票号，中间没有 "No./Number/#" 这类词，
        # 上面几条都要求这样的词才会匹配，全部落空。不敢把「Tax Invoice」后面
        # 任何词都当发票号（很多发票光是标题后面跟的是公司名、地址），所以限定
        # 紧跟着的词必须长得像发票号：字母开头、中间带一个 - 或 /、后面还有数字，
        # 排除掉普通单词、地名这些误伤。
        r"\bTax\s+Invoice\s+([A-Za-z]{2,6}[\-/][\dA-Za-z\-/]{3,})\b",
    ],
    "invoice_date": [
        # Lonza：标签 "Our invoice" 与 "<号> dated <日期>" 被换行拆开，
        # 而正文更靠前处还有一行孤零零的 "dated 27-Jan-2026"（那是 Your order 的日期）。
        # 所以必须跨行锚定 "Our invoice"，并让它排在通用 Date 规则之前。
        r"Our\s*invoice\b[\s\S]{0,80}?dated\s*(\d{1,2}[-\s]*[A-Za-z]{3,9}[-\s]*\d{4})",
        r"Invoice\s*Date\s*[:.]?\s*([\d]{1,4}[./\-][\d]{1,2}[./\-][\d]{1,4})",
        r"(?<!Due)(?<!Due )\bDate[dD]?\s*[:.]\s*([\d]{1,4}[./\-][\d]{1,2}[./\-][\d]{1,4})",
        r"Invoice\s*Date[ \t]*[:.]?\s*\n?\s*(\d{1,2}\s*[A-Za-z]{3,9}\s*\d{4})",
        # "dated" 前面若是 Your order / Our order，那是订单日期不是发票日期。
        # Lonza 抬头就是 "Your order PONCG... dated 27-Jan-2026"，实际发票日是 05-Feb-2026。
        r"(?<!Due)(?<!Due )(?<!order )(?<!Order )\bDate[dD]?\s*[:.]?\s*(\d{1,2}\s*[A-Za-z]{3,9}\s*\d{4})",
        r"Date\s*of\s*Invoice\s*[:.]?\s*([\d]{2,4}[./\-][\d]{1,2}[./\-][\d]{1,4})",
    ],
    "po_no": [
        r"(?<![A-Za-z0-9])(PONCG\d{9,})(?!\d)",     # Neocytogen 自己的 PO 格式，最可靠
        r"P\.?O\.?\s*(?:NO|No|Number)\s*[:.]?\s*([\w\-/]+)",
        r"Your\s*Ref\.?\s*[:.]?\s*([\w\-/]+)",
        r"Purchase\s*Order\s*[:.]?\s*\n?\s*([\w\-/]+)",
    ],
    "due_date": [
        r"Due\s*Date\s*[:.]?\s*([\d]{2,4}[./\-][\d]{1,2}[./\-][\d]{1,4})",
    ],
}

CURRENCY_HINTS = [
    # USD 必须先判：US$ 里含 S$，否则美元发票会被误判成新币（金额差三成以上）
    (r"\bUSD\b|US\$", "USD"),
    (r"\bSGD\b|(?<![A-Za-z])S\$", "SGD"),
    (r"\bEUR\b|€", "EUR"),   (r"\bJPY\b|¥", "JPY"),
    (r"\bGBP\b|£", "GBP"),
]

QTY_HEADS = {"qty", "quantity", "units", "unit no", "no of units"}
DESC_HEADS = {"description", "item description", "item", "particulars", "product"}


def _money(raw):
    """'S$1,234.56' -> 1234.56 ；空/破折号 -> None"""
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
    if not re.search(r"[A-Za-z]", s):           # 含月份英文缩写的保留空格写法
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
    """发票号必须像发票号：含数字、长度够、不是被误抓的标签词。
    这一条就能拦掉 on / Date / GST / in / DO 这类垃圾值。"""
    if not v:
        return False
    s = str(v).strip()
    if s.upper() in INVOICE_NO_STOPWORDS:
        return False
    if not re.search(r"\d", s):          # 不含任何数字 -> 一定不是发票号
        return False
    if len(re.sub(r"\W", "", s)) < 4:
        return False
    if re.fullmatch(r"(19|20)\d{2}", s):  # 纯年份
        return False
    # 地址被误抓（114AJlnJurongKechil 这类）
    if re.search(r"\b(Jln|Jalan|Road|Rd|Street|St|Ave|Avenue|Blk|Block|Singapore|Lorong)\b", s, re.I):
        return False
    if re.search(r"(Jln|Jalan|Road|Street|Avenue|Singapore|Lorong)", s, re.I):
        return False
    return True


def clean_supplier(name):
    """去掉抬头里混进来的标签、税号、买方名。"""
    if not name:
        return None
    s = re.sub(r"\b(TAX\s+)?INVOICE\b", " ", str(name), flags=re.I)
    s = re.sub(r"\b(Bill|Ship|Sold)\s*To\s*:?", " ", s, flags=re.I)
    s = re.sub(r"\bGST\s*Reg(?:istration)?\s*(?:No|Number)?\s*[:.]?\s*[\w\-]+", " ", s, flags=re.I)
    s = re.sub(r"\b(UEN|Customer\s*Code|Invoice\s*Date)\b.*", " ", s, flags=re.I)
    s = re.sub(r"\s{2,}", " ", s).strip(" ,.:;-")
    if not s or any(b in s.upper() for b in BUYER_HINTS):
        return None                       # 抽成买方 = 抽错了，宁可留空
    if len(s) < 3 or not re.search(r"[A-Za-z]", s):
        return None
    # 只剩公司后缀（LTD / PTE LTD / INC）说明抬头没抓全
    if re.fullmatch(r"(PTE|LTD|LIMITED|INC|LLC|CO|CORP|BHD|SDN|GMBH|[.\s&,]+)+", s, re.I):
        return None
    return s


# 文件名括号里除了供应商，常夹着备注："item 1 out of 3"、"8 bottles MEM"、
# "Tax Invoice No. MCIN060317"。这些不是公司名。
NOTE_PAT = re.compile(
    r"(out\s*of|\bitem\b|\bpc?s\b|\bpieces?\b|\bbottles?\b|\bboxes?\b|\bvials?\b"
    r"|\bfinal\b|\btax\s*invoice\b|\bproforma\b|\binvoice\s*no\b|^no\.|^\d)", re.I)


def _is_note(s):
    s = s.strip()
    if not s or NOTE_PAT.search(s):
        return True
    letters = sum(c.isalpha() for c in s)
    return letters < 3 or letters < len(s) * 0.4        # 数字/符号占比过高


def parse_filename(path):
    """'PONCG202512185 - 2026-01-14 Tax Invoice (Genscript).pdf'
       -> {'po_no', 'invoice_date', 'supplier'}"""
    stem = os.path.splitext(os.path.basename(path))[0]
    out = {}
    # 不能用 \b：下划线是单词字符，"PONCG202601008_-_..." 在数字与 _ 之间
    # 没有词边界，\b 永远匹配不上，文件名 PO 回退因此长期失效。
    m = re.search(r"(?<![A-Za-z0-9])(PONCG\d{9,})(?!\d)", stem)
    if m:
        out["po_no"] = m.group(1)
    m = re.search(r"(\d{4}-\d{2}-\d{2})", stem)
    if m:
        out["invoice_date"] = _date(m.group(1))
    m = re.search(r"\(([^)]+)\)\s*$", stem)
    if m:
        inside = m.group(1).strip()
        # 括号里可能是 "Agilent, item 1 out of 3"（供应商在前）
        # 也可能是 "Tax Invoice No. MCIN060317, AsiaMedEnviro"（供应商在后）
        # 所以按内容判断，不能按位置取
        parts = [x.strip() for x in inside.split(",") if x.strip()]
        cands = [x for x in parts if not _is_note(x)]
        out["supplier"] = (cands[-1] if cands else (parts[0] if parts else inside))
        mi = re.search(r"Invoice\s*(?:No\.?|Number|#)\s*([A-Za-z0-9][\w\-/]{3,})", inside, re.I)
        if mi and valid_invoice_no(mi.group(1)):
            out["invoice_no_hint"] = mi.group(1).strip()
    # 文件名尾部形如 "Tax Invoice Final 7010416250"
    if "invoice_no_hint" not in out:
        mt = re.search(r"Invoice\s+(?:Final\s+)?([A-Z0-9][\w\-]{5,})\s*$", stem, re.I)
        if mt and valid_invoice_no(mt.group(1)):
            out["invoice_no_hint"] = mt.group(1).strip()
    out["is_proforma"] = bool(re.search(r"proforma|pro\s*forma", stem, re.I))
    return out


def _label_below(layout_text, label_pat, max_lines=3):
    """标签在上、值在下一行同一列的排版（Xero/Zoho 这类模板常见）。
    同一行右侧往往还有别的栏位（地址等），所以只能靠列位置对齐，不能靠取整行。"""
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
    """把「表头一行 / 值一行」的横向表格拍平成 {标签: 值}。
    GenScript 的 PO NO / DUE DATE / TERMS 就是这种结构，正则按行找不到。"""
    kv = {}
    pairs = []
    for tbl in tables:
        rows = [r for r in tbl if r]
        for a in range(0, len(rows) - 1, 2):      # (0,1) (2,3) ... 逐段配对
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
    """币种必须从合计那一行判断。正文里的银行账户信息常同时列出 SGD/USD 账号，
    按全文匹配会把新币发票误判成美元。"""
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
    for text in texts:                      # 兜底：符号，USD 先判（US$ 内含 S$）
        if not text:
            continue
        for pat, code in CURRENCY_HINTS:
            if re.search(pat, text):
                return code
    return None


# 付款条款段落里全是带 "Pte Ltd" 的公司名（收款户名、银行名），
# 但那不是抬头。实测误抽出 "1. Cheque crossed and made payable to..."、
# "Banker : United Overseas Bank"、"All cheque shall be made payable to..."。
SUPPLIER_NOISE = re.compile(
    r"(cheque|payable|remit|bank|banker|account|swift|paynow|giro|beneficiary"
    r"|payment|please|kindly|refer to|made out|branch\s*code|reg(?:istration)?\s*no"
    r"|property of|finance charge|terms|www\.|@"
    r"|defect|liability|sold shall|limited to the|overdue|crossed)", re.I)

# 抬头常被拆成两行："Agilent Technologies" / "Singapore (Sales) Pte Ltd. 199904761K"。
# 带后缀的是第二行，单看它会得到 "Singapore (Sales) Pte Ltd." 这种没有品牌名的结果。
_SUFFIX_ONLY = re.compile(
    r"^\s*(?:[A-Z][a-z]+\s+)?\(?(?:Singapore|Asia|SG)?\)?\s*"
    r"\(?(?:Sales|Trading|Distribution)?\)?\s*(?:Pte|Sdn|Co)\b", re.I)


def _guess_supplier(text):
    """抬头公司名。取正文前几行里第一个带公司后缀的，
    并砍掉同一行右侧的标签块（如 'Invoice NO:97809367'）。"""
    lines = text.splitlines()[:60]
    for idx, line in enumerate(lines):
        if SUPPLIER_NOISE.search(line):
            continue
        line = re.split(r"\s{2,}|\s(?=(?:Invoice|Order|Customer|Quote|GST)\s*(?:NO|No|Date))",
                        line.strip())[0].strip()
        if any(b in line.upper() for b in BUYER_HINTS):
            continue                                   # 买方，跳过
        if re.search(r"\b(PTE|LTD|LIMITED|INC|LLC|GMBH|CORP|BHD|SDN)\b\.?", line, re.I) \
                or re.search(r"(Pte\.?\s*Ltd|Co\.,\s*Ltd)", line, re.I):
            # 只剩地区+后缀（缺品牌名）时，把上一行的品牌名接回来
            if _SUFFIX_ONLY.match(line) and idx > 0:
                prev = lines[idx - 1].strip()
                if (prev and not SUPPLIER_NOISE.search(prev)
                        and re.search(r"[A-Za-z]{3}", prev)
                        and len(prev) < 60
                        and not any(b in prev.upper() for b in BUYER_HINTS)):
                    return f"{prev} {line}".strip()
            return line
    return None


# 各家对「税前 / 税额 / 税后」的叫法，按可靠度排序
TOTAL_PATTERNS = {
    "amount_excl": [
        r"Total\s*Amount\s*payable\s*excluding\s*GST[^\d\-]{0,12}([\d,]+\.\d{2})",
        r"Net\s*(?:Total|Amount)[^\d\-]{0,14}([\d,]+\.\d{2})",
        r"Subtotal[^\n]{0,60}?([\d,]+\.\d{2})\s*$",
        r"Subtotal[^\d\-]{0,20}([\d,]+\.\d{2})",
        r"^\s*Total\s*[:：][^\d\-]{0,12}([\d,]+\.\d{2})",
        # 坑（Roylab，Odoo 生成的发票模板常见叫法）："Untaxed Amount" 才是税前小计。
        # 上面几条标签（Net Amount / Subtotal）它一个都不用，之前完全没被认出来，
        # 于是税前金额永远是 None，被后面「税前=税后、GST=0」那条零税率兜底误判成
        # 零税发票——这张明明是 9% GST。
        r"Untaxed\s*Amount[^\d\-]{0,14}([\d,]+\.\d{2})",
    ],
    "gst": [
        r"Add\s*GST\s*\(?\s*[\d.]+\s*\)?\s*%[^\d\-]{0,14}([\d,]+\.\d{2})",
        r"Add\s*[\d.]+\s*%?\s*GST[^\d\-]{0,12}([\d,]+\.\d{2})",
        r"Tax\s*Amount\s*(?:[\d.]+\s*%)?[^\d\-]{0,12}([\d,]+\.\d{2})",
        r"Total\s*GST\s*(?:Amount)?\s*(?:[\d.]+\s*%)?[^\d\-]{0,12}([\d,]+\.\d{2})",
        # 坑（Roylab）：税额小计印的是 "TAX 9% 47.16"，不带 "GST" 字样，上面几条
        # 全部要求出现 GST 都对不上。行项目里倒是有一堆 "Sales Tax S$ 247.50"，
        # 但那是单行税额、后面没有紧跟百分比，不会被这条误伤——这条要求 "TAX"
        # 后面立刻跟一个百分号数字（"TAX 9%"），是发票底部汇总行的固定写法。
        r"\bTAX\s*[\d.]+\s*%[^\d\-]{0,14}([\d,]+\.\d{2})",
        # 坑（调试今天这 5 张时顺带发现，跟 Lonza 那份「税率异常：91.7%」的待核查
        # 记录对得上号）：这条是全表里最宽松的一条，只要求出现 "GST" 字样。
        # Lonza 印的是 "TOTAL EXCL. GST 1,176.00" 和 "GST 9% 105.84" 两行，前者排在
        # 前面，这条不加排除的话会先撞上 "EXCL. GST 1,176.00"，把税前小计当成税额，
        # 税后总额固定，于是等式反推把税前金额算成 105.84 —— 税率显示成 91.7%，
        # 正是台账里看到的那条「待核查」记录。加排除后交给等式校验去补 GST。
        r"(?<!EXCL\. )(?<!EXCL )(?<!EXCLUDING )"
        r"GST\s*(?:@|Amount)?\s*(?:\(?\s*[\d.]+\s*\)?\s*%)?[^\d\-]{0,12}([\d,]+\.\d{2})",
    ],
    "amount_incl": [
        # 顺序即优先级。放在最前的是明确无歧义的「应付总额」标签。
        r"Balance\s*Payable[^\d\-]{0,20}([\d,]+\.\d{2})",
        r"Amount\s*Due[^\d\-]{0,12}([\d,]+\.\d{2})",
        r"Grand\s*Total[^\d\-]{0,12}([\d,]+\.\d{2})",
        r"TOTAL\s*AMOUNT\s*(?:in\s*)?(?:[A-Z]{3})?[^\d\-]{0,12}([\d,]+\.\d{2})",
        r"SAY\s*TOTAL\s*[:：]?[^\d\-]{0,16}([\d,]+\.\d{2})",
        r"Balance\s*\([A-Z]{3}\)[^\d\-]{0,12}([\d,]+\.\d{2})",
        # 坑：'Sub' 不能漏。"SubTotal SGD 416.00" 里含 "Total SGD 416.00"，
        # 少了 (?<!Sub) 就会把税前小计当成税后总额 —— Linde/Aik moh 全中招。
        # 'Untaxed'（Roylab）同理。
        # 币种可能带括号或冒号："TOTAL (SGD) 305.20"、"TOTAL SGD 453.44"、"TOTAL: SGD 1,234.00"
        r"(?<!Sub)(?<!SUB)(?<!Sub )(?<!SUB )(?<!Untaxed )\bTOTAL\b[^\d\n]{0,10}?\(?\s*(?:SGD|USD|EUR|GBP|JPY)\s*\)?[^\d\-]{0,14}([\d,]+\.\d{2})",
        # 坑：Acoerela 这类原生 PDF 里 pdfplumber 抽出来的文字词间没有空格
        # （不是扫描件，是这份 PDF 本身字符间距太紧，pdfplumber 的分词阈值判断不出
        # 空格），"TOTAL SGD" 被读成粘在一起的 "TOTALSGD"。\bTOTAL\b 要求 TOTAL
        # 后面是词边界，但紧跟着的 S 也是单词字符，两者之间没有边界，上面所有
        # \bTOTAL\b 开头的规则全部失效，这张发票之前完全没抽到合计。单独加一条
        # 只认「TOTAL 紧贴着币种代码」这种粘连写法，不去改分词阈值（那样风险面太大，
        # 会牵动其他所有字段的抽取）。
        r"\bTOTAL(?:SGD|USD|EUR|GBP|JPY)\b[^\d\-]{0,14}([\d,]+\.\d{2})",
        # 坑：Lonza 印 "TOTAL EXCL. GST 1,176.00" 与 "TOTAL 1,281.84" 两行，
        # 通用 ^TOTAL 会先撞上前者。标签里出现 EXCL/BEFORE 的一律不是税后额。
        # 坑：Genomax 印 "Total Discount 0.00"，通用 ^TOTAL 会取到 0.00。
        # 标签里出现 DISCOUNT/EXCL/BEFORE/PAID/UNITS 的都不是应付总额。
        # 坑（Roylab）：这条行首锚定的规则本来就该认「Total S$ 571.06」——
        # 币种符号 S$ 落在 [^\d\n\-]{0,40} 允许的字符集里，真正拦住它的是结尾
        # 的 \s*$：OCR 把金额后面多认出一个孤立句点（"571.06 ."），句点不是空白，
        # \s*$ 卡在那个点上就是不收尾，整条规则失效，Roylab 三行明细全部落空、
        # 发票合计也没抽到。放宽收尾，容许金额后面跟一小段句点/短横这类 OCR 噪点。
        r"^\s*(?<!Sub)(?<!Sub )TOTAL\b(?![^\d\n]{0,20}(?:EXCL|BEFORE|EXCLUDING|DISCOUNT|PAID|UNITS|QTY))[^\d\n\-]{0,40}([\d,]+\.\d{2})[\s.\-]{0,5}$",
        # 兜底：OCR 把同一横向位置的两栏内容读成了一整行，"Total" 后面直接跟着的不是
        # 数字而是另一栏文字，等真正的金额出现时前面已经不是行首、后面也没有币种。
        # 实测 Vazyme 扫描件被读成 "...Singapore Branch Total 279.04"（银行信息那栏
        # 和金额那栏被拼在了一起）。放在最后一条，排除词沿用上面同一套，避免复发
        # Lonza/Genomax 那两个坑；也不要求行首/币种，只要求数字紧跟在 Total 后面。
        # (?<!Sub-) 这条是测试 Sigma 样张时才发现要补的：Sigma 印的是 "Items Sub-Total
        # 551.70"，Sub 和 Total 中间是连字符而不是无分隔或空格，原来只防了
        # "Subtotal"/"Sub Total" 两种写法，"Sub-Total" 会从 \bTotal\b 的词边界缝隙里钻
        # 过去（连字符是非单词字符，Total 前依然算词边界），把行项目小计当成了总额。
        r"(?<!Sub)(?<!SUB)(?<!Sub )(?<!SUB )(?<!Sub-)(?<!SUB-)(?<!Untaxed )\bTotal\b"
        r"(?![^\d\n]{0,20}(?:EXCL|BEFORE|EXCLUDING|DISCOUNT|PAID|UNITS|QTY))"
        r"[ \t:]{0,10}([\d,]+\.\d{2})(?!\s*%)",
    ],
}


def _totals_from_text(*texts):
    """明细表抽不到时，直接从正文找发票合计。
    对账只需要发票号 + 合计，明细仅影响输出的详细程度。"""
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
    # 零税率/境外发票：没有 GST 行，税前与税后相等
    if out.get("gst") is None and out.get("amount_excl") is not None \
            and out.get("amount_incl") is not None \
            and abs(out["amount_excl"] - out["amount_incl"]) <= 0.01:
        out["gst"] = 0.0
    if out.get("amount_excl") is None and out.get("amount_incl") is not None \
            and out.get("gst") is None:
        out["amount_excl"] = out["amount_incl"]
        out["gst"] = 0.0

    # 税前 + 税额 = 税后。不成立说明某个数抓错了（常见是抓到税率）
    e, g, i = out.get("amount_excl"), out.get("gst"), out.get("amount_incl")
    if None not in (e, g, i) and abs(e + g - i) > 0.05:
        out["gst"] = round(i - e, 2)     # 税前与税后一般印得明确，税额最容易抓错

    # 合计为 0 或负数一定是抓错了行（多为 "Total Discount 0.00" 这类标签）。
    # 宁可判为「未抽到」交由后续兜底，也不能让一个假数字进对账。
    for k in ("amount_incl", "amount_excl"):
        v = out.get(k)
        if v is not None and v <= 0:
            out[k] = None
            out.pop("gst", None)

    # 三者缺一可由另外两者补齐
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
    """判断是否该走 OCR。

    坑：不能用 `not full_text.strip()`。扫描件常会被 pdfplumber 读出 1~3 个
    杂字符（页边框线、扫描噪点被当成字形），strip() 后非空，OCR 分支就被跳过，
    结果整张发票所有字段都是 None 且毫无提示。
    实测 Roylab 两张分别读到 1 和 3 个字符 —— 按「每页平均字符数」判定才稳。
    """
    if not pages_text:
        return True
    total = sum(len((t or "").strip()) for t in pages_text)
    return total < MIN_CHARS_PER_PAGE * len(pages_text)


def _from_ai(data):
    """把 ai_extract 的 schema 映射成本模块的字段结构。
    只做字段搬运，不做任何判断 —— 金额校验、匹配仍由确定性代码负责。"""
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
    """挑出 line item 表：表头里同时出现「数量」和「描述」两类词的那张"""
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
    """返回 {header 字段..., 'line_items': [...], 'warnings': [...]}"""
    warnings = []
    with pdfplumber.open(path) as pdf:
        pages_text = [p.extract_text() or "" for p in pdf.pages]
        full_text = "\n".join(pages_text)
        # 排版式文本保留列间距，很多发票的标签/值在同一行不同列，普通模式会粘在一起
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
                # Tesseract 在部分斜体/紧凑字体下会把印刷体的连字符 "-" 认成波浪号 "~"，
                # 常见于 "INV-26-03285" 被读成 "INV-26~03285"。发票号正则的字符集只认
                # \w - /，遇到 ~ 就整个截断，抽出来的号码少了后半截还看不出异常
                # （valid_invoice_no 照样能过，因为截断后的 "INV-26" 本身形状合法）。
                # 波浪号在真实发票文本里基本不会出现，数字/字母之间夹一个就判定是连字符误读，
                # 在正则匹配前先做文本级归一化，比逐个放宽每条正则的字符集更不容易引入副作用。
                o_plain = re.sub(r"(?<=[0-9A-Za-z])~(?=[0-9A-Za-z])", "-", o_plain)
                o_layout = re.sub(r"(?<=[0-9A-Za-z])~(?=[0-9A-Za-z])", "-", o_layout)
                full_text, layout_text = o_plain, o_layout
                ocr_used = True
                warnings.append("扫描件，内容由 OCR 识别 —— 数字可能有误，请人工复核")
            else:
                warnings.append("扫描件，OCR 未能识别出文字")
        else:
            warnings.append("PDF 无可读文字（扫描件），且本机未安装 OCR")

    header = {}
    for field, patterns in LABEL_PATTERNS.items():
        val = None
        if field == "invoice_date":          # 标签在上值在下的排版，优先按列取
            c = _label_below(layout_text, r"Invoice\s*Date\b")
            if c and _date(c):
                val = c
        for text in (full_text, layout_text):        # 普通文本优先，排版文本兜底
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

    # 表格里的 key-value 补齐正则漏掉的字段（标签与值不同行）
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

    # 供应商：正文第一行通常是抬头；文件名括号内的名字作为兜底与校验
    header["supplier"] = clean_supplier(_guess_supplier(full_text))

    fn = parse_filename(path)
    header["is_proforma"] = fn.get("is_proforma", False)
    header["supplier_from_filename"] = fn.get("supplier")

    if header.get("po_no") in (None, "") and fn.get("po_no"):
        header["po_no"] = fn["po_no"]
    elif fn.get("po_no") and str(header.get("po_no")) != str(fn["po_no"]):
        # 正文抽到的 PO 不是 PONCG 格式，而文件名是——大概率正文那个根本不是
        # Neocytogen 自己的 PO，是文档里恰好也叫「Reference No. / PO No」的别的号码。
        # 实测 UPS 的 Import Tax Invoice：正文没有真正的 PO 字段，货运明细表里的
        # "Reference No."（货运公司自己的运单参考号 5282293306）被 po_no 的通用
        # 兜底正则当成了 PO。PONCG 格式本来就是代码里公认「最可靠」的一条，
        # 文件名给的是这个格式而正文给的不是，就该信文件名。
        body_po = str(header.get("po_no") or "")
        fn_po = str(fn["po_no"])
        if not re.fullmatch(r"PONCG\d{9,}", body_po, re.I) and re.fullmatch(r"PONCG\d{9,}", fn_po, re.I):
            warnings.append(f"正文 PO 号 {header['po_no']} 不是 PONCG 格式，很可能抓错了字段，已改用文件名里的 {fn_po}")
            header["po_no"] = fn_po
        else:
            warnings.append(f"PO 号不一致：正文 {header['po_no']} vs 文件名 {fn['po_no']}")

    if not valid_invoice_no(header.get("invoice_no")) and fn.get("invoice_no_hint"):
        header["invoice_no"] = fn["invoice_no_hint"]
        warnings.append("发票号取自文件名")

    # 供应商优先用文件名括号里的简称 —— 台账 Company Ordered From 也是简称，
    # 正文抬头往往是全称甚至夹着地址/税号，反而更难对上
    if fn.get("supplier"):
        if header.get("supplier") and normalize_cmp(header["supplier"]) != normalize_cmp(fn["supplier"]):
            header["supplier_from_body"] = header["supplier"]
        header["supplier"] = fn["supplier"]

    # 日期只在正文抽不到时才用文件名，且必须提示 —— 文件名日期有手误
    # （如 PONCG202511160 写成 2025-01-03，实为 2026）
    if header.get("invoice_date") is None and fn.get("invoice_date"):
        header["invoice_date"] = fn["invoice_date"]
        warnings.append("发票日期取自文件名，文件名日期常有笔误，请人工确认")

    if header.get("is_proforma"):
        warnings.append("这是 Proforma（形式发票），非付款依据")

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
            # 合计行：第一格写 subtotal/total 且没有描述
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
        warnings.append("未找到明细表格")

    if not subtotal or subtotal.get("amount_incl") is None:
        fallback = _totals_from_text(full_text, layout_text)
        if fallback.get("amount_incl") is not None:
            subtotal = {**fallback, **{k: v for k, v in (subtotal or {}).items() if v is not None}}
            if not items:
                warnings.append("未抽到明细行，已按发票合计参与对账")

    # 再兜一层：TOTAL_PATTERNS 全部要求标签和数字在同一行，但「标签在上、值在下
    # 一行同一列」是发票版式里常见的另一种布局（Sigma 这类系统生成的发票尤其明显：
    # 一行印 "... TOTAL CURRENCY"，数值在下一行同列 "... 601.35 SGD"）。
    # 复用 _label_below 的列对齐逻辑去找 TOTAL 标签正下方那个数字。
    if (subtotal or {}).get("amount_incl") is None:
        tok = _label_below(layout_text, r"\bTOTAL\b")
        val = _money(tok)
        if val is not None and val > 0:
            subtotal = dict(subtotal or {})
            subtotal["amount_incl"] = val
            warnings.append("发票合计取自「TOTAL 标签下一行同列」的版式，请人工确认")

    # 校验：行加总 vs 发票合计（行级 GST 有自己的舍入，只在合计层面卡，容差 0.05）
    priced = [i for i in items if i.get("amount_incl") is not None]
    if items and not priced:
        warnings.append("明细行未取到金额，已改用发票合计对账")
        items = []
    elif priced and len(priced) == len(items) and subtotal.get("amount_incl") is not None:
        s = round(sum(i["amount_incl"] for i in priced), 2)
        if abs(s - subtotal["amount_incl"]) > 0.05:
            warnings.append(f"行加总 {s} 与发票合计 {subtotal['amount_incl']} 不符")

    # OCR 的发票号常见 O/0、l/1、S/5、B/8 混淆。只在校验没过时试一次纠正，
    # 且只纠正发票号这一个字段 —— 全文替换会把公司名里的 O 变成 0。
    if ocr_used and not valid_invoice_no(header.get("invoice_no")) and header.get("invoice_no"):
        try:
            import ocr as _ocr_mod
            fixed = _ocr_mod.fix_ocr_digits(header["invoice_no"])
        except Exception:
            fixed = None
        if fixed and valid_invoice_no(fixed):
            header["invoice_no"] = fixed
            warnings.append(f"发票号经 OCR 字符纠正（原读作 {header['invoice_no']}），请人工确认")

    # ---- 视觉模型兜底 ----
    # 触发条件：扫描件，或有文字但发票号 / 合计任一没抽到。
    # 原代码只在 full_text 为空时走 OCR，且从未调用过 ai_extract —— 扫描件
    # 因此全军覆没。这里让模型只负责「读」，读回来仍走同一套金额等式校验。
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
                warnings.append(f"视觉模型调用失败：{data['_error']}")
            else:
                mapped = _from_ai(data)
                if mapped:
                    ai_header, ai_items, ai_sub = mapped
                    used = []
                    for k, v in ai_header.items():
                        if v not in (None, "") and header.get(k) in (None, "", False):
                            header[k] = v
                            used.append(k)
                    # 发票号：模型读的优先级高于「取自文件名」的猜测值
                    if valid_invoice_no(ai_header.get("invoice_no")):
                        if normalize_cmp(ai_header["invoice_no"]) != normalize_cmp(header.get("invoice_no")):
                            if not valid_invoice_no(header.get("invoice_no")):
                                header["invoice_no"] = ai_header["invoice_no"]
                                used.append("invoice_no")
                            else:
                                warnings.append(
                                    f"发票号存疑：正则读作 {header['invoice_no']}，"
                                    f"模型读作 {ai_header['invoice_no']}")
                    # 日期若是「取自文件名」的猜测值，模型读到的实际印刷日期优先
                    if (ai_header.get("invoice_date")
                            and any("日期取自文件名" in w for w in warnings)
                            and ai_header["invoice_date"] != header.get("invoice_date")):
                        header["invoice_date"] = ai_header["invoice_date"]
                        warnings = [w for w in warnings if "日期取自文件名" not in w]
                        used.append("invoice_date")
                    if subtotal.get("amount_incl") is None and ai_sub.get("amount_incl") is not None:
                        subtotal = ai_sub
                        used.append("金额合计")
                    if not items and ai_items:
                        items = ai_items
                    if used:
                        warnings.append("以下字段由视觉模型读取，请人工复核：" + "、".join(used))
        elif is_scanned:
            warnings.append("扫描件且视觉模型不可用（缺 ANTHROPIC_API_KEY 或未装 anthropic）")

    # 模型/OCR 读回来的金额同样过等式校验，不因为来源不同就免检。
    #
    # 纠正方向取决于来源：
    #   原生 PDF —— 三个数都是确定性读出的，总额是权威，用它倒推税额。
    #   OCR      —— 反过来。分项（63.20 / 5.69）位数少、字形简单，比总额更容易读对；
    #                总额往往是版面最大最花的一行，反而最容易错。
    #                实测 Lip Laundry 收据总额 68.90 被读成 68.50，
    #                而 63.20 + 5.69 = 68.89，与台账仅差 1 分的 rounding。
    #                若仍固定相信总额，就会凭空造出 0.39 的差额报成金额不符。
    e, g, i = subtotal.get("amount_excl"), subtotal.get("gst"), subtotal.get("amount_incl")
    if None not in (e, g, i) and abs(e + g - i) > 0.05:
        if ocr_used:
            subtotal["amount_incl"] = round(e + g, 2)
            warnings.append(
                f"OCR 金额等式不成立（{e} + {g} ≠ {i}），已改用分项加总 {round(e + g, 2)} 作为合计"
                f"（原读作 {i}），请人工核对")
        else:
            subtotal["gst"] = round(i - e, 2)
            warnings.append(f"金额等式不成立（{e} + {g} ≠ {i}），已按税后-税前重算税额")

    if not valid_invoice_no(header.get("invoice_no")):
        header["invoice_no"] = None
        warnings.append("未抽到发票号 —— 无法参与匹配")
    if not header.get("supplier"):
        warnings.append("未抽到供应商")

    return {**header, "ocr": ocr_used, "line_items": items, "subtotal": subtotal,
            "source_file": os.path.basename(path), "warnings": warnings}


if __name__ == "__main__":
    import sys, json
    r = extract_invoice(sys.argv[1])
    print(json.dumps(r, indent=2, ensure_ascii=False, default=str))
