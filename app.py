"""
app.py — 采购发票对账工具（界面版）

    streamlit run app.py

左边传 Procurement Tracking List，右边传发票 PDF（可多选）。
传一张 = 单张核对，传一整个月 = 批量对账，走的是同一套逻辑。
"""

__version__ = "2026-09-11.1"

import os
import tempfile
from io import BytesIO

import pandas as pd
import streamlit as st

# Streamlit Community Cloud 没有 .env 文件，key 存在 App -> Settings -> Secrets 里。
# ai_extract.py 只认 os.environ / 本地 .env，这里把 Secrets 桥接成环境变量，
# 本地开发（用 .env）和云端部署（用 st.secrets）走的是同一份代码不用改。
if "ANTHROPIC_API_KEY" not in os.environ:
    try:
        os.environ["ANTHROPIC_API_KEY"] = st.secrets["ANTHROPIC_API_KEY"]
    except (KeyError, FileNotFoundError):
        pass

from reconcile import reconcile, find_candidates, write_output, OUT_COLS


def show(df, **kw):
    """安全显示表格：object 列一律转字符串，避免 Arrow 类型冲突。"""
    df = pd.DataFrame(df).copy()
    for c in df.columns:
        if df[c].dtype == "object":
            df[c] = df[c].map(lambda v: "" if v is None or (isinstance(v, float) and pd.isna(v))
                              else (str(int(v)) if isinstance(v, float) and v.is_integer() else str(v)))
    st.dataframe(df, width="stretch", hide_index=True, **kw)

st.set_page_config(page_title="采购发票对账", layout="wide")
st.title("采购发票对账")
st.caption("Excel 直读，PDF 抽取，按发票号精确匹配")

# 各模块版本必须一致。文件没换全是最常见的故障，放在最显眼处
import excel_loader as _el, invoice_extractor as _ie, reconcile as _rc
_vers = {"app": __version__, "reconcile": getattr(_rc, "__version__", "?"),
         "invoice_extractor": getattr(_ie, "__version__", "?"),
         "excel_loader": getattr(_el, "__version__", "?")}
try:
    import ocr as _ocr
    _vers["ocr"] = getattr(_ocr, "__version__", "?")
    _ocr_ok, _ocr_path = _ocr.available(), _ocr.where()
except ImportError:
    _vers["ocr"] = "未安装"
    _ocr_ok, _ocr_path = False, None

if len(set(_vers.values())) > 1:
    st.error("文件版本不一致，请把所有 .py 重新覆盖一遍： "
             + "；".join(f"{k} {v}" for k, v in _vers.items()))
else:
    st.caption(f"版本 {__version__} · OCR "
               + (f"就绪（{_ocr_path}）" if _ocr_ok else "不可用"))

uploads = st.file_uploader(
    "把台账 Excel 和发票 PDF 一起拖进来（可多选，顺序无所谓）",
    type=["xlsx", "xlsm", "pdf"], accept_multiple_files=True,
)

xlsx_ups = [f for f in (uploads or []) if f.name.lower().endswith((".xlsx", ".xlsm"))]
pdf_ups = [f for f in (uploads or []) if f.name.lower().endswith(".pdf")]

c1, c2 = st.columns(2)
c1.metric("台账 Excel", len(xlsx_ups))
c2.metric("发票 PDF", len(pdf_ups))

if not uploads:
    st.info("台账和发票一起拖进来即可。PDF 可以只传一张做单笔核对，也可以整个月一起传。")
    st.stop()

if not xlsx_ups:
    st.warning("还缺台账 Excel（.xlsx）。")
    st.stop()
if not pdf_ups:
    st.warning("还缺发票 PDF。")
    st.stop()

if len(xlsx_ups) > 1:
    names = [f.name for f in xlsx_ups]
    pick = st.selectbox("传了多份 Excel，用哪一份作为台账？", names)
    xlsx_file = next(f for f in xlsx_ups if f.name == pick)
else:
    xlsx_file = xlsx_ups[0]
pdf_files = pdf_ups

if not st.button("开始对账", type="primary"):
    st.stop()

# 上传的是内存对象，落到临时目录后复用命令行版的同一套逻辑
with tempfile.TemporaryDirectory() as tmp:
    xlsx_path = os.path.join(tmp, "tracking.xlsx")
    with open(xlsx_path, "wb") as f:
        f.write(xlsx_file.getbuffer())

    pdf_paths = []
    for up in pdf_files:
        p = os.path.join(tmp, up.name)
        with open(p, "wb") as f:
            f.write(up.getbuffer())
        pdf_paths.append(p)

    with st.spinner(f"正在处理 {len(pdf_paths)} 份发票…"):
        sheets, pending, ctx = reconcile(xlsx_path, pdf_paths)

    # 兼容旧版 reconcile.py：缺字段时降级，不让整个页面崩掉
    ctx.setdefault("suspicious", [])
    ctx.setdefault("superseded", [])
    ctx.setdefault("unmatched_pdf", [])
    ctx.setdefault("failed", [])

    buf = BytesIO()
    write_output(sheets, pending, buf, ctx["suspicious"], ctx.get("superseded"))
    buf.seek(0)

n_inv = len(ctx["invoices"])
m1, m2, m3, m4, m5 = st.columns(5)
m1.metric("读入发票", n_inv)
m2.metric("匹配成功", len({r["Invoice Number"] for r in sheets["matched"]}))
m3.metric("金额不符", len({r["Invoice Number"] for r in sheets["discrepancy"]}))
m4.metric("台账查无此单", len(ctx["unmatched_pdf"]))
m5.metric("台账有单缺发票",
          len({r["Invoice Number"] for r in sheets["Excel to PDF - not match"]}))

st.download_button("下载 outcome.xlsx", buf,
                   file_name="Neocytogen - outcome.xlsx",
                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# 待核查：程序自己算出哪些结果可疑并排序，避免人工逐张点开核对
if ctx["superseded"]:
    with st.expander(f"已忽略的 Proforma（{len(ctx['superseded'])} 张）"):
        st.caption("同一 PO 已有 Tax Invoice，按规则只算 Tax Invoice。列在此处仅供查证。")
        show(ctx["superseded"])

susp = ctx["suspicious"]
if susp:
    st.subheader(f"待核查 · {len(susp)} 张（共 {n_inv} 张）")
    st.caption("按疑点数量排序。只看这几张即可，其余已通过金额等式、格式、税率等自动校验。")
    show(susp)
else:
    st.success(f"{n_inv} 张发票全部通过自动校验，无需人工核查。")

# 抽取告警：字段没抽到、行加总对不上等，先让人知道哪些结果不可全信
warned = [i for i in ctx["invoices"].values() if i["warnings"]]
if warned or ctx["failed"]:
    with st.expander(f"提取告警（{len(warned) + len(ctx['failed'])} 份）", expanded=False):
        for inv in warned + ctx["failed"]:
            st.write(f"**{inv.get('invoice_no') or inv['source_file']}** — "
                     + "；".join(inv["warnings"]))

tabs = st.tabs(["matched", "discrepancy", "台账查无此单", "台账有单缺发票", "未开票"])

for tab, key in zip(tabs[:4], ["matched", "discrepancy",
                               "PDF to Excel - not match", "Excel to PDF - not match"]):
    with tab:
        rows = sheets[key]
        if rows:
            show(pd.DataFrame(rows)[OUT_COLS])
        else:
            st.write("无")

with tabs[4]:
    st.caption("货未到或未开票，不算 mismatch，仅供了解还有哪些单悬着")
    show(pending.drop(columns=[c for c in pending.columns if c.startswith("_")]
                      + ["invoice_key", "po_key", "supplier_key"], errors="ignore"))

# 台账里找不到时，列出最接近的几条供人工判断
if ctx["unmatched_pdf"]:
    st.divider()
    st.subheader("人工核对建议")
    st.caption("以下仅为参考排序，程序不会据此自动匹配 —— 发票号必须完全相等才算 matched")
    for inv in ctx["unmatched_pdf"]:
        total = (inv.get("subtotal") or {}).get("amount_incl")
        with st.expander(f"发票 {inv.get('invoice_no')} · {inv.get('supplier')} · {total}"):
            cands = find_candidates(ctx["expected"], inv)
            if cands:
                show(cands)
            else:
                st.write("台账里没有相近的记录。")
            for w in inv["warnings"]:
                st.warning(w)
