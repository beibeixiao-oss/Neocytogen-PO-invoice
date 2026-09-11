"""
diagnose.py — 单张 PDF 排查工具

    python diagnose.py "路径\\某张扫描件.pdf"

依次检查：pdfplumber 能读到多少字 -> OCR 是否可用 -> OCR 读到多少字 ->
最终抽到哪些字段。哪一步断了一眼就看得出来。
"""

import sys
import os


def main(path):
    print("=" * 60)
    print("文件:", os.path.basename(path))
    print("存在:", os.path.exists(path))
    if not os.path.exists(path):
        return

    # 1. 原生文字
    try:
        import pdfplumber
        with pdfplumber.open(path) as pdf:
            native = "\n".join((p.extract_text() or "") for p in pdf.pages)
        print(f"\n[1] pdfplumber 原生文字: {len(native)} 字符  ({len(pdf.pages) if False else ''})")
        if native.strip():
            print("    -> 有文字，不会走 OCR。前 120 字:")
            print("   ", repr(native[:120]))
    except Exception as e:
        native = ""
        print("\n[1] pdfplumber 失败:", e)

    # 2. OCR 环境
    try:
        import ocr
        ok = ocr.available()
        print(f"\n[2] OCR 可用: {ok}   路径: {ocr.where()}")
    except ImportError as e:
        print("\n[2] 无法 import ocr:", e)
        return
    except Exception as e:
        print("\n[2] OCR 检测异常:", e)
        return

    # 3. 实际识别
    if not native.strip() and ok:
        print("\n[3] 正在 OCR，请稍候…")
        try:
            plain, layout = ocr.ocr_pdf(path)
            print(f"    OCR 纯文本: {len(plain)} 字符 / 排版文本: {len(layout)} 字符")
            if plain.strip():
                print("    前 400 字:")
                print("   ", repr(plain[:400]))
            else:
                print("    !! 一个字都没识别出来 —— 图片可能太糊、倾斜或是纯图形")
        except Exception as e:
            print("    OCR 报错:", type(e).__name__, e)
    elif native.strip():
        print("\n[3] 跳过 OCR（本身就有文字）")

    # 4. 完整抽取结果
    print("\n[4] 抽取结果:")
    try:
        from invoice_extractor import extract_invoice
        r = extract_invoice(path)
        for k in ["invoice_no", "invoice_date", "po_no", "due_date",
                  "supplier", "currency", "is_proforma", "ocr"]:
            print(f"    {k:14} {r.get(k)}")
        print(f"    subtotal       {r.get('subtotal')}")
        print(f"    明细行数        {len(r.get('line_items') or [])}")
        print(f"    warnings       {r.get('warnings')}")
        if "ocr" not in r:
            print("\n    !! 返回结果里没有 'ocr' 字段 —— invoice_extractor.py 是旧版，请更新")
    except Exception as e:
        import traceback
        traceback.print_exc()


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print('用法: python diagnose.py "某张扫描件.pdf"')
    else:
        main(sys.argv[1])
