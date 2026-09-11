"""
ocr.py — 扫描件 PDF 的文字识别

只在 pdfplumber 抽不到文字时启用。用 PyMuPDF 把页面渲染成高分辨率图片，
再交给 Tesseract 识别。

依赖：
    pip install pytesseract pymupdf pillow
    另需安装 Tesseract 本体（Windows: https://github.com/UB-Mannheim/tesseract/wiki）

注意：OCR 出来的文字准确率远低于原生 PDF，数字尤其容易错
（0/O、1/l/I、5/S、8/B）。所以 OCR 结果必须经过金额等式校验，
校验不过的一律进「待核查」，不能直接采信。
"""

import os
import re
import shutil

__version__ = "2026-09-11.3"

DPI = 300          # 低于 300 识别率明显下降；再高收益有限且很慢
_TESS_OK = None


def _find_tesseract():
    """定位 tesseract.exe。

    Windows 装完常常不在 PATH 里，开始菜单里那个 Tesseract-OCR 文件夹只是快捷方式，
    不是安装目录。所以依次尝试：PATH -> 注册表 -> 常见安装路径 -> 全盘常见位置搜索。
    """
    import pytesseract

    if shutil.which("tesseract"):
        return True

    cands = []

    # UB Mannheim 安装包会把安装目录写进注册表
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

    # 还是找不到就在常见根目录下搜一层
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
    """返回实际使用的 tesseract 路径，方便排查。"""
    import pytesseract
    return (shutil.which("tesseract")
            or getattr(pytesseract.pytesseract, "tesseract_cmd", None))


def available():
    """OCR 是否可用。不可用时上层应跳过而不是报错。"""
    global _TESS_OK
    if _TESS_OK is None:
        try:
            import pytesseract, fitz          # noqa: F401
            _TESS_OK = _find_tesseract()
        except ImportError:
            _TESS_OK = False
    return _TESS_OK


def ocr_pdf(path, max_pages=4):
    """返回 (纯文本, 排版文本)，与 pdfplumber 的两种模式对应，
    这样下游解析逻辑完全不用改。识别失败返回 ("", "")。"""
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
                img = img.convert("L")          # 灰度，票据识别更稳
            # psm 6：整页当作一个统一文本块，比默认模式更适合发票版式
            plain.append(pytesseract.image_to_string(img, config="--psm 6"))
            # 保留字符位置，供「标签在上、值在下」的列对齐逻辑使用
            layout.append(_layout_from_data(
                pytesseract.image_to_data(img, config="--psm 6",
                                          output_type=pytesseract.Output.DICT)))
        except Exception:
            continue
    doc.close()
    return "\n".join(plain), "\n".join(layout)


def _layout_from_data(data, char_w=9):
    """把 Tesseract 的逐词坐标还原成等宽排版文本，
    让 _label_below 的列对齐能在扫描件上同样生效。"""
    rows = {}
    n = len(data.get("text", []))
    for i in range(n):
        word = (data["text"][i] or "").strip()
        if not word:
            continue
        try:
            if int(data["conf"][i]) < 30:      # 置信度过低的词直接丢
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
    """只在「应当是数字」的字段上纠正常见混淆，绝不全文替换 ——
    全文替换会把公司名里的 O 变成 0，制造新的错误。"""
    if s is None:
        return None
    return (str(s).replace("O", "0").replace("o", "0")
            .replace("l", "1").replace("I", "1")
            .replace("S", "5").replace("B", "8"))
