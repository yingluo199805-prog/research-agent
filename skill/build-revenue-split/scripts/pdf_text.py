# -*- coding: utf-8 -*-
"""年报 PDF -> 版面文本。

优先用 pdftotext -layout（若本机有）；否则用 PyMuPDF 按词坐标重排，模拟 -layout 的
"同一基线的词排在同一行、按 x 间距补空格"效果，保证附注表格的『科目  本期  上期』落在同一行，
parse_notes_generic / extract_headcount 的正则可直接使用。
"""
import os, shutil, subprocess


def _layout_text_fitz(page, ytol=3.0, char_w=4.0):
    words = page.get_text("words")  # x0, y0, x1, y1, word, block, line, wordno
    if not words:
        return ""
    words.sort(key=lambda w: (w[3], w[0]))
    lines, cur, cury = [], [], None
    for w in words:
        if cury is None or abs(w[3] - cury) <= ytol:
            cur.append(w)
            cury = w[3] if cury is None else cury
        else:
            lines.append(cur)
            cur, cury = [w], w[3]
    if cur:
        lines.append(cur)
    out = []
    for ln in lines:
        ln.sort(key=lambda w: w[0])
        s, lastx = "", 0.0
        for w in ln:
            nsp = max(1, int((w[0] - lastx) / char_w)) if s else int(w[0] / char_w)
            s += " " * nsp + w[4]
            lastx = w[2]
        out.append(s)
    return "\n".join(out)


def pdf_to_text(pdf, out_txt):
    """把 pdf 转成版面文本写到 out_txt。已存在则跳过。返回 out_txt。"""
    if os.path.exists(out_txt) and os.path.getsize(out_txt) > 0:
        return out_txt
    os.makedirs(os.path.dirname(out_txt), exist_ok=True)
    if shutil.which("pdftotext"):
        subprocess.run(["pdftotext", "-layout", pdf, out_txt], capture_output=True)
        if os.path.exists(out_txt) and os.path.getsize(out_txt) > 0:
            return out_txt
    import fitz  # PyMuPDF
    doc = fitz.open(pdf)
    with open(out_txt, "w", encoding="utf-8") as f:
        for page in doc:
            f.write(_layout_text_fitz(page))
            f.write("\n\f\n")
    return out_txt


if __name__ == "__main__":
    import sys
    print(pdf_to_text(sys.argv[1], sys.argv[2]))
