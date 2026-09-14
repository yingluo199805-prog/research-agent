# -*- coding: utf-8 -*-
"""按「页顶编号标题」抓附注页并输出版面文本（Mac 版辅助工具，依赖 PyMuPDF）。

用法: python3 dump_note_pages.py <pdf> <标题关键词1> [<关键词2> ...] [--span N] [--out 文件]
对每个关键词，找页面前 200 字符内出现「编号+关键词」标题的页（如 "6 物業、廠房及設備"、"18 政府補貼"），
连同其后 span 页一并输出（表格常跨页），每页按词坐标重排成 pdftotext -layout 风格。
找不到标题时退回全文首次出现该关键词且含表格特征词的页。输出页码为 PDF 物理页码（1 起）。
"""
import sys, re, os
import fitz

sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "..", "revenue-split", "scripts"))
from pdf_text import _layout_text_fitz

MARK = ["期初", "期末", "年初", "年末", "累計", "累计", "賬面", "账面", "添置", "增加", "減少", "转入", "轉入", "撥備", "折舊", "攤銷"]


def find_pages(doc, kw, head_window=220):
    pat = re.compile(r'(^|\n)\s*\d{1,3}\s*[\.．、]?\s*' + re.escape(kw))
    hits = [i for i in range(len(doc)) if pat.search(doc[i].get_text()[:head_window])]
    if hits:
        return hits, "heading"
    fb = [i for i in range(len(doc)) if kw in doc[i].get_text() and sum(m in doc[i].get_text() for m in MARK) >= 4]
    return fb[:3], "fallback"


def main():
    args = sys.argv[1:]
    span = 1; out = None
    if "--span" in args:
        i = args.index("--span"); span = int(args[i + 1]); del args[i:i + 2]
    if "--out" in args:
        i = args.index("--out"); out = args[i + 1]; del args[i:i + 2]
    pdf, kws = args[0], args[1:]
    doc = fitz.open(pdf)
    chunks = []
    for kw in kws:
        pages, how = find_pages(doc, kw)
        pages = sorted(set(p for h in pages for p in range(h, min(h + span + 1, len(doc)))))
        chunks.append(f"\n#################### [{kw}] ({how}) pages(1-based): {[p + 1 for p in pages]}")
        for p in pages:
            chunks.append(f"\n---------- page {p + 1} ----------\n" + _layout_text_fitz(doc[p]))
    text = "\n".join(chunks)
    if out:
        os.makedirs(os.path.dirname(os.path.abspath(out)), exist_ok=True)
        open(out, "w", encoding="utf-8").write(text)
        print(f"written {out} ({len(text)} chars)")
    else:
        print(text)


if __name__ == "__main__":
    main()
