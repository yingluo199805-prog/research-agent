# -*- coding: utf-8 -*-
"""港股年报/招股书一次性抽取：三张主表 + 11 组附注页 → 逐年文本文件（版面重排，去重页，跳过董事酬金/附属公司清单页）。

用法: python3 extract_hk_sections.py <pdf> <out.txt> [--lang cn|en] [--half 0.5] [--extra "关键词1|别名,关键词2"]
  --lang  默认按 PDF 首页是否含 CJK 自动判断
  --half  只在 PDF 后 half 部分找附注（默认 0.5；招股书会计师报告在很后面可给 0.6）
  --extra 追加要抽的附注标题（正则可用 | 给别名），逗号分隔
输出格式与 hk_extract_recipe.md 一致：`#################### [段名] (how) pages [..]` 分段，每页 `---------- page N ----------`。
之后只读这个文本，不要通读 PDF。
"""
import sys, os, re
import fitz

sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "..", "revenue-split", "scripts"))
from pdf_text import _layout_text_fitz

STMT = {"cn": [("IS", "合併損益表|合併損益及其他全面收益表|合併全面收益表|綜合損益表|合并利润表", ["收入", "銷售成本|销售成本", "毛利"]),
               ("BS", "合併資產負債表|合併財務狀況表|綜合財務狀況表|合并资产负债表", ["物業|物业", "無形資產|无形资产"]),
               ("CF", "合併現金流量表|綜合現金流量表|合并现金流量表", ["經營活動|经营活动", "投資活動|投资活动"])],
        "en": [("IS", "CONSOLIDATED INCOME STATEMENT|CONSOLIDATED STATEMENT OF PROFIT OR LOSS|CONSOLIDATED STATEMENTS OF PROFIT OR LOSS", ["Revenue", "Cost of sales", "Gross profit"]),
               ("BS", "CONSOLIDATED BALANCE SHEET|CONSOLIDATED STATEMENT OF FINANCIAL POSITION", ["Property", "Intangible"]),
               ("CF", "CONSOLIDATED STATEMENT OF CASH FLOWS|CONSOLIDATED CASH FLOW STATEMENT", ["operating activities", "investing activities"])]}
NOTES = {"cn": [("其他收入", 1), ("其他虧損淨額|其他收益淨額|其他（虧損）／收益淨額|其他收益／（虧損）淨額|其他（虧損）／收益|其他收益|其他虧損|其他收益及虧損", 1),
                ("按性質劃分之開支|按性質劃分的開支|按性質分類的開支|按性質劃分的費用", 1), ("僱員福利開支|員工成本|僱員成本", 1),
                ("財務收入及成本|財務收入淨額|財務收入|財務成本|融資成本", 1), ("物業、廠房及設備|物業及設備|物業、機器及設備", 2), ("無形資產", 3), ("租賃|使用權資產", 1),
                ("其他非流動負債|遞延收入|遞延收益", 1), ("其他非流動資產", 1), ("所得稅費用|所得稅開支", 1), ("政府補助|政府補貼", 1), ("撥備|保修", 1)],
         "en": [("Other income", 1), ("Other gains|Other losses|Other \\(losses\\)/gains|Other gains/\\(losses\\)", 1), ("Expenses by nature", 1), ("Employee benefit expenses|Staff costs", 1),
                ("Finance income and costs|Finance income, net|Finance income|Finance costs", 1), ("Property, plant and equipment|Property and equipment", 2), ("Intangible assets", 3), ("Leases|Right-of-use assets", 1),
                ("Other non-current liabilities|Deferred income", 1), ("Other non-current assets", 1), ("Income tax expenses|Income tax expense", 1), ("Government grants|Government subsidies", 1), ("Provisions|Warranty", 1)]}
SKIP = re.compile(r"Five highest paid|Benefits and interests of directors|Directors’ termination|Major subsidiaries|董事的利益|五名最高薪|主要附屬公司|董事福利及利益|董事酬金")


def squeeze(t):
    return "\n".join(re.sub(r"  +", "  ", l.rstrip())[:175] for l in t.splitlines() if l.strip())


def norm(s):
    return re.sub(r"\s+", " ", s).upper()


def extract(pdf, out, lang=None, half=0.5, extra=None):
    d = fitz.open(pdf); n = len(d); texts = [d[i].get_text() for i in range(n)]
    if lang is None:
        lang = "cn" if re.search(r"[一-鿿]", texts[1] + texts[2]) else "en"
    chunks, seen = [], set()

    def add_pages(head, pages):
        chunks.append(f"\n#################### {head} pages {[p + 1 for p in pages]}")
        for p in pages:
            if p in seen:
                continue
            seen.add(p)
            body = _layout_text_fitz(d[p])
            if SKIP.search(body):
                chunks.append(f"---------- page {p + 1} ---------- (skipped: directors/subsidiaries)"); continue
            chunks.append(f"---------- page {p + 1} ----------\n" + squeeze(body))
    for tag, title, kws in STMT[lang]:
        tpat = re.compile(title, re.I)
        pages = [i for i in range(int(n * 0.3), n) if tpat.search(norm(texts[i][:500])) and all(re.search(k, texts[i], re.I) for k in kws)]
        if not pages:
            pages = [i for i in range(int(n * 0.3), n) if all(re.search(k, texts[i], re.I) for k in kws) and "附註" not in texts[i][:200] and "NOTES" not in norm(texts[i][:200])][:1]
        add_pages(f"[{tag}]", pages[:2])
    notes = list(NOTES[lang])
    if extra:
        notes += [(e.strip(), 1) for e in extra.split(",") if e.strip()]
    for kw, span in notes:
        pat = re.compile(r'(^|\n)\s*\d{1,2}(\.\d+)?\.?\s+(' + kw + r')(\s*\((Continued|continued)\)|（續）)?\s*(\n|$)', re.I)
        heads = [i for i in range(int(n * half), n) if pat.search(texts[i][:700])]; how = "heading"
        if not heads:
            heads = [i for i in range(int(n * half), n) if re.search(kw, texts[i][:700], re.I)][:2]; how = "head700"
        if not heads:
            heads = [i for i in range(int(n * half), n) if re.search(kw, texts[i], re.I)][:2]; how = "anywhere"
        pages = sorted(set(p for h in heads[:4] for p in range(h, min(h + span + 1, n))))
        add_pages(f"[{kw[:24]}] ({how})", pages)
    os.makedirs(os.path.dirname(os.path.abspath(out)), exist_ok=True)
    open(out, "w", encoding="utf-8").write("\n".join(chunks))
    empties = [c.splitlines()[0] for c in chunks if c.startswith("\n####") and "pages []" in c]
    print(f"{os.path.basename(pdf)} [{lang}] -> {out}: {len(''.join(chunks)) // 1000}K chars; EMPTY: {empties}")


if __name__ == "__main__":
    a = sys.argv[1:]
    if len(a) < 2:
        print(__doc__); sys.exit(1)
    lang = a[a.index("--lang") + 1] if "--lang" in a else None
    half = float(a[a.index("--half") + 1]) if "--half" in a else 0.5
    extra = a[a.index("--extra") + 1] if "--extra" in a else None
    extract(a[0], a[1], lang, half, extra)
