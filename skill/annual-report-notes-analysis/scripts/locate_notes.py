# -*- coding: utf-8 -*-
"""
附注定位器（通用版）—— 在任意公司年报PDF里快速定位各科目附注所在页，避免逐页通读省token。

设计目标：**跨公司通用**。不依赖某家公司特定的科目命名或行名，而是三层叠加：
  ① 先读 PDF 书签目录 get_toc()（多数A股/港股年报都有，最可靠、最省）；
  ② 科目名用「广义同义词集」（简体+繁体+多种叫法）做全文命中；
  ③ 命中页按「通用滚动表特征词」密度排序——期初/期末/本期增加/累计折旧/累计摊销/账面原值/
     减值准备 等，是所有公司这类附注表都会出现的词，与具体公司无关。
输出每科目的候选页（按相关度排序），**供子代理只读这几页确认**，而不是通读整本。

用法：
    python locate_notes.py <某份年报.pdf>          # 单份
    python locate_notes.py <存PDF的文件夹>          # 批量（文件名含年份则按年归类）

注意：
- 这是「缩小范围」工具，不是「确定答案」工具。拿到候选页后，子代理应渲染/读取该页确认是正确的表，
  再喂给 build_digao.py。命中为空时说明该公司用词特殊，回退到读目录页或通读财务附注章节。
- A股 vs 港股差异：「研发投入情况表」是中国证监会规定格式、仅A股有；港股无此表，研发资本化信息散见于
  无形资产/开发成本附注与会计政策。政府补助在港股多并入「其他收入及收益」「递延收入」，用词不同。
"""
import sys, os, json, glob, re
import fitz

# 科目名「广义同义词」：任一命中即认为该页可能涉及此科目。尽量覆盖简繁与不同叫法。
SUBJECT_NAMES = {
    "固定资产": ["固定资产", "固定資產", "物业、厂房及设备", "物業、廠房及設備", "物业厂房及设备",
              "房屋、机器及设备", "房地產、廠房及設備", "不动产、厂房和设备"],
    "在建工程": ["在建工程", "在建工程", "在建工程及预付", "在建工程／工程物资", "在建工程在建"],
    "无形资产": ["无形资产", "無形資產", "无形资产及其他", "其他无形资产"],
    "开发支出": ["开发支出", "開發支出", "开发成本", "開發成本", "研发支出资本化", "資本化開發成本"],
    "研发费用_按性质": ["研发费用", "研發費用", "研究及开发", "研究及開發", "研发及试验"],
    "研发投入情况表": ["研发投入情况表", "本期费用化研发投入", "本期費用化研發投入",
                 "本期资本化研发投入", "本期資本化研發投入", "研发投入资本化的比重", "研發投入資本化的比重"],
    "政府补助_其他收益": ["其他收益", "其他收入及收益", "其他收益及收入"],
    "政府补助_递延收益": ["递延收益", "遞延收益", "递延收入", "遞延收入"],
    "政府补助_营业外收入": ["营业外收入", "營業外收入", "其他收入"],
    "政府补助_非经常性损益": ["非经常性损益", "非經常性損益"],
}
# 政府补助类科目，额外要求页面出现「政府补助/政府補助」才算数（否则其他收益/递延收益噪音太多）
NEEDS_GRANT_WORD = {"政府补助_其他收益", "政府补助_递延收益", "政府补助_营业外收入", "政府补助_非经常性损益"}
GRANT_WORDS = ["政府补助", "政府補助", "政府資助", "政府拨款", "政府撥款"]

# 「通用滚动表特征词」：跨公司不变，用于给命中页打分——出现得越多，越像是那张附注明细表而非正文提及。
TABLE_MARKERS = ["期初余额", "期初餘額", "期末余额", "期末餘額", "本期增加", "本年增加", "本期减少", "本年减少",
                 "账面原值", "賬面原值", "账面价值", "賬面價值", "账面净值", "賬面淨值",
                 "累计折旧", "累計折舊", "累计摊销", "累計攤銷", "减值准备", "減值準備",
                 "原值", "购置", "購置", "处置", "處置", "转入", "轉入"]


def has_heading(text, names, head_window=160):
    """该页顶部是否出现「编号 + 科目名」标题（如 14、固定资产 / 16. 物業、廠房及設備）。
    这是附注起始页的强信号，能把"正文顺带提及"和"目录索引"区分开。返回 True/False。
    A股/港股财务附注均按数字编号，故 \\d+[、．.] + 科目名 通用。"""
    head = text[:head_window]
    for n in names:
        # 编号紧跟科目名（中间允许空格/全角空格）
        if re.search(r'\d{1,3}\s*[、．.]\s*' + re.escape(n), head):
            return True
    return False


def score_page(text, names, need_grant):
    name_hits = sum(text.count(n) for n in names)
    if name_hits == 0:
        return 0, 0, False
    if need_grant and not any(g in text for g in GRANT_WORDS):
        return 0, 0, False
    marker_hits = sum(1 for m in TABLE_MARKERS if m in text)
    head = has_heading(text, names)
    # 页顶编号标题命中 → 强加权（+100），确保附注起始页排到最前；其余按名频+表格词密度
    score = (100 if head else 0) + name_hits + marker_hits * 2
    return score, marker_hits, head


def from_toc(doc, names):
    """从书签目录找标题含科目名的条目，返回其页码（1基）。最可靠。"""
    pages = []
    for lvl, title, page in doc.get_toc():
        if any(n in (title or "") for n in names):
            pages.append(page)
    return sorted(set(pages))


def locate(pdf_path, span=1, topn=3):
    doc = fitz.open(pdf_path)
    pages_text = [doc[i].get_text() for i in range(doc.page_count)]
    result = {}
    for key, names in SUBJECT_NAMES.items():
        need_grant = key in NEEDS_GRANT_WORD
        # 层①：目录
        toc_pages = from_toc(doc, names)
        # 层②③：全文打分
        scored = []
        any_heading = False
        for i, t in enumerate(pages_text):
            s, mk, head = score_page(t, names, need_grant)
            if head:
                any_heading = True
            if s > 0 and (mk >= 1 or head or key in ("研发投入情况表",) or need_grant):
                scored.append((i + 1, s, mk, head))
        scored.sort(key=lambda x: (-x[1], x[0]))
        # 若存在「页顶标题」命中页，只保留这些（最可靠）；否则退回按分数取 topn
        if any_heading:
            cand = [p for p, _, _, head in scored if head][:topn]
        else:
            cand = [p for p, _, _, _ in scored[:topn]]
        # 候选 = 目录页 ∪ 高分页，连带跨页
        base = sorted(set(toc_pages) | set(cand))
        expanded = sorted(set(p for b in base for p in range(b, b + span + 1) if p <= len(pages_text)))
        result[key] = {
            "toc_pages": toc_pages,
            "top_pages": cand,
            "with_span": expanded,
            "best": (toc_pages[0] if toc_pages else (cand[0] if cand else None)),
        }
    doc.close()
    return result


def main():
    if len(sys.argv) < 2:
        print("usage: python locate_notes.py <pdf 或 文件夹>", file=sys.stderr); sys.exit(1)
    arg = sys.argv[1]
    if os.path.isdir(arg):
        targets = sorted(glob.glob(os.path.join(arg, "*.pdf")) + glob.glob(os.path.join(arg, "*.PDF")))
    else:
        targets = [arg]
    out = {}
    for p in targets:
        m = re.search(r"(19|20)\d{2}", os.path.basename(p))
        year = m.group(0) if m else os.path.basename(p)
        out[year] = {"pdf": os.path.basename(p), "located": locate(p)}
    print(json.dumps(out, ensure_ascii=False, indent=1))


if __name__ == "__main__":
    main()
