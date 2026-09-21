---
name: company-breakdown
description: >-
  重点公司财报拆分总控（一家公司一套产物、一个 Excel）。给定任一 A 股/港股上市公司，从上市以来（或指定年份）的年报出发，
  依次产出：收入分配拆分表（营业收入按供应商/员工/银行/政府/折旧/摊销/股东分配，供应商为 residual、加总恒=100%）、
  年报附注五科目分析表（固定资产/在建工程/无形资产/资本化研发/政府补助，带勾稽校验、隐含比率反推与预测假设）、
  年报附注截图底稿，并合并进同一个工作簿（目录+三表+附注底稿+收入分配拆分+附注五表）。
  触发：用户说"做XX的财报拆分""把XX拆一下""XX收入分配+附注分析""按小米/零跑那套做XX""重点公司拆分"。
  不触发：只要单张收入分配表或单张附注表时，直接用 engines/ 下对应引擎；穿透叙事/穿透财报的叙事验证 → penetrate-* 技能。
---

# 公司财报拆分（company-breakdown）

一家公司 = 一个公司目录 `outputs/{公司}/` + 一个合并版工作簿 `{公司}_财报拆分_{起}-{终}.xlsx`。
两个引擎收在 `engines/` 下（原 build-revenue-split、annual-report-notes-analysis），本技能只定义**顺序、输入契约、交付契约与验收**，
确定性的生成/合并/自检交给 `scripts/run_breakdown.py`。已按此交付：零跑（2022-2025）、小米（2018-2025）。

## 0. 环境（本机 macOS）

Python 3.9 + openpyxl + PyMuPDF + formulas。无 Excel COM / LibreOffice / pdftotext / web-access：年报文本用 PyMuPDF 词坐标重排，
重算用 formulas，下载用 WebFetch/curl。根目录 `REVENUE_SPLIT_BASE`（默认 `<仓库>/outputs/收入分配`，放三表导出、年报 PDF 与收入分配工作簿）；
公司产物放 `outputs/{公司}/`（含 `研究底稿/`）。全部在 gitignore 内。

## 1. 输入契约（开工前确认，只问这两个）

- **年份范围**：默认"上市以来"到最新已披露年报（当前月份 ≥5 取去年，否则取前年）；附注五表另加一个基准年（起始年前一年，取自起始年报比较列）。
- **年报来源**：先查 `outputs/收入分配/{公司}*年报/` 是否已有；没有则从港交所披露易/巨潮下载，**下载动作须用户点头**。
  港股用标题检索接口一次列全（见 `engines/revenue-split/SKILL.md`「年报获取与文本层」）；中文 PDF 文本层乱码时读英文版同页。
- 三表：A 股必须有 iFinD/Choice 三表导出（`{公司}2015-2025三表/`）；港股可无，引擎按 note_data 生成静态三表。

其余都用默认，开工后在交付摘要说明。

## 2. 七步流程

| 步 | 做什么 | 工具 | 产物 |
|---|---|---|---|
| 1 | 建公司目录，写 `company.json`（模板 `references/company_template.json`） | 手工 | `outputs/{公司}/company.json` |
| 2 | 附注地图：锚定最新年报，列编号标题→页码；抽 IS/BS/CF + 11 组附注页文本 | `engines/notes-analysis/scripts/dump_note_pages.py`（港股 `references/hk_extract_recipe.md` 有整套关键词） | `研究底稿/_text/` |
| 3 | 读数填两份 JSON：收入分配 `revenue_split.json`（港股：note_data 按年；A 股：src_rows/fy_map）、附注 `notes_spec.json`（PP&E/无形/研发/补助）+ 五张表的备注 | 手工（分析师价值所在） | 两份 JSON |
| 4 | 写 `研究底稿/digao_spec.json`（每科目每年页码） | 手工 | spec |
| 5 | 一键：收入分配 → 底稿截图 → 附注五表合并 → 重算自检 → 回写缓存值（预览可见） → 验收清单 | `scripts/run_breakdown.py company.json` | 合并版工作簿 + 底稿 |
| 6 | 自检收口：校验行非零必须能在备注点名；抽样 3 数对底稿；比率区间体检 | 看 recalc 输出 + 备注 | — |
| 7 | 交付摘要 + 更新 `references/companies.md` 登记 | 手工 | — |

A 股与港股的口径差异（类别、净值滚动、在建工程位置、研发资本化、补助口径、金融子公司）见 `references/a_share_vs_hk.md`。

## 3. 交付契约

合并版工作簿固定九张表（顺序不可改）：`目录` / `{公司}三表` / `{公司}附注底稿` / `收入分配拆分` / `固定资产` / `在建工程` / `无形资产` / `资本化研发` / `政府补助分析`。
颜色规范：深蓝=年报实际、深红=可调假设、黑=公式、黄底=合计、绿底=校验。每张附注表末尾必有备注区：【口径】【特殊事项】【关键预测输出】，
非零校验行逐条点名。另交 `研究底稿/{公司}_年报附注底稿.xlsx`（截图，按年份升序）。

## 4. 验收（run_breakdown 自动打印，人工补三条）

- 公式错误 = 0；预测列无空值；校验行归零或已点名。
- 收入分配：供应商占比在行业区间、员工/折摊趋势能讲通、特殊桶（投资公允价值变动、上市年股份支付、优先股公允价值变动）已在备注说明。
- 附注五表：隐含折旧率/摊销率与会计政策年限一致、转固率与产能节奏一致、政府补助三口径能互相解释。
- 三个数抽样与底稿截图一致（这是校验行抓不到的一层，不能省）。

## 5. 常见坑（两家跑下来的）

- 港股 PP&E 按净值滚动，处置原价要倒算；资产类别口径变更年（小米 2024）用 `reclass_cost/reclass_acc` 单列期初重分类，不要让倒算吸收。
- 政府补助：港股只有其他收入/递延收入/现金流三口径；年报不拆数（小米 2024-2025）就留 0 并在备注说明，不要用其他收入总额冒充。
- 递延收入滚动对不平（零跑 2024）是真实的信息，放"其他变动"行并点名，不要硬调。
- 上市年一次性股份支付（小米 2018 年 124 亿）会把员工桶顶高，备注必须说。
- 来源列文字不能以"="开头；空格写 0 或 None，不写空串。
- 新上市公司（奇瑞 2025）：上市前年份从招股章程「附录一 会计师报告」取，三年并列口径一致、附注编号与年报不同（用 `extract_hk_sections.py --half 0.6 --extra` 指定标题）；会计师报告里公司层面表（「貴公司」）与集团表（「貴集團」）相邻，只取集团表。
- 港股无形资产里有「資本化開發支出/開發成本」类别时（奇瑞），在 notes_spec.intangibles 填 `dev_cost_class`：该类不摊销、购置=资本化研发表本年新增、转出=转入 `rd_class`；资本化率预测默认取近两年平均，利润质量判断要还原。
- openpyxl 生成的工作簿没有公式缓存值，手机/网页预览会显示为空——驱动最后一步 `embed_cached_values.py` 把 formulas 重算结果回写成缓存值，交付前必须跑。

## 目录

```
company-breakdown/
├── SKILL.md / WORKFLOW.md
├── scripts/run_breakdown.py          # 总控驱动
├── references/company_template.json  # company.json + 两份引擎 JSON 的模板说明
├── references/a_share_vs_hk.md       # 口径决策表
├── references/hk_extract_recipe.md   # 港股附注抽取关键词与页定位配方
├── references/companies.md           # 已完成公司登记
└── engines/
    ├── revenue-split/                # 收入分配表引擎（原 build-revenue-split）
    └── notes-analysis/               # 附注五表引擎（原 annual-report-notes-analysis）
```
