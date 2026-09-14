# company-breakdown — 上市公司财报拆分 skill（收入分配表 + 年报附注五表，一个 Excel 交付）

给 Claude Code / Claude 桌面版使用的 skill。给定任一 A 股或港股公司的年报，产出：
收入分配拆分表（营业收入按供应商/员工/银行/政府/折旧/摊销/股东分配，加总恒=100%）、
年报附注五科目分析表（固定资产/在建工程/无形资产/资本化研发/政府补助，带勾稽校验与预测假设）、
年报附注页截图底稿，并合并为一个九张表的工作簿。已在零跑（2022-2025）、小米（2018-2025）上跑通。

## 安装

1. 把本目录放到 Claude Code 能读到的 skill 位置：项目级 `.claude/skills/company-breakdown/`，或个人级 `~/.claude/skills/company-breakdown/`。
2. Python 3.9+，依赖：
   ```bash
   pip install openpyxl pymupdf formulas
   ```
   不需要 Excel、LibreOffice、pdftotext（有的话会自动优先用 pdftotext）。
3. 工作根目录（放三表导出、年报 PDF、收入分配工作簿）：设环境变量 `REVENUE_SPLIT_BASE`，
   或每次命令加 `--base <目录>`；默认为 `<skill 上四级目录>/outputs/收入分配`。公司产物放 `<根目录同级>/outputs/{公司}/`。

## 怎么用

- 在 Claude 里直接说"做 XX 的财报拆分 / 按小米那套把 XX 拆一下"，Claude 会按 `SKILL.md` 的七步流程走。
- 一家公司三份 JSON：`company.json`（入口）、`revenue_split.json`（收入分配引擎）、`notes_spec.json`（附注五表引擎），
  模板与字段说明见 `references/company_template.json`；然后一条命令：
  ```bash
  python3 scripts/run_breakdown.py outputs/XX/company.json
  ```
  它依次跑：收入分配 → 底稿截图 → 附注五表合并 → formulas 重算自检 → 回写缓存值 → 打印验收清单。
- 读数、填 JSON、写每张表的备注是分析师的活（价值所在）；脚本只做确定性的生成与校验。

## 目录

- `SKILL.md` 触发条件、输入/交付契约、七步流程、验收、常见坑（先读这个）
- `WORKFLOW.md` 逐步命令
- `references/` company.json 模板、A 股 vs 港股口径决策表、港股附注抽取配方、已完成公司登记
- `scripts/run_breakdown.py` 总控驱动
- `engines/revenue-split/` 收入分配表引擎（A 股 `run_v2_batch.py --config`、港股 `run_h_share_batch.py --config`，各有 SKILL.md/WORKFLOW.md/规则文档）
- `engines/notes-analysis/` 附注五表引擎（港股数据驱动生成器 `build_hk_notes_model.py`、截图底稿 `build_digao.py`、重算自检 `recalc_check.py`、缓存值回写 `embed_cached_values.py`；A 股按 `references/*.md` 规格手工建表）

## 已知边界

- A 股附注五表生成器尚未抽象，目前 A 股只有收入分配全自动。
- 港股部分中文年报 PDF 文本层乱码（字体无 Unicode 映射），需下载英文版读数，截图底稿也用英文版。
- 格式按广发汽车研报三线表规范（思源黑体、深蓝表头），可在生成器常量里改。
