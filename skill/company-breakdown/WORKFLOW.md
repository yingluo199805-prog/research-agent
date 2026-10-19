# company-breakdown 逐步命令（以港股为例；A 股差异见文末）

新交付默认使用 [比亚迪表达模板](references/byd-expression.md)。完成取数与勾稽后，以 `assets/byd-expression-template.json` 为布局参考建立公司的 `workbook_spec.json`，将路径写入 `company.json.workbook_spec`，再运行总控。以下 Python 引擎步骤保留数据抽取和旧版复现；不得把旧版三线表当成新模板。


所有命令在仓库根目录执行；`$B` = `outputs/收入分配`（三表/年报/收入分配工作簿根目录），`$C` = `outputs/{公司}`（公司目录）。

## 第 1 步：公司目录与 company.json

```bash
mkdir -p "outputs/小米/研究底稿/_text"
cp skill/company-breakdown/references/company_template.json outputs/小米/company.json   # 改 company/short/market/年份/文件名
```

## 第 2 步：年报入库 + 附注地图

年报放 `$B/{公司}{起}-{终}年报/`，文件名前缀 `YYYY-MM-DD：`。港股用披露易标题检索接口列全直链（见 engines/revenue-split/SKILL.md）；
先用 PyMuPDF 抽任意附注页看文本是否乱码，乱码年份另下英文版到 `_en/`。

```bash
# 锚定年附注地图：编号标题 → 页码（脚本示例见 references/hk_extract_recipe.md）
# 抽 11 组附注 + 三张主表到逐年文本（每年一个文件，之后只读这些文件，不通读 PDF）
python3 skill/company-breakdown/engines/notes-analysis/scripts/dump_note_pages.py "$B/小米2018-2025年报/2026-04-28：小米集团2025年度報告.pdf" \
   "物業及設備" "無形資產" "其他收入" "遞延" "按性質劃分之開支" "僱員福利開支" --span 1 --out "$C/研究底稿/_text/notes_2025.txt"
```

## 第 3 步：读数 → 两份 JSON（分析师手工，价值所在）

- `$C/revenue_split.json`：港股按 `references/company_template.json` 的 `revenue_split_h_share` 段填 note_data（百万元，逐年当年列），
  特殊桶（投资公允价值变动、优先股公允价值变动、联营）放 operating_items；A 股填 src_rows/fy_map（用 detect_src_rows.py 自动定位）。
- `$C/notes_spec.json`：字段见 `engines/notes-analysis/scripts/build_hk_notes_model.py` 末尾 SPEC_DOC（千元）。
  港股 PP&E 按净值滚动：cost/accdep 期末 + additions/transfer_in/disposal_nbv/depreciation/other_nbv；类别口径变更年填 reclass_cost/reclass_acc。
  **五张表的 remarks 必须写**：口径、特殊事项（带数字）、校验差异点名、关键预测输出。

## 第 4 步：底稿 spec

`$C/研究底稿/digao_spec.json`：每科目每年的 PDF 物理页码（PP&E 2-3 页、无形 2-3 页、利润表+按性质页、其他收入+递延收入页）。乱码年份用英文版 PDF。

## 第 5 步：一键生成 + 自检

```bash
python3 skill/company-breakdown/scripts/run_breakdown.py outputs/小米/company.json
#   --skip-split / --skip-digao / --skip-notes 跳过某步；--only-check 只重算自检；--base <目录> 改根目录
```

产物：`$B/{公司}H股.xlsx`（收入分配单件）、`$C/{公司}_财报拆分_{起}-{终}.xlsx`（合并版，九张表）、`$C/研究底稿/{公司}_年报附注底稿.xlsx`、`*_recalc.xlsx`（重算副本，仅自检用）。

## 第 6 步：收口

看驱动打印的验收清单：非零校验行去对应表备注点名；每表挑 3 个硬编码值对底稿截图；比率区间体检；收入分配占比讲得通。

## 第 7 步：交付 + 登记

发合并版工作簿（>6MB 的底稿手机端会传不上，说明桌面端路径），在 `references/companies.md` 登记年份、口径特点、遗留问题。

---

## A 股差异

- 收入分配：`revenue_split.json` 用 `references/company_template.json` 的 `revenue_split_a_share` 段（prefix/folder/fy_map/src_rows/warranty），
  引擎自动 PDF→文本→解析销管研附注→抽人数→建表→勾稽；三表必须先用 `combine_three_statements.py` 合并。
- 附注五表：A 股附注按原价/累计折旧/减值滚动、有独立在建工程附注与研发投入情况表、政府补助有非经常性损益与增值税加计抵减口径，
  目前按 `engines/notes-analysis/references/*.md` 规格手工建表（用 openpyxl），建完同样用 `recalc_check.py` 自检，再用 `build_hk_notes_model.py` 的
  merge 思路合并（或直接把五张表建在收入分配工作簿里并加目录页）。A 股生成器待抽象（下一家 A 股重点公司时做）。
