# Lightweight Acceptance Test

Use this when a full annual-report run would consume too many tokens.

## Goal

Learn whether the skill preserves the reference model's structure, applies disclosure-specific line-item rules, and validates revenue reconciliation without reading every annual report page.

## Low-Token Test Procedure

1. Select only two years:
   - one early year with sparse notes
   - one recent year with detailed notes

2. Extract only these source snippets:
   - income statement page/table
   - sales expense note
   - management expense note
   - R&D expense or R&D spending note
   - payroll note
   - depreciation/amortization note if used by the bridge

3. Build a mini output workbook with:
   - the copied reference revenue-split sheet structure
   - only the two selected years populated
   - all other historical and forecast cells blank
   - the same section labels, styles, and formulas as the reference model where applicable

4. Run acceptance checks:
   - output sheet visually matches the reference model's structure
   - section order matches the reference model
   - formulas are present in calculated rows
   - sheet can scroll normally; no deep copied freeze pane blocks navigation
   - A1/A2 hyperlinks jump to the actual output sheet's total/detail sections
   - forecast/yellow cells are blank
   - detailed rows match disclosed note items
   - no undisclosed "其他" row was invented
   - no three-statement formula contains `/100`
   - statement references are quoted and wrapped in `N(...)`
   - revenue split difference is within 2%, or a reason is documented

## Token-Saving Prompt

Ask for a smoke test like this:

```text
使用 $build-revenue-split 做低 token smoke test。只读取 2018 年和 2024 年年报中利润表、销售费用、管理费用、研发费用/研发支出、职工薪酬、折旧摊销相关页；用比亚迪模型收入拆分表作为强制模板，只生成两年数据的迷你收入拆分表，并输出结构/公式/收入校验结果。
```

## Pass/Fail Standard

Pass only if the mini sheet still looks like the BYD/reference revenue-split sheet. If it becomes a generic extracted-note table, the skill failed.
