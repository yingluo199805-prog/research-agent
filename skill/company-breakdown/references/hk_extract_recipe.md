# 港股年报抽取配方（零跑、小米实测）

## 1. 附注地图（锚定最新年报）

```python
import fitz, re
d = fitz.open(pdf)
cn = re.compile(r'(?:^|\n)\s*(\d{1,2})\s+([一-鿿][一-鿿、，（）／—－\s]{1,40}?)(?:（續）)?\s*(?:\n|$)')
en = re.compile(r'(?:^|\n)\s*(\d{1,2})\s+([A-Z][A-Za-z ,&\-/()]{2,60}?)(?:\s*\(Continued\))?\s*(?:\n|$)')
for i in range(len(d)//2, len(d)):
    for m in cn.finditer(d[i].get_text()[:600]): print(m.group(1), m.group(2), i+1)
```
输出即 {附注号 → 标题 → 首页}。同一家公司逐年编号至多偏移 1-2 位，页码平移。

## 2. 要抽的 3 张主表 + 11 组附注（中/英关键词）

| 目标 | 中文（繁） | 英文 | 用途 |
|---|---|---|---|
| IS | 合併損益表 / 合併全面收益表 | CONSOLIDATED INCOME STATEMENT | 收入分配全部损益行；研发开支；收入 |
| BS | 合併資產負債表 | CONSOLIDATED BALANCE SHEET | PP&E/无形资产 BS 校验行 |
| CF | 合併現金流量表 | CONSOLIDATED STATEMENT OF CASH FLOWS | 收取的政府補貼（零跑有、小米无） |
| 其他收入 | 其他收入 | Other income | 政府补贴（小米 2024 起不拆数） |
| 其他收益 | 其他（虧損）／收益淨額 | Other gains/(losses), net | 经营损失桶 |
| 按性质开支 | 按性質劃分之開支 / 按性質劃分的開支 | Expenses by nature | 员工总额、折旧摊销合计、保修、研发资本化声明 |
| 雇员福利 | 僱員福利開支 | Employee benefit expenses | 股份支付 |
| 财务收支 | 財務收入及成本 / 財務收入淨額 | Finance income and costs | 银行桶 |
| PP&E | 物業、廠房及設備 / 物業及設備 | Property, plant and equipment / Property and equipment | 固定资产+在建工程表（3 页：本年、上年、折旧按功能） |
| 无形资产 | 無形資產 | Intangible assets | 无形资产表（含摊销按功能、商誉减值测试） |
| 租赁 | 租賃 | Leases | 使用权资产折旧（折旧总口径） |
| 其他非流动负债/其他应付 | 其他非流動負債 / 其他應付款項及應計費用 | Other non-current liabilities / Other payables and accruals | 递延政府补助（流动+非流动） |
| 递延收入 | 遞延收入 | Deferred income | 零跑式独立附注 |

标题跨行（"CONSOLIDATED INCOME \nSTATEMENT"）→ 匹配前先把空白压成单空格。附注标题找不到时退回 fallback：页内含关键词且表格特征词 ≥4。

## 3. 净值滚动 → spec 字段映射（PP&E）

| 附注行 | spec 字段 | 备注 |
|---|---|---|
| 於年初/年末：成本、累計折舊 | cost / accdep（期末） | 基准年取起始年报比较列 |
| 添置 | additions | 在建工程列 → cip.additions |
| 轉讓 / Transfers | transfer_in（各类别正数）；cip.transfer_out（正数） | 只算转入 PP&E 类别的部分 |
| 出售 / Disposals | disposal_nbv（正数） | 在建工程列的处置 → cip.other_dec |
| 折舊費用 | depreciation | |
| 匯兌差額、收購附屬公司、轉至/自投資物業 | other_nbv（正=增加） | 在建工程列 → cip.other_dec（正=减少） |
| 类别口径变更年 | reclass_cost / reclass_acc | 新口径期初 − 旧口径上年期末，各类加总=0 |
| 折舊按功能 | dep_by_function | 校验行=计提合计 |

无形资产同法：商誉进 `indefinite`；收购进 `combination`；汇兑进 `other_nbv`。

## 4. 已知坑

- 港交所中文 PDF 乱码（私用区字符）：PyMuPDF 抽任意附注页，若出现 `㛮⸶⹨` 类字符即乱码，改英文版；截图底稿也用英文版并在 note 里注明。
- 英文版年报 IS 标题跨行；2024+ 小米 PP&E 表头两行（"Factories and / production equipment"），列名靠位置对齐。
- 年报只给"上年/本年"两列，第 N 年的期初=第 N-1 年报的期末，两份年报的重叠列应一致，不一致（重述/重分类）时以当年报为准并记 reclass。
