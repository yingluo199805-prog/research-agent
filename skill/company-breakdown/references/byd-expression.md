# 比亚迪表达模板（byd-261011）

用户于 2026-10-19 选择所上传的《比亚迪模型-261011》作为财报拆分的默认表达。这里固化财报拆分相关表的布局与计算层级；不复制销售预测、估值、行业月度数据库等其他模块，也不把比亚迪的资产类别、假设或财务数值移植给其他公司。

## 收入分配

顺序：总项金额 → 总项占比 → 分组细项金额 → 同序细项占比 → 底部附表。顶部十四项顺序沿用主技能；细项中，经营损失先列各收益/减值构成，业务费用下分销售/管理/研发第三方费用及制造其他，员工薪酬下依次列制造/销售/管理/研发工资、股份支付、人数和人均工资，随后列供应商、银行、政府、折旧、摊销和两类股东利润。

- A 列成本属性、B 列计算方法备注、C 列项目、D 列起年份升序；冻结 D3。
- 金额为亿元，人数为人，人均工资为万元/年。人数和人均工资不除以营业收入、不进入占比区；细项占比与货币细项同序。
- 微软雅黑 10pt，行高 18；A/B/C 宽 11/27/57，年份宽 14.33。长计算方法换行并增高行高。
- 粉色段标题 `F2DCDB`；收入/加总标签黄底 `FFFF00`；桥接科目标签浅蓝底 `DBE5F1`、红色粗字。普通数字黑色居中，金额两位小数、占比一位百分数。占比蓝色渐变数据条按 0—100% 同尺展示。
- 红字桥接为层级强调，不表示假设。可调输入以黄底标识，附注表假设仍为深红字。不要把收入表全体数字套成“实际蓝、公式黑”的附注样式。
- 来源备注须与科目一致：少数股东行对应少数股东损益，母公司股东行对应归母净利润。
- 附表放间接费用折摊、期初薪酬、制造其他估算比例、营业成本供应商复核、费用勾稽和口径说明。主表总项不增加“其他”来平差。

## 附注五表

固定资产/无形资产按期末净值、原价、累计折摊、减值、当期折摊、隐含比率、预测假设、历史滚动附表和备注组织。假设按驱动因素分组，每组再列资产类别。在建工程先总额和项目，再滚动、预算、转固桥接、假设及预测。资本化研发保留投入、性质、费用化/资本化及开发支出滚动和无形资产联动；披露不足标“未单列”。

B 科目、C 来源、D 起年份；蓝表头 `1F4E78`、浅蓝章节 `D9E2F3`，数值合计黄底 `FFE699`，校验绿底 `E2EFDA`。实际数深蓝、公式黑、可调假设深红。千元整数展示，底层数值保留精度。资产类别、项目和非零差额按本公司披露填写。

政府补助独立采用 A 科目、B 起年份。顶部汇总亿元，下设其他收益、补助损益去向、递延滚动、经常/非经常口径、非补助税费优惠、现金六块千元明细；蓝色大标题、浅蓝汇总条，每块重复年份表头。资产相关/经常性不能互相等同。不得将未披露视作零，也不能把增值税加计抵减混入补助。

## 生成和接入

表达规则与空白模板不依赖某个模型或库。在 Codex 环境可使用下述 artifact 渲染路径；其他环境可用可用的表格工具编辑空白模板，按同一规则填充并重算验证。不具备 artifact 依赖时不能静默退回旧样式。

- 精确样式参数：`../assets/byd-expression-v1.json`。
- 可填空白模板：`../assets/byd-expression-template.xlsx`；其生成输入为 `../assets/byd-expression-template.json`，仅含布局与无公司数值的公式，使用前按目标公司替换科目与项目。
- `scripts/build_workbook.mjs` 是表达层，使用 `@oai/artifact-tool`。在支持该库的 Node 环境运行：`node scripts/build_workbook.mjs workbook_spec.json output.xlsx`。不把本机 runtime 路径写死到仓库。单独调用渲染器后，再用 `python3 scripts/finalize_navigation.py workbook_spec.json output.xlsx` 写入内部跳转；总控已自动接入此步，直接修改 OOXML 保留公式缓存和数据条。
- `company.json` 增加 `"workbook_spec": "workbook_spec.json"`，总控将直接走新表达层；可用 `ARTIFACT_NODE` 指向 artifact runtime 的 Node。原数据抽取 JSON 和原始截图底稿仍按主技能准备。未设置该字段的旧配置走原生成器，供复现旧交付，不能据此声称已应用新模板。

`workbook_spec.json` 顶层：`sheets`、`actualEnd`、`output`（也可 CLI 指定）、可选 `chainTests`/`previews`。固定九表顺序。各表包含 `name`、`years`、`rows`（完整二维值/公式）、`styles`（行号→语义样式）、`formats`（行号→Excel 数字格式）；收入表 `layout=revenue,labelColumn=3,sourceColumn=2,yearColumn=4`，附注 `layout=notes,labelColumn=2,sourceColumn=3,yearColumn=4`，补助 `layout=government,labelColumn=1,yearColumn=2`。收入另给 `detailRange`、`headerRows`、`ratioRanges`；补助给 `headerRows`、`summaryRows`。公式已按最终位置建立，渲染器不猜测科目映射。

样式名包括 `title/subtitle/header/section/total/check/assumption/ratio/note`；收入另有 `pink/plainHeader/revenueTitle/revenueSummary/revenueBridge/subtotal/summaryRatio/detailRatioBridge/detailRatioSubtotal/detailRatio`；补助 `govTitle/govSection/govSummary`。不要让布局调整悄悄改变会计口径。

## 验收

重构时对照重构前缓存值，逐年逐项核对保留的金额/比率；抽查原始截图；检查公式错误和缓存缺失。改变一个在建工程输入，确认固定资产跟随变化；改变资本化率，确认开发支出及无形资产跟随变化，再恢复原值。渲染收入金额/占比/细项、预测假设和补助顶部，核实小数位、层级、换行和导航。

原披露差异继续显示并在备注解释；“收入占比加总 100%”只是 residual 恒等关系，不能代替费用和成本独立复核。后续若拟改变此表达，先展示具体改法并取得用户确认。
