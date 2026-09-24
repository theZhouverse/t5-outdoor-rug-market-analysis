# 户外地垫市场分析

当前有效规范是 [SPEC 3.7](docs/SPEC.md)。父体主指标按月份+父ASIN计算Q/T平均；BSR按月份+BSR值计算Q/T平均，再按三档输出。当前Excel、HTML、JSON、Markdown与构建清单已经按同一SPEC 3.7批次重建并完成本地验收；旧版本只作历史追溯。详细处理步骤和人工复算方法见 [数据处理流程与人工验算](docs/数据处理流程与人工验算-20260922.md)，执行与审查记录见 [TASKS](docs/TASKS.md) 和 [REVIEWS](docs/REVIEWS.md)。

## 当前有效口径

- 24个源文件共48,000行。`小类目`只作为回勾字段，不作为筛选条件；按“小类BSR”解析，只要存在1—100（含100）的有效排名就进入候选池。本批次候选行数为46,780，父体月份组为1,334。
- 多值“小类BSR”全部保留并标记 `MULTI_BSR`；每个有效排名形成BSR观察记录，BSR键为月份+BSR值。
- 统计单元是“月份+父ASIN”，父ASIN缺失时回退ASIN，再回退源行ID；当前形成1,334个父体月份组合。同月ASIN重复将阻止发布，本批次为0组。
- 父体Q列月销量和T列月销售额($)不相加，分别对同一月份+父ASIN内的有效非负值求平均；保留有效数、覆盖率、最小值、最大值和 `AVERAGE_CONFLICT`。
- 有效子体销售额仍只累加“子体销量”和“子体销售额”同时为有效非负数的子体行，但只作为覆盖诊断，不替代Q/T父体主销售额。
- 父体主均价=父体平均T÷父体平均Q；子体诊断均价=有效子体销售额÷有效子体销量。空白保留为空，不转换成0；0销量但正销售额视为无效矛盾。
- PP按父体候选子体标题中完整单词 `plastic` 的多数归类，高客单价市场是其补集；整体市场=PP+高客单价市场。Genimo是同一父体池中的品牌视角。
- MOM=本月÷去年同月；YOY=本年与上一年共同覆盖月份汇总相除，均为比较倍数，不再减1。月度YOY使用截至当前月的共同月份累计值；当前只有2024.08—2026.07数据，因此2025从8月开始、2026从1月开始，05页年度YOY分别比较8—12月和1—7月，不是完整自然年度YOY。
- BSR分层是独立支线：同月同BSR的Q/T取平均，再汇总为头部1—20、中部21—50、尾部51—100。
- BSR分层按业务范围分别计算：整体、PP、高客单价、Genimo和Genimo PP先过滤本范围观察记录，再按月份+BSR值取平均，避免不同市场在同一排名上的记录互相混入。
- 不计算利润；供应方数据属于估算，程序能证明处理、公式与交付一致，不能证明估算等于真实成交。

## 当前 SPEC 3.7 输出

- `outputs/20260920-new-source-parent-model/户外地垫市场分析-SPEC3.7-父体口径.xlsx`：16个子表，包含父体Q/T平均、BSR月份+BSR值平均、四部分固定指标和公式回勾。
- `outputs/20260920-new-source-parent-model/户外地垫市场分析-SPEC3.7-父体口径.html`：与Excel同批次的交互报告，含四部分、1.1—4.7定位、下拉导航、折叠、筛选、趋势图、父体回查和Excel下载。
- `outputs/20260920-new-source-parent-model/户外地垫市场分析-SPEC3.7-父体口径.json`：供网页使用的公开汇总数据，不含原始明细快照。
- `outputs/20260920-new-source-parent-model/完整审计结果-spec37.json`：源范围、模型、工作簿公式和网页的一致性审计摘要。
- `outputs/20260920-new-source-parent-model/公式审计-spec37.json`：独立公式解释器对9,260个保存公式的求值结果，以及Q/T负值、空白和明确0共六类输入扰动结果。
- `outputs/20260920-new-source-parent-model/构建清单-spec37.json`：本批次核心文件指纹。

线上交付页：<https://thezhouverse.github.io/t5-outdoor-rug-market-analysis/>。`gh-pages`提交为`9350d83bf671582f541181fa10f7ebf886db1faa`；线上HTML与本地LF规范化内容的SHA-256均为`32b788ac8f0186a1da8c4586471440c658a64bcf24d3ba66730cfda9e79edb6e`，线上下载XLSX与本地文件的SHA-256均为`72d19d4fdd1395f8ffc8373f0b0ad593fffc44f58d1a60b2817cece3c4bbe5c6`。

旧的 `户外地垫市场分析-SPEC3-父体口径*` 文件保留为历史输出，不能与当前 SPEC 3.7 数字混用。

旧的 `outputs/20260911-spec2-market-analysis`、`交付/` 和REVIEWS中SPEC 3.0以前的数字只用于历史追溯，不能与当前批次混用。

## 构建与检验

```powershell
npm run build:parent
npm run test:parent
npm run test:parent:formula
```

`npm run test:parent`会独立重新读取24个源文件，核对46,780条候选行、月份+父ASIN的Q/T平均、46,854条BSR观察记录、月份+BSR值平均、三档汇总、PP/高客单价回加、去年同月MOM、共同月份年度YOY、XLSX公式缓存和HTML批次。

本机完整工作簿通过XlsxWriter生成，当前文件为8,339,360字节。独立公式解释器逐项核对公式与模型；WPS表格12.0已在临时副本上执行完整重算并保存，保存后的副本再次通过9,260个公式审计。Artifact Tool对16个工作表的顶部数据区逐表渲染，浏览器另验收HTML主题、折叠、年份筛选和父体检索。本机没有Microsoft Excel，因此WPS目标引擎验证不表述为Microsoft Excel桌面版验收。
