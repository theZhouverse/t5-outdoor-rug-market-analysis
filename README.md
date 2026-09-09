# T5 户外地垫市场数据 -> SQLite 转换工程

> Excel (卖家精灵 / 飞书多维表格导出) -> 完整无损 -> SQLite (单文件 DB)
>
> 遵循 [`.config/单个项目开发规范手册`](../.config/单个项目开发规范手册：AI 驱动开发的工程化标准.md) 三文档闭环 (SPEC / TASKS / REVIEWS)。

## 1. 项目概述

- **目的**: 把 44 MB 的多 sheet Excel 文件（50 张月度明细 + 4 张 TOP 汇总）无损导入到 SQLite, 便于后续 BI / 二次分析。
- **数据源**: `data/raw/地垫-卖家精灵市场数据.xlsx` (卖家精灵市场数据, 美国站, 2022.06 - 2026.07)
- **输出**: `data/processed/market.db` (单文件 SQLite, ~48 MB)；领导验收结果工作簿 `outputs/20260909-formula-market-analysis/户外地垫市场分析-公式版-20260909.xlsx`
- **工作表口径**: Excel 界面可见 23 张；工作簿内部另有 32 张隐藏表（31 张历史月度明细 + 1 张无有效区域的遗留 `Sheet6`）
- **规模**: 主 Excel 基础导入 73,812 条；应用 2026.01-07 竞品代表行后当前为 71,451 条，54 张业务表（50 月度 + 4 TOP）+ meta/sheet_catalog/analysis_replacements 元数据
- **2026 口径（2026-09-08 审计更新）**: 2026.01-06 竞品快照为**全市场父体级导出**（每月 1,038-1,993 个父体），2025 则为包含 ASIN/父ASIN 富文本标识的行级导出（含变体行）。两者统计单元不同，跨年数值只作方向性参考，不构成严格同口径同比；2026.07 的 94 父体小样本只展示规模，结果工作簿的 MOM/YOY留空。销量/销售额空值保留为缺失，已知值合计不把空值当业务零值；加权成交均价仅按同时具备销量和销售额的记录计算并提供覆盖率。BSR 支持数值、整数文本和 `N.0` 文本，非零小数按缺失处理；2026.01-07 缺少竞品名次增强时分析直接失败；并列名次保留，分层数量可能超过区间宽度。当前交付对所有月份统一采用父ASIN优先的确定性代表记录，90页保留原始行数和去重行数。趋势结论与 GENIMO 量化建议统一使用 2026.01-06 实绩。竞品库可由 `src/build_competitor_db.js` 确定性重建。

> 原始需求要求“历年”分析及年度/月度 YoY、MoM，因此 31 张隐藏历史月度表属于有效数据并纳入转换。55 是工作簿内部登记总数，不是 Excel 界面可见子表数；隐藏的空白 `Sheet6` 被跳过。

## 2. 目录结构

```
.
├── .env                       # 实际配置 (gitignore)
├── .env.example               # 配置模板 (提交到 git)
├── .gitignore                 # 屏蔽 *.db, node_modules, .env（data/raw 及 xlsx 数据文件纳入版本控制）
├── README.md                  # 本文件
├── package.json               # npm 元数据
├── docs/
│   ├── SPEC.md                # 唯一事实来源 (schema / 验收标准)
│   ├── TASKS.md               # 任务清单 (全部 [x])
│   └── REVIEWS.md             # 审查草稿 (默认空)
├── data/
│   ├── raw/                             # 主 Excel + 7 份 2026 竞品快照
│   └── processed/                       # market.db + competitor_809440.db
├── src/
│   ├── import_xlsx.js                   # 主 Excel 基础导入
│   ├── build_competitor_db.js           # 竞品快照 -> competitor_809440.db（raw + dedup，可复跑）
│   ├── apply_competitor_2026.js         # 2026.01-07 父体代表行确定性重放
│   ├── analyze_market.js                # JSON / Markdown / HTML 分析生成
│   └── build_formula_market_xlsx.mjs    # 结果 XLSX 生成器（公式/缓存值/自动重算标记）
├── sandbox/
│   └── test_xlsx.js           # PoC (验证 xlsx 能读取)
├── tests/
│   ├── verify.js                        # 两阶段快速验收
│   ├── full_audit.js                    # 主库全量逐单元格审计
│   ├── competitor_audit.js              # 7 份竞品原表/raw/dedup审计
│   ├── analysis_audit.js                # MOM/环比/Top100/分层回勾
│   ├── frontend_audit.js                # HTML表格/图表/文本数字逐项溯源
│   └── overwrite_guard.js               # 覆盖保护
├── 交付/                                # 优化版HTML/Markdown、极速版、数据JSON
├── outputs/20260909-formula-market-analysis/
│   └── 户外地垫市场分析-公式版-20260909.xlsx # 结果与公式集中在一份工作簿
├── tmp/                       # 临时文件 (可随时清空)
└── node_modules/              # 依赖 (gitignore)
```

## 3. 快速开始

```bash
# 1. 安装依赖
npm install

# 2. 运行 PoC (验证 xlsx 能读取)
npm run poc

# 3. 执行完整转换（基础 Excel 导入 + 竞品库重建 + 2026.01-07 竞品代表行替换）
npm run import

# 4. 生成优化版结构化数据、完整 Markdown、极速版 Markdown 和独立 HTML
npm run analyze

# 5. 一次运行全部数据、分析和覆盖保护验收
npm test
```

### 3.1 领导验收结果工作簿

`outputs/20260909-formula-market-analysis/户外地垫市场分析-公式版-20260909.xlsx` 是当前给领导验收的单一文件。旧版结构已经废弃：新版把原始要求逐表展开为 20 个工作表。整体市场只由 PP 管和非PP高客单两类组成；GENIMO 是整体市场中的品牌视角。

工作表映射如下：`00_总览与结论` 导航、关键回勾和公式趋势判断；`01/02/03` 整体市场的月度年度、BSR前100、粗细分层；`04/05/06` PP管同样三张；`07/08/09` 非PP高客单同样三张；`10/11/12` GENIMO品牌份额、进退层、2027链接规划（含链接数、库存、广告、利润及晋级/退出门槛）；`13_周报汇总` 按参考周报表头组织领导可读的市场、分层、品牌和利润摘要；`14_利润测试` 按 PD17 参考表头和公式链支持成本录入、单件利润、利润率、SKU利润及测款动作；`15_经营计划与实际` 复制计划部经营底表并追加 H1 达成、7 月差额和执行状态；`90_输入_明细`、`91_聚合输入`、`92_校验`、`93_来源与规则` 分别记录可追溯输入、中间聚合、数学校验和操作边界。

结果页的销量、销售额、两种均价、覆盖率、MOM（当前月 vs 上一自然月）、YOY（当前月 vs 去年同月）、份额、BSR分层和校验均写成 Excel 公式，并同时写入缓存值。当前工作簿包含 63,130 条明细、1,763 条聚合和 9,479 个公式；2026H1 整体销量 `1,737,288` = PP `984,309` + 非PP `752,979`，销量/销售额差额均为 0；Top100最大Listing数100。核心周期为 2026.01-06，2026.07 是 94 父体展示样本，两种增速留空。`13` 页将计划部 BI/BSR 正向验收基准与 market.db 父体快照并列并标注统计单元；`14` 页保留 PD17 利润测试公式链，成本和 7 天销量未填时保持空值并提示补录。

90 页按父ASIN优先、缺省时用ASIN形成统计键；组内按最小可解析BSR、字段完整度、源行ID确定代表行，同时保留原始行数和去重行数。缺失销量/销售额不按0计入分母；当前分别有 37/120 条缺失，92页标为“请复核”。GENIMO主份额以整体市场为分母，PP内份额只作辅助。

如需重建该文件，使用工作区 Node 运行 `src/build_formula_market_xlsx.mjs`；脚本只读 `market.db` 与 `competitor_809440.db`，生成 90/91 输入及结果公式页，不修改 `data/raw/` 或写回数据库。

## 4. DB Schema 概览

### 4.1 命名规则

- 月度明细表: `monthly_<YYYYMM>` (例如 `monthly_202206`, `monthly_202602`)
- TOP 销量表: `top_sales_volume` / `top_sales_volume_ratio` / `top_total_sales` / `top_avg_price`
- 元数据表: `meta` (记录每次导入)
- 工作表目录: `sheet_catalog`（记录顺序、可见性、有效区域、分类、目标表、行数和跳过原因）
- 替换审计表: `analysis_replacements`（记录 2026.01-07 基础/竞品行数、源库 SHA-256、代表行规则和应用时间）

### 4.2 通用列

每张业务表都额外包含:

- `row_id` (INTEGER PRIMARY KEY AUTOINCREMENT)
- `month_label` (TEXT, 规范化月份标签)

详见 `docs/SPEC.md` 第 3 节。

### 4.3 示例查询

```sql
-- 查 2024 年 5 月销量最高的 10 个商品
SELECT 品牌, 商品标题, 月销量, 月销售额, 价格
FROM monthly_202405
ORDER BY 月销量 DESC
LIMIT 10;

-- 看 meta
SELECT * FROM meta;
```

### 4.4 跨月聚合 (注意: 上架时间列类型不一致, 需 CAST)

```sql
-- 跨月销量聚合 (无类型冲突)
SELECT month_label, SUM(月销量) AS sales, SUM(月销售额) AS revenue
FROM monthly_202206 GROUP BY month_label
UNION ALL
SELECT month_label, SUM(月销量), SUM(月销售额) FROM monthly_202207 GROUP BY month_label
...;

-- 上架时间过滤: 旧 TEXT vs 新 INTEGER Excel 序列号
SELECT 品牌, 商品标题, 上架时间
FROM monthly_202206 WHERE 上架时间 > '2024-01-01'
UNION ALL
SELECT 品牌, 商品标题, 上架时间
FROM monthly_202406 WHERE CAST(上架时间 AS TEXT) > '2024-01-01';

-- 上架日期统一为 ISO 格式
SELECT 品牌, 商品标题, 上架时间,
  CASE
    WHEN typeof(上架时间) = 'integer' THEN date(上架时间 - 25569, 'unixepoch')
    ELSE 上架时间
  END AS 上架日期
FROM monthly_202406;

-- 重量 / 重量_2 (旧 sheet 重复列, 第一个磅, 第二个千克)
SELECT 品牌, 商品标题, 重量 AS 重量_磅, 重量_2 AS 重量_千克
FROM monthly_202206
WHERE 重量 IS NOT NULL LIMIT 10;
```

### 4.5 数据质量注意事项

| 列 | 类型不一致 | 原因 |
|---|---|---|
| `上架时间` | TEXT (旧) vs INTEGER (新, Excel 序列号) | 早期文本, 后期 Excel 日期 |
| `子体销量` / `子体销售额` | TEXT (旧, 多空值) vs INTEGER (新) | 旧 sheet 未填充 |
| `Coupon` | TEXT (旧, 多空值) vs REAL (新) | 同上 |
| `LQS` | INTEGER (45) vs REAL (5 张 2026.3+) | 精度差异 |
| `月销量增长率` | 4 张 (2026.4-2026.7) 缺失 | 源 Excel 无此列 |
| `买家运费` | 仅 26 张 (2025.6+) | 源 Excel 新版加列 |
| `小类BSR` | 部分值含换行 | 单元格多行内容 |
| `卖家信息` | 326/494 行含换行 | 同上 |
| `重量` / `体积` | 旧 sheet 重复列名 | 磅 + 千克 |

详见 `docs/REVIEWS.md` 完整记录.

## 5. 关键决策

| 决策点 | 选型 | 理由 |
|---|---|---|
| DB 引擎 | SQLite | 零依赖、单文件、易 gitignore |
| SQLite 驱动 | Node.js 22.5+ 内建 `node:sqlite` | 避免原生编译 (better-sqlite3 需要 Python + node-gyp) |
| Excel 库 | xlsx (SheetJS) | 成熟、社区大、支持流式 |
| Schema | 动态生成 | 新增 sheet / 新增列无需改脚本 |
| 空值 | 全部 NULL | 避免空字符串歧义 |
| 月份标签 | `2025.6` -> `202506` | 统一规范, 便于 ORDER BY / GROUP BY |

## 6. 验收结果

### 6.1 标准验收 (tests/verify.js)

验证脚本已按当前有效数据源分流：2026.01-07 对 `competitor_809440.db/dedup_YYYYMM`，其余表对主 Excel；正式验收不再保留“预期失败”。

```
========== VERIFY SUMMARY ==========
PASS: 128
FAIL: 0
```

### 6.2 全表逐 cell 校验（`tests/full_audit.js`）

```
Sheets checked: 54
Cells checked: 4,261,173
Structural issues: 0
Value mismatches: 0
Numeric type changes: 69,027 (informational)
```

### 6.3 竞品源链路与分析公式验收

```text
competitor_audit: 7 snapshots / 819,910 cells / 51 checks / 0 failures
analysis_audit: 101 checks / 0 failures
frontend_audit: 43 tables / 32,795 cells / 32 checks / 0 failures
```

竞品审计逐月确认：源快照（2026.01-06 全市场父体级 1135/1039/2002/1745/1691/2002 行，2026.07 子体级 3000 行）与 `raw_YYYYMM` 零值差异、父/ASIN键覆盖完整、1134/1038/1766/1744/1690/1993/94 条 `dedup_YYYYMM` 均为原始表中的精确代表行。代表行规则为：同 Listing 键先取最小可解析小类 BSR；同名次优先月销量/月销售额字段完整行，再按源表顺序稳定决胜。分析审计确认：Top100 每类每月不超过100、头中尾/五档精确回勾、PP+high 精确回勾整体、空价格不按0计入平均标价、核心截止 202606、202602 四个验收值复现、历史 BSR 行代理质量诊断和 2025 父体重复敏感性完整、2026.01-07 整体趋势合并且 7 月跨口径指标禁算、预测范围与来源 SHA-256、可复现时间戳、`meta.db_size_bytes` 均有回归检查。前端审计逐格比对 43 张表共 32,795 个数据单元格，并校验三张 SVG 趋势图、四个月预测区间、四类趋势及所有 2026 洞察卡。

### 6.4 自审修复历史

详见 `docs/REVIEWS.md`. 自审发现并修复了 4 个 bug:

| # | Bug | 修复 | 影响范围 |
|---|---|---|---|
| #1 | TOP 表 headerRow 错位 (1→0) | import_xlsx.js | 4 张 TOP 表共缺 1 行 / 表 + 列名错位 |
| #2 | 空 header 列数据被丢弃 | import_xlsx.js 用 `_col_N` 占位 | TOP销量(倍率) 丢失 3697 cell |
| #3 | verify.js TOP header 检测错 | verify.js | 验证脚本误报 |
| #4 | 抽样对比范围不一致 (前6 vs 前3+后3) | verify.js getDbSample | 验证脚本误报 |

## 7. 性能指标

| 阶段 | 耗时 |
|---|---|
| Excel 加载 (44 MB) | 17.5 s |
| 基础 DB 写入 (73,812 行 × 54 表) | ~3 s |
| 2026.01-07 父体代表行重放 | <1 s |
| 基础导入总耗时 | 约20.2 s |
| 当前有效记录 | 71,451 行 / 4,261,173 数据单元格 |

## 8. 后续可扩展

- **增量更新**: 增量读取新月份 sheet, 跳过已存在表 (本期未实现)
- **索引优化**: 对 `month_label`、`品牌`、`小类BSR` 加索引 (本期未加, 数据量不大)
- **导出 Parquet/CSV**: 用 DuckDB / sqlite-utils 做下游分发
- **API 暴露**: 用 better-sqlite3 + Express 起 RESTful 查询服务

## 9. 规范遵循

| 规范要求 | 落地情况 |
|---|---|
| docs/SPEC.md 唯一事实来源 | ✅ 含数据源 / schema / PoC / 验收 |
| docs/TASKS.md 原子化任务 | ✅ 12 个 Task 全部 [x] |
| docs/REVIEWS.md 审查草稿 | ✅ 已建 (默认空) |
| data/raw（含竞品快照、参考 workbook、周报等 xlsx） | ✅ 纳入版本控制，便于复现与回溯 |
| data/processed/*.db 不入 git | ✅ .gitignore 已配置 |
| .env.example 配置唯一来源 | ✅ 含 RAW_EXCEL_PATH / DATABASE_URL 等 |
| sandbox/ 跑通 PoC | ✅ sandbox/test_xlsx.js 已验证 55 sheet |
| tmp/ 可随时清空 | ✅ 已清空 |
| 验证命令 inline 在每个 Task | ✅ 每个 Task 都标了验证方式 |
