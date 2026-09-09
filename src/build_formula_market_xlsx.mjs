import fs from 'node:fs/promises';
import path from 'node:path';
import { DatabaseSync } from 'node:sqlite';
import { SpreadsheetFile, Workbook } from '@oai/artifact-tool';

const ROOT = path.resolve('.');
const DB_PATH = path.resolve(ROOT, 'data/processed/market.db');
const COMP_DB_PATH = path.resolve(ROOT, 'data/processed/competitor_809440.db');
const JSON_PATH = path.resolve(ROOT, '交付/户外地垫市场分析数据.json');
const OUT_DIR = path.resolve(ROOT, 'outputs/20260909-formula-market-analysis');
const OUT_PATH = path.join(OUT_DIR, '户外地垫市场分析-公式版-20260909.xlsx');
const PREVIEW_PATH = path.join(OUT_DIR, '00_概览-预览.png');
const PLASTIC_WORD_RE = /\bplastic\b/i;
const CATEGORIES = ['overall', 'pp', 'high', 'genimo'];
const CATEGORY_LABELS = { overall: '整体市场', pp: 'PP/Plastic', high: '高客单非PP', genimo: 'GENIMO' };
const TIERS = [
  { key: '头部（1-20）', min: 1, max: 20 },
  { key: '中部（21-50）', min: 21, max: 50 },
  { key: '尾部（51-100）', min: 51, max: 100 },
];
const q = (s) => '"' + String(s).replaceAll('"', '""') + '"';
const txt = (v) => v === null || v === undefined ? '' : String(v);
const num = (v) => v !== null && v !== undefined && v !== '' && Number.isFinite(Number(v)) ? Number(v) : null;
const present = (v) => num(v) !== null;
const keyOf = (r) => txt(r.parent ?? r['父ASIN']).trim() || txt(r.asin ?? r.ASIN).trim() || ('row-' + r.row_id);
function parseBsr(v) {
  if (v === null || v === undefined || v === '') return null;
  const matches = String(v).match(/(?<![\d.,])\d[\d,]*(?:\.0+)?(?![\d.])/g) || [];
  const ranks = matches.map((x) => Number(x.replaceAll(',', ''))).filter((x) => Number.isInteger(x) && x > 0);
  return ranks.length ? Math.min(...ranks) : null;
}

const db = new DatabaseSync(DB_PATH, { readOnly: true });
const cdb = new DatabaseSync(COMP_DB_PATH, { readOnly: true });
const meta = db.prepare('SELECT * FROM meta ORDER BY id DESC LIMIT 1').get() || {};
const replacements = db.prepare('SELECT * FROM analysis_replacements ORDER BY month').all();
const replacementMap = new Map(replacements.map((r) => [String(r.month), r]));
const catalog = db.prepare("SELECT target_table FROM sheet_catalog WHERE classification='monthly' ORDER BY sheet_order").all();
const sourceMonths = catalog.map((r) => String(r.target_table).replace('monthly_', '')).filter((m) => /^\d{6}$/.test(m));
// The leadership-facing workbook focuses on the requested H1 comparison and
// July display-only sample; the full 2022-2026 source remains in market.db.
const months = sourceMonths.filter((m) => m >= '202501' && m <= '202607');

const profilesByMonth = new Map();
for (const month of sourceMonths.filter((m) => m >= '202601' && m <= '202607')) {
  const rawTable = 'raw_' + month;
  if (!cdb.prepare("SELECT 1 FROM sqlite_master WHERE type='table' AND name=?").get(rawTable)) throw new Error('Missing ' + rawTable);
  const profiles = new Map();
  const rows = cdb.prepare('SELECT row_id, ASIN asin, "父ASIN" parent, 品牌 brand, 商品标题 title, 小类BSR bsr FROM ' + q(rawTable)).all();
  for (const row of rows) {
    const key = keyOf(row);
    if (!key || key.startsWith('row-')) continue;
    const p = profiles.get(key) || { plastic: false, genimo: false, ranks: { overall: null, pp: null, high: null, genimo: null } };
    const rank = parseBsr(row.bsr);
    const plastic = PLASTIC_WORD_RE.test(txt(row.title));
    const genimo = txt(row.brand).trim().toLowerCase() === 'genimo';
    p.plastic ||= plastic; p.genimo ||= genimo;
    for (const [f, ok] of [['overall', true], ['pp', plastic], ['high', !plastic], ['genimo', genimo]]) {
      if (ok && rank !== null && (p.ranks[f] === null || rank < p.ranks[f])) p.ranks[f] = rank;
    }
    profiles.set(key, p);
  }
  profilesByMonth.set(month, profiles);
}

const detailRows = [];
for (const month of months) {
  const table = 'monthly_' + month;
  const rows = db.prepare('SELECT row_id, ASIN asin, SKU sku, "父ASIN" parent, 品牌 brand, 商品标题 title, 小类BSR bsr, 月销量 sales, 月销售额 revenue, 价格 price FROM ' + q(table)).all();
  const profiles = profilesByMonth.get(month);
  for (const row of rows) {
    const key = keyOf(row); const profile = profiles?.get(key); const parsed = parseBsr(row.bsr);
    const pp = profile ? profile.plastic : PLASTIC_WORD_RE.test(txt(row.title));
    const genimo = profile ? profile.genimo : txt(row.brand).trim().toLowerCase() === 'genimo';
    const sales = num(row.sales); const revenue = num(row.revenue); const price = num(row.price);
    const ranks = profile?.ranks || {};
    detailRows.push({ month, listingKey: key, parent: txt(row.parent).trim(), asin: txt(row.asin).trim(), brand: txt(row.brand).trim(), title: txt(row.title).trim(),
      ppFlag: pp ? 1 : 0, highFlag: pp ? 0 : 1, genimoFlag: genimo ? 1 : 0, sales, revenue, price,
      rankOverall: ranks.overall ?? parsed, rankPP: ranks.pp ?? (pp ? parsed : null), rankHigh: ranks.high ?? (!pp ? parsed : null), rankGenimo: ranks.genimo ?? (genimo ? parsed : null),
      salesValid: sales === null ? 0 : 1, revenueValid: revenue === null ? 0 : 1, pairedValid: sales !== null && revenue !== null ? 1 : 0, priceValid: price === null ? 0 : 1,
      source: replacementMap.has(month) ? '2026竞品原始镜像→确定性去重代表行' : '主源工作簿月度明细行', sourceRowId: Number(row.row_id), coreFlag: month >= '202601' && month <= '202606' ? 1 : 0, julyFlag: month === '202607' ? 1 : 0,
      topOverall: 0, topPP: 0, topHigh: 0, topGenimo: 0 });
  }
}

for (const month of months) {
  const rows = detailRows.filter((r) => r.month === month);
  for (const c of CATEGORIES) {
    const eligible = rows.filter((r) => {
      const flag = c === 'overall' ? true : c === 'pp' ? r.ppFlag : c === 'high' ? r.highFlag : r.genimoFlag;
      const rank = c === 'overall' ? r.rankOverall : c === 'pp' ? r.rankPP : c === 'high' ? r.rankHigh : r.rankGenimo;
      return flag && Number.isInteger(rank) && rank >= 1 && rank <= 100;
    }).sort((a, b) => {
      const ar = c === 'overall' ? a.rankOverall : c === 'pp' ? a.rankPP : c === 'high' ? a.rankHigh : a.rankGenimo;
      const br = c === 'overall' ? b.rankOverall : c === 'pp' ? b.rankPP : c === 'high' ? b.rankHigh : b.rankGenimo;
      return ar - br || a.listingKey.localeCompare(b.listingKey) || a.sourceRowId - b.sourceRowId;
    }).slice(0, 100);
    const flagName = { overall: 'topOverall', pp: 'topPP', high: 'topHigh', genimo: 'topGenimo' }[c];
    for (const row of eligible) row[flagName] = 1;
  }
}

function stats(rows) {
  const salesRows = rows.filter((r) => present(r.sales)); const revenueRows = rows.filter((r) => present(r.revenue)); const paired = rows.filter((r) => present(r.sales) && present(r.revenue)); const prices = rows.filter((r) => present(r.price));
  return { listingCount: rows.length, salesValueCount: salesRows.length, revenueValueCount: revenueRows.length, pairedCount: paired.length, priceCount: prices.length,
    sales: salesRows.reduce((s, r) => s + r.sales, 0), revenue: revenueRows.reduce((s, r) => s + r.revenue, 0), pairedSales: paired.reduce((s, r) => s + r.sales, 0), pairedRevenue: paired.reduce((s, r) => s + r.revenue, 0), priceSum: prices.reduce((s, r) => s + r.price, 0) };
}
const aggregateRows = [];
for (const month of months) {
  const rows = detailRows.filter((r) => r.month === month);
  for (const c of CATEGORIES) {
    const flagName = { overall: null, pp: 'ppFlag', high: 'highFlag', genimo: 'genimoFlag' }[c];
    const topName = { overall: 'topOverall', pp: 'topPP', high: 'topHigh', genimo: 'topGenimo' }[c];
    const cat = rows.filter((r) => !flagName || r[flagName] === 1);
    const topRows = cat.filter((r) => r[topName] === 1);
    aggregateRows.push({ month, category: c, categoryLabel: CATEGORY_LABELS[c], tier: '全部', ...stats(cat), topCount: 0 });
    aggregateRows.push({ month, category: c, categoryLabel: CATEGORY_LABELS[c], tier: 'Top100', ...stats(topRows), topCount: topRows.length });
    for (const tier of TIERS) {
      const tierRows = topRows.filter((r) => { const rank = c === 'overall' ? r.rankOverall : c === 'pp' ? r.rankPP : c === 'high' ? r.rankHigh : r.rankGenimo; return Number.isInteger(rank) && rank >= tier.min && rank <= tier.max; });
      aggregateRows.push({ month, category: c, categoryLabel: CATEGORY_LABELS[c], tier: tier.key, ...stats(tierRows), topCount: tierRows.length });
    }
  }
  const genimoPpRows = rows.filter((r) => r.genimoFlag === 1 && r.ppFlag === 1);
  aggregateRows.push({ month, category: 'genimo_pp', categoryLabel: 'GENIMO PP', tier: '全部', ...stats(genimoPpRows), topCount: 0 });
}

const ref = JSON.parse(await fs.readFile(JSON_PATH, 'utf8'));
const benchmark = ref.leadershipBenchmark || {};
const inputHeaders = ['月份', 'Listing键', '父ASIN', 'ASIN', '品牌', '商品标题', 'PP标记', '高客单非PP标记', 'GENIMO标记', '月销量', '月销售额', '价格', '整体BSR', 'PP BSR', '高客单非PP BSR', 'GENIMO BSR', '销量有效', '销售额有效', '配对有效', '价格有效', '来源', '源行ID', '核心标记', '2026.07展示标记', '整体Top100', 'PP Top100', '高客单非PP Top100', 'GENIMO Top100'];
const inputValues = detailRows.map((r) => [r.month, r.listingKey, r.parent || null, r.asin || null, r.brand || null, r.title || null, r.ppFlag, r.highFlag, r.genimoFlag, r.sales, r.revenue, r.price, r.rankOverall, r.rankPP, r.rankHigh, r.rankGenimo, r.salesValid, r.revenueValid, r.pairedValid, r.priceValid, r.source, r.sourceRowId, r.coreFlag, r.julyFlag, r.topOverall, r.topPP, r.topHigh, r.topGenimo]);

const wb = Workbook.create();
const S = (n) => wb.worksheets.add(n);
const overview = S('00_概览'), monthly = S('01_整体月度'), top = S('02_BSR整体'), tiers = S('03_BSR头中尾'), cats = S('04_分类月度'), brands = S('05_品牌份额'), decisions = S('06_决策建议'), core = S('07_核心汇总'), checks = S('08_校验'), input = S('90_输入_明细'), agg = S('91_聚合输入'), rules = S('92_来源与规则');
for (const sh of [overview, monthly, top, tiers, cats, brands, decisions, core, checks, input, agg, rules]) sh.showGridLines = false;
const navy = '#1F4E78', blue = '#5B9BD5', yellow = '#FFF2CC', green = '#E2F0D9', red = '#FCE4D6';
function title(sh, range, value) { sh.mergeCells(range); sh.getRange(range.split(':')[0]).values = [[value]]; sh.getRange(range).format = { fill: navy, font: { bold: true, color: '#FFFFFF', size: 15 }, rowHeight: 28 }; }
function hdr(sh, range) { sh.getRange(range).format = { fill: blue, font: { bold: true, color: '#FFFFFF' }, wrapText: true, rowHeight: 30, borders: { preset: 'all', style: 'thin', color: '#D9E2F3' } }; }
function widths(sh, arr, rows) { arr.forEach((w, i) => sh.getRangeByIndexes(0, i, rows, 1).format.columnWidth = w); }
function body(sh, range, fmts = {}) { sh.getRange(range).format.borders = { insideHorizontal: { style: 'thin', color: '#E7E6E6' } }; for (const [r, f] of Object.entries(fmts)) sh.getRange(r).format.numberFormat = f; }

title(input, 'A1:AB1', '原始明细输入（值）· 结果页由公式引用聚合输入'); input.getRange('A2:AB2').values = [inputHeaders]; hdr(input, 'A2:AB2');
for (let i = 0; i < inputValues.length; i += 5000) { const part = inputValues.slice(i, i + 5000); input.getRange(`A${i + 3}:AB${i + 2 + part.length}`).values = part; }
const inputEnd = inputValues.length + 2; input.freezePanes.freezeRows(2); input.freezePanes.freezeColumns(6); input.getRange(`J3:K${inputEnd}`).format.numberFormat = '#,##0;[Red]-#,##0'; input.getRange(`L3:L${inputEnd}`).format.numberFormat = '$#,##0.00;[Red]-$#,##0.00'; widths(input, [10, 17, 15, 14, 16, 52, 9, 12, 10, 12, 14, 11, 10, 9, 13, 11, 9, 10, 10, 9, 28, 9, 9, 12, 11, 9, 14, 12], inputEnd + 1);

title(agg, 'A1:Q1', '聚合输入（由脚本按明细与同一规则预汇总，供Excel公式重算）');
agg.getRange('A2:Q2').values = [['月份', '分类键', '分类', '层级', 'Listing行数', '销量', '销售额', '价格合计', '价格有效数', '配对销量', '配对销售额', '销量有效数', '销售额有效数', '配对有效数', 'Top100行数', '来源', '说明']]; hdr(agg, 'A2:Q2');
const aggValues = aggregateRows.map((r) => [r.month, r.category, r.categoryLabel, r.tier, r.listingCount, r.sales, r.revenue, r.priceSum, r.priceCount, r.pairedSales, r.pairedRevenue, r.salesValueCount, r.revenueValueCount, r.pairedCount, r.topCount, replacementMap.has(r.month) ? '竞品替换链' : '主源明细', '明细值已保留在90_输入_明细；本页仅为可追溯的中间聚合']);
const aggEnd = aggValues.length + 2; agg.getRange(`A3:Q${aggEnd}`).values = aggValues; agg.getRange(`E3:G${aggEnd}`).format.numberFormat = '#,##0;[Red]-#,##0'; agg.getRange(`H3:K${aggEnd}`).format.numberFormat = '$#,##0.00;[Red]-$#,##0.00'; widths(agg, [10, 11, 16, 15, 12, 13, 15, 13, 12, 13, 15, 12, 13, 12, 11, 16, 48], aggEnd + 1); agg.freezePanes.freezeRows(2);

const AR = `'91_聚合输入'!$A$3:$A$${aggEnd}`, BR = `'91_聚合输入'!$B$3:$B$${aggEnd}`, DR = `'91_聚合输入'!$D$3:$D$${aggEnd}`;
function aSum(col, mcell, cat, tier) { return `SUMIFS('91_聚合输入'!$${col}$3:$${col}$${aggEnd},${AR},${mcell},${BR},"${cat}",${DR},"${tier}")`; }
function aCore(col, cat, start, end) { return `SUMIFS('91_聚合输入'!$${col}$3:$${col}$${aggEnd},${AR},">=${start}",${AR},"<=${end}",${BR},"${cat}",${DR},"全部")`; }
function aPrice(mcell, cat, tier) { return `IFERROR(${aSum('H', mcell, cat, tier)}/${aSum('I', mcell, cat, tier)},"")`; }
function aWeighted(mcell, cat, tier) { return `IFERROR(${aSum('K', mcell, cat, tier)}/${aSum('J', mcell, cat, tier)},"")`; }
function corePrice(cat, start, end) { return `IFERROR(${aCore('H', cat, start, end)}/${aCore('I', cat, start, end)},"")`; }
function coreWeighted(cat, start, end) { return `IFERROR(${aCore('K', cat, start, end)}/${aCore('J', cat, start, end)},"")`; }

title(monthly, 'A1:O1', '整体市场月度汇总 · 公式页'); monthly.mergeCells('A2:O2'); monthly.getRange('A2').values = [['整体市场=全部产品；MOM/环比=当前月 vs 去年同月；2026.07仅展示。']]; monthly.getRange('A2:O2').format = { fill: yellow, font: { italic: true, color: '#7F6000' }, wrapText: true }; monthly.getRange('A4:O4').values = [['月份', '范围', 'Listing行数', '销量', '销售额', '平均标价', '配对销量', '配对销售额', '加权成交均价', '销量覆盖率', '销售额覆盖率', '配对覆盖率', 'MOM基准月', '销量MOM/环比', '销售额MOM/环比']]; hdr(monthly, 'A4:O4');
const ms = 5, me = ms + months.length - 1; monthly.getRange(`A${ms}:B${me}`).values = months.map((m) => [m, m <= '202606' ? '核心' : (m === '202607' ? '展示' : '历史')]);
for (let i = 0; i < months.length; i++) { const r = ms + i, m = `A${r}`, basis = i >= 12 ? months[i - 12] : ''; monthly.getRange(`C${r}:L${r}`).formulas = [[`=${aSum('E', m, 'overall', '全部')}`, `=${aSum('F', m, 'overall', '全部')}`, `=${aSum('G', m, 'overall', '全部')}`, `=${aPrice(m, 'overall', '全部')}`, `=${aSum('J', m, 'overall', '全部')}`, `=${aSum('K', m, 'overall', '全部')}`, `=${aWeighted(m, 'overall', '全部')}`, `=IFERROR(${aSum('L', m, 'overall', '全部')}/C${r},"")`, `=IFERROR(${aSum('M', m, 'overall', '全部')}/C${r},"")`, `=IFERROR(${aSum('N', m, 'overall', '全部')}/C${r},"")`]]; monthly.getRange(`M${r}`).values = [[basis]]; monthly.getRange(`N${r}:O${r}`).formulas = [[basis ? `=IFERROR(D${r}/SUMIFS($D$${ms}:$D$${me},$A$${ms}:$A$${me},M${r})-1,"")` : '=""', basis ? `=IFERROR(E${r}/SUMIFS($E$${ms}:$E$${me},$A$${ms}:$A$${me},M${r})-1,"")` : '=""']]; }
body(monthly, `A${ms}:O${me}`, { [`C${ms}:E${me}`]: '#,##0;[Red]-#,##0', [`F${ms}:I${me}`]: '$#,##0.00;[Red]-$#,##0.00', [`J${ms}:L${me}`]: '0.0%', [`N${ms}:O${me}`]: '0.0%;[Red]-0.0%' }); widths(monthly, [11, 9, 12, 13, 15, 12, 13, 15, 14, 12, 12, 12, 12, 13, 13], me + 1); monthly.freezePanes.freezeRows(4);

function writeCatSheet() { title(cats, 'A1:M1', '分类月度汇总 · 公式页'); cats.mergeCells('A2:M2'); cats.getRange('A2').values = [['PP=标题完整单词plastic；高客单非PP=排除PP后的全部产品；GENIMO份额在05页以整体为分母。']]; cats.getRange('A2:M2').format = { fill: yellow, font: { italic: true, color: '#7F6000' }, wrapText: true }; cats.getRange('A4:M4').values = [['月份', '分类', '范围', 'Listing行数', '销量', '销售额', '平均标价', '加权成交均价', '整体销量份额', '整体销售额份额', 'MOM基准月', '销量MOM/环比', '销售额MOM/环比']]; hdr(cats, 'A4:M4'); const start = 5, rows = []; for (const m of months) for (const c of CATEGORIES) rows.push([m, CATEGORY_LABELS[c], m <= '202606' ? '核心' : (m === '202607' ? '展示' : '历史')]); const end = start + rows.length - 1; cats.getRange(`A${start}:C${end}`).values = rows; for (let i = 0; i < rows.length; i++) { const r = start + i, m = `A${r}`, c = CATEGORIES[i % 4], basis = months.indexOf(rows[i][0]) >= 12 ? months[months.indexOf(rows[i][0]) - 12] : ''; cats.getRange(`D${r}:J${r}`).formulas = [[`=${aSum('E', m, c, '全部')}`, `=${aSum('F', m, c, '全部')}`, `=${aSum('G', m, c, '全部')}`, `=${aPrice(m, c, '全部')}`, `=${aWeighted(m, c, '全部')}`, c === 'overall' ? '=""' : `=IFERROR(E${r}/SUMIFS($E$${start}:$E$${end},$A$${start}:$A$${end},A${r},$B$${start}:$B$${end},"整体市场"),"")`, c === 'overall' ? '=""' : `=IFERROR(F${r}/SUMIFS($F$${start}:$F$${end},$A$${start}:$A$${end},A${r},$B$${start}:$B$${end},"整体市场"),"")`]]; cats.getRange(`K${r}`).values = [[basis]]; cats.getRange(`L${r}:M${r}`).formulas = [[basis ? `=IFERROR(E${r}/SUMIFS($E$${start}:$E$${end},$A$${start}:$A$${end},K${r},$B$${start}:$B$${end},B${r})-1,"")` : '=""', basis ? `=IFERROR(F${r}/SUMIFS($F$${start}:$F$${end},$A$${start}:$A$${end},K${r},$B$${start}:$B$${end},B${r})-1,"")` : '=""']]; } body(cats, `A${start}:M${end}`, { [`D${start}:F${end}`]: '#,##0;[Red]-#,##0', [`G${start}:H${end}`]: '$#,##0.00;[Red]-$#,##0.00', [`I${start}:J${end}`]: '0.0%', [`L${start}:M${end}`]: '0.0%;[Red]-0.0%' }); widths(cats, [11, 16, 9, 12, 13, 15, 12, 14, 13, 14, 12, 13, 13], end + 1); cats.freezePanes.freezeRows(4); return { start, end }; }
const catSheet = writeCatSheet();

function writeTopSheet() { title(top, 'A1:L1', 'BSR Top100 整体汇总 · 公式页'); top.mergeCells('A2:L2'); top.getRange('A2').values = [['每类别、每月独立从rank 1-100池截取不超过100条；2026.07为94父体样本，仅展示。']]; top.getRange('A2:L2').format = { fill: yellow, font: { italic: true, color: '#7F6000' }, wrapText: true }; top.getRange('A4:L4').values = [['月份', '分类', '范围', 'Top100行数', '销量', '销售额', '平均标价', '加权成交均价', '配对覆盖率', 'MOM基准月', '销量MOM/环比', '销售额MOM/环比']]; hdr(top, 'A4:L4'); const start = 5, rows = []; for (const m of months) for (const c of CATEGORIES) rows.push([m, CATEGORY_LABELS[c], m <= '202606' ? '核心' : (m === '202607' ? '展示' : '历史')]); const end = start + rows.length - 1; top.getRange(`A${start}:C${end}`).values = rows; for (let i = 0; i < rows.length; i++) { const r = start + i, m = `A${r}`, c = CATEGORIES[i % 4], basis = months.indexOf(rows[i][0]) >= 12 ? months[months.indexOf(rows[i][0]) - 12] : ''; top.getRange(`D${r}:I${r}`).formulas = [[`=${aSum('O', m, c, 'Top100')}`, `=${aSum('F', m, c, 'Top100')}`, `=${aSum('G', m, c, 'Top100')}`, `=${aPrice(m, c, 'Top100')}`, `=${aWeighted(m, c, 'Top100')}`, `=IFERROR(${aSum('N', m, c, 'Top100')}/D${r},"")`]]; top.getRange(`J${r}`).values = [[basis]]; top.getRange(`K${r}:L${r}`).formulas = [[basis ? `=IFERROR(E${r}/SUMIFS($E$${start}:$E$${end},$A$${start}:$A$${end},J${r},$B$${start}:$B$${end},B${r})-1,"")` : '=""', basis ? `=IFERROR(F${r}/SUMIFS($F$${start}:$F$${end},$A$${start}:$A$${end},J${r},$B$${start}:$B$${end},B${r})-1,"")` : '=""']]; } body(top, `A${start}:L${end}`, { [`D${start}:F${end}`]: '#,##0;[Red]-#,##0', [`G${start}:H${end}`]: '$#,##0.00;[Red]-$#,##0', [`I${start}:I${end}`]: '0.0%', [`K${start}:L${end}`]: '0.0%;[Red]-0.0%' }); widths(top, [11, 16, 9, 12, 13, 15, 12, 14, 13, 12, 13, 13], end + 1); top.freezePanes.freezeRows(4); return { start, end }; }
const topSheet = writeTopSheet();

function writeTierSheet() { title(tiers, 'A1:M1', 'BSR Top100 头部/中部/尾部 · 公式页'); tiers.mergeCells('A2:M2'); tiers.getRange('A2').values = [['头部=1-20；中部=21-50；尾部=51-100。分层使用同一Top100池，合计可回到Top100。']]; tiers.getRange('A2:M2').format = { fill: yellow, font: { italic: true, color: '#7F6000' }, wrapText: true }; tiers.getRange('A4:M4').values = [['月份', '分类', '分层', '范围', 'Listing行数', '销量', '销售额', '平均标价', '加权成交均价', 'Top100销量份额', 'Top100销售额份额', 'MOM基准月', '销量MOM/环比']]; hdr(tiers, 'A4:M4'); const start = 5, rows = []; for (const m of months) for (const c of CATEGORIES) for (const t of TIERS) rows.push([m, CATEGORY_LABELS[c], t.key, m <= '202606' ? '核心' : (m === '202607' ? '展示' : '历史')]); const end = start + rows.length - 1; tiers.getRange(`A${start}:D${end}`).values = rows; for (let i = 0; i < rows.length; i++) { const r = start + i, m = `A${r}`, c = CATEGORIES[Math.floor(i / 3) % 4], t = TIERS[i % 3], basis = months.indexOf(rows[i][0]) >= 12 ? months[months.indexOf(rows[i][0]) - 12] : ''; tiers.getRange(`E${r}:K${r}`).formulas = [[`=${aSum('O', m, c, t.key)}`, `=${aSum('F', m, c, t.key)}`, `=${aSum('G', m, c, t.key)}`, `=${aPrice(m, c, t.key)}`, `=${aWeighted(m, c, t.key)}`, `=IFERROR(F${r}/${aSum('F',m,c,'Top100')},"")`, `=IFERROR(G${r}/${aSum('G',m,c,'Top100')},"")`]]; tiers.getRange(`L${r}`).values = [[basis]]; tiers.getRange(`M${r}`).formulas = [[basis ? `=IFERROR(F${r}/SUMIFS($F$${start}:$F$${end},$A$${start}:$A$${end},L${r},$B$${start}:$B$${end},B${r},$C$${start}:$C$${end},C${r})-1,"")` : '=""']]; } body(tiers, `A${start}:M${end}`, { [`E${start}:G${end}`]: '#,##0;[Red]-#,##0', [`H${start}:I${end}`]: '$#,##0.00;[Red]-$#,##0', [`J${start}:K${end}`]: '0.0%', [`M${start}:M${end}`]: '0.0%;[Red]-0.0%' }); widths(tiers, [11, 16, 15, 9, 12, 13, 15, 12, 14, 15, 16, 12, 13], end + 1); tiers.freezePanes.freezeRows(4); return { start, end }; }
const tierSheet = writeTierSheet();

title(brands, 'A1:J1', 'GENIMO 品牌份额 · 公式页'); brands.mergeCells('A2:J2'); brands.getRange('A2').values = [['主指标=GENIMO/整体市场；PP内份额仅作辅助。']]; brands.getRange('A2:J2').format = { fill: yellow, font: { italic: true, color: '#7F6000' }, wrapText: true }; brands.getRange('A4:J4').values = [['月份', '整体销量', 'GENIMO销量', '整体销量份额', '整体销售额', 'GENIMO销售额', '整体销售额份额', 'PP销量', 'GENIMO PP销量', 'PP内销量份额']]; hdr(brands, 'A4:J4'); const bs = 5, be = bs + months.length - 1; brands.getRange(`A${bs}:A${be}`).values = months.map((m) => [m]); for (let i = 0; i < months.length; i++) { const r = bs + i, m = `A${r}`; brands.getRange(`B${r}:J${r}`).formulas = [[`=${aSum('F',m,'overall','全部')}`, `=${aSum('F',m,'genimo','全部')}`, `=IFERROR(C${r}/B${r},"")`, `=${aSum('G',m,'overall','全部')}`, `=${aSum('G',m,'genimo','全部')}`, `=IFERROR(F${r}/E${r},"")`, `=${aSum('F',m,'pp','全部')}`, `=${aSum('F',m,'genimo_pp','全部')}`, `=IFERROR(I${r}/H${r},"")`]]; } body(brands, `A${bs}:J${be}`, { [`B${bs}:C${be}`]: '#,##0;[Red]-#,##0', [`D${bs}:D${be}`]: '0.0%', [`E${bs}:F${be}`]: '$#,##0;[Red]-#,##0', [`G${bs}:G${be}`]: '0.0%', [`H${bs}:I${be}`]: '#,##0;[Red]-#,##0', [`J${bs}:J${be}`]: '0.0%' }); widths(brands, [11, 14, 14, 14, 15, 16, 15, 12, 15, 14], be + 1); brands.freezePanes.freezeRows(4);

title(core, 'A1:M1', '核心周期汇总（2025.01-06 vs 2026.01-06）· 公式页'); core.mergeCells('A2:M2'); core.getRange('A2').values = [['核心所有指标从91_聚合输入公式计算；2026.07不参与核心周期。']]; core.getRange('A2:M2').format = { fill: yellow, font: { italic: true, color: '#7F6000' }, wrapText: true }; core.getRange('A4:M4').values = [['分类', '2025H1销量', '2026H1销量', '销量MOM/环比', '2025H1销售额', '2026H1销售额', '销售额MOM/环比', '2025H1平均标价', '2026H1平均标价', '标价MOM/环比', '2025H1加权成交均价', '2026H1加权成交均价', '加权均价MOM/环比']]; hdr(core, 'A4:M4'); const cs = 5, ce = cs + 3; core.getRange(`A${cs}:A${ce}`).values = CATEGORIES.map((c) => [CATEGORY_LABELS[c]]); for (let i = 0; i < CATEGORIES.length; i++) { const r = cs + i, c = CATEGORIES[i]; core.getRange(`B${r}:M${r}`).formulas = [[`=${aCore('F',c,'202501','202506')}`, `=${aCore('F',c,'202601','202606')}`, `=IFERROR(C${r}/B${r}-1,"")`, `=${aCore('G',c,'202501','202506')}`, `=${aCore('G',c,'202601','202606')}`, `=IFERROR(F${r}/E${r}-1,"")`, `=${corePrice(c,'202501','202506')}`, `=${corePrice(c,'202601','202606')}`, `=IFERROR(I${r}/H${r}-1,"")`, `=${coreWeighted(c,'202501','202506')}`, `=${coreWeighted(c,'202601','202606')}`, `=IFERROR(L${r}/K${r}-1,"")`]]; } body(core, `A${cs}:M${ce}`, { [`B${cs}:C${ce}`]: '#,##0;[Red]-#,##0', [`D${cs}:D${ce}`]: '0.0%', [`E${cs}:F${ce}`]: '$#,##0;[Red]-$#,##0', [`G${cs}:G${ce}`]: '0.0%', [`H${cs}:L${ce}`]: '$#,##0.00;[Red]-$#,##0.00', [`J${cs}:J${ce}`]: '0.0%', [`M${cs}:M${ce}`]: '0.0%' }); widths(core, [16, 14, 14, 14, 15, 15, 15, 15, 15, 14, 18, 18, 16], ce + 1);

title(checks, 'A1:F1', '公式校验与数据完整性'); checks.getRange('A2:F2').values = [['校验项', '结果', '期望/阈值', '实际值', '是否通过', '说明']]; hdr(checks, 'A2:F2'); const checkRows = [['整体=PP+高客单非PP（2026H1销量）', null, 0, null, null, '分类加总必须回到整体。'], ['整体=PP+高客单非PP（2026H1销售额）', null, 0, null, null, '销售额也必须回到整体。'], ['BSR Top100每类别每月≤100', null, 100, null, null, '每月每类独立封顶100。'], ['核心明细行数（2026.01-06）', null, 9365, null, null, '控制值来自当前分析 JSON。'], ['明细总行数（本工作簿输入）', null, detailRows.length, null, null, '90_输入_明细的有效记录数。'], ['缺失销量记录数', null, null, null, null, '缺失值保留为缺失，不按0计入。'], ['缺失销售额记录数', null, null, null, null, '缺失值保留为缺失，不按0计入。']]; checks.getRange('A3:F9').values = checkRows;
checks.getRange('B3:B9').formulas = [[`='07_核心汇总'!C5-('07_核心汇总'!C6+'07_核心汇总'!C7)`], [`='07_核心汇总'!F5-('07_核心汇总'!F6+'07_核心汇总'!F7)`], [`=MAX('02_BSR整体'!D5:D${topSheet.end})`], [`=COUNTIFS('90_输入_明细'!$A$3:$A$${inputEnd},">=202601",'90_输入_明细'!$A$3:$A$${inputEnd},"<=202606")`], [`=COUNTA('90_输入_明细'!$A$3:$A$${inputEnd})`], [`=COUNTIF('90_输入_明细'!$Q$3:$Q$${inputEnd},0)`], [`=COUNTIF('90_输入_明细'!$R$3:$R$${inputEnd},0)`]];
checks.getRange('D3:D9').formulas = [['=ABS(B3)'], ['=ABS(B4)'], ['=B5'], ['=B6'], ['=B7'], ['=B8'], ['=B9']]; checks.getRange('E3:E9').formulas = [['=IF(D3<0.5,"通过","失败")'], ['=IF(D4<0.5,"通过","失败")'], ['=IF(D5<=C5,"通过","失败")'], ['=IF(D6=C6,"通过","请复核")'], ['=IF(D7=C7,"通过","失败")'], ['=IF(D8>=0,"通过","失败")'], ['=IF(D9>=0,"通过","失败")']]; checks.getRange('E3:E9').conditionalFormats.add('containsText', { text: '失败', format: { fill: red, font: { color: '#9C0006', bold: true } } }); checks.getRange('E3:E9').conditionalFormats.add('containsText', { text: '通过', format: { fill: green, font: { color: '#006100', bold: true } } }); body(checks, 'A3:F9'); widths(checks, [34, 15, 14, 15, 12, 56], 10);

title(decisions, 'A1:E1', '经营决策提示 · 公式联动'); decisions.mergeCells('A2:E2'); decisions.getRange('A2').values = [['行动提示来自核心汇总、分层和品牌份额；链接数量需输入目标后再分配，预测不等于历史实绩。']]; decisions.getRange('A2:E2').format = { fill: yellow, font: { italic: true, color: '#7F6000' }, wrapText: true }; decisions.getRange('A4:E4').values = [['主题', '公式结果', '动态提示', '建议动作', '限制/复核点']]; hdr(decisions, 'A4:E4'); decisions.getRange('A5:E10').values = [['整体市场销量方向', null, null, '整体结论先看全市场销量和销售额，再看分类、Top100和分层。', '2025行级与2026父体快照统计单元不同，市场DB结果只作方向性参考。'], ['PP销量/销售额方向', null, null, '销量和销售额若均下降，先拆价格、转化、广告和SKU结构；不要用单一均价解释。', '加权成交均价=配对销售额÷配对销量；平均标价是另一指标。'], ['高客单非PP方向', null, null, '高客单非PP是排除PP后的全部剩余产品，用整体份额与层级决定投入。', '分类必须与整体加总一致。'], ['BSR层级优先测试', null, null, '优先测试销量/销售额同比为正且份额扩大的层级；尾部若低基数，先小额验证。', '历史行级Top100可能有重复Listing。'], ['GENIMO整体市场份额', null, null, '以整体市场为分母监控份额；下滑时优先补强有增长层级的中流链接。', 'PP内份额只作辅助。'], ['链接规划输入', 0, null, '在B10输入计划链接总数，再按核心销量份额分配头、中、尾链接。', '0表示尚未给定目标；还需复核利润、库存、广告容量。']]; decisions.getRange('B5:B9').formulas = [[`='07_核心汇总'!D5`], [`='07_核心汇总'!D6`], [`='07_核心汇总'!D7`], [`=MAX('03_BSR头中尾'!M5:M${tierSheet.end})`], [`=IFERROR('07_核心汇总'!C8/'07_核心汇总'!C5,"")`]]; decisions.getRange('C5:C10').formulas = [['=IF(B5>0,"2026H1销量较2025H1为正","2026H1销量较2025H1为负或不可比")'], ['=IF(B6>0,"PP销量增长","PP销量下降；均价需分开看")'], ['=IF(B7>0,"高客单非PP销量增长","高客单非PP销量下降")'], ['=IF(B8>0,"至少一个层级出现正向分层MOM，优先做小额测款","未发现正向层级，先查数据与样本")'], ['=IF(B9>=0.2,"GENIMO整体份额达到20%或以上","GENIMO整体份额低于20%，以份额防守和中流链接为先")'], ['=IF(B10>0,"已输入目标，可按份额分配","请输入计划链接总数后再计算分配")']]; decisions.getRange('B10').format = { fill: yellow, font: { bold: true, color: '#7F6000' }, numberFormat: '#,##0' }; body(decisions, 'A5:E10'); decisions.getRange('B5:B9').format.numberFormat = '0.0%;[Red]-0.0%'; widths(decisions, [24, 18, 44, 60, 58], 12);

title(overview, 'A1:H1', '户外地垫市场分析 · 公式版结果'); overview.mergeCells('A2:H2'); overview.getRange('A2').values = [['领导验收先看本页；参考 workbook 只作独立参考，不作为本项目明细真值。']]; overview.getRange('A2:H2').format = { fill: yellow, font: { italic: true, color: '#7F6000' }, wrapText: true }; overview.getRange('A4:H4').values = [['核心指标', '2025H1', '2026H1', 'MOM/环比', '数据状态', '公式/来源', '领导验收参考', '参考结果']]; hdr(overview, 'A4:H4'); overview.getRange('A5:H10').values = [['整体市场销量', null, null, null, 'market.db明细参考', '07_核心汇总!B5:D5', '计划部BI全类目销量', (benchmark.industry?.growthPct ?? null) / 100], ['整体市场销售额', null, null, null, 'market.db明细参考', '07_核心汇总!E5:G5', '计划部未提供销售额', null], ['PP销量', null, null, null, 'market.db分类计算', '07_核心汇总!B6:D6', '分类拆解', null], ['高客单非PP销量', null, null, null, 'market.db分类计算', '07_核心汇总!B7:D7', '分类拆解', null], ['GENIMO整体销量份额', null, null, null, '整体市场为分母', '05_品牌份额', '品牌主口径', null], ['BSR Top100整体销量', null, null, null, '独立Top100池', '02_BSR整体', '计划部BSR参考', (benchmark.bsrTop100?.growthPct ?? null) / 100]]; overview.getRange('B5:D8').formulas = [[`='07_核心汇总'!B5`,`='07_核心汇总'!C5`,`='07_核心汇总'!D5`],[`='07_核心汇总'!E5`,`='07_核心汇总'!F5`,`='07_核心汇总'!G5`],[`='07_核心汇总'!B6`,`='07_核心汇总'!C6`,`='07_核心汇总'!D6`],[`='07_核心汇总'!B7`,`='07_核心汇总'!C7`,`='07_核心汇总'!D7`]]; overview.getRange('B9:C9').formulas = [[`=IFERROR('07_核心汇总'!B8/'07_核心汇总'!B5,"")`,`=IFERROR('07_核心汇总'!C8/'07_核心汇总'!C5,"")`]]; overview.getRange('D9').formulas = [['=IFERROR(C9/B9-1,"")']]; overview.getRange('B10:D10').formulas = [[`=SUMIFS('02_BSR整体'!$E$${topSheet.start}:$E$${topSheet.end},'02_BSR整体'!$A$${topSheet.start}:$A$${topSheet.end},">=202501",'02_BSR整体'!$A$${topSheet.start}:$A$${topSheet.end},"<=202506",'02_BSR整体'!$B$${topSheet.start}:$B$${topSheet.end},"整体市场")`, `=SUMIFS('02_BSR整体'!$E$${topSheet.start}:$E$${topSheet.end},'02_BSR整体'!$A$${topSheet.start}:$A$${topSheet.end},">=202601",'02_BSR整体'!$A$${topSheet.start}:$A$${topSheet.end},"<=202606",'02_BSR整体'!$B$${topSheet.start}:$B$${topSheet.end},"整体市场")`, '=IFERROR(C10/B10-1,"")']]; overview.getRange('H5:H10').format.numberFormat = '0.0%;[Red]-0.0%'; overview.getRange('B5:C8').format.numberFormat = '#,##0;[Red]-#,##0'; overview.getRange('D5:D8').format.numberFormat = '0.0%;[Red]-0.0%'; overview.getRange('B9:D9').format.numberFormat = '0.0%;[Red]-0.0%'; overview.getRange('B10:C10').format.numberFormat = '#,##0;[Red]-#,##0'; overview.getRange('D10').format.numberFormat = '0.0%;[Red]-0.0%'; overview.getRange('A12:F12').values = [['关键逻辑校验', '结果', '阈值/基准', '实际差额', '状态', '说明']]; hdr(overview, 'A12:F12'); overview.getRange('A13:F16').values = [['整体=PP+高客单非PP（销量）', null, 0, null, null, '分类加总必须回到整体。'], ['整体=PP+高客单非PP（销售额）', null, 0, null, null, '销售额同样复核。'], ['BSR Top100最大行数', null, 100, null, null, '每类别每月不超过100。'], ['核心明细行数', null, 9365, null, null, '与当前JSON控制值对比。']]; overview.getRange('B13:B16').formulas = [[`='08_校验'!B3`],[`='08_校验'!B4`],[`='08_校验'!B5`],[`='08_校验'!B6`]]; overview.getRange('D13:D16').formulas = [[`='08_校验'!D3`],[`='08_校验'!D4`],[`='08_校验'!D5`],[`='08_校验'!D6`]]; overview.getRange('E13:E16').formulas = [[`='08_校验'!E3`],[`='08_校验'!E4`],[`='08_校验'!E5`],[`='08_校验'!E6`]]; overview.getRange('A18:H18').values = [['当前阅读顺序', null, null, null, null, null, null, null]]; overview.mergeCells('A18:H18'); overview.getRange('A18:H18').format = { fill: navy, font: { bold: true, color: '#FFFFFF' } }; overview.getRange('A19:H22').values = [['1', '先看整体与分类加总，再看02/03的Top100和头中尾。', null, null, null, null, null, null], ['2', '如果两个分类均为负而整体为正，先查看08_校验，不能用均价上升掩盖销量/销售额矛盾。', null, null, null, null, null, null], ['3', 'GENIMO主份额在05页以整体市场为分母；PP内份额是辅助指标。', null, null, null, null, null, null], ['4', '2026.01-06是核心实绩；2026.07为94父体展示样本。', null, null, null, null, null, null]]; for (let r = 19; r <= 22; r++) overview.mergeCells(`B${r}:H${r}`); overview.getRange('A19:H22').format = { wrapText: true, fill: '#F7FBFF', borders: { insideHorizontal: { style: 'thin', color: '#D9E2F3' } } }; widths(overview, [28, 15, 15, 14, 18, 28, 22, 16], 24);

title(rules, 'A1:H1', '来源、数据操作和规则登记'); rules.getRange('A3:H3').values = [['项目', '当前值/规则', '来源文件/表', '计算或操作', '是否写入原始数据', '可复核位置', '限制', '备注']]; hdr(rules, 'A3:H3'); const ruleRows = [['主源工作簿', meta.source_file || '地垫-卖家精灵市场数据.xlsx', 'data/raw/地垫-卖家精灵市场数据.xlsx', '只读导入market.db；本次不改写源文件', '否', '90_输入_明细', '2025行级含变体；与2026父体快照统计单元不同', '缺失保留为空'], ['2026替换链', '2026.01-07', 'Competitor-US-2026.*.xlsx + competitor_809440.db', '保留raw镜像；父ASIN优先、缺省ASIN确定性去重；代表行进入market.db', '否（仅派生DB）', 'analysis_replacements + 90/91页', '2026.07为94父体样本', '当前结果使用替换后输入'], ['去重逻辑', '父ASIN优先，缺省ASIN；最小可解析小类BSR；同名次优先指标完整行；再按源行', 'src/build_competitor_db.js', '只对2026替换快照执行；历史行级不自动父体去重', '否', '90页Listing键 / 92页说明', '历史父体语义待业务确认', '原始raw仍保留'], ['BSR解析', '正整数；允许N.0；拒绝非零小数；多值取最小正整数', 'src/analyze_market.js / src/build_competitor_db.js', '2026类别排名使用同父体变体最佳小类BSR；Top100稳定截取100', '否', '90页M:P / 91页层级', '历史排名可能重复Listing', '分层不跨池'], ['分类', 'PP=标题完整单词plastic；高客单非PP=排除PP后的全部产品；GENIMO=品牌', 'SPEC 1.1/7.5', '类别结果由91页公式引用；PP+高客单必须回到整体', '否', '04/07/08页', '不要将高客单缩成丙纶/三明治', '品牌主份额用整体分母'], ['均价', '平均标价 + 配对加权成交均价', 'SPEC 1.2', '平均标价=价格合计/有效价格数；加权成交均价=配对销售额/配对销量', '否', '01/02/03/04/07公式', '缺失值不按0进入分母', '二者不混称客单价'], ['MOM/环比', '当前月 vs 去年同月', 'SPEC 0.3/1.2', '例如2026.02 vs 2025.02；不展示本月vs上月连续环比', '否', '01/02/03/04公式', '首年月份无基准为空', '保留基准月'], ['核心截止', '202606', 'SPEC 1.2', '2026.01-06为核心；2026.07仅展示', '否', '各页范围列', '不得静默延长', '与JSON一致'], ['领导参考', '参考workbook不是本项目明细验收真值', '新增参考的材料和内容/销量预测计划部底表-户外地垫.xlsx', '独立参考的分子/分母/增速保留在00的H列；不覆盖market.db', '否', '00概览/92本行', '未提供销售额/均价，不推导', '正负结果并列呈现'], ['本次数据操作', '只读market.db与竞品DB；写入90明细、91聚合和结果公式页', '本次生成脚本', '未修改data/raw；未删除审计记录', '否', '90/91/92', 'Excel打开后可重算公式', '结果集中在一份xlsx']]; rules.getRange(`A4:H${3 + ruleRows.length}`).values = ruleRows; rules.getRange(`A4:H${3 + ruleRows.length}`).format = { wrapText: true, borders: { insideHorizontal: { style: 'thin', color: '#E7E6E6' } } }; widths(rules, [18, 34, 46, 58, 16, 28, 44, 34], 4 + ruleRows.length); rules.freezePanes.freezeRows(3);

await wb.recalculate();
const check = await wb.inspect({ kind: 'table', range: '00_概览!A1:H22', include: 'values,formulas', tableMaxRows: 22, tableMaxCols: 8, maxChars: 22000 }); console.log(check.ndjson.slice(0, 18000));
const errors = await wb.inspect({ kind: 'match', searchTerm: '#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A', options: { useRegex: true, maxResults: 200 }, summary: 'formula error scan' }); console.log(errors.ndjson);
await fs.mkdir(OUT_DIR, { recursive: true }); const preview = await wb.render({ sheetName: '00_概览', range: 'A1:H22', scale: 1.2, format: 'png' }); await fs.writeFile(PREVIEW_PATH, new Uint8Array(await preview.arrayBuffer())); const out = await SpreadsheetFile.exportXlsx(wb); await out.save(OUT_PATH);
console.log(JSON.stringify({ outputPath: OUT_PATH, previewPath: PREVIEW_PATH, detailRows: detailRows.length, aggregateRows: aggregateRows.length, months, benchmark: { industry: benchmark.industry?.growthPct ?? null, bsrTop100: benchmark.bsrTop100?.growthPct ?? null } }, null, 2)); db.close(); cdb.close();
