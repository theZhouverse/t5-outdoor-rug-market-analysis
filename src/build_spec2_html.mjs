import fs from 'node:fs';
import path from 'node:path';
import crypto from 'node:crypto';
import XLSX from 'xlsx';

const ROOT = process.cwd();
const SOURCE = path.join(ROOT, 'data', 'raw', '地垫-卖家精灵市场数据.xlsx');
const OUTPUT = path.join(ROOT, '交付', '户外地垫市场分析报告-优化版.html');
const SOURCE_HASH = fs.existsSync(SOURCE) ? crypto.createHash('sha256').update(fs.readFileSync(SOURCE)).digest('hex') : '';

const MONTHS = [];
const BANDS = [
  { key: 'head', name: '头部 1-20', lo: 1, hi: 20 },
  { key: 'middle', name: '中部 21-50', lo: 21, hi: 50 },
  { key: 'tail', name: '尾部 51-100', lo: 51, hi: 100 }
];
const FINE_BANDS = [
  { key: '1_5', name: '1-5', lo: 1, hi: 5 },
  { key: '6_10', name: '6-10', lo: 6, hi: 10 },
  { key: '11_20', name: '11-20', lo: 11, hi: 20 },
  { key: '21_50', name: '21-50', lo: 21, hi: 50 },
  { key: '51_100', name: '51-100', lo: 51, hi: 100 }
];

function normaliseMonth(value) {
  const s = clean(value);
  const compact = s.match(/^(\d{4})(\d{2})$/);
  if (compact) return compact[1] + compact[2];
  const dotted = s.match(/^(\d{4})[.．](\d{1,2})$/);
  if (dotted) return dotted[1] + String(Number(dotted[2])).padStart(2, '0');
  return null;
}

function clean(value) {
  if (value === null || value === undefined) return '';
  const s = String(value).replace(/\u00a0/g, ' ').replace(/\r?\n/g, ' ').trim();
  if (!s || /^[-—–]$/.test(s) || /^(n\/a|na|null|undefined)$/i.test(s)) return '';
  return s;
}

function decodeXml(value) {
  return String(value)
    .replace(/<[^>]*>/g, '')
    .replace(/&amp;/g, '&')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&#39;/g, "'");
}

function display(cell) {
  if (!cell) return '';
  if (cell.v !== null && cell.v !== undefined && clean(cell.v) !== '') return String(cell.v);
  if (cell.l && cell.l.display !== null && cell.l.display !== undefined && clean(cell.l.display) !== '') {
    return String(cell.l.display);
  }
  if (cell.r !== null && cell.r !== undefined && clean(cell.r) !== '') return decodeXml(cell.r);
  if (cell.w !== null && cell.w !== undefined && clean(cell.w) !== '') return String(cell.w);
  return '';
}

function normalHeader(value) {
  return clean(value).toLowerCase().replace(/\s/g, '').replace(/[（）()]/g, '');
}

function findColumn(headers, names) {
  const wanted = names.map(normalHeader);
  for (const name of wanted) {
    const idx = headers.findIndex((h) => normalHeader(h) === name);
    if (idx >= 0) return idx;
  }
  return -1;
}

function parseNumber(value) {
  const s = clean(value);
  if (!s) return null;
  const negative = /^\(.*\)$/.test(s);
  const pct = s.includes('%');
  const normalized = s.replace(/[,$￥\s]/g, '').replace(/^\((.*)\)$/, '$1').replace(/%$/, '');
  const n = Number(normalized);
  if (!Number.isFinite(n)) return null;
  return negative ? -n : n;
}

function parseBsr(value) {
  const s = clean(value).replace(/,/g, '');
  if (!s) return null;
  const matches = s.match(/\d+(?:\.0+)?/g) || [];
  const values = matches.map((v) => Number(v)).filter((v) => Number.isInteger(v) && v > 0);
  return values.length ? Math.min(...values) : null;
}

function isPlasticTitle(value) {
  return /(?:^|[^a-z])plastic(?:[^a-z]|$)/i.test(clean(value));
}

function isGenimo(value) {
  return clean(value).toLowerCase() === 'genimo';
}

function sourceId(row) {
  return row.month + '#' + row.sourceRow;
}

function listingKey(row) {
  return row.parent || row.asin || 'source:' + sourceId(row);
}

function readRawRows() {
  if (!fs.existsSync(SOURCE)) throw new Error('找不到原始主源: ' + SOURCE);
  const wb = XLSX.readFile(SOURCE, {
    cellFormula: true,
    cellStyles: true,
    cellNF: true,
    cellText: false
  });
  const sheets = wb.SheetNames.map((name) => ({ original: String(name), month: normaliseMonth(name) }))
    .filter((item) => item.month)
    .sort((a, b) => a.month.localeCompare(b.month));
  if (!sheets.length) throw new Error('原始工作簿没有找到 YYYYMM 月度明细子表');
  const allRows = [];
  const sheetStats = [];
  for (const sheet of sheets) {
    const month = sheet.month;
    MONTHS.push(month);
    const ws = wb.Sheets[sheet.original];
    const range = XLSX.utils.decode_range(ws['!ref'] || 'A1');
    const headers = [];
    for (let c = range.s.c; c <= range.e.c; c += 1) {
      headers.push(display(ws[XLSX.utils.encode_cell({ r: 1, c })]));
    }
    const idx = {
      asin: findColumn(headers, ['ASIN']),
      sku: findColumn(headers, ['SKU']),
      brand: findColumn(headers, ['品牌']),
      title: findColumn(headers, ['商品标题', '标题']),
      parent: findColumn(headers, ['父ASIN', '父体ASIN']),
      smallBsr: findColumn(headers, ['小类BSR', '小类目BSR', '小类排名']),
      sales: findColumn(headers, ['月销量', '月销售量']),
      revenue: findColumn(headers, ['月销售额($)', '月销售额']),
      price: findColumn(headers, ['价格($)', '价格']),
      fba: findColumn(headers, ['FBA($)', 'FBA']),
      margin: findColumn(headers, ['毛利率']),
      coupon: findColumn(headers, ['Coupon'])
    };
    let rowCount = 0;
    let rankCount = 0;
    for (let r = 2; r <= range.e.r; r += 1) {
      const values = [];
      let hasValue = false;
      for (let c = range.s.c; c <= range.e.c; c += 1) {
        const value = display(ws[XLSX.utils.encode_cell({ r, c })]);
        values.push(value);
        if (clean(value)) hasValue = true;
      }
      if (!hasValue) continue;
      const row = {
        month,
        sourceRow: r + 1,
        asin: clean(values[idx.asin]),
        sku: clean(values[idx.sku]),
        brand: clean(values[idx.brand]),
        title: clean(values[idx.title]),
        parent: clean(values[idx.parent]),
        rank: parseBsr(values[idx.smallBsr]),
        sales: parseNumber(values[idx.sales]),
        revenue: parseNumber(values[idx.revenue]),
        price: parseNumber(values[idx.price]),
        fba: parseNumber(values[idx.fba]),
        margin: clean(values[idx.margin]),
        coupon: clean(values[idx.coupon])
      };
      if (!row.asin && !row.parent && !row.title && !row.brand && row.rank === null) continue;
      allRows.push(row);
      rowCount += 1;
      if (row.rank !== null) rankCount += 1;
    }
    sheetStats.push({ month, rowCount, rankCount });
  }
  return { rows: allRows, sheetStats };
}

function buildFamilyFlags(rows) {
  const flags = new Map();
  for (const row of rows) {
    const key = row.month + '|' + listingKey(row);
    const current = flags.get(key) || { plastic: false, genimo: false };
    current.plastic = current.plastic || isPlasticTitle(row.title);
    current.genimo = current.genimo || isGenimo(row.brand);
    flags.set(key, current);
  }
  return flags;
}

function representativeCompare(a, b) {
  const ar = a.rank === null ? Number.POSITIVE_INFINITY : a.rank;
  const br = b.rank === null ? Number.POSITIVE_INFINITY : b.rank;
  if (ar !== br) return ar - br;
  const ac = (Number.isFinite(a.sales) ? 1 : 0) + (Number.isFinite(a.revenue) ? 1 : 0);
  const bc = (Number.isFinite(b.sales) ? 1 : 0) + (Number.isFinite(b.revenue) ? 1 : 0);
  if (ac !== bc) return bc - ac;
  if (Number.isFinite(a.price) !== Number.isFinite(b.price)) return Number.isFinite(b.price) - Number.isFinite(a.price);
  return sourceId(a).localeCompare(sourceId(b));
}

function dedup(rows, flags) {
  const groups = new Map();
  for (const row of rows) {
    const key = row.month + '|' + listingKey(row);
    if (!groups.has(key)) groups.set(key, []);
    groups.get(key).push(row);
  }
  const result = [];
  for (const [key, items] of groups.entries()) {
    const rep = items.slice().sort(representativeCompare)[0];
    const family = flags.get(key) || { plastic: false, genimo: false };
    result.push({
      ...rep,
      familyKey: key,
      plastic: family.plastic,
      genimo: family.genimo,
      rawRows: items.length,
      sourceRows: items.map(sourceId)
    });
  }
  return result.sort((a, b) => (a.month + a.familyKey).localeCompare(b.month + b.familyKey));
}

function percent(current, previous) {
  if (!Number.isFinite(current) || !Number.isFinite(previous) || previous === 0) return null;
  return (current - previous) / Math.abs(previous);
}

function sum(rows, field) {
  const values = rows.map((r) => r[field]).filter(Number.isFinite);
  return values.length ? values.reduce((a, b) => a + b, 0) : null;
}

function avg(rows, field) {
  const values = rows.map((r) => r[field]).filter(Number.isFinite);
  return values.length ? values.reduce((a, b) => a + b, 0) / values.length : null;
}

function summarise(rows) {
  const salesRows = rows.filter((r) => Number.isFinite(r.sales));
  const revenueRows = rows.filter((r) => Number.isFinite(r.revenue));
  const paired = rows.filter((r) => Number.isFinite(r.sales) && Number.isFinite(r.revenue) && r.sales > 0);
  const sales = sum(rows, 'sales');
  const revenue = sum(rows, 'revenue');
  const pairedSales = sum(paired, 'sales');
  const pairedRevenue = sum(paired, 'revenue');
  return {
    count: rows.length,
    sales,
    revenue,
    avgPrice: avg(rows, 'price'),
    pairedAsp: pairedSales && pairedRevenue !== null ? pairedRevenue / pairedSales : null,
    salesCoverage: rows.length ? salesRows.length / rows.length : null,
    revenueCoverage: rows.length ? revenueRows.length / rows.length : null,
    missingSales: rows.length - salesRows.length,
    missingRevenue: rows.length - revenueRows.length,
    missingPrice: rows.length - rows.filter((r) => Number.isFinite(r.price)).length
  };
}

function previousMonth(month) {
  const y = Number(month.slice(0, 4));
  const m = Number(month.slice(4, 6));
  const date = new Date(Date.UTC(y, m - 2, 1));
  return String(date.getUTCFullYear()) + String(date.getUTCMonth() + 1).padStart(2, '0');
}

function previousYear(month) {
  return String(Number(month.slice(0, 4)) - 1) + month.slice(4, 6);
}

function enrichTrend(months, rows) {
  const byMonth = new Map();
  for (const month of months) byMonth.set(month, []);
  for (const row of rows) {
    if (!byMonth.has(row.month)) byMonth.set(row.month, []);
    byMonth.get(row.month).push(row);
  }
  const base = new Map();
  for (const month of months) base.set(month, summarise(byMonth.get(month) || []));
  return months.map((month) => {
    const current = base.get(month);
    const mom = base.get(previousMonth(month)) || {};
    const yoy = base.get(previousYear(month)) || {};
    return {
      month,
      ...current,
      momSales: percent(current.sales, mom.sales),
      momRevenue: percent(current.revenue, mom.revenue),
      momAvgPrice: percent(current.avgPrice, mom.avgPrice),
      yoySales: percent(current.sales, yoy.sales),
      yoyRevenue: percent(current.revenue, yoy.revenue),
      yoyAvgPrice: percent(current.avgPrice, yoy.avgPrice)
    };
  });
}

function annualRows(months, rows) {
  const byYear = new Map();
  for (const row of rows) {
    const year = row.month.slice(0, 4);
    if (!byYear.has(year)) byYear.set(year, []);
    byYear.get(year).push(row);
  }
  const years = Array.from(byYear.keys()).sort();
  const summaries = years.map((year) => ({ year, ...summarise(byYear.get(year)) }));
  return summaries.map((current, i) => {
    const previous = summaries[i - 1] || {};
    const yoySales = percent(current.sales, previous.sales);
    const yoyRevenue = percent(current.revenue, previous.revenue);
    return { ...current, yoySales, yoyRevenue };
  });
}

function filterBand(rows, band) {
  return rows.filter((r) => r.rank !== null && r.rank >= band.lo && r.rank <= band.hi);
}

function buildCategory(title, fullRows, topRows) {
  const tierRows = {};
  for (const band of BANDS) tierRows[band.key] = filterBand(topRows, band);
  const fineRows = {};
  for (const band of FINE_BANDS) fineRows[band.key] = filterBand(topRows, band);
  return {
    title,
    fullRows,
    topRows,
    monthly: enrichTrend(MONTHS, fullRows),
    topMonthly: enrichTrend(MONTHS, topRows),
    annual: annualRows(MONTHS, fullRows),
    topAnnual: annualRows(MONTHS, topRows),
    tiers: Object.fromEntries(BANDS.map((band) => [band.key, enrichTrend(MONTHS, tierRows[band.key])])),
    fine: Object.fromEntries(FINE_BANDS.map((band) => [band.key, enrichTrend(MONTHS, fineRows[band.key])]))
  };
}

function esc(value) {
  return String(value === null || value === undefined ? '' : value)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;').replace(/'/g, '&#39;');
}

function fmt(value, digits = 0) {
  if (!Number.isFinite(value)) return '—';
  return value.toLocaleString('zh-CN', { maximumFractionDigits: digits, minimumFractionDigits: digits });
}

function fmtPct(value) {
  return Number.isFinite(value) ? (value * 100).toFixed(1) + '%' : '—';
}

function lastNonEmpty(rows) {
  return rows.slice().reverse().find((r) => r.count > 0) || rows[rows.length - 1];
}

function svgBarChart(title, rows, field, color, digits) {
  const values = rows.map((r) => Number.isFinite(r[field]) ? r[field] : 0);
  const max = Math.max(...values, 0);
  const width = 900;
  const height = 260;
  const left = 54;
  const bottom = 42;
  const plotH = 170;
  const plotW = width - left - 18;
  const barW = rows.length ? Math.max(2, plotW / rows.length - 2) : 2;
  let body = '<div class="chart"><h4>' + esc(title) + '</h4><svg viewBox="0 0 ' + width + ' ' + height + '" role="img" aria-label="' + esc(title) + '">';
  body += '<line x1="' + left + '" y1="' + (height - bottom) + '" x2="' + width + '" y2="' + (height - bottom) + '" stroke="#94a3b8"/>';
  rows.forEach((row, i) => {
    const value = values[i];
    const h = max > 0 ? (value / max) * plotH : 0;
    const x = left + i * (plotW / Math.max(rows.length, 1)) + 1;
    const y = height - bottom - h;
    body += '<rect x="' + x.toFixed(1) + '" y="' + y.toFixed(1) + '" width="' + barW.toFixed(1) + '" height="' + h.toFixed(1) + '" rx="2" fill="' + color + '"><title>' + esc(row.month + '：' + fmt(row[field], digits)) + '</title></rect>';
    if (i % Math.max(1, Math.ceil(rows.length / 10)) === 0) {
      body += '<text x="' + (x + barW / 2).toFixed(1) + '" y="' + (height - 17) + '" text-anchor="middle" class="axis-label">' + esc(row.month.slice(0, 4) + '.' + Number(row.month.slice(4, 6))) + '</text>';
    }
  });
  body += '<text x="8" y="18" class="axis-label">' + esc(max > 0 ? fmt(max, digits) : '—') + '</text></svg></div>';
  return body;
}

function svgLineChart(title, rows, series, percentMode = false) {
  const width = 900;
  const height = 280;
  const left = 54;
  const right = 18;
  const top = 28;
  const bottom = 44;
  const plotW = width - left - right;
  const plotH = height - top - bottom;
  const values = [];
  for (const item of series) for (const row of rows) if (Number.isFinite(row[item.field])) values.push(row[item.field]);
  let min = values.length ? Math.min(...values) : 0;
  let max = values.length ? Math.max(...values) : 1;
  if (min === max) { min -= 1; max += 1; }
  const scaleX = (i) => left + (rows.length <= 1 ? 0 : i * plotW / (rows.length - 1));
  const scaleY = (v) => top + (max - v) * plotH / (max - min);
  let body = '<div class="chart"><h4>' + esc(title) + '</h4><svg viewBox="0 0 ' + width + ' ' + height + '" role="img" aria-label="' + esc(title) + '">';
  if (min < 0 && max > 0) {
    const y0 = scaleY(0);
    body += '<line x1="' + left + '" y1="' + y0.toFixed(1) + '" x2="' + (width - right) + '" y2="' + y0.toFixed(1) + '" stroke="#cbd5e1" stroke-dasharray="4 4"/>';
  }
  body += '<line x1="' + left + '" y1="' + top + '" x2="' + left + '" y2="' + (height - bottom) + '" stroke="#94a3b8"/>';
  for (const item of series) {
    const points = rows.map((row, i) => Number.isFinite(row[item.field]) ? scaleX(i).toFixed(1) + ',' + scaleY(row[item.field]).toFixed(1) : null).filter(Boolean);
    if (points.length > 1) body += '<polyline fill="none" stroke="' + item.color + '" stroke-width="2.5" points="' + points.join(' ') + '"/>';
    rows.forEach((row, i) => {
      if (Number.isFinite(row[item.field])) {
        body += '<circle cx="' + scaleX(i).toFixed(1) + '" cy="' + scaleY(row[item.field]).toFixed(1) + '" r="2.8" fill="' + item.color + '"><title>' + esc(row.month + '：' + (percentMode ? fmtPct(row[item.field]) : fmt(row[item.field], 2))) + '</title></circle>';
      }
    });
  }
  const step = Math.max(1, Math.ceil(rows.length / 10));
  rows.forEach((row, i) => {
    if (i % step === 0) body += '<text x="' + scaleX(i).toFixed(1) + '" y="' + (height - 17) + '" text-anchor="middle" class="axis-label">' + esc(row.month.slice(0, 4) + '.' + Number(row.month.slice(4, 6))) + '</text>';
  });
  series.forEach((item, i) => {
    const x = left + i * 155;
    body += '<rect x="' + x + '" y="8" width="12" height="12" fill="' + item.color + '"/><text x="' + (x + 17) + '" y="18" class="legend-label">' + esc(item.name) + '</text>';
  });
  body += '<text x="8" y="' + (top + 5) + '" class="axis-label">' + esc(percentMode ? fmtPct(max) : fmt(max, 0)) + '</text>';
  body += '<text x="8" y="' + (height - bottom) + '" class="axis-label">' + esc(percentMode ? fmtPct(min) : fmt(min, 0)) + '</text>';
  body += '</svg></div>';
  return body;
}

function table(headers, rows) {
  let out = '<div class="table-wrap"><table><thead><tr>';
  for (const h of headers) out += '<th>' + esc(h) + '</th>';
  out += '</tr></thead><tbody>';
  for (const row of rows) {
    out += '<tr>';
    for (const cell of row) out += '<td>' + (cell && cell.__html ? cell.value : esc(cell)) + '</td>';
    out += '</tr>';
  }
  return out + '</tbody></table></div>';
}

function trendTable(data, topData) {
  const rows = data.map((r, i) => [
    r.month.slice(0, 4) + '.' + Number(r.month.slice(4, 6)),
    fmt(r.count),
    fmt(r.sales),
    fmt(r.revenue, 2),
    fmt(r.avgPrice, 2),
    fmt(r.pairedAsp, 2),
    fmtPct(r.yoySales),
    fmtPct(r.momSales),
    fmtPct(r.yoyRevenue),
    fmtPct(r.momRevenue),
    fmt(topData[i] ? topData[i].count : null)
  ]);
  return table(['月份', '整体商品数', '销量', '销售额($)', '平均标价($)', '加权成交均价($)', '销量同比', '销量环比', '销售额同比', '销售额环比', 'BSR前100商品数'], rows);
}

function topTrendTable(data) {
  const rows = data.map((r) => [
    r.month.slice(0, 4) + '.' + Number(r.month.slice(4, 6)),
    fmt(r.count),
    fmt(r.sales),
    fmt(r.revenue, 2),
    fmt(r.avgPrice, 2),
    fmt(r.pairedAsp, 2),
    fmtPct(r.yoySales),
    fmtPct(r.momSales),
    fmtPct(r.yoyRevenue),
    fmtPct(r.momRevenue)
  ]);
  return table(['月份', 'Top100商品数', '销量', '销售额($)', '平均标价($)', '加权成交均价($)', '销量同比', '销量环比', '销售额同比', '销售额环比'], rows);
}

function annualTable(data, topData) {
  const rows = data.map((r, i) => {
    const top = topData[i] || {};
    return [
      r.year,
      fmt(r.count),
      fmt(r.sales),
      fmt(r.revenue, 2),
      fmt(r.avgPrice, 2),
      fmtPct(r.yoySales),
      fmtPct(r.yoyRevenue),
      fmt(top.count),
      fmt(top.sales),
      fmt(top.revenue, 2),
      fmtPct(top.yoySales),
      fmtPct(top.yoyRevenue)
    ];
  });
  return table(['年份', '整体商品数', '整体销量', '整体销售额($)', '整体平均标价($)', '整体销量同比', '整体销售额同比', 'Top100商品数', 'Top100销量', 'Top100销售额($)', 'Top100销量同比', 'Top100销售额同比'], rows);
}

function tierTable(data, title) {
  const rows = data.map((r) => [
    r.month.slice(0, 4) + '.' + Number(r.month.slice(4, 6)),
    fmt(r.count),
    fmt(r.sales),
    fmt(r.revenue, 2),
    fmtPct(r.yoySales),
    fmtPct(r.momSales)
  ]);
  return '<h4>' + esc(title) + '</h4>' + table(['月份', '商品数', '销量', '销售额($)', '销量同比', '销量环比'], rows);
}

function narrative(category, overall, pp, nonpp) {
  const latest = lastNonEmpty(category.monthly);
  const prior = category.monthly[category.monthly.findIndex((r) => r.month === latest.month) - 1] || {};
  const top = category.topMonthly.find((r) => r.month === latest.month) || {};
  const pieces = [];
  pieces.push('截至 ' + latest.month.slice(0, 4) + '.' + Number(latest.month.slice(4, 6)) + '，' + category.title + '整体盘共有 ' + fmt(latest.count) + ' 个去重 Listing，销量 ' + fmt(latest.sales) + '，销售额 $' + fmt(latest.revenue, 2) + '，平均标价 $' + fmt(latest.avgPrice, 2) + '。');
  pieces.push('与上月相比，销量' + (latest.momSales === null ? '缺少可比月份' : (latest.momSales >= 0 ? '增长 ' : '下降 ') + fmtPct(Math.abs(latest.momSales))) + '；与去年同月相比，销量' + (latest.yoySales === null ? '缺少可比月份' : (latest.yoySales >= 0 ? '增长 ' : '下降 ') + fmtPct(Math.abs(latest.yoySales)) ) + '。');
  pieces.push('BSR 1—100（包含100）在该月覆盖 ' + fmt(top.count) + ' 个 Listing，销量 ' + fmt(top.sales) + '，销售额 $' + fmt(top.revenue, 2) + '；Top100与整体盘的差异用于判断头部集中度。');
  if (category.title === '整体市场') {
    const p = lastNonEmpty(pp.monthly);
    const n = lastNonEmpty(nonpp.monthly);
    pieces.push('整体市场按完整单词 plastic 划分为 PP 与非PP，两者并集构成整体市场；最新月 PP 销量为 ' + fmt(p.sales) + '，非PP销量为 ' + fmt(n.sales) + '，请结合覆盖率和缺失值复核方向。');
  }
  pieces.push('原始销量覆盖率 ' + fmtPct(latest.salesCoverage) + '，销售额覆盖率 ' + fmtPct(latest.revenueCoverage) + '；空值保留为空，不把缺失当作零。');
  return pieces;
}

function marketSection(id, number, category, overall, pp, nonpp) {
  const latest = lastNonEmpty(category.monthly);
  const topLatest = category.topMonthly.find((r) => r.month === latest.month) || {};
  const fineCombined = MONTHS.map((month, i) => {
    const row = { month };
    for (const band of FINE_BANDS) row['mom_' + band.key] = category.fine[band.key][i] ? category.fine[band.key][i].momSales : null;
    return row;
  });
  const tierCombined = MONTHS.map((month, i) => {
    const row = { month };
    for (const band of BANDS) row['mom_' + band.key] = category.tiers[band.key][i] ? category.tiers[band.key][i].momSales : null;
    return row;
  });
  let out = '<section id="' + id + '"><h2>' + number + '｜' + esc(category.title) + '</h2>';
  out += '<p class="lead">本部分先展示整体盘，再展示 BSR 1—100（包含100）的 Top100 视图和头部/中部/尾部层级。同比为去年同月，环比为上一个自然月。</p>';
  out += '<div class="cards"><div class="card"><b>最新月整体销量</b><strong>' + fmt(latest.sales) + '</strong><small>' + esc(latest.month) + '</small></div>';
  out += '<div class="card"><b>最新月整体销售额</b><strong>$' + fmt(latest.revenue, 2) + '</strong><small>原始月销售额合计</small></div>';
  out += '<div class="card"><b>最新月 Top100 销量</b><strong>' + fmt(topLatest.sales) + '</strong><small>' + fmt(topLatest.count) + ' 个 Listing</small></div>';
  out += '<div class="card"><b>最新月平均标价</b><strong>$' + fmt(latest.avgPrice, 2) + '</strong><small>按有价格记录算术平均</small></div></div>';
  out += '<h3>月度汇总与可回勾数据</h3>' + trendTable(category.monthly, category.topMonthly);
  out += '<h4>BSR Top100 月度明细</h4>' + topTrendTable(category.topMonthly);
  out += '<h4>年度与同周期同比</h4>' + annualTable(category.annual, category.topAnnual);
  out += '<div class="charts">' + svgBarChart(category.title + '月销量', category.monthly, 'sales', '#2563eb', 0);
  out += svgBarChart(category.title + '月销售额', category.monthly, 'revenue', '#0f766e', 2);
  out += svgLineChart(category.title + '平均标价趋势', category.monthly, [{ name: '平均标价($)', field: 'avgPrice', color: '#7c3aed' }]);
  out += svgLineChart(category.title + '销量同比与环比', category.monthly, [{ name: '同比', field: 'yoySales', color: '#dc2626' }, { name: '环比', field: 'momSales', color: '#2563eb' }], true);
  out += '</div><h3>BSR Top100 与五档环比</h3>';
  out += '<p>Top100 先按可解析的小类BSR筛选 1—100（含100），再按父ASIN优先、否则ASIN去重；五档互斥为 1-5、6-10、11-20、21-50、51-100。</p>';
  out += '<div class="charts">';
  out += svgLineChart(category.title + ' Top100 五档销量环比', fineCombined, FINE_BANDS.map((band, i) => ({ name: band.name, field: 'mom_' + band.key, color: ['#1d4ed8', '#0891b2', '#059669', '#f59e0b', '#dc2626'][i] })), true);
  out += svgLineChart(category.title + ' Top100 头/中/尾销量环比', tierCombined, BANDS.map((band, i) => ({ name: band.name, field: 'mom_' + band.key, color: ['#2563eb', '#059669', '#f97316'][i] })), true);
  out += '</div>';
  out += '<h3>头部 / 中部 / 尾部</h3>';
  for (const band of BANDS) out += tierTable(category.tiers[band.key], band.name);
  out += '<h3>数据驱动文字分析</h3><div class="analysis">';
  for (const paragraph of narrative(category, overall, pp, nonpp)) out += '<p>' + esc(paragraph) + '</p>';
  out += '</div></section>';
  return out;
}

function genimoSection(category, overall, pp, genimoPP) {
  const rows = category.monthly.map((r, i) => {
    const base = overall.monthly[i] || {};
    const top = category.topMonthly[i] || {};
    return [
      r.month.slice(0, 4) + '.' + Number(r.month.slice(4, 6)),
      fmt(r.count),
      fmt(r.sales),
      fmt(r.revenue, 2),
      fmt(r.avgPrice, 2),
      fmtPct(percent(r.sales, base.sales)),
      fmtPct(percent(r.revenue, base.revenue)),
      fmt(top.count)
    ];
  });
  let out = '<section id="genimo"><h2>第四部分｜GENIMO 品牌分析</h2>';
  out += '<p class="lead">GENIMO 是整体市场中的品牌视角。本部分同时给出 GENIMO 在整体市场的份额、BSR Top100表现、头中尾分层与基于原始数据的 2027 年行动建议。</p>';
  out += '<h3>品牌月度表现与整体市场份额</h3>';
 out += table(['月份', 'GENIMO Listing数', '销量', '销售额($)', '平均标价($)', '销量占整体', '销售额占整体', 'Top100 Listing数'], rows);
  const ppRows = genimoPP.monthly.map((r, i) => {
    const base = pp.monthly[i] || {};
    return [
      r.month.slice(0, 4) + '.' + Number(r.month.slice(4, 6)),
      fmt(r.count),
      fmt(r.sales),
      fmt(r.revenue, 2),
      fmt(r.avgPrice, 2),
      fmtPct(percent(r.sales, base.sales)),
      fmtPct(percent(r.revenue, base.revenue)),
      fmtPct(r.yoySales),
      fmtPct(r.momSales)
    ];
  });
  out += '<h3>GENIMO 在 PP 市场中的表现</h3>';
  out += '<p>PP市场份额分母为PP整体盘；GENIMO在PP中的统计仍沿用父ASIN优先/ASIN去重和Top100先筛选规则。</p>';
  out += table(['月份', 'GENIMO PP Listing数', 'PP内销量', 'PP内销售额($)', '平均标价($)', '销量占PP', '销售额占PP', '销量同比', '销量环比'], ppRows);
  out += '<div class="charts">' + svgBarChart('GENIMO PP 月销量', genimoPP.monthly, 'sales', '#be185d', 0) + svgBarChart('GENIMO PP 月销售额', genimoPP.monthly, 'revenue', '#9d174d', 2) + '</div>';
  out += '<div class="charts">' + svgBarChart('GENIMO 月销量', category.monthly, 'sales', '#9333ea', 0);
  out += svgBarChart('GENIMO 月销售额', category.monthly, 'revenue', '#c026d3', 2);
  out += svgLineChart('GENIMO 销量同比与环比', category.monthly, [{ name: '同比', field: 'yoySales', color: '#dc2626' }, { name: '环比', field: 'momSales', color: '#2563eb' }], true);
  out += '</div><h3>GENIMO BSR 头中尾</h3>';
 for (const band of BANDS) out += tierTable(category.tiers[band.key], 'GENIMO ' + band.name);
  out += '<h3>GENIMO 年度与同周期同比</h3>' + annualTable(category.annual, category.topAnnual);
  const latest = lastNonEmpty(category.monthly);
  const latestTop = category.topMonthly.find((r) => r.month === latest.month) || {};
  out += '<h3>2027 年建议（仅基于当前原始字段）</h3><div class="analysis">';
  out += '<p>先稳定能够进入 BSR 1—100 的链接：最新月 GENIMO Top100 有 ' + fmt(latestTop.count) + ' 个 Listing，销量 ' + fmt(latestTop.sales) + '。产品测试和链接补充应按 1-20、21-50、51-100 三个层级分别记录，不用单一总排名替代层级判断。</p>';
  out += '<p>用整体市场份额和 PP/非PP拆分制定资源优先级：当 GENIMO 在某一市场的销量或销售额占比连续上升时，优先补充该市场相同字段完整且 BSR 可追踪的链接；当份额下降时，先核对缺失销量、销售额、价格和 BSR 的覆盖率，再决定是否调整产品组合。</p>';
  out += '<p>建立月度复盘表：销量、销售额、平均标价、同比、环比、Top100数量、头中尾占比和标题中 plastic 标记必须同批次留痕。报告只使用原始工作簿已有字段，不推导利润、成本、广告或库存指标。</p>';
  out += '</div></section>';
  return out;
}

function rawFieldCoverage(rows) {
  const fields = [
    ['商品标题', 'title'],
    ['品牌', 'brand'],
    ['父ASIN', 'parent'],
    ['小类BSR', 'rank'],
    ['月销量', 'sales'],
    ['月销售额($)', 'revenue'],
    ['价格($)', 'price'],
    ['毛利率（原始）', 'margin'],
    ['FBA($)（原始）', 'fba'],
    ['Coupon（原始）', 'coupon']
  ];
  const result = fields.map(([name, field]) => {
    const present = rows.filter((row) => field === 'rank' ? row.rank !== null : (field === 'sales' || field === 'revenue' || field === 'price' || field === 'fba') ? Number.isFinite(row[field]) : clean(row[field]) !== '').length;
    return [name, fmt(present), fmtPct(rows.length ? present / rows.length : null), field === 'margin' || field === 'fba' || field === 'coupon' ? '仅保留原始值/覆盖率，不参与利润推导' : '参与对应的市场统计或筛选'];
  });
  return table(['原始字段', '非空/有效行数', '覆盖率', '处理方式'], result);
}

function buildHtml(raw, categories) {
  const rawCount = raw.rows.length;
  const ranked = raw.rows.filter((r) => r.rank !== null).length;
  const topCandidates = raw.rows.filter((r) => r.rank !== null && r.rank >= 1 && r.rank <= 100).length;
  const generated = new Date().toISOString().replace('T', ' ').replace(/\.\d+Z$/, ' UTC');
  let html = '<!doctype html><html lang="zh-CN"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">';
  html += '<title>户外地垫市场分析报告｜SPEC 2.0</title><style>';
  html += 'body{margin:0;background:#f8fafc;color:#172033;font:14px/1.6 -apple-system,BlinkMacSystemFont,"Segoe UI","Microsoft YaHei",sans-serif}header{background:linear-gradient(130deg,#0f172a,#1e3a8a);color:white;padding:34px 5vw 28px}h1{margin:0 0 8px;font-size:30px}h2{font-size:25px;color:#0f172a;margin:10px 0 15px}h3{font-size:19px;color:#1e3a8a;margin:28px 0 12px}h4{margin:8px 0;color:#334155}main{max-width:1440px;margin:auto;padding:20px 5vw 80px}nav{position:sticky;top:0;z-index:3;background:#fff;border-bottom:1px solid #e2e8f0;padding:11px 5vw;display:flex;gap:9px;flex-wrap:wrap}nav a{color:#1e40af;text-decoration:none;border:1px solid #bfdbfe;border-radius:18px;padding:4px 12px;background:#eff6ff}.lead{font-size:15px;color:#475569}.meta,.notice{background:#fff;border:1px solid #dbeafe;border-radius:12px;padding:14px 16px;margin:14px 0}.notice{border-left:5px solid #2563eb;background:#eff6ff}.cards{display:grid;grid-template-columns:repeat(auto-fit,minmax(190px,1fr));gap:12px;margin:16px 0}.card{background:#fff;border:1px solid #e2e8f0;border-radius:13px;padding:15px;box-shadow:0 2px 8px #0f172a0b}.card b,.card small{display:block;color:#64748b}.card strong{display:block;font-size:24px;color:#1d4ed8;margin:5px 0}.table-wrap{overflow:auto;background:#fff;border:1px solid #e2e8f0;border-radius:10px;margin:9px 0 18px}table{border-collapse:collapse;width:100%;min-width:900px}th,td{border-bottom:1px solid #eef2f7;padding:7px 9px;text-align:right;white-space:nowrap}th:first-child,td:first-child{text-align:left}th{position:sticky;top:47px;background:#eef4ff;color:#1e3a8a;font-weight:700}tr:nth-child(even) td{background:#fbfdff}.charts{display:grid;grid-template-columns:repeat(auto-fit,minmax(430px,1fr));gap:13px}.chart{background:#fff;border:1px solid #e2e8f0;border-radius:12px;padding:8px;overflow:auto}.chart svg{width:100%;height:auto;min-width:390px}.axis-label,.legend-label{font-size:10px;fill:#64748b}.analysis{background:#fff;border:1px solid #e2e8f0;border-radius:12px;padding:13px 17px}.analysis p{margin:8px 0}footer{margin-top:36px;color:#64748b;font-size:12px;border-top:1px solid #cbd5e1;padding-top:16px}@media(max-width:700px){h1{font-size:24px}main{padding-left:3vw;padding-right:3vw}.charts{grid-template-columns:1fr}}';
  html += '</style></head><body><header><h1>户外地垫市场分析报告</h1><div>SPEC 2.0｜原始主源直读｜四部分交付｜生成时间 ' + esc(generated) + '</div></header>';
  html += '<nav><a href="#overall">一、整体市场</a><a href="#pp">二、PP市场</a><a href="#nonpp">三、非PP市场</a><a href="#genimo">四、GENIMO品牌</a></nav><main>';
  html += '<div class="notice"><b>口径说明：</b>整体市场 = PP市场 ∪ 非PP市场；PP 使用商品标题中完整单词 plastic 匹配，非PP为其补集。BSR Top100 先解析并标记小类BSR 1—100（包含100），再做去重和汇总。月度同比比较去年同月，月度环比比较上一个自然月。</div>';
  html += '<div class="notice"><b>数据边界：</b>本版本只读取原始工作簿已有字段，不计算利润、成本、广告、库存或其他外部指标；原始空值保留为空，不当作零。原始毛利率、FBA、Coupon 如需查看，应回到输入工作簿核对。</div>';
  html += '<div class="meta"><b>来源与处理记录</b><br>主源：' + esc(SOURCE) + '<br>主源 SHA-256：' + esc(SOURCE_HASH) + '<br>月度子表：' + MONTHS.length + ' 张（' + esc(MONTHS[0]) + '—' + esc(MONTHS[MONTHS.length - 1]) + '）<br>原始有效行：' + fmt(rawCount) + '；可解析小类BSR：' + fmt(ranked) + '；BSR 1—100候选行：' + fmt(topCandidates) + '<br>统计单元：独立 Listing（父ASIN优先，否则ASIN；均无则保留源行）；代表行按最小可解析BSR、销量/销售额完整度、价格完整度、源行ID确定。</div>';
  html += '<h3>原始字段覆盖（只读）</h3>' + rawFieldCoverage(raw.rows);
  html += marketSection('overall', '第一部分', categories.overall, categories.overall, categories.pp, categories.nonpp);
  html += marketSection('pp', '第二部分', categories.pp, categories.overall, categories.pp, categories.nonpp);
  html += marketSection('nonpp', '第三部分', categories.nonpp, categories.overall, categories.pp, categories.nonpp);
  html += genimoSection(categories.genimo, categories.overall, categories.pp, categories.genimoPP);
  html += '<footer>报告由 src/build_spec2_html.mjs 从原始工作簿生成。数值表保留可回勾字段；如源表发生更换，请先确认文件身份、月份覆盖和列名，再重新生成。</footer></main></body></html>';
  return html;
}

const raw = readRawRows();
const flags = buildFamilyFlags(raw.rows);
const fullDedup = dedup(raw.rows, flags);
const topCandidates = raw.rows.filter((r) => r.rank !== null && r.rank >= 1 && r.rank <= 100);
const topDedup = dedup(topCandidates, flags);
const categories = {
  overall: buildCategory('整体市场', fullDedup, topDedup),
  pp: buildCategory('PP市场', fullDedup.filter((r) => r.plastic), topDedup.filter((r) => r.plastic)),
  nonpp: buildCategory('非PP市场', fullDedup.filter((r) => !r.plastic), topDedup.filter((r) => !r.plastic)),
  genimo: buildCategory('GENIMO品牌', fullDedup.filter((r) => r.genimo), topDedup.filter((r) => r.genimo)),
  genimoPP: buildCategory('GENIMO PP市场', fullDedup.filter((r) => r.genimo && r.plastic), topDedup.filter((r) => r.genimo && r.plastic))
};
function addMetric(a, b) {
  if (Number.isFinite(a) && Number.isFinite(b)) return a + b;
  if (Number.isFinite(a)) return a;
  if (Number.isFinite(b)) return b;
  return null;
}
function metricEqual(a, b) {
  if (!Number.isFinite(a) && !Number.isFinite(b)) return true;
  return Number.isFinite(a) && Number.isFinite(b) && Math.abs(a - b) < 1e-6;
}
for (let i = 0; i < MONTHS.length; i += 1) {
  const overallMonth = categories.overall.monthly[i];
  const ppMonth = categories.pp.monthly[i];
  const nonppMonth = categories.nonpp.monthly[i];
  const overallTop = categories.overall.topMonthly[i];
  const ppTop = categories.pp.topMonthly[i];
  const nonppTop = categories.nonpp.topMonthly[i];
  const fullSales = addMetric(ppMonth.sales, nonppMonth.sales);
  const fullRevenue = addMetric(ppMonth.revenue, nonppMonth.revenue);
  const topSales = addMetric(ppTop.sales, nonppTop.sales);
  const topRevenue = addMetric(ppTop.revenue, nonppTop.revenue);
  if (overallMonth.count !== ppMonth.count + nonppMonth.count ||
      overallTop.count !== ppTop.count + nonppTop.count ||
      !metricEqual(overallMonth.sales, fullSales) ||
      !metricEqual(overallMonth.revenue, fullRevenue) ||
      !metricEqual(overallTop.sales, topSales) ||
      !metricEqual(overallTop.revenue, topRevenue)) {
    throw new Error('PP/非PP互补校验失败: ' + MONTHS[i]);
  }
}
if (!raw.rows.some((row) => row.rank === 100)) throw new Error('原始主源没有检测到 BSR=100，无法确认包含100');
fs.mkdirSync(path.dirname(OUTPUT), { recursive: true });
fs.writeFileSync(OUTPUT, buildHtml(raw, categories), 'utf8');
console.log(JSON.stringify({
  output: OUTPUT,
  months: MONTHS.length,
  rawRows: raw.rows.length,
  sourceSha256: SOURCE_HASH,
  fullDedup: fullDedup.length,
  top100Dedup: topDedup.length,
  categories: Object.fromEntries(Object.entries(categories).map(([k, v]) => [k, { full: v.fullRows.length, top100: v.topRows.length }]))
}, null, 2));
