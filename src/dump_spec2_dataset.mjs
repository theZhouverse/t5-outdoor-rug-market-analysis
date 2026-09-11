import fs from 'node:fs';
import { buildDataset, MONTHS } from './build_spec2_html.mjs';

const dataset = buildDataset();
const output = process.argv[2] || 'tmp/formula_market_builder/spec2_dataset.json';
fs.writeFileSync(output, JSON.stringify({
  raw: dataset.raw.rows,
  dedup: dataset.fullDedup,
  months: MONTHS,
  sheetStats: dataset.raw.sheetStats
}), 'utf8');
console.log(JSON.stringify({ output, rawRows: dataset.raw.rows.length, dedupRows: dataset.fullDedup.length }));
