import fs from 'node:fs';
import {buildDataset, MONTHS, buildHtml, BANDS, FINE_BANDS} from './build_spec2_html.mjs';
export function prepare(){
  const d=buildDataset();
  const output={...d, months:[...MONTHS], bands:[...BANDS,...FINE_BANDS], dedup:d.fullDedup, raw:d.raw.rows, sheetStats:d.raw.sheetStats};
  fs.mkdirSync('tmp/formula_market_builder',{recursive:true});
  fs.writeFileSync('tmp/formula_market_builder/spec2_dataset.json', JSON.stringify(output));
  return d;
}
if(process.argv[1]?.endsWith('build_spec21_model.mjs')){const d=prepare();console.log(JSON.stringify({raw:d.raw.rows.length,dedup:d.fullDedup.length,months:MONTHS.length}));}
