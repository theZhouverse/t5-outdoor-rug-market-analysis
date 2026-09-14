import fs from 'node:fs';
import path from 'node:path';
import crypto from 'node:crypto';
import {spawn} from 'node:child_process';
import {pathToFileURL} from 'node:url';
import {prepare} from './build_spec21_model.mjs';
import {buildHtml,MONTHS,narrative,fmt,fmtPct,BANDS,FINE_BANDS} from './build_spec2_html.mjs';
const ROOT=process.cwd();
const PY=process.env.SPEC21_PYTHON || path.join(process.env.USERPROFILE || '', '.cache/codex-runtimes/codex-primary-runtime/dependencies/python/python.exe');
function run(file){return new Promise((resolve,reject)=>{const p=spawn(PY,['-u',file],{cwd:ROOT,stdio:'inherit',windowsHide:true,env:{...process.env,PYTHONIOENCODING:'utf-8'}});p.on('error',reject);p.on('exit',code=>code===0?resolve():reject(new Error(file+' failed: '+code)));});}
const mdTable=(headers,rows)=>['|'+headers.join('|')+'|','|'+headers.map(()=>'---').join('|')+'|',...rows.map(r=>'|'+r.map(v=>String(v ?? '—').replaceAll('|',' / ')).join('|')+'|')].join('\n');
export function writeReports(d){
 const publicData={metadata:d.metadata,months:[...MONTHS],sheetStats:d.raw.sheetStats,rawRowCount:d.raw.rows.length,dedupRowCount:d.fullDedup.length,categories:{},movements:d.movements};
 for(const [scope,cat] of Object.entries(d.categories)){const {fullRows,topRows,...rest}=cat;publicData.categories[scope]=rest;}
 const base='交付/户外地垫市场分析';
 fs.mkdirSync('交付',{recursive:true});
 let html=buildHtml(d.raw,d.categories).replaceAll('SPEC 2.0','SPEC 2.1').replace('DATA · VERIFIED','DATA · SOURCE').replace('市场总览｜范围与数据口径','数据总览').replace('整体盘去重 Listing</span>','整体盘 Listing月次</span>').replace('BSR Top100 去重</span>','BSR Top100 Listing月次</span>').replace('GENIMO Listing</span>','GENIMO Listing月次</span>');
 html=html.replace('</head>','<script id="report-metadata" type="application/json">'+JSON.stringify(d.metadata).replaceAll('<','\\u003c')+'</script></head>');
 html=html.replace('原始字段只用于市场统计、筛选和回勾。','原始字段只用于市场统计、筛选和回勾。年度YOY仅取相邻年共同覆盖月份；BSR数值1—100筛选后保留全部去重商品；PP与高客单价范围相加等于整体。');
 const preamble='# 户外地垫市场分析 SPEC 2.1\n\n'+Object.entries(d.metadata).map(([k,v])=>'- '+k+'：'+v).join('\n')+'\n\n整体=PP+高客单价；GENIMO是品牌视角。月度MOM=本月/去年同月−1；年度YOY=相邻年共同月份。无有效值不补零。源估算不等于真实成交。\n';
 let md=preamble,quick=preamble;
 for(const [index,scope] of ['overall','pp','nonpp','genimo'].entries()){
  const cat=d.categories[scope];const analysis=narrative(cat,d.categories.overall,d.categories.pp,d.categories.nonpp);
  const text='\n## '+(index+1)+' '+cat.title+'\n\n'+analysis.map(t=>'- '+t).join('\n')+'\n';md+=text;quick+=text;
  for(const [title,rows] of [['月度汇总',cat.monthly],['BSR Top100月度',cat.topMonthly]]) md+='\n### '+title+'\n\n'+mdTable(['月份','Listing','销量','金额($)','标价($)','配对均价($)','销量MOM','金额MOM','基期Listing','可比性'],rows.map(r=>[r.month,fmt(r.count),fmt(r.sales),fmt(r.revenue,2),fmt(r.avgPrice,2),fmt(r.pairedAsp,2),fmtPct(r.momSales),fmtPct(r.momRevenue),fmt(r.comparisonCount),r.quality]))+'\n';
  md+='\n### 年度总量与共同月份YOY\n\n'+mdTable(['年份','本期月份','同比期间','Listing月次','销量','金额($)','销量YOY','金额YOY','同比本期销量','同比基期销量','同比本期金额','同比基期金额'],cat.annual.map(r=>[r.year,r.coverage,r.yoyPeriod,fmt(r.count),fmt(r.sales),fmt(r.revenue,2),fmtPct(r.yoySales),fmtPct(r.yoyRevenue),fmt(r.currentComparable.sales),fmt(r.previousComparable.sales),fmt(r.currentComparable.revenue,2),fmt(r.previousComparable.revenue,2)]))+'\n';
  for(const band of [...BANDS,...FINE_BANDS])md+='\n### BSR '+band.name+'\n\n'+mdTable(['月份','Listing','销量','金额($)','销量MOM','金额MOM'],(cat.tiers[band.key]||cat.fine[band.key]).map(r=>[r.month,fmt(r.count),fmt(r.sales),fmt(r.revenue,2),fmtPct(r.momSales),fmtPct(r.momRevenue)]))+'\n';
 }
 const planning='\n## 2027 规划使用方法\n\n- 在XLSX的08_2027规划与分析!B3输入非负整数新增链接总量。默认空，不生成虚假的计划数。\n- 权重按最近核心期GENIMO整体榜内头/中/尾销量比例；前两档向下取整，尾档取余；仅为分配情景。\n- 连续两个月金额、份额与BSR改善且样本覆盖稳定，再考虑扩充；覆盖受限先核验，退出榜单不等于下架。\n- 花型、尺寸、颜色尚未可靠提取，不据此编造趋势。仅使用给定源字段，不计算利润。\n';md+=planning;quick+=planning;
 const files={[base+'报告-优化版.html']:html,[base+'数据.json']:JSON.stringify(publicData,null,2),[base+'报告-优化版.md']:md,[base+'报告-极速版.md']:quick};
 for(const [f,content] of Object.entries(files))fs.writeFileSync(f,content,'utf8');
 const workbook='outputs/20260911-spec2-market-analysis/户外地垫市场分析-SPEC2-可回勾版.xlsx';
 const manifest={...d.metadata,files:[...Object.keys(files),workbook].map(f=>({path:f,sha256:crypto.createHash('sha256').update(fs.readFileSync(f)).digest('hex'),bytes:fs.statSync(f).size})),formulaAudit:JSON.parse(fs.readFileSync('tmp/spec21_formula_audit.json','utf8'))};
 fs.writeFileSync('交付/构建清单-SPEC2.1.json',JSON.stringify(manifest,null,2));
 console.log(JSON.stringify({delivered:manifest.files,formulaCount:manifest.formulaAudit.formulaCount},null,2));
}
export async function buildDelivery(){
 const d=prepare();
 await run('src/prepare_spec21_workbook.py');
 await run('src/build_spec21_xlsx.py');
 writeReports(d);
}
if(process.argv[1]&&import.meta.url===pathToFileURL(path.resolve(process.argv[1])).href)await buildDelivery();
