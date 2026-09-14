"""Streaming fallback: full artifact-tool import exhausted the available JS heap.
Formula caches are computed from actual formula text by the independent evaluator.
"""
import json,os,time,math
from pathlib import Path
import xlsxwriter
from formula_engine import Engine
ROOT=Path(__file__).resolve().parent.parent;plan=ROOT/'tmp/formula_market_builder/workbook_plan.jsonl'
sheets={};infos={};current=None;meta=None;names=[]
with open(plan,encoding='utf-8') as f:
 for line in f:
  x=json.loads(line)
  if 'metadata' in x:meta=x['metadata'];names=x['names']
  elif 'sheet' in x:current=x['sheet'];infos[current]=x;sheets[current]={}
  elif 'row' in x:
   for c,v in enumerate(x['values'],1):
    if v is not None:sheets[current][(x['row'],c)]=v
print('Formula plan loaded',flush=True)
engine=Engine(sheets);differences=[];expected_count=0
for sn in ['90_原始输入','91_去重明细','92_聚合输入']+[x for x in names if x not in ['90_原始输入','91_去重明细','92_聚合输入']]:
 for (r,c),v in sheets[sn].items():
  if isinstance(v,dict):
   actual=engine.get(sn,r,c);expected=v.get('v')
   if 'v' in v:
    expected_count+=1
    equal=math.isclose(actual,expected,rel_tol=1e-10,abs_tol=1e-6) if isinstance(expected,(int,float)) and isinstance(actual,(int,float)) else actual in [None,''] if expected is None else actual==expected
    if not equal:differences.append({'sheet':sn,'row':r,'col':c,'formula':v['f'],'actual':actual,'expected':expected})
 print(sn,'evaluated',engine.count,'differences',len(differences),flush=True)
 # Range indexes are transient accelerators, formula result cache remains authoritative.
 engine.index.clear()
audit={'engine':'SPEC2.1 explicit Excel subset evaluator (not desktop Excel)','formulaCount':engine.count,'independentExpectedComparisons':expected_count,'mismatches':differences,'mismatchCount':len(differences),'metadata':meta}
(ROOT/'tmp/spec21_formula_audit.json').write_text(json.dumps(audit,ensure_ascii=False,indent=2),encoding='utf-8')
if differences:raise RuntimeError('Formula evaluation differs from independently computed source model: '+str(len(differences)))
outfile=ROOT/'outputs/20260911-spec2-market-analysis/户外地垫市场分析-SPEC2-可回勾版.xlsx'
staging=outfile.with_name('spec21-staging.xlsx');wb=xlsxwriter.Workbook(str(staging),{'constant_memory':True,'strings_to_urls':False});wb.set_calc_mode('auto')
wb.set_properties({'title':'户外地垫市场分析 SPEC2.1','comments':json.dumps(meta,ensure_ascii=False)})
formats={
 'title':wb.add_format({'bold':True,'font_size':18,'font_color':'#FFFFFF','bg_color':'#1F2937','valign':'vcenter'}),
 'note':wb.add_format({'font_size':11,'text_wrap':True,'valign':'top','bg_color':'#EFF6FF','font_color':'#374151'}),
 'header':wb.add_format({'bold':True,'text_wrap':True,'valign':'vcenter','font_color':'white','bg_color':'#2563EB','border':1,'border_color':'#BFDBFE'}),
 'text':wb.add_format({'font_size':11,'valign':'top'}),
 'input':wb.add_format({'font_size':11,'font_color':'#2563EB','num_format':'#,##0.00;[Red]-#,##0.00','valign':'top'}),
 'formula':wb.add_format({'font_size':11,'font_color':'#111827','num_format':'#,##0.00;[Red]-#,##0.00','valign':'top'}),
 'integer':wb.add_format({'font_size':11,'num_format':'#,##0;[Red]-#,##0','valign':'top'}),
 'percent':wb.add_format({'font_size':11,'num_format':'0.0%;[Red]-0.0%','valign':'top'}),
 'wrapped':wb.add_format({'font_size':11,'text_wrap':True,'valign':'top'})}
for sn in names:
 info=infos[sn];s=wb.add_worksheet(sn);s.hide_gridlines(2);s.freeze_panes(4,2);s.set_column(0,len(info['headers'])-1,19);s.set_column(0,0,15)
 if sn in ['05_年度YOY','06_BSR分层']:
  s.set_column('D:D',30);s.set_column('E:E',60)
 for c,w in info['widths'].items():s.set_column(c+':'+c,w)
 s.merge_range(0,0,0,max(7,len(info['headers'])-1),info['title'],formats['title']);s.set_row(0,34)
 s.merge_range(1,0,1,max(7,len(info['headers'])-1),info['note'],formats['note']);s.set_row(1,55)
 for c,h in enumerate(info['headers']):s.write(3,c,h,formats['header'])
 s.set_row(3,46)
 active_headers=info['headers']
 for (r,c),v in sorted(sheets[sn].items()):
  secondary_header=r>4 and sheets[sn].get((r,1)) in ['月份','年份','候选市场']
  if secondary_header and c==1:active_headers=[sheets[sn].get((r,j),'') for j in range(1,max(cc for rr,cc in sheets[sn] if rr==r)+1)]
  header=active_headers[c-1] if c<=len(active_headers) else ''
  if sn=='08_2027规划与分析' and 12<=r<31:header=''
  fmt=formats['percent'] if (any(x in header for x in ['份额','权重']) or (any(x in header for x in ['MOM','YOY']) and not any(x in header for x in ['期间','基期','本期']))) else formats['integer'] if any(x in header for x in ['数量','有效数','行号','Listing','标记','入池','候选','成员数','索引','BSR下限','BSR上限']) else formats['formula']
  if isinstance(v,dict):
   value=engine.get(sn,r,c)
   if isinstance(value,str):fmt=formats['wrapped'] if len(value)>28 else formats['text']
   s.write_formula(r-1,c-1,v['f'],fmt,'' if value is None else value)
  else:s.write(r-1,c-1,v,formats['header'] if secondary_header else formats['input'] if isinstance(v,(int,float)) else formats['wrapped'] if len(str(v))>28 else formats['text'])
  if secondary_header:s.set_row(r-1,46)
  if sn in ['05_年度YOY','06_BSR分层'] and isinstance(v,str) and c in [4,5] and len(v)>25:s.set_row(r-1,65)
  if not isinstance(v,dict) and isinstance(v,str) and len(v)>70 and sn not in ['90_原始输入','91_去重明细']:s.set_row(r-1,65)
  if sn in ['90_原始输入','91_去重明细'] and c==(6 if sn=='90_原始输入' else 7):
   title=engine.get(sn,r,c) or '';s.set_row(r-1,min(150,max(30,math.ceil(len(str(title))/70)*14+6)))
 if sn=='08_2027规划与分析':
  s.write_blank('B3',None,wb.add_format({'bg_color':'#FEF3C7','font_color':'#2563EB','num_format':'0','border':1}));s.data_validation('B3',{'validate':'integer','criteria':'>=','value':0,'ignore_blank':True,'input_title':'目标新增链接数','input_message':'输入非负整数；空白不生成计划。'})
 if sn in names[1:4]:
  for idx,(c,title) in enumerate([(2,'月销量'),(3,'月销售额($)'),(4,'平均标价($)'),(6,'销量MOM（去年同月）')]):
   chart=wb.add_chart({'type':'line' if idx>=2 else 'column'});chart.add_series({'name':title,'categories':[sn,4,0,53,0],'values':[sn,4,c,53,c]});chart.set_title({'name':title});chart.set_legend({'none':True});chart.set_size({'width':780,'height':310});s.insert_chart(58+idx*17,0,chart)
  scope_index=names.index(sn)-1
  for chart_index,band_indices in enumerate([range(3),range(3,8)]):
   chart=wb.add_chart({'type':'line'})
   for bi in band_indices:
    start=4+scope_index*400+bi*50;chart.add_series({'name':['头部1—20','中部21—50','尾部51—100','1—5','6—10','11—20','21—50','51—100'][bi],'categories':['06_BSR分层',start,0,start+49,0],'values':['06_BSR分层',start,9,start+49,9]})
   chart.set_title({'name':'BSR分层销量MOM（去年同月）'});chart.set_size({'width':780,'height':340});s.insert_chart(128+chart_index*19,0,chart)
 if sn=='04_GENIMO品牌月度':
  for idx,(c,title) in enumerate([(2,'GENIMO月销量'),(3,'GENIMO月销售额($)'),(4,'GENIMO平均标价($)'),(6,'GENIMO销量MOM')]):
   chart=wb.add_chart({'type':'line' if idx>=2 else 'column'});chart.add_series({'name':title,'categories':[sn,4,0,53,0],'values':[sn,4,c,53,c]});chart.set_title({'name':title});chart.set_legend({'none':True});chart.set_size({'width':780,'height':310});s.insert_chart(224+idx*17,0,chart)
  for j,bis in enumerate([range(3),range(3,8)]):
   chart=wb.add_chart({'type':'line'})
   for bi in bis:
    start=1204+bi*50;chart.add_series({'name':['头部1—20','中部21—50','尾部51—100','1—5','6—10','11—20','21—50','51—100'][bi],'categories':['06_BSR分层',start,0,start+49,0],'values':['06_BSR分层',start,9,start+49,9]})
   chart.set_title({'name':'GENIMO BSR分层销量MOM'});chart.set_size({'width':780,'height':340});s.insert_chart(294+j*19,0,chart)
 if sn in ['90_原始输入','91_去重明细','92_聚合输入']:s.autofilter(3,0,max(r for r,c in sheets[sn])-1,len(info['headers'])-1)
wb.close();os.replace(staging,outfile)
# Preserve evaluated ranges for visual review and perturbation without re-evaluating the baseline.
with open(ROOT/'tmp/spec21_evaluated.jsonl','w',encoding='utf-8') as f:
 for sn in names:
  f.write(json.dumps({'sheet':sn,**infos[sn]},ensure_ascii=False)+'\n')
  widths={}
  for r,c in sheets[sn]:widths[r]=max(widths.get(r,0),c)
  rs=sorted(widths)
  for r in rs:
   maxc=widths[r]
   if r<=18 or sn not in ['90_原始输入','91_去重明细']:
    f.write(json.dumps({'row':r,'values':[engine.get(sn,r,c) for c in range(1,maxc+1)]},ensure_ascii=False)+'\n')
print('EXPORTED '+str(outfile),flush=True)
