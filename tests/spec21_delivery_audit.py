"""Reopen saved XLSX: all formula text/cache pairs, public outputs, units and headers."""
import json,zipfile,xml.etree.ElementTree as ET,hashlib,math,re,sys
from pathlib import Path
ROOT=Path(__file__).resolve().parent.parent
N={'s':'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
D=json.loads((ROOT/'交付/户外地垫市场分析数据.json').read_text(encoding='utf-8'));manifest=json.loads((ROOT/'交付/构建清单-SPEC2.1.json').read_text(encoding='utf-8'));assert D['metadata']=={k:manifest[k] for k in D['metadata']}
for item in manifest['files']:assert hashlib.sha256((ROOT/item['path']).read_bytes()).hexdigest()==item['sha256'],item['path']
plans={};sn=None
for line in open(ROOT/'tmp/formula_market_builder/workbook_plan.jsonl',encoding='utf-8'):
 x=json.loads(line)
 if 'sheet' in x:sn=x['sheet'];plans[sn]={}
 elif 'row' in x:
  for c,v in enumerate(x['values'],1):
   if isinstance(v,dict):plans[sn][(x['row'],c)]=v
def addr(a):
 p=re.fullmatch(r'([A-Z]+)(\d+)',a);c=0
 for x in p[1]:c=c*26+ord(x)-64
 return int(p[2]),c
errors=[];checks=0;formula_count=0;public={};snapshots=[]
file=ROOT/'outputs/20260911-spec2-market-analysis/户外地垫市场分析-SPEC2-可回勾版.xlsx'
with zipfile.ZipFile(file) as z:
 strings=[]
 if 'xl/sharedStrings.xml' in z.namelist():strings=[''.join(si.itertext()) for si in ET.fromstring(z.read('xl/sharedStrings.xml')).findall('s:si',N)]
 wb=ET.fromstring(z.read('xl/workbook.xml'));rel={e.attrib['Id']:e.attrib['Target'] for e in ET.fromstring(z.read('xl/_rels/workbook.xml.rels'))};sheets=wb.findall('s:sheets/s:sheet',N)
 assert len(sheets)==14
 styles=ET.fromstring(z.read('xl/styles.xml'));formats={int(x.attrib['numFmtId']):x.attrib['formatCode'] for x in styles.findall('s:numFmts/s:numFmt',N)};xfs=styles.findall('s:cellXfs/s:xf',N)
 def cell_format(tree,a):
  cell=tree.find('.//s:c[@r="'+a+'"]',N);return formats.get(int(xfs[int(cell.attrib.get('s',0))].attrib.get('numFmtId',0)),'General')
 annual=ET.fromstring(z.read('xl/worksheets/sheet6.xml'));bands=ET.fromstring(z.read('xl/worksheets/sheet7.xml'));plan=ET.fromstring(z.read('xl/worksheets/sheet9.xml'))
 assert float(annual.find('.//s:row[@r="6"]',N).attrib.get('ht',0))>=65
 assert '%' not in cell_format(bands,'J1610') and '%' in cell_format(bands,'K1610')
 assert all('%' in cell_format(plan,a) for a in ['D32','E32','F32','G32'])
 assert '%' not in cell_format(plan,'F13')
 for sheet in sheets:
  sn=sheet.attrib['name'];target=rel[sheet.attrib['{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id']];target=target.lstrip('/') if target.startswith('/') else 'xl/'+target
  public[sn]={};seen=0
  with z.open(target) as stream:
   for ev,e in ET.iterparse(stream,events=['end']):
    if e.tag.endswith('}c'):
     a=e.attrib.get('r');r,c=addr(a);f=e.find('s:f',N);v=e.find('s:v',N);t=e.attrib.get('t');value=v.text if v is not None else None
     if t=='inlineStr':value=''.join(e.find('s:is',N).itertext())
     elif t=='s':value=strings[int(value)]
     elif t=='b':value=value=='1'
     elif t=='str' and value is None:value=''
     elif t not in ['str','e'] and value is not None:value=float(value)
     if t=='e':errors.append([sn,a,'Excel error',value])
     if f is not None:
      seen+=1;formula_count+=1;p=plans[sn].get((r,c));checks+=1
      if not p or p['f'][1:]!=f.text:errors.append([sn,a,'formula serialization'])
      if p and 'v' in p:
       want=p['v'];checks+=1
       same=value in [None,''] if want in [None,''] else math.isclose(value,want,abs_tol=1e-6,rel_tol=1e-10) if isinstance(want,(int,float)) and isinstance(value,(int,float)) else value==want
       if not same:errors.append([sn,a,'cache/source model',value,want])
      if sn!='93_数据校验' and '93_数据校验' in (f.text or ''):errors.append([sn,a,'terminal check dependency'])
     if sn not in ['90_原始输入','91_去重明细'] or r<=18:public[sn][a]=value
     e.clear()
    elif e.tag.endswith('}row'):e.clear()
  assert seen==len(plans[sn]),(sn,seen,len(plans[sn]))
def same(a,b):return a in [None,''] if b is None else math.isclose(a,b,abs_tol=1e-6,rel_tol=1e-9) if isinstance(a,(int,float)) and isinstance(b,(int,float)) else a==b
for scope,sn in [('overall','01_整体市场月度'),('pp','02_PP市场月度'),('nonpp','03_高客单价市场月度')]:
 for i,m in enumerate(D['months'],5):
  for level,columns in [('monthly',{'count':'B','sales':'C','revenue':'D','avgPrice':'E','pairedAsp':'F','momSales':'G','momRevenue':'H','quality':'S'}),('topMonthly',{'count':'L','sales':'M','revenue':'N','avgPrice':'O','pairedAsp':'P','momSales':'Q','momRevenue':'R'})]:
   rec=D['categories'][scope][level][i-5]
   for field,c in columns.items():checks+=1;assert same(public[sn].get(c+str(i)),rec[field]),(sn,m,c,public[sn].get(c+str(i)),rec[field])
for block,(scope,level) in enumerate([('genimo','monthly'),('genimoPP','monthly'),('genimo','topMonthly'),('genimoIndependent','topMonthly')]):
 for i,rec in enumerate(D['categories'][scope][level],5+block*54):
  for field,c in [('count','B'),('sales','C'),('revenue','D'),('avgPrice','E'),('pairedAsp','F'),('momSales','G'),('momRevenue','H')]:checks+=1;assert same(public['04_GENIMO品牌月度'].get(c+str(i)),rec[field]),(scope,i,c)
for r in range(5,1605):
 if public['06_BSR分层'].get('F'+str(r))==0:
  assert public['06_BSR分层'].get('G'+str(r)) in [None,''];assert public['06_BSR分层'].get('H'+str(r)) in [None,'']
assert all(v=='通过' for k,v in public['93_数据校验'].items() if re.fullmatch(r'D\d+',k) and int(k[1:])>=5)
for scope in ['01_整体市场月度','02_PP市场月度','03_高客单价市场月度']:assert public[scope]['G4']=='销量MOM'
assert public['05_年度YOY']['F4']=='Listing月次'
assert public['08_2027规划与分析'].get('B3') is None
for f in ['户外地垫市场分析报告-优化版.md','户外地垫市场分析报告-极速版.md','户外地垫市场分析报告-优化版.html']:
 text=(ROOT/'交付'/f).read_text(encoding='utf-8');assert D['metadata']['batchId'] in text;assert D['metadata']['generatedAt'] in text;assert 'market.db' not in text;assert 'replacementMetadata' not in text
html=(ROOT/'交付/户外地垫市场分析报告-优化版.html').read_text(encoding='utf-8')
ids=re.findall(r'\bid="([^"]+)"',html);assert len(ids)==len(set(ids));assert 'genimo-4-7' in ids
for target in re.findall(r'href="#([^"]+)"',html):assert target in ids,target
for no,scope in enumerate(['overall','pp','nonpp'],1):
 for part in range(1,8):assert scope+'-1-'+str(part) in ids
assert html.count('data-interactive-chart=')>=26
assert not any(x in html for x in ['>NaN<','>undefined<','Infinity%'])
result={'savedFormulaCount':formula_count,'checks':checks,'errors':errors,'xlsxBytes':file.stat().st_size,'sheets':len(public),'metadata':D['metadata']}
(ROOT/'tmp/spec21_saved_audit.json').write_text(json.dumps(result,ensure_ascii=False,indent=2),encoding='utf-8');(ROOT/'tmp/spec21_public_cells.json').write_text(json.dumps(public,ensure_ascii=False),encoding='utf-8');print(json.dumps({**result,'errors':errors[:10],'errorCount':len(errors)},ensure_ascii=False));assert not errors,errors[:10]
