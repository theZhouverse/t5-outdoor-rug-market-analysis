"""Independent ZIP/XML extraction and model reconciliation; never writes source XLSX."""
import json,zipfile,re,hashlib,math,collections,xml.etree.ElementTree as ET
from pathlib import Path
ROOT=Path(__file__).resolve().parent.parent
D=json.loads((ROOT/'tmp/formula_market_builder/spec2_dataset.json').read_text(encoding='utf-8'))
NS={'s':'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
def clean(v):
 s=re.sub(r'\r?\n',' ',str(v or '').replace('\xa0',' ')).strip()
 return '' if s.lower() in ['-','—','–','n/a','na','null','undefined'] else s
def num(v):
 s=clean(v)
 if not s:return None
 negative=s.startswith('(') and s.endswith(')');s=re.sub(r'[,\s$￥%()]','',s)
 try:return -float(s) if negative else float(s)
 except:return None
def bsr(v):
 s=clean(v).replace(',','');values=[]
 # Tokenize whole signed decimal numbers, then accept positive integers.
 for token in re.findall(r'(?<![\d.])[-−]?\s*\d+(?:\.\d+)?(?![\d.])',s):
  n=num(token.replace('−','-'))
  if n is not None and n>0 and n==int(n):values.append(int(n))
 return min(values) if values else None
def col(ref):return re.match('[A-Z]+',ref)[0]
source=ROOT/'data/raw/地垫-卖家精灵市场数据.xlsx';sha=hashlib.sha256(source.read_bytes()).hexdigest();assert sha==D['metadata']['sourceSha256']
raw=[]
with zipfile.ZipFile(source) as z:
 shared=[]
 if 'xl/sharedStrings.xml' in z.namelist():
  shared=[''.join(e.itertext()) for e in ET.fromstring(z.read('xl/sharedStrings.xml')).findall('s:si',NS)]
 relationships={e.attrib['Id']:e.attrib['Target'] for e in ET.fromstring(z.read('xl/_rels/workbook.xml.rels'))}
 for sh in ET.fromstring(z.read('xl/workbook.xml')).findall('s:sheets/s:sheet',NS):
  name=sh.attrib['name'];match=re.fullmatch(r'(\d{4})[.．]?(\d{1,2})',name)
  if not match:continue
  month=match[1]+match[2].zfill(2);target=relationships[sh.attrib['{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id']];target=target.lstrip('/') if target.startswith('/') else 'xl/'+target
  root=ET.fromstring(z.read(target));links={x.attrib['ref']:x.attrib.get('display','') for x in root.findall('s:hyperlinks/s:hyperlink',NS)};header=None
  for rr in root.findall('s:sheetData/s:row',NS):
   values={}
   for cell in rr:
    a=cell.attrib['r'];v=cell.find('s:v',NS);v=v.text if v is not None else None;t=cell.attrib.get('t')
    if t=='s':v=shared[int(v)] if v is not None else ''
    if t=='inlineStr':v=''.join(cell.find('s:is',NS).itertext())
    if not clean(v):v=links.get(a,'')
    values[col(a)]=clean(v)
   if header is None:
    if 'ASIN' in values.values() and any(x in values.values() for x in ['商品标题','品牌']):header={v:k for k,v in values.items()}
    continue
   def get(*keys):return next((values.get(header[k],'') for k in keys if k in header),'')
   row={'month':month,'sourceSheet':name,'sourceRow':int(rr.attrib['r']),'asin':get('ASIN'),'parent':get('父ASIN','父体ASIN'),'sku':get('SKU'),'brand':get('品牌'),'title':get('商品标题','标题'),'sourceBsr':get('小类BSR','小类目BSR','小类排名')}
   row.update(rank=bsr(row['sourceBsr']),sales=num(get('月销量','月销售量')),revenue=num(get('月销售额($)','月销售额')),price=num(get('价格($)','价格')))
   hasmetric=any(row[k] is not None for k in ['rank','sales','revenue','price']) or num(get('FBA($)','FBA')) is not None
   if not any(row[k] for k in ['asin','parent','sku','brand','title']) and not hasmetric:continue
   if not re.fullmatch(r'[A-Z0-9]{8,15}',row['asin'],re.I) and not any(row[k] for k in ['parent','sku','brand','title']) and not hasmetric:continue
   row['plastic']=bool(re.search(r'\bplastic\b',row['title'],re.I|re.ASCII));row['genimo']=row['brand'].lower()=='genimo';raw.append(row)
expected={(r['month'],r['sourceRow']):r for r in D['raw']};errors=[];checks=0
for r in raw:
 e=expected.get((r['month'],r['sourceRow']))
 if e is None:errors.append(['unexpected-row',r['month'],r['sourceRow']]);continue
 for k,v in r.items():
  checks+=1
  if e[k]!=v:errors.append([r['month'],r['sourceRow'],k,v,e[k]])
assert len(raw)==len(expected)
groups=collections.defaultdict(list)
for r in raw:groups[(r['month'],r['parent'] or r['asin'] or 'source:'+r['month']+'#'+str(r['sourceRow']))].append(r)
reps=[]
for key,rs in groups.items():
 rep=min(rs,key=lambda r:(r['rank'] if r['rank'] is not None else math.inf,-sum(r[k] is not None for k in ['sales','revenue']),-(r['price'] is not None),r['sourceRow'])).copy()
 rep.update(plastic=any(r['plastic'] for r in rs),genimo=any(r['genimo'] for r in rs),key=key[1]);reps.append(rep)
def metrics(rs):
 def sm(k):vs=[r[k] for r in rs if r[k] is not None];return sum(vs) if vs else None
 p=[r for r in rs if r['sales'] is not None and r['sales']>0 and r['revenue'] is not None];prices=[r['price'] for r in rs if r['price'] is not None]
 return dict(count=len(rs),sales=sm('sales'),revenue=sm('revenue'),avgPrice=sum(prices)/len(prices) if prices else None,pairedAsp=sum(r['revenue'] for r in p)/sum(r['sales'] for r in p) if p else None)
def same(a,b):return a is b if a is None or b is None else math.isclose(a,b,abs_tol=1e-6,rel_tol=1e-10)
for m in D['months']:
 full=[r for r in reps if r['month']==m];candidates=sorted([r for r in full if r['rank'] is not None and r['rank']<=100],key=lambda r:(r['rank'],r['key'],r['sourceRow']));top=candidates
 for scope in ['overall','pp','nonpp','genimo','genimoPP']:
  def eligible(r):return scope=='overall' or scope=='pp' and r['plastic'] or scope=='nonpp' and not r['plastic'] or scope=='genimo' and r['genimo'] or scope=='genimoPP' and r['genimo'] and r['plastic']
  pool=[r for r in full if eligible(r)];tp=[r for r in (top if scope.startswith('genimo') else candidates) if eligible(r)]
  for kind,rs in [('monthly',pool),('topMonthly',tp)]:
   a=metrics(rs);e=next(r for r in D['categories'][scope][kind] if r['month']==m)
   for k,v in a.items():
    checks+=1
    if not same(v,e[k]):errors.append([scope,m,kind,k,v,e[k]])
out={'sourceSha256':sha,'sourceRows':len(raw),'listingMonthRows':len(reps),'independentChecks':checks,'errors':errors}
(ROOT/'tmp/spec21_source_audit.json').write_text(json.dumps(out,ensure_ascii=False,indent=2),encoding='utf-8');print(json.dumps({k:v for k,v in out.items() if k!='errors'},ensure_ascii=False));assert not errors,errors[:10]
