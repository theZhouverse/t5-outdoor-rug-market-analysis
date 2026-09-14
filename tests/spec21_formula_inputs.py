"""Input perturbation on the actual formula plan, whose saved XLSX text is verified separately."""
import sys,json,math,copy,collections
from pathlib import Path
ROOT=Path(__file__).resolve().parent.parent;sys.path.insert(0,str(ROOT/'src'))
from formula_engine import Engine,isnum
D=json.loads((ROOT/'tmp/formula_market_builder/spec2_dataset.json').read_text(encoding='utf-8'));sheets={};sn=None
for line in open(ROOT/'tmp/formula_market_builder/workbook_plan.jsonl',encoding='utf-8'):
 x=json.loads(line)
 if 'sheet' in x:sn=x['sheet'];sheets[sn]={}
 elif 'row' in x:
  for c,v in enumerate(x['values'],1):
   if v is not None:sheets[sn][(x['row'],c)]=v
def eq(a,b):return a in [None,''] if b is None else isinstance(a,(int,float)) and math.isclose(a,b,rel_tol=1e-9,abs_tol=1e-6)
def key(r):return 'source:'+r['month']+'#'+str(r['sourceRow'])
def summarize(rs):
 def sm(k):vs=[r[k] for r in rs if isnum(r[k])];return sum(vs) if vs else None
 ps=[r['price'] for r in rs if isnum(r['price'])];paired=[r for r in rs if isnum(r['sales']) and isnum(r['revenue']) and r['sales']>0]
 return [len(rs),sm('sales'),sm('revenue'),sum(ps)/len(ps) if ps else None,sum(r['revenue'] for r in paired)/sum(r['sales'] for r in paired) if paired else None]
def oracle(raw):
 # The current rule is deliberately row-level: first filter BSR 1..100,
 # then keep every candidate source row. Parent ASIN is never a grouping key.
 full=[r.copy() for r in raw if isnum(r['rank']) and 1<=r['rank']<=100]
 pools={}
 for m in [D['months'][-1],str(int(D['months'][-1][:4])-1)+D['months'][-1][4:]]:
  f=[r for r in full if r['month']==m];cand=sorted(f,key=lambda r:(r['rank'],key(r),r['sourceRow']));tp=cand
  for scope in ['overall','pp','nonpp','genimo','genimoPP']:
   def ok(r):return scope=='overall' or scope=='pp' and r['plastic'] or scope=='nonpp' and not r['plastic'] or scope=='genimo' and r['genimo'] or scope=='genimoPP' and r['genimo'] and r['plastic']
   pools[(scope,m,'full')]=[r for r in f if ok(r)];pools[(scope,m,'top')]=[r for r in (tp if scope.startswith('genimo') else cand) if ok(r)]
 return pools
latest=D['months'][-1];lr=len(D['months'])+4;results=[];checks=0
def verify(label,changes,plan_input=None):
 global checks
 work=[r.copy() for r in D['raw']]
 originals={}
 for idx,field,value in changes:
  work[idx][field]=value;c={'rank':9,'sales':11,'revenue':12,'price':13}[field];cell=(idx+5,c);originals[cell]=sheets['90_原始输入'].get(cell);sheets['90_原始输入'][cell]=value
 sheets['08_2027规划与分析'][(3,2)]=plan_input
 en=Engine(sheets);pools=oracle(work);readback={}
 for scope,sn,row in [('overall','01_整体市场月度',lr),('pp','02_PP市场月度',lr),('nonpp','03_高客单价市场月度',lr),('genimo','04_GENIMO品牌月度',lr),('genimoPP','04_GENIMO品牌月度',lr+len(D['months'])+4)]:
  want=summarize(pools[(scope,latest,'full')]);actual=[en.get(sn,row,c) for c in range(2,7)]
  for c,(a,b) in enumerate(zip(actual,want),2):checks+=1;assert eq(a,b),(label,sn,row,c,a,b)
  prior=summarize(pools[(scope,str(int(latest[:4])-1)+latest[4:],'full')])
  for c,j in [(7,1),(8,2)]:
   b=want[j]/prior[j]-1 if isnum(want[j]) and isnum(prior[j]) and prior[j]>0 else None;a=en.get(sn,row,c);checks+=1;assert eq(a,b),(label,'growth',scope,a,b)
  if scope in ['overall','pp','nonpp']:
   wanttop=summarize(pools[(scope,latest,'top')])
   for c,b in enumerate(wanttop,12):a=en.get(sn,row,c);checks+=1;assert eq(a,b),(label,'top',scope,c,a,b)
  else:
   for c,metric in [(9,1),(10,2)]:
    base=summarize(pools[('overall' if scope=='genimo' else 'pp',latest,'full')])[metric];target=want[metric]/base if isnum(want[metric]) and isnum(base) and base>0 else None;a=en.get(sn,row,c);checks+=1;assert eq(a,target),(label,'share',scope,a,target)
  readback[scope]=actual
 allocations=[en.get('08_2027规划与分析',r,7) for r in [5,6,7]]
 if plan_input is None:assert all(x in [None,''] for x in allocations)
 else:assert all(isnum(x) and int(x)==x and x>=0 for x in allocations) and sum(allocations)==plan_input
 checks+=1
 for cell,v in originals.items():sheets['90_原始输入'][cell]=v
 results.append({'case':label,'changes':changes,'latestMetrics':readback,'allocations':allocations});print(label,'OK',en.count,flush=True)
raw=D['raw'];idx=next(i for i,r in enumerate(raw) if r['month']==latest and isnum(r['rank']) and 1<=r['rank']<100);same_parent=[i for i,r in enumerate(raw) if r['month']==latest and r.get('parent')==raw[idx].get('parent') and i!=idx];other=same_parent[0] if same_parent else idx+1
verify('基线',[])
verify('销量金额标价联动',[(idx,'sales',raw[idx]['sales']+123),(idx,'revenue',raw[idx]['revenue']+4567),(idx,'price',raw[idx]['price']+7)])
verify('单行移出BSR候选池',[(idx,'rank',101)])
verify('同父ASIN其他行独立计入',[(other,'sales',raw[other]['sales']+7654),(other,'revenue',raw[other]['revenue']+456789),(other,'price',raw[other]['price']+59)])
verify('单行空值与真实零',[(idx,'sales',None),(idx,'revenue',None),(idx,'price',0)])
verify('恢复并输入23链接',[],23)
verify('零链接输入',[],0)
verify('恢复空输入',[])
(ROOT/'tmp/spec21_perturbation_audit.json').write_text(json.dumps({'checks':checks,'cases':results,'note':'独立公式求值；非桌面Excel。原始文件与正式XLSX均未修改。'},ensure_ascii=False,indent=2),encoding='utf-8')
