"""Bounded independent evaluator for the nonvolatile Excel formula subset in SPEC 2.1.

Unsupported syntax/functions fail closed. This is not a claim of desktop Excel validation.
"""
import re,math,bisect
from collections import defaultdict
TOKEN=re.compile(r'\s*(?:(\d+(?:\.\d+)?)|("(?:[^"]|"")*")|((?:\$?[A-Z]+\$?\d+)(?::\$?[A-Z]+\$?\d+)?)|([A-Z][A-Z0-9_]*)|(\'(?:[^\']|\'\')*\')|(<=|>=|<>|[=<>+*/&(),!:\-]))')
def ci(s):
 n=0
 for ch in s:n=n*26+ord(ch)-64
 return n
def addr(s):
 a=re.fullmatch(r'\$?([A-Z]+)\$?(\d+)',s);return int(a[2]),ci(a[1])
class Error(Exception):pass
def number(x):
 if x is None:return 0
 if isinstance(x,(float,int)):return x
 try:return float(x)
 except:raise Error('VALUE '+str(x))
def isnum(x):return isinstance(x,(float,int)) and not isinstance(x,bool)
class Engine:
 def __init__(self,sheets):self.sheets=sheets;self.cache={};self.busy=set();self.index={};self.count=0
 def get(self,s,r,c):
  key=(s,r,c)
  if key in self.cache:return self.cache[key]
  v=self.sheets.get(s,{}).get((r,c))
  if not isinstance(v,dict):return v
  if key in self.busy:raise Error('CYCLE '+str(key))
  self.busy.add(key)
  try:value=self.eval(self.parse(v['f'][1:]),s)
  except Exception as ex:raise Error(f'{s}!{r},{c} {v["f"][:150]}: {ex}') from ex
  self.busy.remove(key);self.cache[key]=value;self.count+=1;return value
 def parse(self,f):
  ts=[];p=0
  while p<len(f):
   m=TOKEN.match(f,p)
   if not m:raise Error('SYNTAX '+f[p:p+40])
   ts.append(next((i+1,v) for i,v in enumerate(m.groups()) if v is not None));p=m.end()
  i=0
  def expression(minp=0):
   nonlocal i
   kind,v=ts[i];i+=1
   if kind==1:left=('lit',float(v) if '.' in v else int(v))
   elif kind==2:left=('lit',v[1:-1].replace('""','"'))
   elif kind==5:
    assert ts[i][1]=='!';i+=1;left=('ref',v[1:-1].replace("''","'"),ts[i][1]);i+=1
   elif kind==3:left=('ref',None,v)
   elif v=='(':
    left=expression();assert ts[i][1]==')';i+=1
   elif v=='-':left=('neg',expression(6))
   elif v in ['TRUE','FALSE'] and (i==len(ts) or ts[i][1]!='('):left=('lit',v=='TRUE')
   elif kind==4:
    assert ts[i][1]=='(';i+=1;args=[]
    if ts[i][1]!=')':
     while True:
      args.append(expression())
      if ts[i][1]!=',':break
      i+=1
    assert ts[i][1]==')';i+=1;left=('fn',v,args)
   else:raise Error('TOKEN '+v)
   prec={'=':1,'<>':1,'<':1,'>':1,'<=':1,'>=':1,'&':2,'+':3,'-':3,'*':4,'/':4}
   while i<len(ts) and ts[i][1] in prec and prec[ts[i][1]]>=minp:
    op=ts[i][1];i+=1;left=('op',op,left,expression(prec[op]+1))
   return left
  tree=expression()
  if i!=len(ts):raise Error('TRAILING '+str(ts[i:]))
  return tree
 def range(self,node,s):
  if node[0]!='ref':raise Error('EXPECTED RANGE')
  s=node[1] or s;parts=node[2].split(':');a=addr(parts[0]);b=addr(parts[-1]);return s,a[0],b[0],a[1],b[1]
 def values(self,node,s):
  if node[0]=='ref' and ':' in node[2]:
   sn,a,b,c,e=self.range(node,s);return [self.get(sn,r,col) for r in range(a,b+1) for col in range(c,e+1)]
  return [self.eval(node,s)]
 def compare(self,a,b,op):
  if a is None:a='' if isinstance(b,str) else 0
  if b is None:b='' if isinstance(a,str) else 0
  # Excel criteria comparisons coerce numeric text (for example YYYYMM
  # strings) when the criterion is numeric. Preserve ordinary text ordering
  # for nonnumeric values.
  if isinstance(a,str) and isnum(b):
   try:a=float(a)
   except ValueError:
    return op == '<>'
  if isnum(a) and isinstance(b,str):
   try:b=float(b)
   except ValueError:
    return op == '<>'
  if isinstance(a,str) and isinstance(b,str):a=a.casefold();b=b.casefold()
  if type(a)!=type(b) and isinstance(a,str)!=isinstance(b,str):
   if op=='=':return False
   if op=='<>':return True
   a=(1,a) if isinstance(a,str) else (0,a);b=(1,b) if isinstance(b,str) else (0,b)
  return {'=':lambda:a==b,'<>':lambda:a!=b,'<':lambda:a<b,'>':lambda:a>b,'<=':lambda:a<=b,'>=':lambda:a>=b}[op]()
 def eval(self,n,s):
  typ=n[0]
  if typ=='lit':return n[1]
  if typ=='ref':
   if ':' in n[2]:raise Error('RANGE AS SCALAR')
   return self.get(n[1] or s,*addr(n[2]))
  if typ=='neg':return -number(self.eval(n[1],s))
  if typ=='op':
   op=n[1];a=self.eval(n[2],s);b=self.eval(n[3],s)
   if op in ['=','<>','<','>','<=','>=']:return self.compare(a,b,op)
   if op=='&':return str(a if a is not None else '')+str(b if b is not None else '')
   a=number(a);b=number(b)
   if op=='+':return a+b
   if op=='-':return a-b
   if op=='*':return a*b
   if op=='/':
    if b==0:raise Error('DIV/0')
    return a/b
  name,args=n[1:]
  if name=='IF':return self.eval(args[1] if self.eval(args[0],s) else args[2],s)
  if name=='ISNUMBER':return isnum(self.eval(args[0],s))
  if name=='AND':return all(self.eval(x,s) for x in args)
  if name=='OR':return any(self.eval(x,s) for x in args)
  if name=='ABS':return abs(number(self.eval(args[0],s)))
  if name=='INT':return math.floor(number(self.eval(args[0],s)))
  if name=='ROW':return self.range(args[0],s)[1]
  if name in ['SUM','MIN','MAX','COUNTA']:
   vals=[v for x in args for v in self.values(x,s)]
   if name=='COUNTA':return sum(v is not None for v in vals)
   vals=[v for v in vals if isnum(v)]
   return sum(vals) if name=='SUM' else min(vals,default=0) if name=='MIN' else max(vals,default=0)
  if name=='INDEX':
   sn,a,b,c,e=self.range(args[0],s);idx=int(self.eval(args[1],s))
   if idx<1 or idx>b-a+1:raise Error('REF INDEX')
   return self.get(sn,a+idx-1,c)
  if name=='MATCH':
   target=self.eval(args[0],s);sn,a,b,c,e=self.range(args[1],s);k=(sn,a,b,c,'match')
   if k not in self.index:
    ix={}
    for r in range(a,b+1):ix.setdefault(self.get(sn,r,c),r-a+1)
    self.index[k]=ix
   if target not in self.index[k]:raise Error('N/A MATCH')
   return self.index[k][target]
  if name in ['SUMIFS','COUNTIFS','AVERAGEIFS']:
   offset=1 if name in ['SUMIFS','AVERAGEIFS'] else 0;criteria=[]
   for i in range(offset,len(args),2):
    rg=self.range(args[i],s);criterion=self.eval(args[i+1],s);op='='
    if isinstance(criterion,str):
     m=re.match(r'^(<=|>=|<>|<|>|=)(.*)$',criterion)
     if m:op,criterion=m.groups()
     try:criterion=float(criterion)
     except:pass
    criteria.append((rg,op,criterion))
   rg0=criteria[0][0];length=rg0[2]-rg0[1]+1
   matches=None
   # Cache per bounded column, then intersect row positions; eligible pools are small.
   for (sn,a,b,c,e),op,criterion in criteria:
    if b-a+1!=length:raise Error('RANGE LENGTH')
    k=(sn,a,b,c,'index')
    if len(self.index)>1200:self.index.clear()
    if k not in self.index:
     ix=defaultdict(set)
     for r in range(a,b+1):ix[self.get(sn,r,c)].add(r-a)
     self.index[k]=ix
    ix=self.index[k];selected=set()
    if op=='=':
     selected=ix.get(criterion,set())
     if isnum(criterion):
      selected=set(selected)
      for key,positions in ix.items():
       if isinstance(key,str):
        try:
         if float(key)==criterion:selected.update(positions)
        except ValueError:pass
     if isinstance(criterion,str):
      selected=set().union(*(v for key,v in ix.items() if isinstance(key,str) and key.casefold()==criterion.casefold()))
    else:
     selected=set().union(*(v for key,v in ix.items() if self.compare(key,criterion,op)))
    matches=set(selected) if matches is None else matches.intersection(selected)
    if not matches:break
   if name=='COUNTIFS':return len(matches or [])
   sn,a,b,c,e=self.range(args[0],s)
   vals=[v for pos in (matches or []) if isnum(v:=self.get(sn,a+pos,c))]
   if name=='AVERAGEIFS':
    if not vals:raise Error('DIV/0')
    return sum(vals)/len(vals)
   return sum(vals)
  raise Error('UNSUPPORTED FUNCTION '+name)
