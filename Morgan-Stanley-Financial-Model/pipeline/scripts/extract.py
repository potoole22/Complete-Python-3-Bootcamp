import re, json, sys, glob, os, datetime
sys.path.insert(0, os.path.dirname(__file__))
from tables import parse
MON={m:i+1 for i,m in enumerate(['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'])}
DATE_RE=re.compile(r'(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\.? ?(\d{1,2}), ?(\d{4})')
NUM_RE=re.compile(r'^\(?-?\$?\s*[\d,]*\.?\d+\s*\)?%?$')
def qtag(m,y):
    return f"{y}q{(m-1)//3+1}"
def tonum(c):
    c=c.strip().replace('$','').replace(' ','')
    if c in ('—','–','-','--','— ','—%'): return 0.0
    if c.endswith('%'): c=c[:-1]
    neg=c.startswith('(')
    c=c.strip('()')
    c=c.replace(',','')
    try: v=float(c)
    except: return None
    return -v if neg else v
def is_num(c):
    c=c.strip()
    if c in ('—','–','-','--'): return True
    return bool(NUM_RE.match(c.replace(' ','')))
def extract(path):
    out=[]; pending=[]
    for ti,rows in enumerate(parse(path)):
        title=[]; periods=None; qcount=0; data=[]; ctx=[]
        for r in rows:
            joined=' | '.join(r)
            dates=DATE_RE.findall(joined)
            if periods is None:
                if len(dates)>=2 and all(DATE_RE.search(c) or c.lower().startswith(('quarter','three','nine','six','twelve','year','percentage','change')) for c in r):
                    ds=[(MON[m[:3]],int(y)) for m,d,y in dates]
                    periods=ds
                    continue
                title.append(joined); continue
            label=r[0]
            vals=r[1:]
            if not vals or not any(is_num(v) for v in vals):
                if not is_num(label): ctx.append(label)
                continue
            nums=[]
            for v in vals:
                if v.strip()=='*': nums.append(None); continue
                if is_num(v): nums.append(tonum(v))
                else: nums.append('X')
            data.append({'label':label,'ctx':ctx[-3:],'nums':nums,'raw':vals})
        if periods and data:
            t=[x for x in title if not x.lower().startswith(('quarter','(unaudited','three','nine','six','twelve'))]
            if not t: t=pending
            # quarter columns: leading distinct dates until first repeat
            seen=[]
            for d in periods:
                if d in seen: break
                seen.append(d)
            out.append({'ti':ti,'title':' / '.join(t[:3]),'periods':periods,'q':seen[:3] if len(seen)>=3 else seen,'rows':data})
            pending=[]
        elif not periods:
            txt=[' | '.join(r) for r in rows if len(' | '.join(r))<150]
            if txt: pending=txt[-3:]
    return out
if __name__=='__main__':
    res={}
    for f in sorted(glob.glob(sys.argv[1]+'/*.htm')):
        res[os.path.basename(f)[:-4]]=extract(f)
    json.dump(res,open(sys.argv[2],'w'))
