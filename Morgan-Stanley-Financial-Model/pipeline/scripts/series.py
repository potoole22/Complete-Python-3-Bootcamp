import json, re, sys, os
sys.path.insert(0, os.path.dirname(__file__))
from classify import classify
D=json.load(open(os.path.join(os.path.dirname(__file__),'..','data','extracted.json')))
def norm(l):
    l=re.sub(r'\(\s*\d\s*\)','',l)          # footnotes
    l=re.sub(r'\s+',' ',l).strip().lower()
    l=l.replace('‐','-').replace('’',"'")
    return l
DOCS=sorted(D.keys(), key=lambda k:(k[3:], k[:2]))   # by quarter, er before fs
CLS={k:classify(D[k]) for k in D}
def qkey(m,y): return f"{y}q{(m-1)//3+1}"
def find_rows(doc, cls, pat, occ=0, ctxpat=None):
    """return (table, row) for occ-th row whose normalized label matches regex pat (fullmatch) in first table of class cls"""
    tabs=[t for t,c in zip(D[doc],CLS[doc]) if c==cls]
    if not tabs: return None
    t=tabs[0]
    hits=[r for r in t['rows'] if re.fullmatch(pat, norm(r['label'])) and (ctxpat is None or any(re.search(ctxpat, norm(c)) for c in r['ctx']))]
    if len(hits)<=occ: return None
    return t, hits[occ]
def vals(t, r):
    out={}
    for i,(m,y) in enumerate(t['q']):
        if i < len(r['nums']):
            v=r['nums'][i]
            if v=='X': v=None
            out[qkey(m,y)]=v
    return out
def get_series(rules, docs=None, prefer='latest'):
    """rules: list of (doc_prefix, cls, pat, occ, ctxpat, transform) ; returns {q:(val, doc)}"""
    res={}
    for doc in (docs or DOCS):
        for rule in rules:
            pref,cls,pat,occ=rule[:4]; ctxpat=rule[4] if len(rule)>4 else None
            if not doc.startswith(pref): continue
            if isinstance(pat, tuple):   # combo: sum of several patterns
                parts=[find_rows(doc,cls,p,o,cp) for p,o,cp in pat]
                if any(x is None for x in parts): continue
                vs=[vals(*x) for x in parts]
                tot={}
                for q in vs[0]:
                    if all(q in v and v[q] is not None for v in vs): tot[q]=sum(v[q] for v in vs)
                v=tot
            else:
                x=find_rows(doc,cls,pat,occ,ctxpat)
                if x is None: continue
                v=vals(*x)
            for q,val in v.items():
                if val is None: continue
                if prefer=='latest' or q not in res: res[q]=(val,doc)
            break
    return res
