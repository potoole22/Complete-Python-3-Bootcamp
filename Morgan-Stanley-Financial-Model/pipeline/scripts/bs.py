import sys, os, re, json, glob
sys.path.insert(0, os.path.dirname(__file__))
from tables import parse
from extract import tonum
MAP=[ # (key, regex on lowercase label)  -- summed if several match
 ('cash', r'^(cash and due from banks|interest bearing deposits with banks|restricted cash|cash and cash equivalents)'),
 ('trading_assets', r'^trading assets at fair value'),
 ('inv_sec', r'^(investment securities|available-for-sale at fair value|held-to-maturity)'),
 ('resale', r'^securities purchased under agreements to resell'),
 ('sec_borrowed', r'^securities borrowed'),
 ('cust_recv', r'^customer and other receivables'),
 ('loans_hfi', r'^held for investment'),
 ('loans_hfs', r'^held for sale'),
 ('goodwill', r'^goodwill'),
 ('intangibles', r'^intangible assets'),
 ('other_assets', r'^other assets'),
 ('total_assets', r'^total assets'),
 ('deposits', r'^deposits'),
 ('trading_liab', r'^trading liabilities at fair value'),
 ('repo', r'^securities sold under agreements to repurchase'),
 ('sec_loaned', r'^securities loaned'),
 ('other_sec_fin', r'^other secured financings'),
 ('cust_pay', r'^customer and other payables'),
 ('other_liab', r'^other liabilities and accrued expenses'),
 ('borrowings', r'^borrowings'),
 ('total_liab', r'^total liabilities$'),
 ('pref_stock', r'^preferred stock'),
 ('ms_equity', r'^total morgan stanley shareholders'),
 ('nci', r'^noncontrolling interests'),
 ('total_equity', r'^total equity'),
 ('total_le', r'^total liabilities and equity'),
]
def parse_bs(path):
    rows=parse(path)[0]
    out={}
    for r in rows[1:]:
        lab=r[0].lower().replace('’',"'")
        if len(r)<2: continue
        v=tonum(r[1])
        if v is None: continue
        for k,pat in MAP:
            if re.search(pat, lab):
                out[k]=out.get(k,0)+v; break
    return out
if __name__=='__main__':
    res={}
    for f in sorted(glob.glob('data/rfiles/*_bs.htm')):
        rep=os.path.basename(f)[:10]
        y,m=int(rep[:4]),int(rep[5:7]); q=f"{y}q{(m-1)//3+1}"
        res[q]=parse_bs(f)
    json.dump(res,open('data/bs.json','w'),indent=0)
    keys=[k for k,_ in MAP]
    for q,d in res.items():
        a=sum(d.get(k,0) for k in keys[:11]); l=sum(d.get(k,0) for k in keys[12:20])
        print(q, 'A chk', round(a-d['total_assets']), 'L chk', round(l-d['total_liab']), 'E chk', round(d['total_liab']+d['total_equity']-d['total_le']), 'MS eq', d['ms_equity']+d['nci']-d['total_equity'], 'missing', [k for k in keys if k not in d])
