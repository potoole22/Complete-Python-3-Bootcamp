import json
S=json.load(open('data/series.json'))
L=S['latest']; SRC={k[:-5]:v for k,v in L.items() if k.endswith('__src')}
BS=json.load(open('data/bs.json'))
CR=json.load(open('data/capret.json'))
Q=[f"{y}q{q}" for y in range(2018,2027) for q in range(1,5)][:35]
END={1:'03-31',2:'06-30',3:'09-30',4:'12-31'}
def qe(q): return f"{q[:4]}-{END[int(q[-1])]}"
H={}
def put(name, vals): H[name]=[vals.get(q) for q in Q]
for k,v in L.items():
    if not k.endswith('__src'): put(k, v)
# provision zero pre-2020 (embedded in other revenues pre-CECL)
for n in ['c_provision','is_provision','wm_provision']:
    H[n]=[0.0 if v is None else v for v in H[n]]
H['c_discops']=[0.0 if v is None else v for v in H['c_discops']]
# percent -> fraction
for n in ['roe','rotce','etr','cet1_ratio_std','cet1_ratio_adv','t1_ratio_std','t1_ratio_adv','t1_lev','slr']:
    H[n]=[None if v is None else v/100 for v in H[n]]
# derived eliminations
def g(n,i): return H[n][i] or 0
H['elim_exp']=[round(g('c_nonint_exp',i)-g('is_exp',i)-g('wm_exp',i)-g('im_exp',i),1) for i in range(35)]
H['elim_ni']=[round(g('c_ni_ms',i)-g('is_ni',i)-g('wm_ni',i)-g('im_ni',i),1) for i in range(35)]
# balance sheet (10-Q/10-K); 3Q26 estimated below
for k in ['cash','trading_assets','inv_sec','resale','sec_borrowed','cust_recv','goodwill','intangibles','other_assets','total_assets','deposits','trading_liab','repo','sec_loaned','other_sec_fin','cust_pay','other_liab','borrowings','total_liab','pref_stock','ms_equity','nci','total_equity']:
    H['bs_'+k]=[BS.get(q,{}).get(k) for q in Q]
H['bs_loans']=[(BS[q]['loans_hfi']+BS[q]['loans_hfs']) if q in BS else None for q in Q]
H['bs_common_eq']=[(BS[q]['ms_equity']-BS[q]['pref_stock']) if q in BS else None for q in Q]
# capital return
bb={k:v/1e6 for k,v in CR['bb'].items()}; dv={k:v/1e6 for k,v in CR['dv'].items()}; dps=CR['dps']
dv.update({'2018-06-30':441.0,'2018-09-30':524.0,'2018-12-31':516.0})   # 2018 10-Q/10-K equity statements (total less preferred)
bb['2026-09-30']=1500.0                     # 3Q26 earnings release
dps['2026-09-30']=1.15                       # declared Jul-2026 (2Q26 release)
dv['2026-09-30']=round(1.15*1572,0)          # ESTIMATE: DPS x 2Q26 period-end shares (10-Q not yet filed)
H['buyback']=[bb.get(qe(q)) for q in Q]
H['div_common']=[dv.get(qe(q)) for q in Q]
H['dps']=[round(dps[qe(q)],3) if qe(q) in dps else None for q in Q]
# 3Q26 balance sheet estimate (10-Q not filed as of model date): anchor to supplement totals
i=34; p=33
ta=L['total_assets']['2026q3']; dep=L['deposits']['2026q3']; ce=L['common_equity']['2026q3']; pref=L['pref_equity']['2026q3']
gwi=-L['gw_intang']['2026q3']; gw=H['bs_goodwill'][p]; intang=round(gwi-gw,0)
loans=round(H['bs_loans'][p]*L['loans_bs']['2026q3']/L['loans_bs']['2026q2'],0)
borr=round(H['bs_borrowings'][p]*L['ltd']['2026q3']/L['ltd']['2026q2'],0)
nci=H['bs_nci'][p]
teq=ce+pref+nci; tl=ta-teq
oth_l=['trading_liab','repo','sec_loaned','other_sec_fin','cust_pay','other_liab']
base=sum(H['bs_'+k][p] for k in oth_l); tgt=tl-dep-borr; f=tgt/base
EST={}
for k in oth_l: EST[k]=round(H['bs_'+k][p]*f,0)
EST[oth_l[-1]]+= tgt-sum(EST[k] for k in oth_l)
oth_a=['cash','trading_assets','inv_sec','resale','sec_borrowed','cust_recv','other_assets']
base=sum(H['bs_'+k][p] for k in oth_a); tgt=ta-loans-gw-intang; f2=tgt/base
for k in oth_a: EST[k]=round(H['bs_'+k][p]*f2,0)
EST['cash']+= tgt-sum(EST[k] for k in oth_a)
EST.update({'loans':loans,'goodwill':gw,'intangibles':intang,'deposits':dep,'borrowings':borr,'pref_stock':pref,'common_eq':ce,'nci':nci,'total_assets':ta})
for k,v in EST.items(): H['bs_'+k][i]=v
H['bs_ms_equity'][i]=ce+pref; H['bs_total_liab'][i]=tl; H['bs_total_equity'][i]=teq
BS_EST={'liab_scale':f,'asset_scale':f2,'keys':list(EST.keys())}
SOURCES=SRC
if __name__=='__main__':
    json.dump({'Q':Q,'H':H,'BS_EST':BS_EST,'SRC':SRC},open('data/dataset.json','w'))
    print('3Q26 est', EST, f, f2)
