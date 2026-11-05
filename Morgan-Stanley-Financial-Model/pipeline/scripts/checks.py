import json
S=json.load(open('data/series.json'))['latest']   # run from pipeline/
Q=[f"{y}q{q}" for y in range(2018,2027) for q in range(1,5)][:35]
def g(n,q): return S.get(n,{}).get(q,0) or 0
# Intersegment expense / pre-tax / net income eliminations are derived in dataset.py
# (consolidated less sum of segments), so they tie by construction and are not checked here.
checks={
 'seg sum vs net_rev': lambda q: g('seg_rev_is',q)+g('seg_rev_wm',q)+g('seg_rev_im',q)+g('seg_rev_elim',q)-g('net_rev',q),
 'c_net_rev vs net_rev': lambda q: g('c_net_rev',q)-g('net_rev',q),
 'IS ib parts': lambda q: g('is_advisory',q)+g('is_eq_uw',q)+g('is_fi_uw',q)-g('is_ib',q),
 'IS rev parts': lambda q: g('is_ib',q)+g('is_eq_st',q)+g('is_fi_st',q)+g('is_other',q)-g('is_net_rev',q),
 'IS rev vs summary': lambda q: g('is_net_rev',q)-g('seg_rev_is',q),
 'IS pretax': lambda q: g('is_net_rev',q)-g('is_provision',q)-g('is_comp',q)-g('is_noncomp',q)-g('is_pretax',q),
 'IS exp': lambda q: g('is_comp',q)+g('is_noncomp',q)-g('is_exp',q),
 'WM rev parts': lambda q: g('wm_am',q)+g('wm_trans',q)+g('wm_nii',q)+g('wm_other',q)-g('wm_net_rev',q),
 'WM rev vs summary': lambda q: g('wm_net_rev',q)-g('seg_rev_wm',q),
 'WM pretax': lambda q: g('wm_net_rev',q)-g('wm_provision',q)-g('wm_comp',q)-g('wm_noncomp',q)-g('wm_pretax',q),
 'IM rev parts': lambda q: g('im_am',q)+g('im_perf_other',q)-g('im_net_rev',q),
 'IM rev vs summary': lambda q: g('im_net_rev',q)-g('seg_rev_im',q),
 'IM pretax': lambda q: g('im_net_rev',q)-g('im_comp',q)-g('im_noncomp',q)-g('im_pretax',q),
 'C nonint rev': lambda q: sum(g(n,q) for n in ['c_ib','c_trading','c_investments','c_commissions','c_asset_mgmt','c_other_rev'])-g('c_nonint_rev',q),
 'C net int': lambda q: g('c_int_inc',q)-g('c_int_exp',q)-g('c_net_int',q),
 'C net rev': lambda q: g('c_nonint_rev',q)+g('c_net_int',q)-g('c_net_rev',q),
 'C noncomp': lambda q: sum(g(n,q) for n in ['c_brokerage','c_infoproc','c_profserv','c_occupancy','c_marketing','c_other_exp'])-g('c_noncomp',q),
 'C exp': lambda q: g('c_comp',q)+g('c_noncomp',q)-g('c_nonint_exp',q),
 'C pretax': lambda q: g('c_net_rev',q)-g('c_provision',q)-g('c_nonint_exp',q)-g('c_pretax',q),
 'C NI': lambda q: g('c_pretax',q)-g('c_tax',q)+g('c_discops',q)-g('c_net_income',q),
 'C NI MS': lambda q: g('c_net_income',q)-g('c_nci',q)-g('c_ni_ms',q),
 'C common': lambda q: g('c_ni_ms',q)-g('c_pref_div',q)-g('c_common_earn',q),
 'seg prov vs C': lambda q: g('is_provision',q)+g('wm_provision',q)-g('c_provision',q),
 'TCE': lambda q: g('common_equity',q)+g('gw_intang',q)-g('tce',q),
}
for name,f in checks.items():
    bad=[(q,round(f(q),1)) for q in Q if abs(f(q))>1.01]
    print(f"{name:22s}", 'OK' if not bad else bad[:12])
