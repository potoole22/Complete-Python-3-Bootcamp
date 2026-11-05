import json, sys, os, datetime
sys.path.insert(0, os.path.dirname(__file__))
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.comments import Comment
from openpyxl.utils import get_column_letter as CL
from engine import *

HERE = os.path.dirname(__file__)
D = json.load(open(os.path.join(HERE, '..', 'data', 'dataset.json')))
H = D['H']
OUT = sys.argv[1] if len(sys.argv) > 1 else os.path.join(HERE, '..', 'MS_Financial_Model.xlsx')

# ----------------------------------------------------------------------------- helpers
def neg(name):
    H[name + '_neg'] = [None if v is None else -v for v in H[name]]
    return name + '_neg'

def yr(d):
    return lambda y, q, k: d[max(x for x in d if x <= y)]

def qq(d):
    return lambda y, q, k: d[q]

def c(v):
    return lambda y, q, k: v

def last4(fn):
    vals = [fn(i) for i in range(NH - 4, NH)]
    return sum(vals) / 4

def g(n, i):
    return H[n][i]

DEF = 'Illustrative default set by model builder from recent trend — not guidance or consensus. Overwrite with your own view.'

# mixes (last-4-quarter averages)
mix = {n: round(last4(lambda i: g(n, i) / g('c_nonint_rev', i)), 4) for n in ['c_ib', 'c_trading', 'c_investments', 'c_commissions', 'c_asset_mgmt']}
ib_ratio = round(last4(lambda i: g('c_ib', i) / (g('is_advisory', i) + g('is_eq_uw', i) + g('is_fi_uw', i))), 4)
am_ratio = round(last4(lambda i: g('c_asset_mgmt', i) / (g('wm_am', i) + g('im_am', i))), 4)
mix_other = round(last4(lambda i: g('c_other_rev', i) / g('c_nonint_rev', i)), 4)
ncm = {n: round(last4(lambda i: g(n, i) / g('c_noncomp', i)), 4) for n in ['c_brokerage', 'c_infoproc', 'c_profserv', 'c_occupancy', 'c_marketing']}

def dps_path(y, q, k):
    steps = y - 2026 - (1 if q < 3 else 0)
    return [1.15, 1.25, 1.35, 1.45, 1.55][steps]

EST_Q = {NH - 1}          # 3Q26 balance sheet estimate column index
EST_NOTE = ('ESTIMATE — Morgan Stanley had not filed its 3Q26 Form 10-Q as of this model date (5-Nov-2026). '
            'Total assets, deposits, common equity, preferred equity and goodwill + intangibles are as reported in the 3Q26 '
            'Financial Supplement (8-K, 14-Oct-2026). Loans and borrowings are scaled from 2Q26 by the change in supplement loans / '
            'long-term debt; all other lines are 2Q26 balances scaled pro rata to tie to reported totals. Replace with 10-Q values when filed.')

# ============================================================================= DRIVERS
drivers = [
    Sec('GLOBAL SETTINGS (value in column B)'),
    Row('g_bb_method', 'Buyback method: 1 = $ input per quarter, 2 = solve to target CET1', '', 'int'),
    Row('g_cet1_target', 'Target Standardized CET1 ratio (used when method = 2)', '%', 'pct'),
    Row('g_cet1_req', 'Regulatory Standardized CET1 requirement (memo)', '%', 'pct'),
    Blank(),
    Sec('INSTITUTIONAL SECURITIES'),
    Row('d_is_adv', 'Advisory revenue — YoY growth', '%', 'pct', hist=('f', '{Segments!is_advisory}/{Segments!is_advisory@-4}-1'),
        fcst=('in', yr({2026: 0.10, 2027: 0.05, 2028: 0.04})), guard=True, fcst_note=DEF),
    Row('d_is_equw', 'Equity underwriting — YoY growth', '%', 'pct', hist=('f', '{Segments!is_eq_uw}/{Segments!is_eq_uw@-4}-1'),
        fcst=('in', yr({2026: 0.10, 2027: 0.05, 2028: 0.04})), guard=True, fcst_note=DEF),
    Row('d_is_fiuw', 'Fixed income underwriting — YoY growth', '%', 'pct', hist=('f', '{Segments!is_fi_uw}/{Segments!is_fi_uw@-4}-1'),
        fcst=('in', yr({2026: 0.05, 2027: 0.03})), guard=True, fcst_note=DEF),
    Row('d_is_eqst', 'Equity sales & trading — YoY growth', '%', 'pct', hist=('f', '{Segments!is_eq_st}/{Segments!is_eq_st@-4}-1'),
        fcst=('in', yr({2026: 0.20, 2027: 0.00, 2028: 0.03})), guard=True, fcst_note=DEF),
    Row('d_is_fist', 'Fixed income sales & trading — YoY growth', '%', 'pct', hist=('f', '{Segments!is_fi_st}/{Segments!is_fi_st@-4}-1'),
        fcst=('in', yr({2026: 0.05, 2027: 0.02, 2028: 0.03})), guard=True, fcst_note=DEF),
    Row('d_is_other', 'IS Other revenue', '$mm', 'mm', hist=('f', '{Segments!is_other}'), fcst=('in', c(100)), fcst_note=DEF),
    Row('d_is_prov', 'IS provision — annualized, on avg IS loans', 'bps', 'bps',
        hist=('f', '{Segments!is_provision}*4/(({Segments!is_loans@-1}+{Segments!is_loans})/2*1000)*10000'),
        fcst=('in', c(35.0)), guard=True, fcst_note=DEF),
    Row('d_is_comp', 'IS compensation ratio (% of IS net revenues)', '%', 'pct', hist=('f', '{Segments!is_comp}/{Segments!is_net_rev}'),
        fcst=('in', qq({1: 0.305, 2: 0.28, 3: 0.27, 4: 0.255})), guard=True, fcst_note=DEF + ' Seasonal by quarter (1Q highest).'),
    Row('d_is_nc', 'IS non-compensation expense — YoY growth', '%', 'pct', hist=('f', '{Segments!is_noncomp}/{Segments!is_noncomp@-4}-1'),
        fcst=('in', yr({2026: 0.10, 2027: 0.05, 2028: 0.04})), guard=True, fcst_note=DEF),
    Row('d_is_loan_g', 'IS loans — QoQ growth', '%', 'pct', hist=('f', '{Segments!is_loans}/{Segments!is_loans@-1}-1'),
        fcst=('in', c(0.02)), guard=True, fcst_note=DEF),
    Blank(),
    Sec('WEALTH MANAGEMENT'),
    Row('d_wm_mkt', 'Market appreciation — client & fee-based assets (QoQ)', '%', 'pct',
        hist=('f', '{Segments!wm_ca_mkt}/{Segments!wm_client_assets@-1}'), fcst=('in', c(0.015)), guard=True,
        fcst_note=DEF + ' ~6% annualized.', hist_note='Historical value = (change in client assets less net new assets) / beginning client assets; includes acquisitions (E*TRADE 4Q20) and, before 4Q19, net new assets (not reported).'),
    Row('d_wm_nna', 'Net new assets', '$bn', 'bn', hist=('f', '{Segments!wm_nna}'), fcst=('in', c(100.0)), guard=True, fcst_note=DEF),
    Row('d_wm_ffl', 'Fee-based asset flows', '$bn', 'bn', hist=('f', '{Segments!wm_fee_flows}'), fcst=('in', c(45.0)), fcst_note=DEF),
    Row('d_wm_fee', 'Asset management fee rate — annualized, on avg fee-based assets', 'bps', 'bps',
        hist=('f', '{Segments!wm_am}*4/(({Segments!wm_fee_assets@-1}+{Segments!wm_fee_assets})/2*1000)*10000'),
        fcst=('in', c(72.0)), guard=True, fcst_note=DEF),
    Row('d_wm_trans', 'Transactional revenue — annualized, on avg total client assets', 'bps', 'bps',
        hist=('f', '{Segments!wm_trans}*4/(({Segments!wm_client_assets@-1}+{Segments!wm_client_assets})/2*1000)*10000'),
        fcst=('in', c(6.0)), guard=True, fcst_note=DEF),
    Row('d_wm_nii', 'WM net interest income — annualized, on avg WM deposits', 'bps', 'bps',
        hist=('f', '{Segments!wm_nii}*4/(({Segments!wm_deposits@-1}+{Segments!wm_deposits})/2*1000)*10000'),
        fcst=('in', c(200.0)), guard=True, fcst_note=DEF),
    Row('d_wm_other', 'WM Other revenue', '$mm', 'mm', hist=('f', '{Segments!wm_other}'), fcst=('in', c(150)), fcst_note=DEF),
    Row('d_wm_prov', 'WM provision — annualized, on avg WM loans', 'bps', 'bps',
        hist=('f', '{Segments!wm_provision}*4/(({Segments!wm_loans@-1}+{Segments!wm_loans})/2*1000)*10000'),
        fcst=('in', c(3.0)), guard=True, fcst_note=DEF),
    Row('d_wm_comp', 'WM compensation ratio (% of WM net revenues)', '%', 'pct', hist=('f', '{Segments!wm_comp}/{Segments!wm_net_rev}'),
        fcst=('in', c(0.535)), guard=True, fcst_note=DEF),
    Row('d_wm_nc', 'WM non-compensation expense — YoY growth', '%', 'pct', hist=('f', '{Segments!wm_noncomp}/{Segments!wm_noncomp@-4}-1'),
        fcst=('in', c(0.05)), guard=True, fcst_note=DEF),
    Row('d_wm_loan_g', 'WM loans — QoQ growth', '%', 'pct', hist=('f', '{Segments!wm_loans}/{Segments!wm_loans@-1}-1'),
        fcst=('in', c(0.025)), guard=True, fcst_note=DEF),
    Row('d_wm_dep_share', 'WM deposits as % of firm deposits', '%', 'pct', hist=('f', '{Segments!wm_deposits}*1000/{BS!bs_deposits}'),
        fcst=('in', c(0.972)), guard=True, fcst_note='Held at 3Q26 actual (97.2%).'),
    Blank(),
    Sec('INVESTMENT MANAGEMENT'),
    Row('d_im_mkt', 'Market appreciation — long-term AUM (QoQ)', '%', 'pct', hist=('f', '{Segments!im_lt_mkt}/{Segments!im_lt_aum@-1}'),
        fcst=('in', c(0.015)), guard=True, fcst_note=DEF, hist_note='Historical value = (change in long-term AUM less long-term net flows) / beginning LT AUM; includes acquisitions (Eaton Vance 1Q21), FX and reclassifications.'),
    Row('d_im_ltf', 'Long-term net flows', '$bn', 'bn', hist=('f', '{Segments!im_lt_flows}'), fcst=('in', c(8.0)), fcst_note=DEF),
    Row('d_im_liqf', 'Liquidity & overlay services net flows', '$bn', 'bn', hist=('f', '{Segments!im_liq_flows}'), fcst=('in', c(8.0)), fcst_note=DEF),
    Row('d_im_fee', 'Asset management fee rate — annualized, on avg total AUM', 'bps', 'bps',
        hist=('f', '{Segments!im_am}*4/(({Segments!im_total_aum@-1}+{Segments!im_total_aum})/2*1000)*10000'),
        fcst=('in', c(31.0)), guard=True, fcst_note=DEF),
    Row('d_im_perf', 'IM performance-based income & other', '$mm', 'mm', hist=('f', '{Segments!im_perf}'), fcst=('in', c(100)), fcst_note=DEF),
    Row('d_im_comp', 'IM compensation ratio (% of IM net revenues)', '%', 'pct', hist=('f', '{Segments!im_comp}/{Segments!im_net_rev}'),
        fcst=('in', c(0.37)), guard=True, fcst_note=DEF),
    Row('d_im_nc', 'IM non-compensation expense — YoY growth', '%', 'pct', hist=('f', '{Segments!im_noncomp}/{Segments!im_noncomp@-4}-1'),
        fcst=('in', c(0.05)), guard=True, fcst_note=DEF),
    Blank(),
    Sec('FIRM-LEVEL INCOME STATEMENT'),
    Row('d_elim_rev', 'Intersegment eliminations — revenues', '$mm', 'mm', hist=('f', '{Segments!elim_rev}'), fcst=('in', c(-220)), fcst_note=DEF),
    Row('d_elim_exp', 'Intersegment eliminations — non-compensation expenses', '$mm', 'mm', hist=('f', '{Segments!elim_exp}'), fcst=('in', c(-210)), fcst_note=DEF),
    Row('d_etr', 'Effective tax rate', '%', 'pct', hist=('f', '{IS!c_tax}/{IS!c_pretax}'), fcst=('in', c(0.23)), guard=True, fcst_note=DEF),
    Row('d_nci', 'Net income applicable to noncontrolling interests', '$mm', 'mm', hist=('f', '{IS!c_nci}'), fcst=('in', c(60)), fcst_note=DEF),
    Row('d_pref', 'Preferred stock dividends', '$mm', 'mm', hist=('f', '{IS!c_pref_div}'),
        fcst=('in', qq({1: 156, 2: 145, 3: 158, 4: 147})), fcst_note='Default = same quarter of prior year (semi-annual payment pattern on some series).'),
    Row('d_nonwm_ni', 'Net interest outside Wealth Management (IS, IM, Parent)', '$mm', 'mm', hist=('f', '{IS!c_net_int}-{Segments!wm_nii}'),
        fcst=('in', c(700)), fcst_note=DEF + ' ~last-4-quarter average.'),
    Row('d_ii_yield', 'Gross interest income yield — annualized, on avg total assets', 'bps', 'bps',
        hist=('f', '{IS!c_int_inc}*4/(({BS!bs_total_assets@-1}+{BS!bs_total_assets})/2)*10000'), fcst=('in', c(385.0)), guard=True,
        fcst_note='Held at 3Q26 level. Grosses up interest income / expense only; net interest is driven by the WM and non-WM lines above.'),
    Sub('GAAP non-interest revenue lines ("Trading" is the plug to total non-interest revenues)'),
    Row('d_ib_ratio', 'Investment banking (GAAP) / IS investment banking revenues', 'x', 'ratio', hist=('f', '{IS!c_ib}/{Segments!is_ib}'), fcst=('in', c(ib_ratio)), guard=True, indent=1,
        fcst_note='Default = last-4-quarter average. GAAP line also captures WM investment banking (within WM transactional).'),
    Row('d_am_ratio', 'Asset management (GAAP) / (WM asset management + IM fees)', 'x', 'ratio', hist=('f', '{IS!c_asset_mgmt}/({Segments!wm_am}+{Segments!im_am})'), fcst=('in', c(am_ratio)), guard=True, indent=1,
        fcst_note='Default = last-4-quarter average (very stable ~1.02x).'),
    Row('d_mix_inv', 'Investments — % of total non-interest revenues', '%', 'pct', hist=('f', '{IS!c_investments}/{IS!c_nonint_rev}'), fcst=('in', c(mix['c_investments'])), guard=True, indent=1, fcst_note='Default = last-4-quarter average.'),
    Row('d_mix_comm', 'Commissions and fees — % of total non-interest revenues', '%', 'pct', hist=('f', '{IS!c_commissions}/{IS!c_nonint_rev}'), fcst=('in', c(mix['c_commissions'])), guard=True, indent=1, fcst_note='Default = last-4-quarter average.'),
    Row('d_mix_other', 'Other — % of total non-interest revenues', '%', 'pct', hist=('f', '{IS!c_other_rev}/{IS!c_nonint_rev}'), fcst=('in', c(mix_other)), guard=True, indent=1, fcst_note='Default = last-4-quarter average.'),
    Sub('Non-compensation expense mix (% of total non-comp; "Other" is the plug)'),
    Row('d_nc_brok', 'Brokerage, clearing and exchange fees', '%', 'pct', hist=('f', '{IS!c_brokerage}/{IS!c_noncomp}'), fcst=('in', c(ncm['c_brokerage'])), guard=True, indent=1, fcst_note='Default = last-4-quarter average.'),
    Row('d_nc_info', 'Information processing and communications', '%', 'pct', hist=('f', '{IS!c_infoproc}/{IS!c_noncomp}'), fcst=('in', c(ncm['c_infoproc'])), guard=True, indent=1, fcst_note='Default = last-4-quarter average.'),
    Row('d_nc_prof', 'Professional services', '%', 'pct', hist=('f', '{IS!c_profserv}/{IS!c_noncomp}'), fcst=('in', c(ncm['c_profserv'])), guard=True, indent=1, fcst_note='Default = last-4-quarter average.'),
    Row('d_nc_occ', 'Occupancy and equipment', '%', 'pct', hist=('f', '{IS!c_occupancy}/{IS!c_noncomp}'), fcst=('in', c(ncm['c_occupancy'])), guard=True, indent=1, fcst_note='Default = last-4-quarter average.'),
    Row('d_nc_mkt', 'Marketing and business development', '%', 'pct', hist=('f', '{IS!c_marketing}/{IS!c_noncomp}'), fcst=('in', c(ncm['c_marketing'])), guard=True, indent=1, fcst_note='Default = last-4-quarter average.'),
    Blank(),
    Sec('BALANCE SHEET (QoQ growth unless noted; cash is the balancing item)'),
    Row('d_g_trading_assets', 'Trading assets at fair value', '%', 'pct', hist=('f', '{BS!bs_trading_assets}/{BS!bs_trading_assets@-1}-1'), fcst=('in', c(0.01)), guard=True, fcst_note=DEF),
    Row('d_g_inv_sec', 'Investment securities (AFS + HTM)', '%', 'pct', hist=('f', '{BS!bs_inv_sec}/{BS!bs_inv_sec@-1}-1'), fcst=('in', c(0.005)), guard=True, fcst_note=DEF),
    Row('d_g_resale', 'Securities purchased under agreements to resell', '%', 'pct', hist=('f', '{BS!bs_resale}/{BS!bs_resale@-1}-1'), fcst=('in', c(0.01)), guard=True, fcst_note=DEF),
    Row('d_g_sec_borrowed', 'Securities borrowed', '%', 'pct', hist=('f', '{BS!bs_sec_borrowed}/{BS!bs_sec_borrowed@-1}-1'), fcst=('in', c(0.01)), guard=True, fcst_note=DEF),
    Row('d_g_cust_recv', 'Customer and other receivables', '%', 'pct', hist=('f', '{BS!bs_cust_recv}/{BS!bs_cust_recv@-1}-1'), fcst=('in', c(0.01)), guard=True, fcst_note=DEF),
    Row('d_g_other_assets', 'Other assets', '%', 'pct', hist=('f', '{BS!bs_other_assets}/{BS!bs_other_assets@-1}-1'), fcst=('in', c(0.01)), guard=True, fcst_note=DEF),
    Row('d_amort', 'Intangible asset amortization (reduction in intangibles)', '$mm', 'mm', hist=('f', '{BS!bs_intang@-1}-{BS!bs_intang}'), fcst=('in', c(85)), guard=True, fcst_note=DEF),
    Row('d_g_deposits', 'Deposits', '%', 'pct', hist=('f', '{BS!bs_deposits}/{BS!bs_deposits@-1}-1'), fcst=('in', c(0.02)), guard=True, fcst_note=DEF),
    Row('d_g_trading_liab', 'Trading liabilities at fair value', '%', 'pct', hist=('f', '{BS!bs_trading_liab}/{BS!bs_trading_liab@-1}-1'), fcst=('in', c(0.01)), guard=True, fcst_note=DEF),
    Row('d_g_repo', 'Securities sold under agreements to repurchase', '%', 'pct', hist=('f', '{BS!bs_repo}/{BS!bs_repo@-1}-1'), fcst=('in', c(0.01)), guard=True, fcst_note=DEF),
    Row('d_g_sec_loaned', 'Securities loaned', '%', 'pct', hist=('f', '{BS!bs_sec_loaned}/{BS!bs_sec_loaned@-1}-1'), fcst=('in', c(0.01)), guard=True, fcst_note=DEF),
    Row('d_g_other_sec_fin', 'Other secured financings', '%', 'pct', hist=('f', '{BS!bs_other_sec_fin}/{BS!bs_other_sec_fin@-1}-1'), fcst=('in', c(0.01)), guard=True, fcst_note=DEF),
    Row('d_g_cust_pay', 'Customer and other payables', '%', 'pct', hist=('f', '{BS!bs_cust_pay}/{BS!bs_cust_pay@-1}-1'), fcst=('in', c(0.01)), guard=True, fcst_note=DEF),
    Row('d_g_other_liab', 'Other liabilities and accrued expenses', '%', 'pct', hist=('f', '{BS!bs_other_liab}/{BS!bs_other_liab@-1}-1'), fcst=('in', c(0.01)), guard=True, fcst_note=DEF),
    Row('d_g_borrowings', 'Borrowings', '%', 'pct', hist=('f', '{BS!bs_borrowings}/{BS!bs_borrowings@-1}-1'), fcst=('in', c(0.01)), guard=True, fcst_note=DEF),
    Row('d_pref_iss', 'Preferred stock issued / (redeemed)', '$mm', 'mm', hist=('f', '{BS!bs_pref}-{BS!bs_pref@-1}'), fcst=('in', c(0)), guard=True, fcst_note='No preferred issuance or redemption assumed.'),
    Row('d_g_nci', 'Noncontrolling interests (equity) — QoQ growth', '%', 'pct', hist=('f', '{BS!bs_nci}/{BS!bs_nci@-1}-1'), fcst=('in', c(0.0)), guard=True, fcst_note='Held flat.'),
    Blank(),
    Sec('CAPITAL, SHAREHOLDER RETURNS & SHARE COUNT'),
    Row('d_dps', 'Dividend declared per common share', '$/sh', 'eps', hist=('f', '{IS!dps}'), fcst=('in', dps_path),
        fcst_note='$1.15 declared Jul-2026 (2Q26 release, +15c). Default assumes ~8% annual increases each 3Q (MS has raised its dividend every 3Q since 2021). Illustrative.'),
    Row('d_bb', 'Common share repurchases (used when method = 1)', '$mm', 'mm', hist=('f', '{BS!bb_amount}'),
        fcst=('in', yr({2026: 1500, 2027: 2000, 2028: 2500})), fcst_note='3Q26 actual $1.5bn (release). $20bn multi-year authorization re-approved Jun-2026 (2Q26 10-Q). ' + DEF),
    Row('d_bb_px', 'Average repurchase price', '$/sh', 'px', fcst=('in', c(215.0)),
        fcst_note='June 2026 average price paid $214.98 (2Q26 10-Q, Issuer Purchases of Equity Securities). Update to current share price.'),
    Row('d_sh_iss', 'Shares issued for employee compensation (net)', 'mm', 'sh1', fcst=('in', qq({1: 6.0, 2: 1.0, 3: 1.0, 4: 1.0})),
        fcst_note='Calibrated to recent net share-count change after buybacks; 1Q higher on annual award delivery. Illustrative.'),
    Row('d_basic_ratio', 'Avg basic shares / avg period-end shares', 'x', 'ratio', hist=('f', '{IS!sh_basic}/(({BS!sh_pe@-1}+{BS!sh_pe})/2)'),
        fcst=('in', c(0.986)), guard=True, fcst_note='Held near recent average.'),
    Row('d_dil', 'Dilutive shares (diluted less basic)', 'mm', 'sh', hist=('f', '{IS!sh_diluted}-{IS!sh_basic}'), fcst=('in', c(17.0)), guard=True, fcst_note='Held near recent average.'),
    Row('d_oth_eq', 'Other changes in common equity (stock comp, OCI, other)', '$mm', 'mm', hist=('f', '{BS!cer_other}'), fcst=('in', c(250)), guard=True, fcst_note=DEF),
    Row('d_g_rwa_std', 'Standardized RWA — QoQ growth', '%', 'pct', hist=('f', '{Capital!cap_rwa_std}/{Capital!cap_rwa_std@-1}-1'), fcst=('in', c(0.02)), guard=True, fcst_note=DEF),
    Row('d_g_rwa_adv', 'Advanced RWA — QoQ growth', '%', 'pct', hist=('f', '{Capital!cap_rwa_adv}/{Capital!cap_rwa_adv@-1}-1'), fcst=('in', c(0.02)), guard=True, fcst_note=DEF),
    Row('d_cet1_adj', 'Other CET1 capital adjustments (beyond change in tangible common equity)', '$bn', 'bn',
        hist=('f', '{Capital!cap_cet1}-{Capital!cap_cet1@-1}-(({BS!bs_common_eq}-{BS!bs_common_eq@-1})-(({BS!bs_goodwill}+{BS!bs_intang})-({BS!bs_goodwill@-1}+{BS!bs_intang@-1})))/1000'),
        fcst=('in', c(0.0)), guard=True, fcst_note='Zero = CET1 moves 1:1 with tangible common equity.'),
    Row('d_lev_exp', 'Tier 1 leverage exposure / period-end total assets', 'x', 'ratio', hist=('f', '({Capital!cap_t1}/{Capital!cap_t1_lev})/({BS!bs_total_assets}/1000)'),
        fcst=('in', c(0.975)), guard=True, fcst_note='Held near recent average.'),
    Row('d_slr_exp', 'Supplementary leverage exposure / period-end total assets', 'x', 'ratio', hist=('f', '({Capital!cap_t1}/{Capital!cap_slr})/({BS!bs_total_assets}/1000)'),
        fcst=('in', c(1.20)), guard=True, fcst_note='Held near recent average.'),
]

# ============================================================================= SEGMENTS
SEG_OTHER_NOTE = ('1Q18–4Q19: Institutional Securities "Other" = sales & trading "Other" + Investments + Other (prior presentation). '
                  'From 1Q20 MS reports Other as a single line (restated in 2021 supplements). Pre-2020 also includes the loan-loss provision (pre-CECL).')
PROV_NOTE = 'Before CECL adoption (1-Jan-2020) the provision for loan losses was reported within Other revenues; shown as zero for 2018–2019.'
segments = [
    Sec('INSTITUTIONAL SECURITIES ($mm)'),
    Row('is_advisory', 'Advisory', '$mm', hist=('data', 'is_advisory'), fcst=('f', '{is_advisory@-4}*(1+{Drivers!d_is_adv})'), ann='sum', indent=1),
    Row('is_eq_uw', 'Equity underwriting', '$mm', hist=('data', 'is_eq_uw'), fcst=('f', '{is_eq_uw@-4}*(1+{Drivers!d_is_equw})'), ann='sum', indent=1),
    Row('is_fi_uw', 'Fixed income underwriting', '$mm', hist=('data', 'is_fi_uw'), fcst=('f', '{is_fi_uw@-4}*(1+{Drivers!d_is_fiuw})'), ann='sum', indent=1),
    Row('is_uw', 'Underwriting', '$mm', f='{is_eq_uw}+{is_fi_uw}', ann='sum', indent=1),
    Row('is_ib', 'Investment banking', '$mm', f='{is_advisory}+{is_uw}', ann='sum', bold=True),
    Row('is_eq_st', 'Equity', '$mm', hist=('data', 'is_eq_st'), fcst=('f', '{is_eq_st@-4}*(1+{Drivers!d_is_eqst})'), ann='sum'),
    Row('is_fi_st', 'Fixed income', '$mm', hist=('data', 'is_fi_st'), fcst=('f', '{is_fi_st@-4}*(1+{Drivers!d_is_fist})'), ann='sum'),
    Row('is_other', 'Other', '$mm', hist=('data', 'is_other'), fcst=('f', '{Drivers!d_is_other}'), ann='sum', note=SEG_OTHER_NOTE),
    Row('is_net_rev', 'IS net revenues', '$mm', f='{is_ib}+{is_eq_st}+{is_fi_st}+{is_other}', ann='sum', total=True),
    Row('is_provision', 'Provision for credit losses', '$mm', hist=('data', 'is_provision'),
        fcst=('f', '{Drivers!d_is_prov}/10000*({is_loans@-1}+{is_loans})/2*1000/4'), ann='sum', hist_note=PROV_NOTE),
    Row('is_comp', 'Compensation and benefits', '$mm', hist=('data', 'is_comp'), fcst=('f', '{Drivers!d_is_comp}*{is_net_rev}'), ann='sum'),
    Row('is_noncomp', 'Non-compensation expenses', '$mm', hist=('data', 'is_noncomp'), fcst=('f', '{is_noncomp@-4}*(1+{Drivers!d_is_nc})'), ann='sum'),
    Row('is_exp', 'Total non-interest expenses', '$mm', f='{is_comp}+{is_noncomp}', ann='sum', total=True),
    Row('is_pretax', 'IS income before taxes', '$mm', f='{is_net_rev}-{is_provision}-{is_exp}', ann='sum', bold=True),
    Row('is_ni', 'IS net income applicable to MS', '$mm', hist=('data', 'is_ni'), fcst=('f', '{is_pretax}*(1-{Drivers!d_etr})-{Drivers!d_nci}'), ann='sum',
        fcst_note='Forecast: segment pre-tax x (1 - firm ETR) less all NCI (NCI sits predominantly in IS).'),
    Row('is_margin', 'Pre-tax margin', '%', 'pct', f='IFERROR({is_pretax}/{is_net_rev},0)', ann=('f', 'IFERROR({is_pretax}/{is_net_rev},0)'), italic=True, indent=1),
    Row('is_compr', 'Compensation ratio', '%', 'pct', f='IFERROR({is_comp}/{is_net_rev},0)', ann=('f', 'IFERROR({is_comp}/{is_net_rev},0)'), italic=True, indent=1),
    Row('is_yoy', 'Net revenues — YoY growth', '%', 'pct', f='IFERROR({is_net_rev}/{is_net_rev@-4}-1,0)', ann=('f', 'IFERROR({is_net_rev}/{is_net_rev@-1}-1,0)'), italic=True, indent=1),
    Row('is_loans', 'IS loans (period end)', '$bn', 'bn', hist=('data', 'is_loans'), fcst=('f', '{is_loans@-1}*(1+{Drivers!d_is_loan_g})'), ann='end'),
    Blank(),
    Sec('WEALTH MANAGEMENT ($mm)'),
    Row('wm_am', 'Asset management', '$mm', hist=('data', 'wm_am'), fcst=('f', '{Drivers!d_wm_fee}/10000*({wm_fee_assets@-1}+{wm_fee_assets})/2*1000/4'), ann='sum'),
    Row('wm_trans', 'Transactional', '$mm', hist=('data', 'wm_trans'), fcst=('f', '{Drivers!d_wm_trans}/10000*({wm_client_assets@-1}+{wm_client_assets})/2*1000/4'), ann='sum'),
    Row('wm_nii', 'Net interest income', '$mm', hist=('data', 'wm_nii'), fcst=('f', '{Drivers!d_wm_nii}/10000*({wm_deposits@-1}+{wm_deposits})/2*1000/4'), ann='sum'),
    Row('wm_other', 'Other', '$mm', hist=('data', 'wm_other'), fcst=('f', '{Drivers!d_wm_other}'), ann='sum'),
    Row('wm_net_rev', 'WM net revenues', '$mm', f='{wm_am}+{wm_trans}+{wm_nii}+{wm_other}', ann='sum', total=True),
    Row('wm_provision', 'Provision for credit losses', '$mm', hist=('data', 'wm_provision'),
        fcst=('f', '{Drivers!d_wm_prov}/10000*({wm_loans@-1}+{wm_loans})/2*1000/4'), ann='sum', hist_note=PROV_NOTE),
    Row('wm_comp', 'Compensation and benefits', '$mm', hist=('data', 'wm_comp'), fcst=('f', '{Drivers!d_wm_comp}*{wm_net_rev}'), ann='sum'),
    Row('wm_noncomp', 'Non-compensation expenses', '$mm', hist=('data', 'wm_noncomp'), fcst=('f', '{wm_noncomp@-4}*(1+{Drivers!d_wm_nc})'), ann='sum'),
    Row('wm_exp', 'Total non-interest expenses', '$mm', f='{wm_comp}+{wm_noncomp}', ann='sum', total=True),
    Row('wm_pretax', 'WM income before taxes', '$mm', f='{wm_net_rev}-{wm_provision}-{wm_exp}', ann='sum', bold=True),
    Row('wm_ni', 'WM net income applicable to MS', '$mm', hist=('data', 'wm_ni'), fcst=('f', '{wm_pretax}*(1-{Drivers!d_etr})'), ann='sum'),
    Row('wm_margin', 'Pre-tax margin', '%', 'pct', f='IFERROR({wm_pretax}/{wm_net_rev},0)', ann=('f', 'IFERROR({wm_pretax}/{wm_net_rev},0)'), italic=True, indent=1),
    Row('wm_compr', 'Compensation ratio', '%', 'pct', f='IFERROR({wm_comp}/{wm_net_rev},0)', ann=('f', 'IFERROR({wm_comp}/{wm_net_rev},0)'), italic=True, indent=1),
    Row('wm_yoy', 'Net revenues — YoY growth', '%', 'pct', f='IFERROR({wm_net_rev}/{wm_net_rev@-4}-1,0)', ann=('f', 'IFERROR({wm_net_rev}/{wm_net_rev@-1}-1,0)'), italic=True, indent=1),
    Sub('Wealth Management client assets ($bn)'),
    Row('wm_ca_beg', 'Total client assets — beginning', '$bn', 'bn', f='{wm_client_assets@-1}', ann='beg', indent=1),
    Row('wm_nna', 'Net new assets', '$bn', 'bn', hist=('data', 'wm_nna'), fcst=('f', '{Drivers!d_wm_nna}'), ann='sum', indent=1,
        hist_note='Net new assets first disclosed for 4Q19 (2020 supplements); not available for earlier quarters.'),
    Row('wm_ca_mkt', 'Market impact & other (incl. acquisitions)', '$bn', 'bn', hist=('f', '{wm_client_assets}-{wm_client_assets@-1}-{wm_nna}'),
        fcst=('f', '{wm_client_assets@-1}*{Drivers!d_wm_mkt}'), ann='sum', indent=1,
        cell_notes={11: '4Q20 includes E*TRADE client assets acquired (closed 2-Oct-2020).'}),
    Row('wm_client_assets', 'Total client assets — ending', '$bn', 'bn', hist=('data', 'wm_client_assets'), fcst=('f', '{wm_ca_beg}+{wm_nna}+{wm_ca_mkt}'), ann='end', total=True),
    Row('wm_fa_beg', 'Fee-based client assets — beginning', '$bn', 'bn', f='{wm_fee_assets@-1}', ann='beg', indent=1),
    Row('wm_fee_flows', 'Fee-based asset flows', '$bn', 'bn', hist=('data', 'wm_fee_flows'), fcst=('f', '{Drivers!d_wm_ffl}'), ann='sum', indent=1),
    Row('wm_fa_mkt', 'Market impact & other', '$bn', 'bn', hist=('f', '{wm_fee_assets}-{wm_fee_assets@-1}-{wm_fee_flows}'),
        fcst=('f', '{wm_fee_assets@-1}*{Drivers!d_wm_mkt}'), ann='sum', indent=1),
    Row('wm_fee_assets', 'Fee-based client assets — ending', '$bn', 'bn', hist=('data', 'wm_fee_assets'), fcst=('f', '{wm_fa_beg}+{wm_fee_flows}+{wm_fa_mkt}'), ann='end', total=True),
    Row('wm_fee_pct', 'Fee-based assets as % of total client assets', '%', 'pct', f='IFERROR({wm_fee_assets}/{wm_client_assets},0)',
        ann=('f', 'IFERROR({wm_fee_assets}/{wm_client_assets},0)'), italic=True, indent=1),
    Row('wm_loans', 'WM loans (period end)', '$bn', 'bn', hist=('data', 'wm_loans'), fcst=('f', '{wm_loans@-1}*(1+{Drivers!d_wm_loan_g})'), ann='end'),
    Row('wm_deposits', 'WM deposits (period end)', '$bn', 'bn', hist=('data', 'wm_deposits'), fcst=('f', '{BS!bs_deposits}/1000*{Drivers!d_wm_dep_share}'), ann='end',
        hist_note='WM deposits first disclosed for 4Q19.'),
    Blank(),
    Sec('INVESTMENT MANAGEMENT ($mm)'),
    Row('im_am', 'Asset management and related fees', '$mm', hist=('data', 'im_am'),
        fcst=('f', '{Drivers!d_im_fee}/10000*({im_total_aum@-1}+{im_total_aum})/2*1000/4'), ann='sum'),
    Row('im_perf', 'Performance-based income and other', '$mm', hist=('data', 'im_perf_other'), fcst=('f', '{Drivers!d_im_perf}'), ann='sum',
        hist_note='2018–2019 (prior presentation): Investments + Other revenues.'),
    Row('im_net_rev', 'IM net revenues', '$mm', f='{im_am}+{im_perf}', ann='sum', total=True),
    Row('im_comp', 'Compensation and benefits', '$mm', hist=('data', 'im_comp'), fcst=('f', '{Drivers!d_im_comp}*{im_net_rev}'), ann='sum'),
    Row('im_noncomp', 'Non-compensation expenses', '$mm', hist=('data', 'im_noncomp'), fcst=('f', '{im_noncomp@-4}*(1+{Drivers!d_im_nc})'), ann='sum'),
    Row('im_exp', 'Total non-interest expenses', '$mm', f='{im_comp}+{im_noncomp}', ann='sum', total=True),
    Row('im_pretax', 'IM income before taxes', '$mm', f='{im_net_rev}-{im_exp}', ann='sum', bold=True),
    Row('im_ni', 'IM net income applicable to MS', '$mm', hist=('data', 'im_ni'), fcst=('f', '{im_pretax}*(1-{Drivers!d_etr})'), ann='sum'),
    Row('im_margin', 'Pre-tax margin', '%', 'pct', f='IFERROR({im_pretax}/{im_net_rev},0)', ann=('f', 'IFERROR({im_pretax}/{im_net_rev},0)'), italic=True, indent=1),
    Row('im_compr', 'Compensation ratio', '%', 'pct', f='IFERROR({im_comp}/{im_net_rev},0)', ann=('f', 'IFERROR({im_comp}/{im_net_rev},0)'), italic=True, indent=1),
    Row('im_yoy', 'Net revenues — YoY growth', '%', 'pct', f='IFERROR({im_net_rev}/{im_net_rev@-4}-1,0)', ann=('f', 'IFERROR({im_net_rev}/{im_net_rev@-1}-1,0)'), italic=True, indent=1),
    Sub('Assets under management or supervision ($bn)'),
    Row('im_lt_beg', 'Long-term AUM — beginning', '$bn', 'bn', f='{im_lt_aum@-1}', ann='beg', indent=1),
    Row('im_lt_flows', 'Long-term net flows', '$bn', 'bn', hist=('data', 'im_lt_flows'), fcst=('f', '{Drivers!d_im_ltf}'), ann='sum', indent=1),
    Row('im_lt_mkt', 'Market impact & other (incl. acquisitions, FX, reclass)', '$bn', 'bn', hist=('f', '{im_lt_aum}-{im_lt_aum@-1}-{im_lt_flows}'),
        fcst=('f', '{im_lt_aum@-1}*{Drivers!d_im_mkt}'), ann='sum', indent=1, cell_notes={12: '1Q21 includes Eaton Vance AUM acquired (closed 1-Mar-2021).'}),
    Row('im_lt_aum', 'Long-term AUM — ending', '$bn', 'bn', hist=('data', 'im_lt_aum'), fcst=('f', '{im_lt_beg}+{im_lt_flows}+{im_lt_mkt}'), ann='end', total=True),
    Row('im_liq_beg', 'Liquidity & overlay services AUM — beginning', '$bn', 'bn', f='{im_liq_aum@-1}', ann='beg', indent=1),
    Row('im_liq_flows', 'Liquidity & overlay services net flows', '$bn', 'bn', hist=('data', 'im_liq_flows'), fcst=('f', '{Drivers!d_im_liqf}'), ann='sum', indent=1),
    Row('im_liq_mkt', 'Market impact & other', '$bn', 'bn', hist=('f', '{im_liq_aum}-{im_liq_aum@-1}-{im_liq_flows}'), fcst=('const', 0), ann='sum', indent=1),
    Row('im_liq_aum', 'Liquidity & overlay services AUM — ending', '$bn', 'bn', hist=('data', 'im_liq_aum'), fcst=('f', '{im_liq_beg}+{im_liq_flows}+{im_liq_mkt}'), ann='end', total=True),
    Row('im_total_aum', 'Total AUM', '$bn', 'bn', f='{im_lt_aum}+{im_liq_aum}', ann='end', bold=True),
    Row('im_total_flows', 'Total net flows', '$bn', 'bn', f='{im_lt_flows}+{im_liq_flows}', ann='sum', indent=1),
    Blank(),
    Sec('FIRM — SEGMENT ROLL-UP ($mm)'),
    Row('elim_rev', 'Intersegment eliminations — net revenues', '$mm', hist=('data', 'seg_rev_elim'), fcst=('f', '{Drivers!d_elim_rev}'), ann='sum'),
    Row('elim_exp', 'Intersegment eliminations — non-interest expenses', '$mm', hist=('data', 'elim_exp'), fcst=('f', '{Drivers!d_elim_exp}'), ann='sum',
        hist_note='Derived: consolidated non-interest expenses less the sum of segment expenses (all eliminations fall in non-compensation expenses).'),
    Row('elim_pretax', 'Intersegment eliminations — income before taxes', '$mm', f='{elim_rev}-{elim_exp}', ann='sum'),
    Row('elim_ni', 'Intersegment eliminations — net income', '$mm', hist=('data', 'elim_ni'), fcst=('f', '{elim_pretax}*(1-{Drivers!d_etr})'), ann='sum',
        hist_note='Derived: consolidated net income applicable to MS less the sum of segment net income.'),
    Row('tot_net_rev', 'Firm net revenues', '$mm', f='{is_net_rev}+{wm_net_rev}+{im_net_rev}+{elim_rev}', ann='sum', total=True),
    Row('tot_provision', 'Firm provision for credit losses', '$mm', f='{is_provision}+{wm_provision}', ann='sum'),
    Row('tot_comp', 'Firm compensation and benefits', '$mm', f='{is_comp}+{wm_comp}+{im_comp}', ann='sum'),
    Row('tot_noncomp', 'Firm non-compensation expenses', '$mm', f='{is_noncomp}+{wm_noncomp}+{im_noncomp}+{elim_exp}', ann='sum'),
    Row('tot_pretax', 'Firm income before taxes', '$mm', f='{is_pretax}+{wm_pretax}+{im_pretax}+{elim_pretax}', ann='sum', total=True),
    Row('tot_ni', 'Firm net income applicable to MS', '$mm', f='{is_ni}+{wm_ni}+{im_ni}+{elim_ni}', ann='sum', bold=True),
    Row('mix_is', 'Revenue mix — Institutional Securities', '%', 'pct', f='IFERROR({is_net_rev}/{tot_net_rev},0)', ann=('f', 'IFERROR({is_net_rev}/{tot_net_rev},0)'), italic=True, indent=1),
    Row('mix_wm', 'Revenue mix — Wealth Management', '%', 'pct', f='IFERROR({wm_net_rev}/{tot_net_rev},0)', ann=('f', 'IFERROR({wm_net_rev}/{tot_net_rev},0)'), italic=True, indent=1),
    Row('mix_im', 'Revenue mix — Investment Management', '%', 'pct', f='IFERROR({im_net_rev}/{tot_net_rev},0)', ann=('f', 'IFERROR({im_net_rev}/{tot_net_rev},0)'), italic=True, indent=1),
    Blank(),
    Sec('MEMO: AS REPORTED (historical tie-out)'),
    Row('rep_is_rev', 'IS net revenues — as reported', '$mm', hist=('data', 'seg_rev_is'), ann='sum'),
    Row('rep_wm_rev', 'WM net revenues — as reported', '$mm', hist=('data', 'seg_rev_wm'), ann='sum'),
    Row('rep_im_rev', 'IM net revenues — as reported', '$mm', hist=('data', 'seg_rev_im'), ann='sum'),
    Row('rep_is_pt', 'IS income before taxes — as reported', '$mm', hist=('data', 'is_pretax'), ann='sum'),
    Row('rep_wm_pt', 'WM income before taxes — as reported', '$mm', hist=('data', 'wm_pretax'), ann='sum'),
    Row('rep_im_pt', 'IM income before taxes — as reported', '$mm', hist=('data', 'im_pretax'), ann='sum'),
    Row('rep_tot_aum', 'Total AUM — as reported', '$bn', 'bn', hist=('data', 'im_total_aum'), ann='end'),
    Row('chk_seg', 'Check: model less reported (should be zero)', '$mm', hist=('f', '({is_net_rev}-{rep_is_rev})+({wm_net_rev}-{rep_wm_rev})+({im_net_rev}-{rep_im_rev})+({is_pretax}-{rep_is_pt})+({wm_pretax}-{rep_wm_pt})+({im_pretax}-{rep_im_pt})+({im_total_aum}-{rep_tot_aum})'), ann=None, italic=True),
]

# ============================================================================= INCOME STATEMENT
is_rows = [
    Sec('CONSOLIDATED INCOME STATEMENT ($mm)'),
    Sub('Revenues'),
    Row('c_ib', 'Investment banking', '$mm', hist=('data', 'c_ib'), fcst=('f', '{Drivers!d_ib_ratio}*{Segments!is_ib}'), ann='sum', indent=1),
    Row('c_trading', 'Trading', '$mm', hist=('data', 'c_trading'), fcst=('f', '{c_nonint_rev}-{c_ib}-{c_investments}-{c_commissions}-{c_asset_mgmt}-{c_other_rev}'), ann='sum', indent=1,
        fcst_note='Forecast: plug so that GAAP revenue lines sum to total non-interest revenues from the segment build.'),
    Row('c_investments', 'Investments', '$mm', hist=('data', 'c_investments'), fcst=('f', '{Drivers!d_mix_inv}*{c_nonint_rev}'), ann='sum', indent=1),
    Row('c_commissions', 'Commissions and fees', '$mm', hist=('data', 'c_commissions'), fcst=('f', '{Drivers!d_mix_comm}*{c_nonint_rev}'), ann='sum', indent=1),
    Row('c_asset_mgmt', 'Asset management', '$mm', hist=('data', 'c_asset_mgmt'), fcst=('f', '{Drivers!d_am_ratio}*({Segments!wm_am}+{Segments!im_am})'), ann='sum', indent=1),
    Row('c_other_rev', 'Other', '$mm', hist=('data', 'c_other_rev'), fcst=('f', '{Drivers!d_mix_other}*{c_nonint_rev}'), ann='sum', indent=1),
    Row('c_nonint_rev', 'Total non-interest revenues', '$mm', hist=('f', '{c_ib}+{c_trading}+{c_investments}+{c_commissions}+{c_asset_mgmt}+{c_other_rev}'),
        fcst=('f', '{c_net_rev}-{c_net_int}'), ann='sum', total=True),
    Row('c_int_inc', 'Interest income', '$mm', hist=('data', 'c_int_inc'), fcst=('f', '{Drivers!d_ii_yield}/10000*({BS!bs_total_assets@-1}+{BS!bs_total_assets})/2/4'), ann='sum', indent=1),
    Row('c_int_exp', 'Interest expense', '$mm', hist=('data', 'c_int_exp'), fcst=('f', '{c_int_inc}-{c_net_int}'), ann='sum', indent=1),
    Row('c_net_int', 'Net interest', '$mm', hist=('f', '{c_int_inc}-{c_int_exp}'), fcst=('f', '{Segments!wm_nii}+{Drivers!d_nonwm_ni}'), ann='sum', total=True),
    Row('c_net_rev', 'Net revenues', '$mm', hist=('f', '{c_nonint_rev}+{c_net_int}'), fcst=('f', '{Segments!tot_net_rev}'), ann='sum', total=True),
    Row('c_provision', 'Provision for credit losses', '$mm', hist=('data', 'c_provision'), fcst=('f', '{Segments!tot_provision}'), ann='sum', hist_note=PROV_NOTE),
    Sub('Non-interest expenses'),
    Row('c_comp', 'Compensation and benefits', '$mm', hist=('data', 'c_comp'), fcst=('f', '{Segments!tot_comp}'), ann='sum', indent=1),
    Row('c_brokerage', 'Brokerage, clearing and exchange fees', '$mm', hist=('data', 'c_brokerage'), fcst=('f', '{Drivers!d_nc_brok}*{c_noncomp}'), ann='sum', indent=2),
    Row('c_infoproc', 'Information processing and communications', '$mm', hist=('data', 'c_infoproc'), fcst=('f', '{Drivers!d_nc_info}*{c_noncomp}'), ann='sum', indent=2),
    Row('c_profserv', 'Professional services', '$mm', hist=('data', 'c_profserv'), fcst=('f', '{Drivers!d_nc_prof}*{c_noncomp}'), ann='sum', indent=2),
    Row('c_occupancy', 'Occupancy and equipment', '$mm', hist=('data', 'c_occupancy'), fcst=('f', '{Drivers!d_nc_occ}*{c_noncomp}'), ann='sum', indent=2),
    Row('c_marketing', 'Marketing and business development', '$mm', hist=('data', 'c_marketing'), fcst=('f', '{Drivers!d_nc_mkt}*{c_noncomp}'), ann='sum', indent=2),
    Row('c_other_exp', 'Other', '$mm', hist=('data', 'c_other_exp'), fcst=('f', '{c_noncomp}-{c_brokerage}-{c_infoproc}-{c_profserv}-{c_occupancy}-{c_marketing}'), ann='sum', indent=2,
        fcst_note='Forecast: plug so that expense lines sum to total non-compensation expenses from the segment build.'),
    Row('c_noncomp', 'Total non-compensation expenses', '$mm', hist=('f', '{c_brokerage}+{c_infoproc}+{c_profserv}+{c_occupancy}+{c_marketing}+{c_other_exp}'),
        fcst=('f', '{Segments!tot_noncomp}'), ann='sum', indent=1, total=True),
    Row('c_nonint_exp', 'Total non-interest expenses', '$mm', f='{c_comp}+{c_noncomp}', ann='sum', total=True),
    Row('c_pretax', 'Income before provision for income taxes', '$mm', f='{c_net_rev}-{c_provision}-{c_nonint_exp}', ann='sum', bold=True),
    Row('c_tax', 'Provision for income taxes', '$mm', hist=('data', 'c_tax'), fcst=('f', '{Drivers!d_etr}*{c_pretax}'), ann='sum'),
    Row('c_discops', 'Gain / (loss) from discontinued operations, after tax', '$mm', hist=('data', 'c_discops'), fcst=('const', 0), ann='sum',
        fcst_note='No discontinued operations assumed.'),
    Row('c_net_income', 'Net income', '$mm', f='{c_pretax}-{c_tax}+{c_discops}', ann='sum', total=True),
    Row('c_nci', 'Net income applicable to noncontrolling interests', '$mm', hist=('data', 'c_nci'), fcst=('f', '{Drivers!d_nci}'), ann='sum'),
    Row('c_ni_ms', 'Net income applicable to Morgan Stanley', '$mm', f='{c_net_income}-{c_nci}', ann='sum', bold=True),
    Row('c_pref_div', 'Preferred stock dividends', '$mm', hist=('data', 'c_pref_div'), fcst=('f', '{Drivers!d_pref}'), ann='sum'),
    Row('c_common_earn', 'Earnings applicable to MS common shareholders', '$mm', f='{c_ni_ms}-{c_pref_div}', ann='sum', total=True),
    Blank(),
    Sec('PER SHARE'),
    Row('sh_basic', 'Average common shares — basic', 'mm', 'sh', hist=('data', 'shares_basic'), fcst=('f', '{BS!sh_basic}'), ann='avg'),
    Row('sh_diluted', 'Average common shares — diluted', 'mm', 'sh', hist=('data', 'shares_diluted'), fcst=('f', '{BS!sh_diluted}'), ann='avg'),
    Row('eps_basic', 'Earnings per basic share', '$/sh', 'eps', hist=('data', 'eps_basic'), fcst=('f', '{c_common_earn}/{sh_basic}'),
        ann=('f', '{c_common_earn}/{sh_basic}'), hist_note='Historical quarters: as reported (two-class method). Annual and forecast: earnings to common / average shares.'),
    Row('eps_diluted', 'Earnings per diluted share', '$/sh', 'eps', hist=('data', 'eps_diluted'), fcst=('f', '{c_common_earn}/{sh_diluted}'),
        ann=('f', '{c_common_earn}/{sh_diluted}'), bold=True),
    Row('dps', 'Dividends declared per common share', '$/sh', 'eps', hist=('data', 'dps'), fcst=('f', '{Drivers!d_dps}'), ann='sum'),
    Blank(),
    Sec('RATIOS & GROWTH'),
    Row('r_rev_yoy', 'Net revenues — YoY growth', '%', 'pct', f='IFERROR({c_net_rev}/{c_net_rev@-4}-1,0)', ann=('f', 'IFERROR({c_net_rev}/{c_net_rev@-1}-1,0)')),
    Row('r_eps_yoy', 'Diluted EPS — YoY growth', '%', 'pct', f='IFERROR({eps_diluted}/{eps_diluted@-4}-1,0)', ann=('f', 'IFERROR({eps_diluted}/{eps_diluted@-1}-1,0)')),
    Row('r_comp', 'Compensation ratio', '%', 'pct', f='IFERROR({c_comp}/{c_net_rev},0)', ann=('f', 'IFERROR({c_comp}/{c_net_rev},0)')),
    Row('r_noncomp', 'Non-compensation ratio', '%', 'pct', f='IFERROR({c_noncomp}/{c_net_rev},0)', ann=('f', 'IFERROR({c_noncomp}/{c_net_rev},0)')),
    Row('r_eff', 'Expense efficiency ratio', '%', 'pct', f='IFERROR({c_nonint_exp}/{c_net_rev},0)', ann=('f', 'IFERROR({c_nonint_exp}/{c_net_rev},0)'), bold=True),
    Row('r_margin', 'Pre-tax margin', '%', 'pct', f='IFERROR({c_pretax}/{c_net_rev},0)', ann=('f', 'IFERROR({c_pretax}/{c_net_rev},0)')),
    Row('r_etr', 'Effective tax rate', '%', 'pct', f='IFERROR({c_tax}/{c_pretax},0)', ann=('f', 'IFERROR({c_tax}/{c_pretax},0)')),
    Row('r_nim', 'Net interest / avg total assets (annualized)', '%', 'pct2', f='IFERROR({c_net_int}*4/(({BS!bs_total_assets@-1}+{BS!bs_total_assets})/2),0)',
        ann=('f', 'IFERROR({c_net_int}/(({BS!bs_total_assets@-1}+{BS!bs_total_assets})/2),0)')),
    Blank(),
    Sec('CHECKS & MEMO: AS REPORTED'),
    Row('chk_rev', 'Check: net revenues vs. segment build', '$mm', f='{c_net_rev}-{Segments!tot_net_rev}', ann='sum', italic=True),
    Row('chk_pretax', 'Check: pre-tax income vs. segment build', '$mm', f='{c_pretax}-{Segments!tot_pretax}', ann='sum', italic=True),
    Row('chk_ni', 'Check: net income to MS vs. segment build', '$mm', f='{c_ni_ms}-{Segments!tot_ni}', ann='sum', italic=True),
    Row('rep_net_rev', 'Net revenues — as reported', '$mm', hist=('data', 'c_net_rev'), ann='sum'),
    Row('rep_pretax', 'Income before taxes — as reported', '$mm', hist=('data', 'c_pretax'), ann='sum'),
    Row('rep_ni_ms', 'Net income applicable to MS — as reported', '$mm', hist=('data', 'c_ni_ms'), ann='sum'),
    Row('rep_common', 'Earnings applicable to common — as reported', '$mm', hist=('data', 'c_common_earn'), ann='sum'),
    Row('chk_rep', 'Check: model less reported (should be zero)', '$mm',
        hist=('f', '({c_net_rev}-{rep_net_rev})+({c_pretax}-{rep_pretax})+({c_ni_ms}-{rep_ni_ms})+({c_common_earn}-{rep_common})'), italic=True),
]

# ============================================================================= BALANCE SHEET
E = EST_Q
bs_rows = [
    Sec('CONSOLIDATED BALANCE SHEET ($mm, period end)'),
    Sub('Assets'),
    Row('bs_cash', 'Cash and cash equivalents', '$mm', hist=('data', 'bs_cash'),
        fcst=('f', '{bs_total_le}-({bs_trading_assets}+{bs_inv_sec}+{bs_resale}+{bs_sec_borrowed}+{bs_cust_recv}+{bs_loans}+{bs_goodwill}+{bs_intang}+{bs_other_assets})'),
        ann='end', indent=1, est_cols=E, fcst_note='Forecast: balancing item (total liabilities & equity less all non-cash assets). Monitor on Checks sheet (must stay positive).',
        hist_note='2018–2019: cash and due from banks + interest bearing deposits with banks + restricted cash (prior presentation).'),
    Row('bs_trading_assets', 'Trading assets at fair value', '$mm', hist=('data', 'bs_trading_assets'), fcst=('f', '{bs_trading_assets@-1}*(1+{Drivers!d_g_trading_assets})'), ann='end', indent=1, est_cols=E),
    Row('bs_inv_sec', 'Investment securities (AFS + HTM)', '$mm', hist=('data', 'bs_inv_sec'), fcst=('f', '{bs_inv_sec@-1}*(1+{Drivers!d_g_inv_sec})'), ann='end', indent=1, est_cols=E),
    Row('bs_resale', 'Securities purchased under agreements to resell', '$mm', hist=('data', 'bs_resale'), fcst=('f', '{bs_resale@-1}*(1+{Drivers!d_g_resale})'), ann='end', indent=1, est_cols=E),
    Row('bs_sec_borrowed', 'Securities borrowed', '$mm', hist=('data', 'bs_sec_borrowed'), fcst=('f', '{bs_sec_borrowed@-1}*(1+{Drivers!d_g_sec_borrowed})'), ann='end', indent=1, est_cols=E),
    Row('bs_cust_recv', 'Customer and other receivables', '$mm', hist=('data', 'bs_cust_recv'), fcst=('f', '{bs_cust_recv@-1}*(1+{Drivers!d_g_cust_recv})'), ann='end', indent=1, est_cols=E),
    Row('bs_loans', 'Loans: held for investment (net of ACL) + held for sale', '$mm', hist=('data', 'bs_loans'),
        fcst=('f', '{bs_loans@-1}*({Segments!is_loans}+{Segments!wm_loans})/({Segments!is_loans@-1}+{Segments!wm_loans@-1})'), ann='end', indent=1, est_cols=E,
        fcst_note='Forecast: grows with IS + WM loans from the segment build (supplement loans also include fair-value loans held in trading assets).'),
    Row('bs_goodwill', 'Goodwill', '$mm', hist=('data', 'bs_goodwill'), fcst=('f', '{bs_goodwill@-1}'), ann='end', indent=1, est_cols=E),
    Row('bs_intang', 'Intangible assets', '$mm', hist=('data', 'bs_intangibles'), fcst=('f', '{bs_intang@-1}-{Drivers!d_amort}'), ann='end', indent=1, est_cols=E),
    Row('bs_other_assets', 'Other assets', '$mm', hist=('data', 'bs_other_assets'), fcst=('f', '{bs_other_assets@-1}*(1+{Drivers!d_g_other_assets})'), ann='end', indent=1, est_cols=E),
    Row('bs_total_assets', 'Total assets', '$mm',
        f='{bs_cash}+{bs_trading_assets}+{bs_inv_sec}+{bs_resale}+{bs_sec_borrowed}+{bs_cust_recv}+{bs_loans}+{bs_goodwill}+{bs_intang}+{bs_other_assets}', ann='end', total=True),
    Sub('Liabilities'),
    Row('bs_deposits', 'Deposits', '$mm', hist=('data', 'bs_deposits'), fcst=('f', '{bs_deposits@-1}*(1+{Drivers!d_g_deposits})'), ann='end', indent=1, est_cols=E),
    Row('bs_trading_liab', 'Trading liabilities at fair value', '$mm', hist=('data', 'bs_trading_liab'), fcst=('f', '{bs_trading_liab@-1}*(1+{Drivers!d_g_trading_liab})'), ann='end', indent=1, est_cols=E),
    Row('bs_repo', 'Securities sold under agreements to repurchase', '$mm', hist=('data', 'bs_repo'), fcst=('f', '{bs_repo@-1}*(1+{Drivers!d_g_repo})'), ann='end', indent=1, est_cols=E),
    Row('bs_sec_loaned', 'Securities loaned', '$mm', hist=('data', 'bs_sec_loaned'), fcst=('f', '{bs_sec_loaned@-1}*(1+{Drivers!d_g_sec_loaned})'), ann='end', indent=1, est_cols=E),
    Row('bs_other_sec_fin', 'Other secured financings', '$mm', hist=('data', 'bs_other_sec_fin'), fcst=('f', '{bs_other_sec_fin@-1}*(1+{Drivers!d_g_other_sec_fin})'), ann='end', indent=1, est_cols=E),
    Row('bs_cust_pay', 'Customer and other payables', '$mm', hist=('data', 'bs_cust_pay'), fcst=('f', '{bs_cust_pay@-1}*(1+{Drivers!d_g_cust_pay})'), ann='end', indent=1, est_cols=E),
    Row('bs_other_liab', 'Other liabilities and accrued expenses', '$mm', hist=('data', 'bs_other_liab'), fcst=('f', '{bs_other_liab@-1}*(1+{Drivers!d_g_other_liab})'), ann='end', indent=1, est_cols=E),
    Row('bs_borrowings', 'Borrowings', '$mm', hist=('data', 'bs_borrowings'), fcst=('f', '{bs_borrowings@-1}*(1+{Drivers!d_g_borrowings})'), ann='end', indent=1, est_cols=E),
    Row('bs_total_liab', 'Total liabilities', '$mm', f='{bs_deposits}+{bs_trading_liab}+{bs_repo}+{bs_sec_loaned}+{bs_other_sec_fin}+{bs_cust_pay}+{bs_other_liab}+{bs_borrowings}', ann='end', total=True),
    Sub('Equity'),
    Row('bs_pref', 'Preferred stock', '$mm', hist=('data', 'bs_pref_stock'), fcst=('f', '{bs_pref@-1}+{Drivers!d_pref_iss}'), ann='end', indent=1, est_cols=E),
    Row('bs_common_eq', 'Common shareholders\' equity', '$mm', hist=('data', 'bs_common_eq'), fcst=('f', '{cer_end}'), ann='end', indent=1, est_cols=E),
    Row('bs_ms_equity', 'Total Morgan Stanley shareholders\' equity', '$mm', f='{bs_pref}+{bs_common_eq}', ann='end', total=True),
    Row('bs_nci', 'Noncontrolling interests', '$mm', hist=('data', 'bs_nci'), fcst=('f', '{bs_nci@-1}*(1+{Drivers!d_g_nci})'), ann='end', indent=1, est_cols=E),
    Row('bs_total_equity', 'Total equity', '$mm', f='{bs_ms_equity}+{bs_nci}', ann='end', total=True),
    Row('bs_total_le', 'Total liabilities and equity', '$mm', f='{bs_total_liab}+{bs_total_equity}', ann='end', total=True),
    Row('bs_check', 'Check: total assets less total liabilities & equity', '$mm', f='{bs_total_assets}-{bs_total_le}', ann='end', italic=True),
    Blank(),
    Sec('COMMON EQUITY ROLL-FORWARD ($mm)'),
    Row('cer_begin', 'Common equity — beginning of period', '$mm', f='{bs_common_eq@-1}', ann='beg'),
    Row('cer_earn', '(+) Earnings applicable to common shareholders', '$mm', f='{IS!c_common_earn}', ann='sum', indent=1),
    Row('cer_div', '(–) Common dividends declared', '$mm', hist=('data', neg('div_common')), fcst=('f', '-{IS!dps}*{sh_pe@-1}'), ann='sum', indent=1,
        est_cols=E, cell_notes={NH - 1: 'ESTIMATE: $1.15 DPS x 2Q26 period-end shares (1,572mm); 3Q26 10-Q not yet filed.'},
        hist_note='Source: XBRL DividendsCommonStock (10-Q/10-K); 2Q18–4Q18 from 10-Q/10-K equity statements.'),
    Row('cer_bb', '(–) Common share repurchases', '$mm', f='-{bb_amount}', ann='sum', indent=1),
    Row('cer_other', '(+/–) Other (stock comp, OCI, other)', '$mm', hist=('f', '{bs_common_eq}-{bs_common_eq@-1}-{cer_earn}-{cer_div}-{cer_bb}'),
        fcst=('f', '{Drivers!d_oth_eq}'), ann='sum', indent=1, hist_note='Historical: derived residual.'),
    Row('cer_end', 'Common equity — end of period', '$mm', f='{bs_common_eq@-1}+{cer_earn}+{cer_div}+{cer_bb}+{cer_other}', ann='end', total=True),
    Sub('Share repurchase calculation'),
    Row('bb_pre_cet1', 'CET1 capital before repurchases', '$bn', 'bn',
        fcst=('f', '{Capital!cap_cet1@-1}+({cer_earn}+{cer_div}+{cer_other})/1000-(({bs_goodwill}+{bs_intang})-({bs_goodwill@-1}+{bs_intang@-1}))/1000+{Drivers!d_cet1_adj}'), indent=1),
    Row('bb_target_cet1', 'CET1 capital at target ratio (target x binding RWA)', '$bn', 'bn', fcst=('f', '{$g_cet1_target}*MAX({Capital!cap_rwa_std},{Capital!cap_rwa_adv})'), indent=1),
    Row('bb_amount', 'Common share repurchases', '$mm', hist=('data', 'buyback'),
        fcst=('f', 'IF({$g_bb_method}=2,MAX(0,({bb_pre_cet1}-{bb_target_cet1})*1000),{Drivers!d_bb})'), ann='sum', bold=True,
        hist_note='Source: XBRL StockRepurchasedDuringPeriodValue / TreasuryStockValueAcquiredCostMethod (program repurchases, excl. tax withholding); 3Q26 from earnings release.'),
    Blank(),
    Sec('SHARE COUNT ROLL-FORWARD (mm)'),
    Row('sh_begin', 'Period-end shares — beginning', 'mm', 'sh', f='{sh_pe@-1}', ann='beg'),
    Row('sh_rep', '(–) Shares repurchased', 'mm', 'sh1', fcst=('f', '-{bb_amount}/{Drivers!d_bb_px}'), ann='sum', indent=1),
    Row('sh_iss', '(+) Shares issued (employee compensation, net)', 'mm', 'sh1', fcst=('f', '{Drivers!d_sh_iss}'), ann='sum', indent=1),
    Row('sh_net', 'Net change in period-end shares', 'mm', 'sh1', hist=('f', '{sh_pe}-{sh_pe@-1}'), fcst=('f', '{sh_rep}+{sh_iss}'), ann='sum', indent=1),
    Row('sh_pe', 'Period-end common shares outstanding', 'mm', 'sh', hist=('data', 'shares_pe'), fcst=('f', '{sh_begin}+{sh_net}'), ann='end', total=True),
    Row('sh_basic', 'Average shares — basic', 'mm', 'sh', hist=('data', 'shares_basic'), fcst=('f', '({sh_begin}+{sh_pe})/2*{Drivers!d_basic_ratio}'), ann='avg'),
    Row('sh_diluted', 'Average shares — diluted', 'mm', 'sh', hist=('data', 'shares_diluted'), fcst=('f', '{sh_basic}+{Drivers!d_dil}'), ann='avg'),
    Blank(),
    Sec('MEMO: AS REPORTED (historical tie-out)'),
    Row('rep_ta', 'Total assets — as reported (10-Q / 10-K)', '$mm', hist=('data', 'bs_total_assets'), ann='end'),
    Row('rep_tl', 'Total liabilities — as reported (10-Q / 10-K)', '$mm', hist=('data', 'bs_total_liab'), ann='end'),
    Row('chk_bs_rep', 'Check: model less reported (should be zero)', '$mm', hist=('f', '({bs_total_assets}-{rep_ta})+({bs_total_liab}-{rep_tl})'), italic=True),
]

# ============================================================================= CAPITAL
cap_rows = [
    Sec('REGULATORY CAPITAL ($bn, period end)'),
    Row('cap_cet1', 'Common Equity Tier 1 capital', '$bn', 'bn', hist=('data', 'cet1_cap'),
        fcst=('f', '{cap_cet1@-1}+(({BS!bs_common_eq}-{BS!bs_common_eq@-1})-(({BS!bs_goodwill}+{BS!bs_intang})-({BS!bs_goodwill@-1}+{BS!bs_intang@-1})))/1000+{Drivers!d_cet1_adj}'),
        ann='end', bold=True, fcst_note='Forecast: prior CET1 + change in tangible common equity + other adjustments (Drivers).'),
    Row('cap_t1', 'Tier 1 capital', '$bn', 'bn', hist=('data', 't1_cap'), fcst=('f', '{cap_t1@-1}+({cap_cet1}-{cap_cet1@-1})+({BS!bs_pref}-{BS!bs_pref@-1})/1000'), ann='end'),
    Row('cap_rwa_std', 'Risk-weighted assets — Standardized', '$bn', 'bn', hist=('data', 'rwa_std'), fcst=('f', '{cap_rwa_std@-1}*(1+{Drivers!d_g_rwa_std})'), ann='end'),
    Row('cap_rwa_adv', 'Risk-weighted assets — Advanced', '$bn', 'bn', hist=('data', 'rwa_adv'), fcst=('f', '{cap_rwa_adv@-1}*(1+{Drivers!d_g_rwa_adv})'), ann='end'),
    Row('cap_cet1_std', 'CET1 ratio — Standardized', '%', 'pct', f='IFERROR({cap_cet1}/{cap_rwa_std},0)', ann=('f', 'IFERROR({cap_cet1}/{cap_rwa_std},0)'), bold=True),
    Row('cap_cet1_adv', 'CET1 ratio — Advanced', '%', 'pct', f='IFERROR({cap_cet1}/{cap_rwa_adv},0)', ann=('f', 'IFERROR({cap_cet1}/{cap_rwa_adv},0)')),
    Row('cap_t1_std', 'Tier 1 ratio — Standardized', '%', 'pct', f='IFERROR({cap_t1}/{cap_rwa_std},0)', ann=('f', 'IFERROR({cap_t1}/{cap_rwa_std},0)')),
    Row('cap_t1_adv', 'Tier 1 ratio — Advanced', '%', 'pct', f='IFERROR({cap_t1}/{cap_rwa_adv},0)', ann=('f', 'IFERROR({cap_t1}/{cap_rwa_adv},0)')),
    Row('cap_t1_lev', 'Tier 1 leverage ratio', '%', 'pct', hist=('data', 't1_lev'), fcst=('f', '{cap_t1}/({BS!bs_total_assets}/1000*{Drivers!d_lev_exp})'), ann='end'),
    Row('cap_slr', 'Supplementary leverage ratio', '%', 'pct', hist=('data', 'slr'), fcst=('f', '{cap_t1}/({BS!bs_total_assets}/1000*{Drivers!d_slr_exp})'), ann='end'),
    Row('cap_cet1_bind', 'Binding CET1 ratio (lower of Standardized / Advanced)', '%', 'pct', f='MIN({cap_cet1_std},{cap_cet1_adv})', ann=('f', 'MIN({cap_cet1_std},{cap_cet1_adv})')),
    Row('cap_req', 'Standardized CET1 requirement (Drivers global)', '%', 'pct', f='{$g_cet1_req}', ann=('f', '{$g_cet1_req}'), italic=True, indent=1,
        note='4.5% minimum + 4.3% stress capital buffer + 3.0% G-SIB surcharge = 11.8% (2Q26 10-Q, Capital Buffer Requirements). Applied to all periods for reference only.'),
    Row('cap_excess_req', 'CET1 capital in excess of requirement', '$bn', 'bn', f='{cap_cet1}-{cap_req}*{cap_rwa_std}', ann=('f', '{cap_cet1}-{cap_req}*{cap_rwa_std}'), indent=1),
    Row('cap_excess_tgt', 'CET1 capital in excess of target (Drivers global)', '$bn', 'bn', f='{cap_cet1}-{$g_cet1_target}*MAX({cap_rwa_std},{cap_rwa_adv})',
        ann=('f', '{cap_cet1}-{$g_cet1_target}*MAX({cap_rwa_std},{cap_rwa_adv})'), indent=1),
    Blank(),
    Sec('RETURNS & BOOK VALUE'),
    Row('tce', 'Tangible common equity (common equity less goodwill & intangibles)', '$mm', f='{BS!bs_common_eq}-{BS!bs_goodwill}-{BS!bs_intang}', ann='end'),
    Row('avg_ce', 'Average common equity (simple avg of period ends)', '$mm', f='({BS!bs_common_eq@-1}+{BS!bs_common_eq})/2', ann='avg'),
    Row('avg_tce', 'Average tangible common equity (simple avg)', '$mm', f='({tce@-1}+{tce})/2', ann='avg'),
    Row('roe', 'Return on average common equity (annualized)', '%', 'pct', hist=('data', 'roe'), fcst=('f', '{IS!c_common_earn}*4/{avg_ce}'),
        ann=('f', 'IFERROR({IS!c_common_earn}/{avg_ce},0)'), hist_note='Historical quarters: as reported (MS uses average monthly balances). Forecast and annual: simple average of period-end balances.'),
    Row('rotce', 'Return on average tangible common equity (annualized)', '%', 'pct', hist=('data', 'rotce'), fcst=('f', '{IS!c_common_earn}*4/{avg_tce}'),
        ann=('f', 'IFERROR({IS!c_common_earn}/{avg_tce},0)'), bold=True),
    Row('bvps', 'Book value per common share', '$/sh', 'eps', hist=('data', 'bvps'), fcst=('f', '{BS!bs_common_eq}/{BS!sh_pe}'), ann='end'),
    Row('tbvps', 'Tangible book value per common share', '$/sh', 'eps', hist=('data', 'tbvps'), fcst=('f', '{tce}/{BS!sh_pe}'), ann='end', bold=True,
        hist_note='Historical: as reported (MS deducts goodwill & intangibles net of allowable mortgage servicing rights).'),
    Blank(),
    Sec('CAPITAL RETURN ($mm)'),
    Row('cr_div', 'Common dividends', '$mm', f='-{BS!cer_div}', ann='sum'),
    Row('cr_bb', 'Share repurchases', '$mm', f='{BS!bb_amount}', ann='sum'),
    Row('cr_total', 'Total capital returned to common shareholders', '$mm', f='{cr_div}+{cr_bb}', ann='sum', total=True),
    Row('cr_payout', 'Total payout ratio (% of earnings to common)', '%', 'pct', f='IFERROR({cr_total}/{IS!c_common_earn},0)', ann=('f', 'IFERROR({cr_total}/{IS!c_common_earn},0)')),
    Row('cr_div_payout', 'Dividend payout ratio', '%', 'pct', f='IFERROR({cr_div}/{IS!c_common_earn},0)', ann=('f', 'IFERROR({cr_div}/{IS!c_common_earn},0)')),
    Blank(),
    Sec('OTHER STATISTICS'),
    Row('employees', 'Worldwide employees', '#', 'sh', hist=('data', 'employees'), ann='end'),
    Row('rev_americas', 'Regional net revenues — Americas', '$mm', hist=('data', 'rev_americas'), ann='sum'),
    Row('rev_emea', 'Regional net revenues — EMEA', '$mm', hist=('data', 'rev_emea'), ann='sum'),
    Row('rev_asia', 'Regional net revenues — Asia', '$mm', hist=('data', 'rev_asia'), ann='sum'),
]

# ============================================================================= BUILD
wb = Workbook()
cover = wb.active
cover.title = 'Cover'
M = Model(wb, H)
M.add('Drivers', 'Forecast Drivers & Assumptions', drivers, annual=False)
M.add('Segments', 'Segment Build', segments)
M.add('IS', 'Consolidated Income Statement', is_rows)
M.add('BS', 'Balance Sheet, Equity & Share Count', bs_rows)
M.add('Capital', 'Regulatory Capital, Returns & Capital Return', cap_rows)

# globals live in Drivers_Q column B
gvals = {'g_bb_method': 1, 'g_cet1_target': 0.14, 'g_cet1_req': 0.118}
for k in gvals:
    M.globals[k] = f"Drivers_Q!$B${M.sheets['Drivers']['rowmap'][k]}"

for base in M.order:
    M.write_quarterly(base)
    M.write_annual(base)

# write global values
wsd = M.sheets['Drivers']['q']
gnotes = {
    'g_bb_method': 'Enter 1 to use the $ repurchase inputs below; enter 2 to size repurchases each quarter so Standardized/Advanced CET1 lands at the target ratio (never negative).',
    'g_cet1_target': 'Management-style target used only when buyback method = 2. Illustrative default.',
    'g_cet1_req': '4.5% minimum + 4.3% SCB + 3.0% G-SIB surcharge (2Q26 10-Q). Memo only.',
}
for k, v in gvals.items():
    r = M.sheets['Drivers']['rowmap'][k]
    cell = wsd.cell(row=r, column=2, value=v)
    cell.font = Font(name=FONT, color=BLUE, bold=True, size=9)
    cell.fill = F_KEY
    cell.number_format = FMT['int'] if k == 'g_bb_method' else FMT['pct']
    cell.comment = Comment(gnotes[k], 'Model')

pickle_rows = {b: M.sheets[b]['rowmap'] for b in M.order}
json.dump(pickle_rows, open(os.path.join(HERE, '..', 'data', 'rowmap.json'), 'w'))

import build_extra
build_extra.finish(wb, M, D)
wb.save(OUT)
print('saved', OUT)
