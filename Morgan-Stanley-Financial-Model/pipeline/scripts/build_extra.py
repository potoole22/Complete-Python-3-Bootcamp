"""Cover, Checks and Sources sheets + workbook finishing."""
import json, os, datetime
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter as CL
from openpyxl.comments import Comment
from engine import FONT, BLUE, BLACK, GREEN, F_IN, F_KEY, F_EST, F_SEC, F_SUB, FMT, NH, NQ, NY, COL0, QLAB, THIN, MED

HERE = os.path.dirname(__file__)


def fnt(color=BLACK, bold=False, italic=False, size=9):
    return Font(name=FONT, color=color, bold=bold, italic=italic, size=size)


def qc(i):
    return CL(COL0 + i)


LAST_H, FIRST_F, LAST_F = qc(NH - 1), qc(NH), qc(NQ - 1)


def finish(wb, M, D):
    rm = {b: M.sheets[b]['rowmap'] for b in M.order}
    build_checks(wb, rm)
    build_sources(wb, D)
    build_cover(wb, rm)
    order = ['Cover', 'Drivers_Q', 'Segments_Q', 'IS_Q', 'BS_Q', 'Capital_Q', 'Segments_A', 'IS_A', 'BS_A', 'Capital_A', 'Checks', 'Sources']
    wb._sheets = [wb[n] for n in order]
    colors = {'Cover': '1F3864', 'Drivers_Q': 'FFC000', 'Checks': 'C00000', 'Sources': '808080'}
    for ws in wb.worksheets:
        ws.sheet_properties.tabColor = colors.get(ws.title, '2F75B5' if ws.title.endswith('_Q') else '548235')
        ws.sheet_view.zoomScale = 85
        ws.page_setup.orientation = 'landscape'
        ws.page_setup.fitToWidth = 1
        ws.sheet_properties.pageSetUpPr.fitToPage = True
        ws.page_setup.fitToHeight = 0
    wb.active = 0
    # input validation + instruction line on Drivers
    from openpyxl.worksheet.datavalidation import DataValidation
    wsd = wb['Drivers_Q']
    r = rm['Drivers']['g_bb_method']
    dv = DataValidation(type='whole', operator='between', formula1='1', formula2='2', allow_blank=False,
                        showErrorMessage=True, errorTitle='Buyback method', error='Enter 1 ($ input) or 2 (solve to target CET1).')
    wsd.add_data_validation(dv)
    dv.add(f'B{r}')
    wsd['A7'] = 'Edit the blue-on-yellow cells (4Q26E–4Q30E) and the yellow global settings in column B. Historical columns show the implied value of each driver.'
    wsd['A7'].font = fnt(italic=True, color='C00000')


# --------------------------------------------------------------------------- checks
def build_checks(wb, rm):
    ws = wb.create_sheet('Checks')
    ws.sheet_view.showGridLines = False
    ws['A1'] = 'Morgan Stanley — Model Integrity Checks'
    ws['A1'].font = fnt(bold=True, size=12)
    ws['A2'] = 'Each check should read OK. Tolerance: $1mm (rounding in reported figures).'
    ws['A2'].font = fnt(italic=True, color='595959')
    hdr = ['Check', 'Scope', 'Value', 'Status']
    for j, h in enumerate(hdr):
        c = ws.cell(row=4, column=1 + j, value=h)
        c.font = fnt('FFFFFF', bold=True)
        c.fill = F_SEC
    A = f'{qc(0)}'
    rows = [
        ('Balance sheet balances (assets = liabilities + equity)', 'All quarters 1Q18–4Q30', f'=SUMPRODUCT(MAX(ABS(BS_Q!{A}{rm["BS"]["bs_check"]}:{LAST_F}{rm["BS"]["bs_check"]})))', 'max'),
        ('Balance sheet balances — annual', 'FY2018–FY2030', f'=SUMPRODUCT(MAX(ABS(BS_A!C{rm["BS"]["bs_check"]}:O{rm["BS"]["bs_check"]})))', 'max'),
        ('Consolidated net revenues = segment build', 'All quarters', f'=SUMPRODUCT(MAX(ABS(IS_Q!{A}{rm["IS"]["chk_rev"]}:{LAST_F}{rm["IS"]["chk_rev"]})))', 'max'),
        ('Consolidated pre-tax income = segment build', 'All quarters', f'=SUMPRODUCT(MAX(ABS(IS_Q!{A}{rm["IS"]["chk_pretax"]}:{LAST_F}{rm["IS"]["chk_pretax"]})))', 'max'),
        ('Net income to MS = segment build', 'All quarters', f'=SUMPRODUCT(MAX(ABS(IS_Q!{A}{rm["IS"]["chk_ni"]}:{LAST_F}{rm["IS"]["chk_ni"]})))', 'max'),
        ('Income statement totals tie to reported', 'Historical 1Q18–3Q26', f'=SUMPRODUCT(MAX(ABS(IS_Q!{A}{rm["IS"]["chk_rep"]}:{LAST_H}{rm["IS"]["chk_rep"]})))', 'max'),
        ('Segment revenues, pre-tax & AUM tie to reported', 'Historical 1Q18–3Q26', f'=SUMPRODUCT(MAX(ABS(Segments_Q!{A}{rm["Segments"]["chk_seg"]}:{LAST_H}{rm["Segments"]["chk_seg"]})))', 'max'),
        ('Balance sheet totals tie to reported', 'Historical 1Q18–3Q26', f'=SUMPRODUCT(MAX(ABS(BS_Q!{A}{rm["BS"]["chk_bs_rep"]}:{LAST_H}{rm["BS"]["chk_bs_rep"]})))', 'max'),
        ('Forecast cash balance stays positive (count of negative quarters)', 'Forecast 4Q26–4Q30', f'=COUNTIF(BS_Q!{FIRST_F}{rm["BS"]["bs_cash"]}:{LAST_F}{rm["BS"]["bs_cash"]},"<0")', 'count'),
        ('Forecast Standardized CET1 at or above requirement (count of breaches)', 'Forecast 4Q26–4Q30',
         f'=SUMPRODUCT(--(Capital_Q!{FIRST_F}{rm["Capital"]["cap_cet1_std"]}:{LAST_F}{rm["Capital"]["cap_cet1_std"]}<Drivers_Q!$B${rm["Drivers"]["g_cet1_req"]}))', 'count'),
        ('Forecast share repurchases non-negative (count of negatives)', 'Forecast 4Q26–4Q30', f'=COUNTIF(BS_Q!{FIRST_F}{rm["BS"]["bb_amount"]}:{LAST_F}{rm["BS"]["bb_amount"]},"<0")', 'count'),
    ]
    r = 5
    for name, scope, fml, kind in rows:
        ws.cell(row=r, column=1, value=name).font = fnt()
        ws.cell(row=r, column=2, value=scope).font = fnt(italic=True, color='595959')
        v = ws.cell(row=r, column=3, value=fml)
        v.font = fnt(GREEN)
        v.number_format = '#,##0.00' if kind == 'max' else '0'
        tol = '1' if kind == 'max' else '0.5'
        s = ws.cell(row=r, column=4, value=f'=IF(C{r}<{tol},"OK","ERROR")')
        s.font = fnt(bold=True)
        r += 1
    ws.cell(row=r + 1, column=1, value='OVERALL MODEL STATUS').font = fnt(bold=True)
    o = ws.cell(row=r + 1, column=4, value=f'=IF(COUNTIF(D5:D{r - 1},"OK")={r - 5},"ALL OK","CHECK ERRORS")')
    o.font = fnt(bold=True)
    o.fill = F_KEY
    ws.column_dimensions['A'].width = 66
    ws.column_dimensions['B'].width = 24
    ws.column_dimensions['C'].width = 14
    ws.column_dimensions['D'].width = 16


# --------------------------------------------------------------------------- sources
def build_sources(wb, D):
    ws = wb.create_sheet('Sources')
    ws.sheet_view.showGridLines = False
    ws['A1'] = 'Morgan Stanley — Historical Data Sources (SEC EDGAR, CIK 0000895421)'
    ws['A1'].font = fnt(bold=True, size=12)
    notes = [
        'Income statement detail, segment results, KPIs, capital and per-share data: quarterly earnings releases (EX-99.1) and financial supplements (EX-99.2) furnished on Form 8-K (Item 2.02).',
        'Each quarter is taken from the most recent supplement/release that reports it (i.e., the year-ago column one year later), so restated presentations are used where available.',
        'Balance sheet: Consolidated Balance Sheet in each Form 10-Q / 10-K (XBRL R-file). 3Q26 is an estimate anchored to supplement totals (10-Q not filed as of 5-Nov-2026).',
        'Dividends per share, common dividends and repurchases: XBRL (CommonStockDividendsPerShareDeclared, DividendsCommonStock, StockRepurchasedDuringPeriodValue, TreasuryStockValueAcquiredCostMethod); 4Q = FY less 9M.',
        'Regulatory capital: supplement capital page ($mm in 2018 supplements converted to $bn; later values reported in $bn).',
    ]
    for k, t in enumerate(notes):
        c = ws.cell(row=3 + k, column=1, value=f'• {t}')
        c.font = fnt()
    hdr = ['Quarter', 'Earnings 8-K filed', '8-K accession', 'Earnings release (EX-99.1)', 'Financial supplement (EX-99.2)',
           'Balance sheet filing', 'Filed', 'Accession', 'Balance sheet (XBRL R-file)']
    r0 = 10
    for j, h in enumerate(hdr):
        c = ws.cell(row=r0, column=1 + j, value=h)
        c.font = fnt('FFFFFF', bold=True)
        c.fill = F_SEC
    docs = json.load(open(os.path.join(HERE, '..', 'data', 'earnings_docs.json')))
    per = json.load(open(os.path.join(HERE, '..', 'data', 'periodic.json')))
    byq = {}
    for d in docs:
        byq[d['tag']] = d
    perq = {}
    for rep, form, fd, acc, prim in per:
        y, m = int(rep[:4]), int(rep[5:7])
        perq[f'{y}q{(m - 1) // 3 + 1}'] = (form, fd, acc)
    r = r0 + 1
    for i in range(NH):
        y, q = QLAB[i]
        tag = f'{y}q{q}'
        d = byq.get(tag)
        ws.cell(row=r, column=1, value=f'{q}Q{str(y)[2:]}').font = fnt(bold=True)
        if d:
            a = d['acc'].replace('-', '')
            ex = {t: l for desc, l, t in d['docs']}
            ws.cell(row=r, column=2, value=d['date']).font = fnt()
            ws.cell(row=r, column=3, value=d['acc']).font = fnt()
            for col, t in ((4, 'EX-99.1'), (5, 'EX-99.2')):
                if t in ex:
                    url = 'https://www.sec.gov' + ex[t]
                    c = ws.cell(row=r, column=col, value=url)
                    c.hyperlink = url
                    c.font = fnt('0563C1')
        if tag in perq:
            form, fd, acc = perq[tag]
            ws.cell(row=r, column=6, value=form).font = fnt()
            ws.cell(row=r, column=7, value=fd).font = fnt()
            ws.cell(row=r, column=8, value=acc).font = fnt()
            url = f'https://www.sec.gov/Archives/edgar/data/895421/{acc.replace("-", "")}/'
            c = ws.cell(row=r, column=9, value=url)
            c.hyperlink = url
            c.font = fnt('0563C1')
        else:
            ws.cell(row=r, column=6, value='Estimate (10-Q pending)').font = fnt(italic=True)
        r += 1
    r += 1
    extra = [
        ('2Q26 Form 10-Q (8/4/2026)', 'June-2026 average repurchase price $214.98; $20bn repurchase authorization; capital buffer requirements (SCB 4.3%, G-SIB 3.0%).',
         'https://www.sec.gov/Archives/edgar/data/895421/000089542126000212/ms-20260630.htm'),
        ('2018 Form 10-Qs / 2018 Form 10-K', 'Common dividends 2Q18–4Q18 (equity statements: total dividends less preferred).',
         'https://www.sec.gov/Archives/edgar/data/895421/000119312519051943/'),
    ]
    ws.cell(row=r, column=1, value='Other sources').font = fnt(bold=True)
    r += 1
    for a, b, u in extra:
        ws.cell(row=r, column=1, value=a).font = fnt()
        ws.cell(row=r, column=2, value=b).font = fnt()
        c = ws.cell(row=r, column=6, value=u)
        c.hyperlink = u
        c.font = fnt('0563C1')
        r += 1
    widths = [10, 16, 22, 70, 72, 22, 12, 22, 62]
    for j, w in enumerate(widths):
        ws.column_dimensions[CL(1 + j)].width = w
    ws.freeze_panes = f'B{r0 + 1}'


# --------------------------------------------------------------------------- cover
def build_cover(wb, rm):
    ws = wb['Cover']
    ws.sheet_view.showGridLines = False
    ws.column_dimensions['A'].width = 2
    ws.column_dimensions['B'].width = 44
    for col in 'CDEFGHIJ':
        ws.column_dimensions[col].width = 12
    ws.column_dimensions['K'].width = 40
    ws['B2'] = 'Morgan Stanley (NYSE: MS)'
    ws['B2'].font = fnt(bold=True, size=16)
    ws['B3'] = 'Three-Statement Operating Model — History (1Q18–3Q26) and Forecast Chain (4Q26–4Q30)'
    ws['B3'].font = fnt(bold=True, size=11)
    ws['B4'] = 'Model date: 5-Nov-2026  |  Last reported quarter: 3Q26 (earnings 8-K, 14-Oct-2026)  |  $ in millions unless noted  |  No valuation components'
    ws['B4'].font = fnt(italic=True, color='595959')

    r = 6
    ws.cell(row=r, column=2, value='TIMELINE').font = fnt('FFFFFF', bold=True)
    for col in range(2, 12):
        ws.cell(row=r, column=col).fill = F_SEC
    tl = [
        ('Historical (actual)', '1Q18–3Q26 quarterly; FY2018–FY2025 annual'),
        ('FY1 = FY2026', '1Q26–3Q26 actual + 4Q26E (next forecast period) — quarterly and annual'),
        ('FY2 = FY2027', '1Q27E–4Q27E quarterly and annual'),
        ('FY3–FY5 = FY2028–FY2030', 'Quarterly 1Q28E–4Q30E (annual roll-ups also provided)'),
    ]
    for a, b in tl:
        r += 1
        ws.cell(row=r, column=2, value=a).font = fnt(bold=True)
        ws.cell(row=r, column=3, value=b).font = fnt()

    r += 2
    ws.cell(row=r, column=2, value='FORECAST CHAIN').font = fnt('FFFFFF', bold=True)
    for col in range(2, 12):
        ws.cell(row=r, column=col).fill = F_SEC
    chain = [
        ('1. Drivers_Q', 'All forecast inputs (blue on yellow) by quarter; historical columns show the implied value of each driver for calibration.'),
        ('2. Segments_Q', 'IS: YoY growth by business line on same quarter last year (keeps seasonality). WM: client-asset & fee-based-asset roll-forwards x fee rates; NII on deposits. IM: AUM roll-forward x fee rate.'),
        ('3. IS_Q', 'Consolidated P&L = segment roll-up; GAAP revenue lines linked to segment lines (IB, asset mgmt) or mix, trading as plug; non-comp lines by mix; tax, NCI, preferred, EPS.'),
        ('4. BS_Q', 'Balance sheet by line (growth drivers; loans from segment loan build); common equity & share count roll-forwards; cash balances.'),
        ('5. Capital_Q', 'CET1 / Tier 1 roll-forward, RWA, leverage ratios, ROE / ROTCE, BVPS / TBVPS, capital return.'),
        ('6. *_A sheets', 'Annual mirrors (same row numbers as the quarterly sheets): flows summed, balances at year end, ratios recalculated.'),
        ('7. Checks / Sources', 'Integrity checks (balance, segment ties, tie-out to reported) and SEC EDGAR source list for every historical quarter.'),
    ]
    for a, b in chain:
        r += 1
        ws.cell(row=r, column=2, value=a).font = fnt(bold=True)
        c = ws.cell(row=r, column=3, value=b)
        c.font = fnt()

    r += 2
    ws.cell(row=r, column=2, value='KEY OUTPUTS (links to annual sheets)').font = fnt('FFFFFF', bold=True)
    for col in range(2, 12):
        ws.cell(row=r, column=col).fill = F_SEC
    r += 1
    years = list(range(2023, 2031))
    for j, y in enumerate(years):
        c = ws.cell(row=r, column=3 + j, value=f'FY{y}{"A" if y <= 2025 else ("A/E" if y == 2026 else "E")}')
        c.font = fnt(bold=True)
        c.alignment = Alignment(horizontal='right')
        c.border = Border(bottom=THIN)
    outs = [
        ('Net revenues ($mm)', 'IS_A', 'c_net_rev', 'mm'),
        ('  YoY growth', 'IS_A', 'r_rev_yoy', 'pct'),
        ('Income before taxes ($mm)', 'IS_A', 'c_pretax', 'mm'),
        ('Net income applicable to MS ($mm)', 'IS_A', 'c_ni_ms', 'mm'),
        ('Diluted EPS ($)', 'IS_A', 'eps_diluted', 'eps'),
        ('Expense efficiency ratio', 'IS_A', 'r_eff', 'pct'),
        ('ROTCE', 'Capital_A', 'rotce', 'pct'),
        ('Standardized CET1 ratio', 'Capital_A', 'cap_cet1_std', 'pct'),
        ('Book value per share ($)', 'Capital_A', 'bvps', 'eps'),
        ('Tangible book value per share ($)', 'Capital_A', 'tbvps', 'eps'),
        ('WM client assets ($bn)', 'Segments_A', 'wm_client_assets', 'bn'),
        ('IM total AUM ($bn)', 'Segments_A', 'im_total_aum', 'bn'),
    ]
    base = {'IS_A': 'IS', 'Capital_A': 'Capital', 'Segments_A': 'Segments'}
    for lab, sh, key, fm in outs:
        r += 1
        ws.cell(row=r, column=2, value=lab).font = fnt()
        rr = rm[base[sh]][key]
        for j, y in enumerate(years):
            col = CL(COL0 + (y - 2018))
            c = ws.cell(row=r, column=3 + j, value=f'={sh}!{col}{rr}')
            c.font = fnt(GREEN)
            c.number_format = FMT[fm]
    r += 1
    ws.cell(row=r, column=2, value='Model status (Checks sheet)').font = fnt(bold=True)
    c = ws.cell(row=r, column=3, value='=Checks!D17')
    c.font = fnt(GREEN, bold=True)

    r += 2
    ws.cell(row=r, column=2, value='COLOR & INPUT LEGEND').font = fnt('FFFFFF', bold=True)
    for col in range(2, 12):
        ws.cell(row=r, column=col).fill = F_SEC
    leg = [
        ('1,234', BLUE, F_IN, 'Forecast input / driver — edit these (Drivers_Q, forecast columns)'),
        ('1,234', BLUE, F_KEY, 'Key global setting (Drivers_Q column B: buyback method, CET1 target / requirement)'),
        ('1,234', BLUE, None, 'Historical hard-coded data from SEC filings'),
        ('1,234', BLUE, F_EST, 'Historical estimate (3Q26 balance sheet and 3Q26 common dividend — 10-Q pending)'),
        ('1,234', BLACK, None, 'Formula (calculation on the same sheet)'),
        ('1,234', GREEN, None, 'Link to another sheet'),
    ]
    for txt, col, fill, desc in leg:
        r += 1
        c = ws.cell(row=r, column=2, value=txt)
        c.font = fnt(col, bold=True)
        c.alignment = Alignment(horizontal='right')
        if fill:
            c.fill = fill
        ws.cell(row=r, column=3, value=desc).font = fnt()
    r += 1
    ws.cell(row=r, column=2, value='Cells with a red corner carry notes (sources, methodology, estimates).').font = fnt(italic=True, color='595959')

    r += 2
    ws.cell(row=r, column=2, value='METHODOLOGY & DATA NOTES').font = fnt('FFFFFF', bold=True)
    for col in range(2, 12):
        ws.cell(row=r, column=col).fill = F_SEC
    notes = [
        'Sources: SEC EDGAR filings only (8-K earnings releases & financial supplements, 10-Q / 10-K XBRL). See Sources sheet for document-level links per quarter.',
        'Restatements: each historical quarter uses the latest presentation in which MS reported it (year-ago column of the following year). E.g. MS updated its presentation in 1Q21 and reclassified 2020 (FY2020 net revenues $48.8bn in the model vs. $48.2bn originally reported), and 2023 interest income / expense were later re-grossed.',
        'Pre-CECL (2018–2019): provision for loan losses was reported within Other revenues; provision lines show zero for those years.',
        'IS "Other" for 2018–2019 = sales & trading Other + Investments + Other (prior presentation). IM 2018–2019: fees = Asset management; performance & other = Investments + Other.',
        'WM net new assets and WM deposits are disclosed from 4Q19; implied market impact before 4Q19 includes net new assets.',
        'Acquisitions affect trends: E*TRADE (closed 2-Oct-2020; WM client assets, deposits) and Eaton Vance (closed 1-Mar-2021; IM AUM, goodwill & intangibles).',
        '3Q26 balance sheet is an ESTIMATE: totals (assets, deposits, common & preferred equity, goodwill + intangibles) per the 3Q26 supplement; other lines scaled from 2Q26. Replace with 10-Q values when filed.',
        'Forecast mechanics: segment pre-tax drives the P&L; firm ETR applied to segment pre-tax for segment net income (NCI charged to IS). Cash is the balance-sheet plug; check it stays positive.',
        'Share repurchases: method 1 uses $ inputs; method 2 sizes buybacks so CET1 lands at the target ratio on the binding (higher) RWA. No circular references.',
        'Quarterly ROE / ROTCE: history as reported (MS averages monthly balances); forecast uses simple averages of period-end balances.',
        'Default forecast drivers are illustrative assumptions set from recent trends by the model builder — not company guidance or consensus. Replace them with your own views.',
    ]
    for t in notes:
        r += 1
        c = ws.cell(row=r, column=2, value=f'• {t}')
        c.font = fnt()
    ws.freeze_panes = None
