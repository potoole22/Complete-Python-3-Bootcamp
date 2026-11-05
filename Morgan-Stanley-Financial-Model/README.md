# Morgan Stanley (NYSE: MS) — Financial Model

`MS_Financial_Model_3Q26.xlsx` is a three-statement operating model for Morgan Stanley. It has history from 1Q18 to 3Q26 and a formula-driven forecast chain from 4Q26 to 4Q30. It has no valuation components.

| | Periods |
|---|---|
| Historical | 1Q18–3Q26 quarterly; FY2018–FY2025 annual |
| FY1 = FY2026 | 1Q–3Q26 actual + **4Q26E**, the next forecast period (quarterly and annual) |
| FY2 = FY2027 | 1Q27E–4Q27E quarterly and annual |
| FY3–FY5 = FY2028–FY2030 | Quarterly, with annual roll-ups |

Model date: 5-Nov-2026. Last reported quarter: 3Q26 (earnings 8-K filed 14-Oct-2026).

## Workbook map (forecast chain)

1. **Drivers_Q**: every forecast input sits here, by quarter (blue on yellow). The historical columns show the implied value of each driver, so you can calibrate against history. Global settings are in column B:
   - buyback method (1 = $ input, 2 = solve to target CET1)
   - target CET1 ratio
   - regulatory CET1 requirement
2. **Segments_Q**:
   - **Institutional Securities:** advisory, equity and fixed income underwriting, and equity and fixed income sales & trading each grow year over year off the same quarter of the prior year, which keeps seasonality. Also: provision as bps of loans, a comp ratio, and non-comp growth.
   - **Wealth Management:** client assets and fee-based assets roll forward (net new assets, flows, market). Fee rate × average fee-based assets gives asset management revenue. Transactional revenue is bps of client assets, and net interest income is bps of WM deposits.
   - **Investment Management:** long-term and liquidity AUM roll forward. Fee rate × average AUM gives fee revenue.
   - Intersegment eliminations, then the firm roll-up.
3. **IS_Q**: the consolidated P&L equals the segment roll-up.
   - GAAP investment banking and asset management revenues link to the segment lines. Investments, commissions and other are a mix of non-interest revenue, and trading is the plug.
   - Non-comp expense lines are a mix of total non-comp.
   - Then tax, NCI, preferred dividends, EPS and DPS.
4. **BS_Q**: balance sheet line by line (growth drivers; loans come from the segment loan build). Includes the common-equity and share-count roll-forwards and the share repurchase engine. Cash balances the sheet.
5. **Capital_Q**: CET1 and Tier 1 roll-forward, Standardized and Advanced RWA, leverage ratio and SLR, excess capital, ROE/ROTCE, BVPS/TBVPS, and capital return.
6. **\*_A sheets**: annual mirrors with the same row numbers as the quarterly sheets. Flows are summed, balances are year-end, and ratios are recalculated.
7. **Checks**: the balance sheet balances, the P&L ties to the segment build, history ties to reported totals, forecast cash stays positive, and CET1 stays at or above the requirement. All checks currently read **OK**.
8. **Sources**: SEC EDGAR links (8-K release and supplement, 10-Q/10-K) for every historical quarter.

**Colour convention:**
- Blue: hard-coded inputs and historical data
- Blue on yellow: forecast drivers
- Blue on orange: historical estimates
- Black: formulas
- Green: links to other sheets

Cells with notes carry sources and methodology. The model has no circular references.

## Data and methodology notes

- **Sources:** SEC EDGAR only. Income statement, segments, KPIs and capital come from the quarterly earnings releases and financial supplements (EX-99.1 / EX-99.2). The balance sheet comes from 10-Q/10-K XBRL statements. Dividends and buybacks come from XBRL.
- **Restatements:** each quarter is taken from the most recent document that reports it, so restated presentations are used. For example, MS reclassified 2020 in 1Q21, so FY2020 net revenues show as $48.8bn here vs. $48.2bn originally reported.
- **3Q26 balance sheet is an estimate:** the 3Q26 10-Q had not been filed as of the model date. Totals (assets, deposits, common and preferred equity, goodwill + intangibles) match the 3Q26 supplement. The other lines are scaled from 2Q26. These cells are flagged in orange; replace them when the 10-Q is filed.
- **2018–2019:**
  - The provision is embedded in other revenues (pre-CECL), so the provision lines show zero.
  - IS "Other" combines sales & trading Other, Investments and Other.
  - IM fees = Asset management, and IM performance & other = Investments + Other.
- WM net new assets and WM deposits are disclosed only from 4Q19.
- **Default forecast drivers are illustrative.** I set them from recent trends; they are not guidance or consensus. Replace them with your own views on `Drivers_Q`.

## Rebuilding the workbook

`pipeline/` holds the code that produced the model, plus the processed data:

```bash
cd pipeline
python3 scripts/checks.py                                   # tie-out checks on the parsed history
python3 scripts/dataset.py                                  # assemble dataset.json (incl. 3Q26 BS estimate)
python3 scripts/build.py ../MS_Financial_Model_3Q26.xlsx    # write the workbook (openpyxl)
```

`spec.py` holds the line-item mapping from the supplements and releases to the model series. `tables.py`, `extract.py`, `classify.py` and `series.py` parse the raw EDGAR HTML exhibits. Those raw exhibits are not included, but you can re-download them from the URLs on the Sources sheet. A freshly built workbook has no cached values until a spreadsheet application recalculates it. Excel does this on open.
