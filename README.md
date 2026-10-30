# Financial Statement Analysis

A home for quarterly financial statement workbooks on individual equities, built directly from SEC EDGAR filings.

Each company gets an Excel workbook (`workbooks/<TICKER>.xlsx`) and matching CSVs (`data/<TICKER>/`) covering its income statement, balance sheet and cash flow statement quarter by quarter, built from its 10-Q filings plus the 10-Ks needed to fill in Q4.

## Companies

| Ticker | Company | Filings | Periods |
|---|---|---|---|
| GOOGL | Alphabet Inc. | Last 12 10-Qs + FY2023–FY2025 10-Ks | Q1 2023 – Q3 2026 (15 quarters) |

## What's in a workbook

| Sheet | Contents |
|---|---|
| Notes | Source filings (with links), methodology, colour key |
| Income Statement | Three months ended each quarter (Q4 derived by formula) |
| Balance Sheet | Quarter-end balances (Q4 = 10-K year-end) |
| Cash Flow | **Standalone quarters**, derived with formulas from the YTD sheet |
| Income Statement (FY) | Full fiscal years exactly as reported in each 10-K |
| Cash Flow (YTD) | Year-to-date cash flows exactly as reported (3M/6M/9M from 10-Qs, 12M from 10-Ks) |

Every sheet ends with check rows (balance sheet balances, cash flow subtotals tie to the change in cash, ending cash ties to the balance sheet) that should all read zero.

Colour key: **blue** = hardcoded from the filing, **black** = formula, **green** = link to another sheet, **yellow fill** = see the cell comment.

## Method

- **Source.** For each 10-Q, `FilingSummary.xml` identifies the primary statements, and EDGAR's rendered XBRL statement pages supply each line item with its XBRL tag, in the company's own presentation order.
- **Line matching.** Rows are matched across filings by XBRL tag rather than by label, so a line keeps its row when the company renames it or note numbers shift. Lines re-tagged over time are merged via `TAG_ALIASES` in `scripts/build_workbook.py`.
- **Basis.** Figures are as originally reported in each quarter's own 10-Q. Later restatements are not back-applied.
- **Cash flow de-cumulation.** 10-Q and 10-K cash flow statements are year-to-date. Standalone quarters are `Q1 = 3M`, `Q2 = 6M − 3M`, `Q3 = 9M − 6M`, `Q4 = 12M − 9M`. Beginning cash for Q2–Q4 is the previous quarter's ending cash. Period-end balances disclosed in the cash flow statement (e.g. capex in accrued liabilities) are not differenced.
- **Q4.** No Q4 report is filed, so Q4 comes from the 10-K: the balance sheet is the 10-K year-end balance sheet, the income statement is `FY − Q1 − Q2 − Q3`, and the cash flow is `12M − 9M`. Quarters therefore sum exactly to the reported full year. Two things are deliberately left blank in Q4, each with a cell comment:
  - **EPS**, because per-share figures can't be derived by subtraction (share counts differ between periods). Q4 EPS is in the company's Q4 earnings release.
  - **Annual-only supplemental disclosures** (e.g. cash paid for income taxes), which would otherwise put the whole year into Q4. The full-year figure is on `Cash Flow (YTD)`.

  A fiscal year's Q4 appears once its 10-K is filed (Alphabet's FY2026 10-K is due early 2027).
- **Reclassifications.** When a line appears in or disappears from the statement partway through a fiscal year, that quarter's derived figure absorbs the shift. These cells are filled yellow with an explanatory comment; subtotals are unaffected.

## Adding a company

```bash
pip install -r requirements.txt
python scripts/build_workbook.py --ticker MSFT --cik 789019 --name "Microsoft Corp." --n 12
```

Look up a company's CIK at [EDGAR company search](https://www.sec.gov/edgar/searchedgar/companysearch). Before running regularly, set `USER_AGENT` in `scripts/sec_fetch.py` to your own name and email, as the SEC asks of automated clients. Check the new workbook's yellow cells and blank rows: other companies may re-tag lines in ways that need adding to `TAG_ALIASES`.

## Layout

```
workbooks/        Excel workbooks, one per ticker
data/<TICKER>/    CSV versions (income_statement, income_statement_fy, balance_sheet, cash_flow_quarterly, cash_flow_ytd)
scripts/
  sec_fetch.py      EDGAR download + statement parsing
  build_workbook.py merge across filings, de-cumulate cash flow, write xlsx + csv
```
