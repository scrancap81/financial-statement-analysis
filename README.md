# Financial Statement Analysis

A home for quarterly financial statement workbooks on individual equities, built directly from SEC EDGAR filings.

Each company gets an Excel workbook (`workbooks/<TICKER>.xlsx`) and matching CSVs (`data/<TICKER>/`) covering its income statement, balance sheet and cash flow statement across its most recent 10-Q filings.

## Companies

| Ticker | Company | Filings | Periods |
|---|---|---|---|
| GOOGL | Alphabet Inc. | Last 12 10-Qs | Q1 2023 – Q3 2026 |

## What's in a workbook

| Sheet | Contents |
|---|---|
| Notes | Source filings (with links), methodology, colour key |
| Income Statement | Three months ended each quarter |
| Balance Sheet | Quarter-end balances |
| Cash Flow | **Standalone quarters**, derived with formulas from the YTD sheet |
| Cash Flow (YTD) | Year-to-date cash flows exactly as reported |

Every sheet ends with check rows (balance sheet balances, cash flow subtotals tie to the change in cash, ending cash ties to the balance sheet) that should all read zero.

Colour key: **blue** = hardcoded from the filing, **black** = formula, **green** = link to another sheet, **yellow fill** = see the cell comment.

## Method

- **Source.** For each 10-Q, `FilingSummary.xml` identifies the primary statements, and EDGAR's rendered XBRL statement pages supply each line item with its XBRL tag, in the company's own presentation order.
- **Line matching.** Rows are matched across filings by XBRL tag rather than by label, so a line keeps its row when the company renames it or note numbers shift. Lines re-tagged over time are merged via `TAG_ALIASES` in `scripts/build_workbook.py`.
- **Basis.** Figures are as originally reported in each quarter's own 10-Q. Later restatements are not back-applied.
- **Cash flow de-cumulation.** 10-Q cash flow statements are year-to-date. Standalone quarters are `Q1 = 3M`, `Q2 = 6M − 3M`, `Q3 = 9M − 6M`. Beginning cash for Q2/Q3 is the previous quarter's ending cash. Period-end balances disclosed in the cash flow statement (e.g. capex in accrued liabilities) are not differenced.
- **Reclassifications.** When a line appears in or disappears from the statement partway through a fiscal year, that quarter's derived figure absorbs the shift. These cells are filled yellow with an explanatory comment; subtotals are unaffected.
- **Q4 gap.** 10-Qs cover fiscal Q1–Q3. Q4 lives in the annual 10-K and is not yet included.

## Adding a company

```bash
pip install -r requirements.txt
python scripts/build_workbook.py --ticker MSFT --cik 789019 --name "Microsoft Corp." --n 12
```

Look up a company's CIK at [EDGAR company search](https://www.sec.gov/edgar/searchedgar/companysearch). Before running regularly, set `USER_AGENT` in `scripts/sec_fetch.py` to your own name and email, as the SEC asks of automated clients. Check the new workbook's yellow cells and blank rows: other companies may re-tag lines in ways that need adding to `TAG_ALIASES`.

## Layout

```
workbooks/        Excel workbooks, one per ticker
data/<TICKER>/    CSV versions (income_statement, balance_sheet, cash_flow_quarterly, cash_flow_ytd)
scripts/
  sec_fetch.py      EDGAR download + statement parsing
  build_workbook.py merge across filings, de-cumulate cash flow, write xlsx + csv
```
