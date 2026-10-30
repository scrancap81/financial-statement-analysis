"""
Build a quarterly financial statements workbook for one company from its 10-Qs.

Usage:
    python scripts/build_workbook.py --ticker GOOGL --cik 1652044 --name "Alphabet Inc." --n 12

Outputs:
    workbooks/<TICKER>.xlsx              Excel workbook (statements + notes)
    data/<TICKER>/*.csv                  the same statements as CSVs for pandas work

Cash flow: 10-Q cash flow statements are year-to-date (3, 6 or 9 months).
Standalone quarters are derived as  Q1 = 3M YTD,  Q2 = 6M − 3M,  Q3 = 9M − 6M.
In the workbook this is done with live formulas pointing at the
"Cash Flow (YTD)" sheet, so every number traces back to the filing.
"""

import argparse
import re
from pathlib import Path

import pandas as pd
from openpyxl import Workbook
from openpyxl.comments import Comment
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

import sec_fetch

ROOT = Path(__file__).resolve().parent.parent

# Line items whose XBRL tag changed between filings but which are the same line.
# alternative tag -> canonical tag. Extend this when adding a new company.
TAG_ALIASES = {
    "us-gaap_RevenueFromContractWithCustomerExcludingAssessedTax": "us-gaap_Revenues",
    "us-gaap_PropertyPlantAndEquipmentNet":
        "us-gaap_PropertyPlantAndEquipmentAndFinanceLeaseRightOfUseAssetAfterAccumulatedDepreciationAndAmortization",
    "us-gaap_LongTermDebtAndCapitalLeaseObligations": "us-gaap_LongTermDebtNoncurrent",
    "us-gaap_PaymentsOfDividends": "us-gaap_PaymentsOfOrdinaryDividends",
    "goog_DepreciationAndImpairmentOnDispositionOfPropertyAndEquipment": "us-gaap_Depreciation",
}

# Cash flow lines that are balances at period end rather than flows: never de-cumulated.
POINT_IN_TIME_TAGS = {"us-gaap_CapitalExpendituresIncurredButNotYetPaid"}
CASH_TAG = "us-gaap_CashCashEquivalentsRestrictedCashAndRestrictedCashEquivalents"
CASH_BEGIN, CASH_END = CASH_TAG + "#1", CASH_TAG + "#2"

STATEMENTS = [("income_statement", "Income Statement"),
              ("balance_sheet", "Balance Sheet"),
              ("cash_flow", "Cash Flow (YTD)")]


# ---------------------------------------------------------------- merging

def clean_label(label: str) -> str:
    label = re.sub(r"\s*\(Note \d+\)", "", label)
    label = label.replace("(in dollars per share)", "($ per share)")
    return label.strip()


def keyed_rows(rows):
    """Attach a stable key (canonical tag + occurrence number) to every row."""
    counts, out = {}, []
    for r in rows:
        # drop XBRL scaffolding rows that carry no information on the face of the statement
        if not r["tag"] or "[Abstract]" in r["label"] or r["tag"].endswith("CommitmentsAndContingencies"):
            continue
        tag = TAG_ALIASES.get(r["tag"], r["tag"])
        counts[tag] = counts.get(tag, 0) + 1
        out.append({**r, "key": f"{tag}#{counts[tag]}"})
    return out


def master_order(per_filing_rows: list[list[dict]]) -> list[dict]:
    """Union of line items across filings, in presentation order.
    Newest filing sets the order; lines only found in older filings are
    slotted in after the line that precedes them in that filing."""
    master = []
    for rows in per_filing_rows:  # newest first
        keys = [m["key"] for m in master]
        for i, r in enumerate(rows):
            if r["key"] in keys:
                continue
            pos = 0
            for prev in reversed(rows[:i]):
                if prev["key"] in keys:
                    pos = keys.index(prev["key"]) + 1
                    break
            master.insert(pos, {"key": r["key"], "label": clean_label(r["label"]),
                                "is_header": r["is_header"]})
            keys.insert(pos, r["key"])
    return master


def pick_column(stmt: dict, period_end: str, duration_prefix: str | None) -> int:
    for i, (dur, date) in enumerate(stmt["columns"]):
        if date == period_end and (duration_prefix is None or dur.startswith(duration_prefix)):
            return i
    raise ValueError(f"No column for {period_end} / {duration_prefix} in {stmt['title']}")


# ---------------------------------------------------------------- styling

FONT = "Arial"
F_BASE = Font(name=FONT, size=10)
F_INPUT = Font(name=FONT, size=10, color="0000FF")          # hardcoded from filing
F_LINK = Font(name=FONT, size=10, color="008000")           # link to another sheet
F_CALC = Font(name=FONT, size=10, color="000000")           # formula
F_HDR = Font(name=FONT, size=10, bold=True, color="FFFFFF")
F_SECTION = Font(name=FONT, size=10, bold=True, italic=True)
F_TITLE = Font(name=FONT, size=13, bold=True)
F_NOTE = Font(name=FONT, size=9, italic=True, color="595959")
FILL_HDR = PatternFill("solid", fgColor="1F3864")
FILL_SUB = PatternFill("solid", fgColor="D9E1F2")
FILL_CHECK = PatternFill("solid", fgColor="F2F2F2")
TOP = Border(top=Side(style="thin"))
FMT_MM = '#,##0;(#,##0);"-"'
FMT_EPS = '0.00;(0.00);"-"'

BOLD_PREFIXES = ("Total", "Net income", "Income from operations", "Income before",
                 "Net cash", "Net increase", "Revenues")


def is_bold(label):
    return label.startswith(BOLD_PREFIXES)


def num_fmt(key):
    return FMT_EPS if "PerShare" in key else FMT_MM


# ---------------------------------------------------------------- build

def build(ticker: str, cik: int, name: str, n: int):
    filings = sec_fetch.list_10q_filings(cik, n)
    data = {}
    for f in filings:
        print(f"  {f['period_end']}  {f['accession']}")
        data[f["accession"]] = sec_fetch.fetch_filing_statements(cik, f)

    # Fiscal quarter from the cash flow's YTD length (3/6/9 months)
    for f in filings:
        cf = data[f["accession"]]["cash_flow"]
        dur = cf["columns"][pick_column(cf, f["period_end"], None)][0]
        f["months"] = int(re.match(r"(\d+)", dur).group(1))
        f["fq"] = f["months"] // 3
        f["fy"] = int(f["period_end"][:4]) if f["fq"] else None

    chron = sorted(filings, key=lambda x: x["period_end"])  # oldest first for columns
    newest_first = list(reversed(chron))

    # values[stmt][key][accession] = number ; present[stmt][accession] = set(keys)
    values, present, masters = {}, {}, {}
    for stmt, _ in STATEMENTS:
        values[stmt], present[stmt] = {}, {}
        rows_by_filing = []
        for f in newest_first:
            st = data[f["accession"]][stmt]
            dur = {"income_statement": "3 Months", "balance_sheet": None, "cash_flow": None}[stmt]
            col = pick_column(st, f["period_end"], dur)
            rows = keyed_rows(st["rows"])
            rows_by_filing.append(rows)
            present[stmt][f["accession"]] = {r["key"] for r in rows}
            for r in rows:
                if not r["is_header"]:
                    values[stmt].setdefault(r["key"], {})[f["accession"]] = r["values"][col]
        masters[stmt] = master_order(rows_by_filing)
        # a line that is a header in one filing but carries values elsewhere is a data line
        for m in masters[stmt]:
            if m["key"] in values[stmt]:
                m["is_header"] = False

    # Prior filing in the same fiscal year (for de-cumulation)
    by_fy_q = {(f["fy"], f["fq"]): f for f in filings}
    for f in filings:
        f["prior"] = by_fy_q.get((f["fy"], f["fq"] - 1)) if f["fq"] > 1 else None

    # ------------------------------------------------ discrete-quarter cash flow (Python, for CSV + flags)
    cf_q, cf_flags = {}, {}
    for m in masters["cash_flow"]:
        if m["is_header"]:
            continue
        k = m["key"]
        cf_q[k] = {}
        for f in chron:
            a, p = f["accession"], f["prior"]
            cur = values["cash_flow"].get(k, {}).get(a)
            if k == CASH_BEGIN:
                cf_q[k][a] = cur  # Q2/Q3 replaced below with prior quarter's ending cash
                continue
            if k in POINT_IN_TIME_TAGS or k.split("#")[0] in POINT_IN_TIME_TAGS or k == CASH_END:
                cf_q[k][a] = cur
                continue
            if p is None:
                cf_q[k][a] = cur
                continue
            prev = values["cash_flow"].get(k, {}).get(p["accession"])
            in_cur = k in present["cash_flow"][a]
            in_prev = k in present["cash_flow"][p["accession"]]
            if cur is None and prev is None:
                cf_q[k][a] = None
                continue
            cf_q[k][a] = (cur or 0) - (prev or 0)
            if in_cur and not in_prev and cur:
                cf_flags[(k, a)] = (f"Line first presented in the Q{f['fq']} {f['fy']} 10-Q, so the earlier "
                                    f"quarter is treated as zero. If this activity also happened earlier in the "
                                    f"year, it was reported within another line and this quarter's figure "
                                    f"includes it (subtotals are unaffected).")
            elif in_prev and not in_cur and prev:
                cf_flags[(k, a)] = (f"Line not presented in the Q{f['fq']} {f['fy']} 10-Q (folded into "
                                    f"another line), so this quarter shows the reversal of the prior YTD amount.")

    # Need CASH_END computed before CASH_BEGIN: recompute begin now that end exists
    for f in chron:
        if f["prior"] and CASH_BEGIN in cf_q:
            cf_q[CASH_BEGIN][f["accession"]] = cf_q[CASH_END][f["prior"]["accession"]]

    # ------------------------------------------------ CSVs
    out_data = ROOT / "data" / ticker
    out_data.mkdir(parents=True, exist_ok=True)
    col_names = [f["period_end"] for f in chron]

    def frame(stmt, vals):
        recs = []
        for m in masters[stmt]:
            if m["is_header"]:
                continue
            recs.append([m["label"], m["key"].split("#")[0]] +
                        [vals.get(m["key"], {}).get(f["accession"]) for f in chron])
        return pd.DataFrame(recs, columns=["line_item", "xbrl_tag"] + col_names)

    frame("income_statement", values["income_statement"]).to_csv(out_data / "income_statement.csv", index=False)
    frame("balance_sheet", values["balance_sheet"]).to_csv(out_data / "balance_sheet.csv", index=False)
    frame("cash_flow", values["cash_flow"]).to_csv(out_data / "cash_flow_ytd.csv", index=False)
    frame("cash_flow", cf_q).to_csv(out_data / "cash_flow_quarterly.csv", index=False)

    # ------------------------------------------------ workbook
    wb = Workbook()
    ws_notes = wb.active
    ws_notes.title = "Notes"

    def qlabel(f):
        return f"Q{f['fq']} {f['fy']}"

    def filing_url(f):
        return (f"https://www.sec.gov/Archives/edgar/data/{cik}/"
                f"{f['accession'].replace('-', '')}/{f['document']}")

    def write_header(ws, title, subtitle, extra_row_label=None, extra_row_vals=None):
        ws["A1"] = title
        ws["A1"].font = F_TITLE
        ws["A2"] = subtitle
        ws["A2"].font = F_NOTE
        hdr_row = 4
        ws.cell(hdr_row, 1, "USD millions (except per share)").font = F_HDR
        ws.cell(hdr_row + 1, 1, "Period end").font = F_HDR
        for j, f in enumerate(chron, start=2):
            c = ws.cell(hdr_row, j, qlabel(f))
            c.font = F_HDR
            c.alignment = Alignment(horizontal="right")
            d = ws.cell(hdr_row + 1, j, pd.Timestamp(f["period_end"]).strftime("%d %b %Y"))
            d.hyperlink = filing_url(f)
            d.font = Font(name=FONT, size=9, color="0563C1", underline="single")
            d.alignment = Alignment(horizontal="right")
        ws.cell(hdr_row + 1, 1).font = Font(name=FONT, size=9, bold=True)
        for j in range(1, len(chron) + 2):
            ws.cell(hdr_row, j).fill = FILL_HDR
            ws.cell(hdr_row + 1, j).fill = FILL_SUB
        r = hdr_row + 2
        if extra_row_label:
            ws.cell(r, 1, extra_row_label).font = F_NOTE
            for j, v in enumerate(extra_row_vals, start=2):
                c = ws.cell(r, j, v)
                c.font = F_NOTE
                c.alignment = Alignment(horizontal="right")
            r += 1
        ws.column_dimensions["A"].width = 62
        for j in range(2, len(chron) + 2):
            ws.column_dimensions[get_column_letter(j)].width = 12.5
        ws.freeze_panes = ws.cell(hdr_row + 2 + (1 if extra_row_label else 0), 2)
        return r

    def write_label(ws, r, m):
        c = ws.cell(r, 1, m["label"])
        if m["is_header"]:
            c.font = F_SECTION
        else:
            c.font = Font(name=FONT, size=10, bold=is_bold(m["label"]))
            c.alignment = Alignment(indent=0 if is_bold(m["label"]) else 1)

    def style_total(ws, r, m):
        if not m["is_header"] and is_bold(m["label"]) and m["label"].startswith(("Total", "Net cash", "Income from")):
            for j in range(1, len(chron) + 2):
                ws.cell(r, j).border = TOP

    rowmap = {}  # (sheet, key) -> row number

    def write_reported(sheet_title, stmt, subtitle, extra=None):
        ws = wb.create_sheet(sheet_title)
        r = write_header(ws, f"{name} ({ticker}) — {sheet_title}", subtitle, *(extra or (None, None)))
        for m in masters[stmt]:
            write_label(ws, r, m)
            rowmap[(sheet_title, m["key"])] = r
            if not m["is_header"]:
                for j, f in enumerate(chron, start=2):
                    v = values[stmt].get(m["key"], {}).get(f["accession"])
                    if v is not None:
                        c = ws.cell(r, j, v)
                        c.font = F_INPUT
                        c.number_format = num_fmt(m["key"])
                ws.cell(r, 1).font = Font(name=FONT, size=10, bold=is_bold(m["label"]))
            style_total(ws, r, m)
            r += 1
        return ws, r

    # Income statement
    ws_is, r = write_reported("Income Statement", "income_statement",
                              "Three months ended each quarter, as originally reported in that quarter's 10-Q.")
    # Balance sheet
    ws_bs, r_bs = write_reported("Balance Sheet", "balance_sheet",
                                 "Balances at each quarter end, as originally reported in that quarter's 10-Q.")
    # Cash flow YTD
    ws_ytd, r_ytd = write_reported(
        "Cash Flow (YTD)", "cash_flow",
        "Year-to-date figures exactly as reported in each 10-Q. The quarterly sheet is calculated from these.",
        ("Months year-to-date", [f"{f['months']}M" for f in chron]))

    # Cash flow quarterly (formulas)
    ws_cf = wb.create_sheet("Cash Flow", 1)
    r = write_header(ws_cf, f"{name} ({ticker}) — Cash Flow (standalone quarters)",
                     "Each column is that quarter alone: Q1 = 3M YTD, Q2 = 6M − 3M, Q3 = 9M − 6M "
                     "(formulas link to the 'Cash Flow (YTD)' sheet).")
    ytd_ref = "'Cash Flow (YTD)'!"
    for m in masters["cash_flow"]:
        write_label(ws_cf, r, m)
        rowmap[("Cash Flow", m["key"])] = r
        if not m["is_header"]:
            yr = rowmap[("Cash Flow (YTD)", m["key"])]
            for j, f in enumerate(chron, start=2):
                col = get_column_letter(j)
                cur = values["cash_flow"].get(m["key"], {}).get(f["accession"])
                p = f["prior"]
                pcol = get_column_letter(chron.index(p) + 2) if p else None
                prev = values["cash_flow"].get(m["key"], {}).get(p["accession"]) if p else None
                c = ws_cf.cell(r, j)
                base = m["key"].split("#")[0]
                if m["key"] == CASH_BEGIN and p:
                    c.value = f"={pcol}{rowmap[('Cash Flow', CASH_END)] if ('Cash Flow', CASH_END) in rowmap else '{END}'}"
                    c.font = F_CALC
                elif m["key"] in (CASH_BEGIN, CASH_END) or base in POINT_IN_TIME_TAGS or p is None:
                    if cur is None:
                        continue
                    c.value = f"={ytd_ref}{col}{yr}"
                    c.font = F_LINK
                else:
                    if cur is None and prev is None:
                        continue
                    c.value = f"={ytd_ref}{col}{yr}-{ytd_ref}{pcol}{yr}"
                    c.font = F_CALC
                c.number_format = num_fmt(m["key"])
                if (m["key"], f["accession"]) in cf_flags:
                    c.comment = Comment(cf_flags[(m["key"], f["accession"])], "build_workbook")
                    c.fill = PatternFill("solid", fgColor="FFF2CC")
        style_total(ws_cf, r, m)
        r += 1
    # patch beginning-of-period formulas that were written before the end row existed
    end_row = rowmap[("Cash Flow", CASH_END)]
    for row in ws_cf.iter_rows():
        for c in row:
            if isinstance(c.value, str) and "{END}" in c.value:
                c.value = c.value.replace("{END}", str(end_row))
    r_cf = r

    # ------------------------------------------------ check rows
    def check_block(ws, r, checks):
        r += 1
        ws.cell(r, 1, "Checks (should all be zero)").font = F_SECTION
        r += 1
        for label, formula_fn in checks:
            ws.cell(r, 1, label).font = F_BASE
            ws.cell(r, 1).fill = FILL_CHECK
            for j, f in enumerate(chron, start=2):
                col = get_column_letter(j)
                c = ws.cell(r, j, formula_fn(col))
                c.font = F_CALC
                c.number_format = FMT_MM
                c.fill = FILL_CHECK
            r += 1
        return r

    def row(sheet, tag, occ=1):
        return rowmap[(sheet, f"{tag}#{occ}")]

    s = "Income Statement"
    r_is = ws_is.max_row + 1
    r_is = check_block(ws_is, r_is, [
        ("Revenues − total costs − operating income",
         lambda c: f"={c}{row(s,'us-gaap_Revenues')}-{c}{row(s,'us-gaap_CostsAndExpenses')}-{c}{row(s,'us-gaap_OperatingIncomeLoss')}"),
        ("Pre-tax income − tax − net income",
         lambda c: f"={c}{row(s,'us-gaap_IncomeLossFromContinuingOperationsBeforeIncomeTaxesExtraordinaryItemsNoncontrollingInterest')}"
                   f"-{c}{row(s,'us-gaap_IncomeTaxExpenseBenefit')}-{c}{row(s,'us-gaap_NetIncomeLoss')}"),
    ])
    s = "Balance Sheet"
    r_bs = check_block(ws_bs, ws_bs.max_row + 1, [
        ("Total assets − total liabilities & equity",
         lambda c: f"={c}{row(s,'us-gaap_Assets')}-{c}{row(s,'us-gaap_LiabilitiesAndStockholdersEquity')}"),
        ("Total liabilities + equity − total liabilities & equity",
         lambda c: f"={c}{row(s,'us-gaap_Liabilities')}+{c}{row(s,'us-gaap_StockholdersEquity')}"
                   f"-{c}{row(s,'us-gaap_LiabilitiesAndStockholdersEquity')}"),
    ])
    s = "Cash Flow"
    fx_tag = next(m["key"] for m in masters["cash_flow"] if m["key"].startswith("us-gaap_EffectOfExchangeRate"))
    chg_tag = next(m["key"] for m in masters["cash_flow"]
                   if m["key"].startswith("us-gaap_CashCashEquivalentsRestrictedCashAndRestrictedCashEquivalentsPeriodIncrease"))
    r_cf = check_block(ws_cf, r_cf, [
        ("Operating + investing + financing + FX − net change in cash",
         lambda c: f"={c}{row(s,'us-gaap_NetCashProvidedByUsedInOperatingActivities')}"
                   f"+{c}{row(s,'us-gaap_NetCashProvidedByUsedInInvestingActivities')}"
                   f"+{c}{row(s,'us-gaap_NetCashProvidedByUsedInFinancingActivities')}"
                   f"+{c}{rowmap[(s, fx_tag)]}-{c}{rowmap[(s, chg_tag)]}"),
        ("Beginning cash + net change − ending cash",
         lambda c: f"={c}{rowmap[(s, CASH_BEGIN)]}+{c}{rowmap[(s, chg_tag)]}-{c}{rowmap[(s, CASH_END)]}"),
        ("Ending cash − balance sheet cash & equivalents",
         lambda c: f"={c}{rowmap[(s, CASH_END)]}-'Balance Sheet'!{c}"
                   f"{rowmap[('Balance Sheet', 'us-gaap_CashAndCashEquivalentsAtCarryingValue#1')]}"),
    ])

    src = (f"Source: {name} Form 10-Q filings on SEC EDGAR (CIK {cik}); each period-end date in the header "
           f"links to its filing. Blue = as reported, black = formula, green = link to another sheet.")
    for ws, rr in [(ws_is, r_is), (ws_bs, r_bs), (ws_cf, r_cf), (ws_ytd, ws_ytd.max_row + 1)]:
        ws.cell(rr + 1, 1, src).font = F_NOTE
    flag_note = ("Yellow cells carry a comment: a line was added to or dropped from the statement partway through "
                 "a fiscal year, so that quarter's derived figure absorbs a reclassification.")
    ws_cf.cell(r_cf + 2, 1, flag_note).font = F_NOTE

    # ------------------------------------------------ notes sheet
    ws = ws_notes
    ws["A1"] = f"{name} ({ticker}) — quarterly financial statements"
    ws["A1"].font = F_TITLE
    notes = [
        ("Source", f"Last {n} Form 10-Q filings, SEC EDGAR (CIK {cik}). Statements parsed from EDGAR's "
                   f"rendered XBRL statement pages."),
        ("Basis", "As originally reported in each quarter's own 10-Q; later restatements or reclassifications "
                  "are not back-applied."),
        ("Coverage", "10-Qs cover fiscal Q1–Q3 only. Q4 is reported in the annual 10-K, so there is a gap each year."),
        ("Cash flow", "10-Q cash flows are year-to-date. 'Cash Flow' derives standalone quarters "
                      "(Q2 = 6M − 3M, Q3 = 9M − 6M) with formulas on 'Cash Flow (YTD)'. Beginning cash for Q2/Q3 "
                      "is the prior quarter's ending cash. 'Property and equipment included in accrued liabilities' "
                      "is a period-end balance, so it is shown as reported, not differenced."),
        ("Line matching", "Rows are matched across filings by XBRL tag, so relabelled lines stay aligned. "
                          "Re-tagged lines merged: Revenues; Property and equipment, net; Long-term debt; "
                          "Dividend payments; Depreciation (2023 label: 'Depreciation and impairment of property "
                          "and equipment'). Labels shown are from the most recent filing."),
        ("Blanks", "A blank cell means the line was not presented in that filing (often because it was combined "
                   "into another line)."),
        ("Colour key", "Blue = hardcoded from the filing · Black = formula · Green = link to another sheet · "
                       "Yellow fill = see cell comment."),
        ("Rebuild", f"python scripts/build_workbook.py --ticker {ticker} --cik {cik} --name \"{name}\" --n {n}"),
    ]
    rr = 3
    for k, v in notes:
        ws.cell(rr, 1, k).font = Font(name=FONT, size=10, bold=True)
        c = ws.cell(rr, 2, v)
        c.font = F_BASE
        c.alignment = Alignment(wrap_text=True, vertical="top")
        ws.cell(rr, 1).alignment = Alignment(vertical="top")
        rr += 1
    rr += 1
    hdrs = ["Quarter", "Period end", "Filed", "Accession no.", "Filing"]
    for j, h in enumerate(hdrs, start=1):
        c = ws.cell(rr, j, h)
        c.font = F_HDR
        c.fill = FILL_HDR
    for f in reversed(chron):
        rr += 1
        vals = [qlabel(f), f["period_end"], f["filed"], f["accession"], "Open 10-Q"]
        for j, v in enumerate(vals, start=1):
            ws.cell(rr, j, v).font = F_BASE
        ws.cell(rr, 5).hyperlink = filing_url(f)
        ws.cell(rr, 5).font = Font(name=FONT, size=10, color="0563C1", underline="single")
    ws.column_dimensions["A"].width = 16
    ws.column_dimensions["B"].width = 95
    ws.column_dimensions["C"].width = 12
    ws.column_dimensions["D"].width = 22
    ws.column_dimensions["E"].width = 12

    # sheet order: Notes, Income Statement, Balance Sheet, Cash Flow, Cash Flow (YTD)
    wb._sheets = [wb["Notes"], wb["Income Statement"], wb["Balance Sheet"], wb["Cash Flow"], wb["Cash Flow (YTD)"]]
    for w in wb.worksheets:
        w.sheet_view.showGridLines = False

    out = ROOT / "workbooks" / f"{ticker}.xlsx"
    out.parent.mkdir(parents=True, exist_ok=True)
    wb.save(out)
    print(f"Saved {out}")
    return out


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--ticker", required=True)
    ap.add_argument("--cik", type=int, required=True)
    ap.add_argument("--name", required=True)
    ap.add_argument("--n", type=int, default=12, help="number of most recent 10-Qs")
    a = ap.parse_args()
    build(a.ticker.upper(), a.cik, a.name, a.n)
