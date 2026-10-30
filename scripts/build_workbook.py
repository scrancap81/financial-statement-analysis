"""
Build a quarterly financial statements workbook for one company from its 10-Qs and 10-Ks.

Usage:
    python scripts/build_workbook.py --ticker GOOGL --cik 1652044 --name "Alphabet Inc." --n 12

Outputs:
    workbooks/<TICKER>.xlsx              Excel workbook (statements + notes)
    data/<TICKER>/*.csv                  the same statements as CSVs for pandas work

Quarters:
    Q1–Q3 come from the 10-Qs. Q4 is not filed on its own, so it is derived from
    the 10-K for every fiscal year whose Q3 10-Q is in range:
      * Balance sheet: the 10-K's year-end balance sheet, as reported.
      * Income statement: Q4 = full year − Q1 − Q2 − Q3.
      * Cash flow: 10-Q/10-K cash flows are year-to-date, so standalone quarters are
        Q1 = 3M, Q2 = 6M − 3M, Q3 = 9M − 6M, Q4 = 12M − 9M.
    In the workbook these are live formulas pointing at the 'Income Statement (FY)'
    and 'Cash Flow (YTD)' sheets, so every derived number traces back to a filing.
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

STATEMENTS = ["income_statement", "balance_sheet", "cash_flow"]

EPS_Q4_NOTE = ("Not derived: per-share figures can't be found by subtracting year-to-date EPS, because "
               "the share count differs between periods. Q4 EPS is in the company's Q4 earnings release.")


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


def is_eps(key):
    return "PerShare" in key


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
FILL_FLAG = PatternFill("solid", fgColor="FFF2CC")
TOP = Border(top=Side(style="thin"))
FMT_MM = '#,##0;(#,##0);"-"'
FMT_EPS = '0.00;(0.00);"-"'

BOLD_PREFIXES = ("Total", "Net income", "Income from operations", "Income before",
                 "Net cash", "Net increase", "Revenues")


def is_bold(label):
    return label.startswith(BOLD_PREFIXES)


def num_fmt(key):
    return FMT_EPS if is_eps(key) else FMT_MM


# ---------------------------------------------------------------- build

def build(ticker: str, cik: int, name: str, n: int, include_10k: bool = True):
    tenqs = sec_fetch.list_filings(cik, "10-Q", n)
    tenks = sec_fetch.list_filings(cik, "10-K") if include_10k else []
    earliest = min(f["period_end"] for f in tenqs)
    tenks = [f for f in tenks if f["period_end"] > earliest]

    data = {}
    for f in tenqs + tenks:
        print(f"  {f['form']}  {f['period_end']}  {f['accession']}")
        data[f["accession"]] = sec_fetch.fetch_filing_statements(cik, f)

    # Fiscal quarter from the cash flow's year-to-date length (3/6/9/12 months)
    for f in tenqs + tenks:
        cf = data[f["accession"]]["cash_flow"]
        dur = cf["columns"][pick_column(cf, f["period_end"], None)][0]
        f["months"] = int(re.match(r"(\d+)", dur).group(1))
        f["fq"] = f["months"] // 3
        f["fy"] = int(f["period_end"][:4])

    # Only keep 10-Ks whose Q1–Q3 are all in range, so Q4 can be derived
    have = {(f["fy"], f["fq"]) for f in tenqs}
    tenks = [f for f in tenks if all((f["fy"], q) in have for q in (1, 2, 3))]

    filings = tenqs + tenks
    chron = sorted(filings, key=lambda x: x["period_end"])      # quarterly columns, oldest first
    annual = sorted(tenks, key=lambda x: x["period_end"])       # FY columns
    newest_first = list(reversed(chron))

    by_fy_q = {(f["fy"], f["fq"]): f for f in filings}
    for f in filings:
        f["prior"] = by_fy_q.get((f["fy"], f["fq"] - 1)) if f["fq"] > 1 else None
        f["year_quarters"] = [by_fy_q[(f["fy"], q)] for q in (1, 2, 3)] if f["fq"] == 4 else []

    # values[stmt][key][accession] = number as reported
    #   income statement: 3 months for 10-Qs, 12 months for 10-Ks
    #   balance sheet: period end ; cash flow: year-to-date
    values, present, masters = {}, {}, {}
    for stmt in STATEMENTS:
        values[stmt], present[stmt] = {}, {}
        rows_by_filing = []
        for f in newest_first:
            st = data[f["accession"]][stmt]
            dur = None
            if stmt == "income_statement":
                dur = "12 Months" if f["fq"] == 4 else "3 Months"
            col = pick_column(st, f["period_end"], dur)
            rows = keyed_rows(st["rows"])
            rows_by_filing.append(rows)
            present[stmt][f["accession"]] = {r["key"] for r in rows}
            for r in rows:
                if not r["is_header"]:
                    values[stmt].setdefault(r["key"], {})[f["accession"]] = r["values"][col]
        masters[stmt] = master_order(rows_by_filing)
        for m in masters[stmt]:
            if m["key"] in values[stmt]:
                m["is_header"] = False

    def reclass_flag(stmt, k, f, p, cur, prev):
        a, pa = f["accession"], p["accession"]
        in_cur, in_prev = k in present[stmt][a], k in present[stmt][pa]
        if in_cur and not in_prev and cur:
            return (f"Line first presented in the {f['form']} for Q{f['fq']} {f['fy']}, so the earlier "
                    f"period is treated as zero. If this activity also happened earlier in the year, it was "
                    f"reported within another line and this quarter's figure includes it (subtotals are unaffected).")
        if in_prev and not in_cur and prev:
            return (f"Line not presented in the {f['form']} for Q{f['fq']} {f['fy']} (folded into another line), "
                    f"so this quarter shows the reversal of the earlier amount.")
        return None

    # ------------------------------------------------ quarterly income statement (Python, for CSV + flags)
    is_q, is_flags = {}, {}
    for m in masters["income_statement"]:
        if m["is_header"]:
            continue
        k = m["key"]
        is_q[k] = {}
        for f in chron:
            a = f["accession"]
            v = values["income_statement"].get(k, {}).get(a)
            if f["fq"] != 4:
                is_q[k][a] = v
                continue
            if is_eps(k):
                is_q[k][a] = None
                continue
            qs = [values["income_statement"].get(k, {}).get(q["accession"]) for q in f["year_quarters"]]
            if v is None and all(x is None for x in qs):
                is_q[k][a] = None
                continue
            is_q[k][a] = (v or 0) - sum(x or 0 for x in qs)
            flag = reclass_flag("income_statement", k, f, f["prior"], v, qs[-1])
            if flag:
                is_flags[(k, a)] = flag

    # ------------------------------------------------ standalone-quarter cash flow
    # Supplemental lines below the cash reconciliation that a 10-K discloses but that year's
    # 10-Qs don't (e.g. cash paid for taxes) are annual-only: differencing them would put
    # the whole year into Q4, so they are left blank in the quarterly view.
    cf_keys = [m["key"] for m in masters["cash_flow"]]
    supplemental = set(cf_keys[cf_keys.index(CASH_END) + 1:])

    def annual_only(k, f):
        return (f["fq"] == 4 and k in supplemental
                and not any(k in present["cash_flow"][q["accession"]] for q in f["year_quarters"]))

    ANNUAL_ONLY_NOTE = ("Disclosed only for the full year in the 10-K, so no Q4 figure can be derived. "
                        "The full-year amount is on 'Cash Flow (YTD)'.")
    cf_q, cf_flags = {}, {}
    for m in masters["cash_flow"]:
        if m["is_header"]:
            continue
        k = m["key"]
        cf_q[k] = {}
        for f in chron:
            a, p = f["accession"], f["prior"]
            cur = values["cash_flow"].get(k, {}).get(a)
            if annual_only(k, f):
                cf_q[k][a] = None
                continue
            if k in (CASH_BEGIN, CASH_END) or k.split("#")[0] in POINT_IN_TIME_TAGS or p is None:
                cf_q[k][a] = cur          # beginning cash for Q2–Q4 replaced below
                continue
            prev = values["cash_flow"].get(k, {}).get(p["accession"])
            if cur is None and prev is None:
                cf_q[k][a] = None
                continue
            cf_q[k][a] = (cur or 0) - (prev or 0)
            flag = reclass_flag("cash_flow", k, f, p, cur, prev)
            if flag:
                cf_flags[(k, a)] = flag
    for f in chron:
        if f["prior"]:
            cf_q[CASH_BEGIN][f["accession"]] = cf_q[CASH_END][f["prior"]["accession"]]

    # ------------------------------------------------ CSVs
    out_data = ROOT / "data" / ticker
    out_data.mkdir(parents=True, exist_ok=True)

    def frame(stmt, vals, cols, colname):
        recs = []
        for m in masters[stmt]:
            if m["is_header"]:
                continue
            recs.append([m["label"], m["key"].split("#")[0]] +
                        [vals.get(m["key"], {}).get(f["accession"]) for f in cols])
        return pd.DataFrame(recs, columns=["line_item", "xbrl_tag"] + [colname(f) for f in cols])

    by_date = lambda f: f["period_end"]
    frame("income_statement", is_q, chron, by_date).to_csv(out_data / "income_statement.csv", index=False)
    frame("income_statement", values["income_statement"], annual,
          lambda f: f"FY{f['fy']}").to_csv(out_data / "income_statement_fy.csv", index=False)
    frame("balance_sheet", values["balance_sheet"], chron, by_date).to_csv(out_data / "balance_sheet.csv", index=False)
    frame("cash_flow", values["cash_flow"], chron, by_date).to_csv(out_data / "cash_flow_ytd.csv", index=False)
    frame("cash_flow", cf_q, chron, by_date).to_csv(out_data / "cash_flow_quarterly.csv", index=False)

    # ------------------------------------------------ workbook
    wb = Workbook()
    ws_notes = wb.active
    ws_notes.title = "Notes"

    def qlabel(f):
        return f"Q{f['fq']} {f['fy']}"

    def filing_url(f):
        return (f"https://www.sec.gov/Archives/edgar/data/{cik}/"
                f"{f['accession'].replace('-', '')}/{f['document']}")

    def write_header(ws, title, subtitle, cols, col_label, extra_rows=()):
        """Title, subtitle, a dark header row of period labels, a row of linked period-end
        dates, then any extra descriptor rows. Returns the first data row."""
        ws["A1"] = title
        ws["A1"].font = F_TITLE
        ws["A2"] = subtitle
        ws["A2"].font = F_NOTE
        h = 4
        ws.cell(h, 1, "USD millions (except per share)").font = F_HDR
        ws.cell(h + 1, 1, "Period end").font = Font(name=FONT, size=9, bold=True)
        for j, f in enumerate(cols, start=2):
            c = ws.cell(h, j, col_label(f))
            c.font = F_HDR
            c.alignment = Alignment(horizontal="right")
            d = ws.cell(h + 1, j, pd.Timestamp(f["period_end"]).strftime("%d %b %Y"))
            d.hyperlink = filing_url(f)
            d.font = Font(name=FONT, size=9, color="0563C1", underline="single")
            d.alignment = Alignment(horizontal="right")
        for j in range(1, len(cols) + 2):
            ws.cell(h, j).fill = FILL_HDR
            ws.cell(h + 1, j).fill = FILL_SUB
        r = h + 2
        for label, fn in extra_rows:
            ws.cell(r, 1, label).font = F_NOTE
            for j, f in enumerate(cols, start=2):
                c = ws.cell(r, j, fn(f))
                c.font = F_NOTE
                c.alignment = Alignment(horizontal="right")
            r += 1
        ws.column_dimensions["A"].width = 62
        for j in range(2, len(cols) + 2):
            ws.column_dimensions[get_column_letter(j)].width = 12.5
        ws.freeze_panes = ws.cell(r, 2)
        return r

    def write_label(ws, r, m):
        c = ws.cell(r, 1, m["label"])
        if m["is_header"]:
            c.font = F_SECTION
        else:
            c.font = Font(name=FONT, size=10, bold=is_bold(m["label"]))
            c.alignment = Alignment(indent=0 if is_bold(m["label"]) else 1)

    def style_total(ws, r, m, ncols):
        if not m["is_header"] and m["label"].startswith(("Total", "Net cash", "Income from")):
            for j in range(1, ncols + 2):
                ws.cell(r, j).border = TOP

    def flag(c, text):
        c.comment = Comment(text, "build_workbook")
        c.comment.width, c.comment.height = 320, 110
        c.fill = FILL_FLAG

    rowmap = {}  # (sheet, key) -> row number
    col_of = {f["accession"]: get_column_letter(j) for j, f in enumerate(chron, start=2)}

    def write_reported(sheet_title, stmt, subtitle, cols, col_label, extra_rows=(), skip_q4=False):
        """Sheet of as-reported values (blue). skip_q4 leaves 10-K columns for formulas."""
        ws = wb.create_sheet(sheet_title)
        r = write_header(ws, f"{name} ({ticker}) — {sheet_title}", subtitle, cols, col_label, extra_rows)
        for m in masters[stmt]:
            write_label(ws, r, m)
            rowmap[(sheet_title, m["key"])] = r
            if not m["is_header"]:
                for j, f in enumerate(cols, start=2):
                    if skip_q4 and f["fq"] == 4:
                        continue
                    v = values[stmt].get(m["key"], {}).get(f["accession"])
                    if v is not None:
                        c = ws.cell(r, j, v)
                        c.font = F_INPUT
                        c.number_format = num_fmt(m["key"])
            style_total(ws, r, m, len(cols))
            r += 1
        return ws

    filing_row = ("Source", lambda f: "10-Q" if f["form"] == "10-Q" else "10-K, derived")

    # Income statement (FY), Cash flow (YTD): the as-reported inputs
    ws_fy = write_reported(
        "Income Statement (FY)", "income_statement",
        "Full fiscal year as reported in each 10-K. Q4 on the 'Income Statement' sheet is derived from these.",
        annual, lambda f: f"FY{f['fy']}")
    ws_ytd = write_reported(
        "Cash Flow (YTD)", "cash_flow",
        "Year-to-date figures exactly as reported in each 10-Q and 10-K. The 'Cash Flow' sheet is calculated from these.",
        chron, qlabel,
        [("Months year-to-date", lambda f: f"{f['months']}M"), ("Source", lambda f: f["form"])])

    # Income statement (quarterly): 10-Q values as reported; Q4 = FY − Q1 − Q2 − Q3
    ws_is = write_reported(
        "Income Statement", "income_statement",
        "Three months ended each quarter. Q1–Q3 as originally reported in each 10-Q; "
        "Q4 = full year (10-K) − Q1 − Q2 − Q3, as formulas.",
        chron, qlabel, [filing_row], skip_q4=True)
    fy_col = {f["accession"]: get_column_letter(j) for j, f in enumerate(annual, start=2)}
    for m in masters["income_statement"]:
        if m["is_header"]:
            continue
        r = rowmap[("Income Statement", m["key"])]
        fr = rowmap[("Income Statement (FY)", m["key"])]
        for f in chron:
            if f["fq"] != 4:
                continue
            a = f["accession"]
            c = ws_is[f"{col_of[a]}{r}"]
            if is_eps(m["key"]):
                if values["income_statement"].get(m["key"], {}).get(a) is not None:
                    c.comment = Comment(EPS_Q4_NOTE, "build_workbook")
                    c.comment.width, c.comment.height = 300, 90
                continue
            if is_q[m["key"]].get(a) is None:
                continue
            qcells = "+".join(f"{col_of[q['accession']]}{r}" for q in f["year_quarters"])
            c.value = f"='Income Statement (FY)'!{fy_col[a]}{fr}-({qcells})"
            c.font = F_CALC
            c.number_format = num_fmt(m["key"])
            if (m["key"], a) in is_flags:
                flag(c, is_flags[(m["key"], a)])

    # Balance sheet: all as reported (Q4 = 10-K year-end balance sheet)
    ws_bs = write_reported(
        "Balance Sheet", "balance_sheet",
        "Balances at each quarter end, as originally reported (Q1–Q3 from 10-Qs, Q4 from the 10-K).",
        chron, qlabel, [("Source", lambda f: f["form"])])

    # Cash flow (standalone quarters): formulas on the YTD sheet
    ws_cf = wb.create_sheet("Cash Flow")
    r = write_header(ws_cf, f"{name} ({ticker}) — Cash Flow (standalone quarters)",
                     "Each column is that quarter alone: Q1 = 3M, Q2 = 6M − 3M, Q3 = 9M − 6M, Q4 = 12M − 9M "
                     "(formulas on the 'Cash Flow (YTD)' sheet).",
                     chron, qlabel, [filing_row])
    ytd = "'Cash Flow (YTD)'!"
    for m in masters["cash_flow"]:
        write_label(ws_cf, r, m)
        rowmap[("Cash Flow", m["key"])] = r
        r += 1
    for m in masters["cash_flow"]:
        if m["is_header"]:
            style_total(ws_cf, rowmap[("Cash Flow", m["key"])], m, len(chron))
            continue
        k, r = m["key"], rowmap[("Cash Flow", m["key"])]
        yr = rowmap[("Cash Flow (YTD)", k)]
        for f in chron:
            a, p = f["accession"], f["prior"]
            col = col_of[a]
            cur = values["cash_flow"].get(k, {}).get(a)
            prev = values["cash_flow"].get(k, {}).get(p["accession"]) if p else None
            c = ws_cf[f"{col}{r}"]
            if annual_only(k, f):
                if cur is not None:
                    c.comment = Comment(ANNUAL_ONLY_NOTE, "build_workbook")
                    c.comment.width, c.comment.height = 300, 80
                continue
            if k == CASH_BEGIN and p:
                c.value = f"={col_of[p['accession']]}{rowmap[('Cash Flow', CASH_END)]}"
                c.font = F_CALC
            elif k in (CASH_BEGIN, CASH_END) or k.split("#")[0] in POINT_IN_TIME_TAGS or p is None:
                if cur is None:
                    continue
                c.value = f"={ytd}{col}{yr}"
                c.font = F_LINK
            else:
                if cur is None and prev is None:
                    continue
                c.value = f"={ytd}{col}{yr}-{ytd}{col_of[p['accession']]}{yr}"
                c.font = F_CALC
            c.number_format = num_fmt(k)
            if (k, a) in cf_flags:
                flag(c, cf_flags[(k, a)])
        style_total(ws_cf, r, m, len(chron))
    r_cf = r

    # ------------------------------------------------ check rows
    def check_block(ws, r, checks, cols):
        r += 1
        ws.cell(r, 1, "Checks (should all be zero)").font = F_SECTION
        r += 1
        for label, formula_fn in checks:
            ws.cell(r, 1, label).font = F_BASE
            ws.cell(r, 1).fill = FILL_CHECK
            for j, f in enumerate(cols, start=2):
                col = get_column_letter(j)
                c = ws.cell(r, j, formula_fn(col))
                c.font = F_CALC
                c.number_format = FMT_MM
                c.fill = FILL_CHECK
            r += 1
        return r

    def row(sheet, tag, occ=1):
        return rowmap[(sheet, f"{tag}#{occ}")]

    PRETAX = "us-gaap_IncomeLossFromContinuingOperationsBeforeIncomeTaxesExtraordinaryItemsNoncontrollingInterest"

    def is_checks(s):
        return [
            ("Revenues − total costs − operating income",
             lambda c: f"={c}{row(s,'us-gaap_Revenues')}-{c}{row(s,'us-gaap_CostsAndExpenses')}"
                       f"-{c}{row(s,'us-gaap_OperatingIncomeLoss')}"),
            ("Pre-tax income − tax − net income",
             lambda c: f"={c}{row(s,PRETAX)}-{c}{row(s,'us-gaap_IncomeTaxExpenseBenefit')}"
                       f"-{c}{row(s,'us-gaap_NetIncomeLoss')}"),
        ]

    end = {}
    end["Income Statement"] = check_block(ws_is, ws_is.max_row + 1, is_checks("Income Statement"), chron)
    end["Income Statement (FY)"] = check_block(ws_fy, ws_fy.max_row + 1, is_checks("Income Statement (FY)"), annual)
    s = "Balance Sheet"
    end[s] = check_block(ws_bs, ws_bs.max_row + 1, [
        ("Total assets − total liabilities & equity",
         lambda c: f"={c}{row(s,'us-gaap_Assets')}-{c}{row(s,'us-gaap_LiabilitiesAndStockholdersEquity')}"),
        ("Total liabilities + equity − total liabilities & equity",
         lambda c: f"={c}{row(s,'us-gaap_Liabilities')}+{c}{row(s,'us-gaap_StockholdersEquity')}"
                   f"-{c}{row(s,'us-gaap_LiabilitiesAndStockholdersEquity')}"),
    ], chron)
    s = "Cash Flow"
    fx_key = next(m["key"] for m in masters["cash_flow"] if m["key"].startswith("us-gaap_EffectOfExchangeRate"))
    chg_key = next(m["key"] for m in masters["cash_flow"]
                   if m["key"].startswith(CASH_TAG + "PeriodIncrease"))
    end[s] = check_block(ws_cf, r_cf, [
        ("Operating + investing + financing + FX − net change in cash",
         lambda c: f"={c}{row(s,'us-gaap_NetCashProvidedByUsedInOperatingActivities')}"
                   f"+{c}{row(s,'us-gaap_NetCashProvidedByUsedInInvestingActivities')}"
                   f"+{c}{row(s,'us-gaap_NetCashProvidedByUsedInFinancingActivities')}"
                   f"+{c}{rowmap[(s, fx_key)]}-{c}{rowmap[(s, chg_key)]}"),
        ("Beginning cash + net change − ending cash",
         lambda c: f"={c}{rowmap[(s, CASH_BEGIN)]}+{c}{rowmap[(s, chg_key)]}-{c}{rowmap[(s, CASH_END)]}"),
        ("Ending cash − balance sheet cash & equivalents",
         lambda c: f"={c}{rowmap[(s, CASH_END)]}-'Balance Sheet'!{c}"
                   f"{row('Balance Sheet', 'us-gaap_CashAndCashEquivalentsAtCarryingValue')}"),
    ], chron)
    end["Cash Flow (YTD)"] = ws_ytd.max_row + 1

    src = (f"Source: {name} 10-Q and 10-K filings on SEC EDGAR (CIK {cik}); each period-end date in the header "
           f"links to its filing. Blue = as reported, black = formula, green = link to another sheet.")
    flag_note = ("Yellow cells carry a comment: a line was added to or dropped from the statement partway through "
                 "a fiscal year, so that quarter's derived figure absorbs a reclassification.")
    for title, rr in end.items():
        wb[title].cell(rr + 1, 1, src).font = F_NOTE
        if title in ("Cash Flow", "Income Statement"):
            wb[title].cell(rr + 2, 1, flag_note).font = F_NOTE
    ws_is.cell(end["Income Statement"] + 3, 1, "Q4 per-share figures are left blank: " + EPS_Q4_NOTE).font = F_NOTE

    # ------------------------------------------------ notes sheet
    ws = ws_notes
    ws["A1"] = f"{name} ({ticker}) — quarterly financial statements"
    ws["A1"].font = F_TITLE
    q4_years = ", ".join(f"FY{f['fy']}" for f in annual) or "none"
    notes = [
        ("Source", f"Last {n} Form 10-Q filings plus the 10-Ks for {q4_years}, SEC EDGAR (CIK {cik}). "
                   f"Statements parsed from EDGAR's rendered XBRL statement pages."),
        ("Basis", "As originally reported in each filing; later restatements or reclassifications are not back-applied."),
        ("Q4", "Companies don't file a Q4 report, so Q4 comes from the 10-K. Balance sheet: the 10-K year-end balance "
               "sheet. Income statement: full year − Q1 − Q2 − Q3 (formulas on 'Income Statement (FY)'). Cash flow: "
               "12M − 9M. Q4 per-share figures are left blank because EPS can't be derived by subtraction. "
               "A year's Q4 is only included once its 10-K is filed."),
        ("Cash flow", "10-Q and 10-K cash flows are year-to-date. 'Cash Flow' derives standalone quarters "
                      "(Q2 = 6M − 3M, Q3 = 9M − 6M, Q4 = 12M − 9M) with formulas on 'Cash Flow (YTD)'. Beginning cash "
                      "for Q2–Q4 is the prior quarter's ending cash. 'Property and equipment included in accrued "
                      "liabilities' is a period-end balance, so it is shown as reported, not differenced."),
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
    for j, h in enumerate(["Quarter", "Filing (click to open)"], start=1):
        c = ws.cell(rr, j, h)
        c.font = F_HDR
        c.fill = FILL_HDR
    for f in reversed(chron):
        rr += 1
        ws.cell(rr, 1, qlabel(f)).font = F_BASE
        c = ws.cell(rr, 2, f"{f['form']}  ·  period ended {f['period_end']}  ·  filed {f['filed']}  ·  "
                           f"accession {f['accession']}")
        c.hyperlink = filing_url(f)
        c.font = Font(name=FONT, size=10, color="0563C1", underline="single")
    ws.column_dimensions["A"].width = 16
    ws.column_dimensions["B"].width = 95

    order = ["Notes", "Income Statement", "Balance Sheet", "Cash Flow", "Income Statement (FY)", "Cash Flow (YTD)"]
    wb._sheets = [wb[t] for t in order]
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
    ap.add_argument("--no-10k", action="store_true", help="skip 10-Ks (no Q4 columns)")
    a = ap.parse_args()
    build(a.ticker.upper(), a.cik, a.name, a.n, include_10k=not a.no_10k)
