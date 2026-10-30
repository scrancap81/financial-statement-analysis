"""
Fetch quarterly financial statements from SEC EDGAR 10-Q filings.

For each 10-Q the script reads FilingSummary.xml to find the three primary
statements, downloads EDGAR's rendered statement pages (R*.htm) and parses
every line item along with its XBRL tag. Line items are matched across
filings by XBRL tag, so a relabelled line (e.g. a changed note number)
still lines up in the same row.

All figures are "as originally reported" in each quarter's own 10-Q.
"""

import json
import re
import time
from pathlib import Path
from datetime import datetime

import requests
from bs4 import BeautifulSoup

# The SEC asks automated clients to identify themselves in the User-Agent.
# Put your own name/email here if you run this regularly.
USER_AGENT = "financial-statement-analysis research script admin@example.com"
HEADERS = {"User-Agent": USER_AGENT, "Accept-Encoding": "gzip, deflate"}

CACHE_DIR = Path(__file__).resolve().parent.parent / "data" / "_cache"

STATEMENT_KEYWORDS = {
    "balance_sheet": "BALANCE SHEET",
    "income_statement": "STATEMENTS OF INCOME",
    "cash_flow": "CASH FLOWS",
}

MONTHS = {m: i for i, m in enumerate(
    ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"], 1)}


def _get(url: str, cache_name: str | None = None) -> str:
    """GET a URL politely (SEC limit is 10 req/s), caching to disk."""
    if cache_name:
        path = CACHE_DIR / cache_name
        if path.exists():
            return path.read_text(encoding="utf-8")
    time.sleep(0.15)
    r = requests.get(url, headers=HEADERS, timeout=30)
    r.raise_for_status()
    text = r.text
    if cache_name:
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_text(text, encoding="utf-8")
    return text


def list_10q_filings(cik: int, n: int = 12) -> list[dict]:
    """Return the n most recent 10-Q filings (newest first)."""
    cik10 = f"{cik:010d}"
    base = json.loads(_get(f"https://data.sec.gov/submissions/CIK{cik10}.json"))
    blocks = [base["filings"]["recent"]]
    # Companies with lots of filings (e.g. Form 4s) spill older ones into extra files
    for f in base["filings"].get("files", []):
        blocks.append(json.loads(_get(f"https://data.sec.gov/submissions/{f['name']}")))

    filings = []
    for b in blocks:
        for i, form in enumerate(b["form"]):
            if form == "10-Q":
                filings.append({
                    "accession": b["accessionNumber"][i],
                    "period_end": b["reportDate"][i],
                    "filed": b["filingDate"][i],
                    "document": b["primaryDocument"][i],
                })
    filings.sort(key=lambda x: x["period_end"], reverse=True)
    return filings[:n]


def _parse_value(text: str):
    text = text.strip().replace("$", "").replace(",", "").strip()
    text = re.sub(r"\[\d+\]", "", text).strip()   # footnote markers
    if text in ("", "—", "-"):
        return None
    neg = text.startswith("(") and text.endswith(")")
    text = text.strip("()").strip()
    try:
        v = float(text)
    except ValueError:
        return None
    return -v if neg else v


def _parse_header_date(s: str):
    m = re.match(r"([A-Z][a-z]{2})\.? (\d{1,2}), (\d{4})", s.strip())
    if not m:
        return None
    return datetime(int(m.group(3)), MONTHS[m.group(1)], int(m.group(2))).date().isoformat()


def parse_statement_page(html: str) -> dict:
    """
    Parse an EDGAR R*.htm statement page.

    Returns {"columns": [(duration_label, date_iso), ...],
             "rows": [{"tag", "label", "is_header", "values": [...]}, ...],
             "units": header text}
    """
    soup = BeautifulSoup(html, "lxml")
    table = soup.find("table")
    trs = table.find_all("tr")

    # --- header rows -> one (duration, date) per value column
    header_rows = [tr for tr in trs if tr.find("th")]
    title = header_rows[0].find("th").get_text(" ", strip=True)
    durations, dates = [], []
    if len(header_rows) >= 2:
        for th in header_rows[0].find_all("th")[1:]:
            durations += [th.get_text(strip=True)] * int(th.get("colspan", 1))
        dates = [th.get_text(strip=True) for th in header_rows[1].find_all("th")]
    else:  # balance sheets: a single header row of dates
        dates = [th.get_text(strip=True) for th in header_rows[0].find_all("th")[1:]]
        durations = ["Instant"] * len(dates)
    columns = list(zip(durations, [_parse_header_date(d) for d in dates]))

    rows = []
    for tr in trs:
        cls = tr.get("class") or []
        if not any(c in ("re", "ro", "reu", "rou", "rh") for c in cls):
            continue
        tds = tr.find_all("td")
        if not tds:
            continue
        a = tds[0].find("a")
        m = re.search(r"defref_([^'\"]+)", a.get("onclick", "")) if a else None
        tag = m.group(1) if m else None
        label = tds[0].get_text(" ", strip=True)
        values = [_parse_value(td.get_text(" ", strip=True)) for td in tds[1:]]
        is_header = all(v is None for v in values)
        rows.append({"tag": tag, "label": label, "is_header": is_header, "values": values})
    return {"title": title, "columns": columns, "rows": rows}


def fetch_filing_statements(cik: int, filing: dict) -> dict:
    """Download and parse the BS, IS and CF for one filing."""
    acc_nodash = filing["accession"].replace("-", "")
    base = f"https://www.sec.gov/Archives/edgar/data/{cik}/{acc_nodash}"
    summary = _get(f"{base}/FilingSummary.xml", f"{cik}/{acc_nodash}/FilingSummary.xml")
    soup = BeautifulSoup(summary, "xml")

    found = {}
    for rep in soup.find_all("Report"):
        if (rep.find("MenuCategory") and rep.find("MenuCategory").text != "Statements"):
            continue
        short = (rep.find("ShortName").text or "").upper()
        if "PARENTHETICAL" in short:
            continue
        for key, kw in STATEMENT_KEYWORDS.items():
            if key not in found and kw in short:
                found[key] = rep.find("HtmlFileName").text

    missing = set(STATEMENT_KEYWORDS) - set(found)
    if missing:
        raise RuntimeError(f"{filing['accession']}: could not find {missing}")

    out = {}
    for key, fname in found.items():
        html = _get(f"{base}/{fname}", f"{cik}/{acc_nodash}/{fname}")
        out[key] = parse_statement_page(html)
    return out
