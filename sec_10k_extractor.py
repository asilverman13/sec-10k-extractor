#!/usr/bin/env python3
"""
SEC EDGAR 10-K Financial Data Extractor

Usage:
    python3 sec_10k_extractor.py AAPL
    python3 sec_10k_extractor.py AAPL,MSFT,GOOG
    python3 sec_10k_extractor.py XOM,CVX --years 3 --export results.xlsx
"""

from __future__ import annotations

import re
import sys
import time
import argparse
import requests
from datetime import datetime
from typing import Optional
from bs4 import BeautifulSoup
from tabulate import tabulate

# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------

HEADERS = {
    "User-Agent": "FinancialResearchTool adogggoat13@icloud.com",
    "Accept-Encoding": "gzip, deflate",
    "Host": "data.sec.gov",
}
HTML_HEADERS = {
    "User-Agent": "FinancialResearchTool adogggoat13@icloud.com",
    "Accept-Encoding": "gzip, deflate",
}

OIL_GAS_SIC = {
    "1311", "1321", "1381", "1382", "1389", "2911", "5171", "5172",
}

# Metrics displayed in order; oil/gas-only ones are filtered later
METRIC_KEYS = [
    "revenue", "ebitda", "net_income", "total_assets",
    "total_debt", "cash", "equity", "pv10", "proved_reserves",
]
METRIC_LABELS = {
    "revenue":         "Revenue",
    "ebitda":          "EBITDA (est.)",
    "net_income":      "Net Income",
    "total_assets":    "Total Assets",
    "total_debt":      "Total Debt",
    "cash":            "Cash",
    "equity":          "Equity",
    "pv10":            "PV-10",
    "proved_reserves": "Proved Reserves",
}

# ---------------------------------------------------------------------------
# EDGAR lookup helpers
# ---------------------------------------------------------------------------

def get_cik(ticker: str) -> tuple[str, str]:
    """Return (cik_padded, company_name) for a ticker."""
    url = "https://www.sec.gov/files/company_tickers.json"
    r = requests.get(url, headers=HTML_HEADERS, timeout=15)
    r.raise_for_status()
    ticker_upper = ticker.upper()
    for entry in r.json().values():
        if entry["ticker"].upper() == ticker_upper:
            return str(entry["cik_str"]).zfill(10), entry["title"]
    raise ValueError(f"Ticker '{ticker}' not found in SEC EDGAR.")


def get_company_info(cik: str) -> dict:
    url = f"https://data.sec.gov/submissions/CIK{cik}.json"
    r = requests.get(url, headers=HEADERS, timeout=15)
    r.raise_for_status()
    return r.json()


def get_10k_filings(submissions: dict, n: int = 3) -> list[dict]:
    """
    Return the most recent `n` distinct fiscal-year 10-K filings as a list of
    {"accession": ..., "date": ..., "original_accession": ..., "fiscal_year": ...}

    When a 10-K/A is the newest filing for a fiscal year, we also record the
    original 10-K accession so we can fall back to it for oil-and-gas HTML.
    """
    filings = submissions.get("filings", {}).get("recent", {})
    forms   = filings.get("form", [])
    accs    = filings.get("accessionNumber", [])
    dates   = filings.get("filingDate", [])
    periods = filings.get("reportDate", dates)   # fiscal year-end date

    # Group by fiscal year-end (reportDate), keeping only 10-K / 10-K/A
    by_year: dict[str, dict] = {}
    for form, acc, date, period in zip(forms, accs, dates, periods):
        if form not in ("10-K", "10-K/A"):
            continue
        fiscal_yr = period[:4] if period else date[:4]
        if fiscal_yr not in by_year:
            by_year[fiscal_yr] = {"accession": acc, "date": date,
                                   "original_accession": None, "fiscal_year": fiscal_yr}
        # Track the first plain 10-K so we can fall back for HTML scraping
        if form == "10-K" and by_year[fiscal_yr]["original_accession"] is None:
            by_year[fiscal_yr]["original_accession"] = acc

    # Sort descending by fiscal year, take most recent n
    years = sorted(by_year.keys(), reverse=True)[:n]
    result = [by_year[y] for y in years]

    # Fill original_accession with accession itself if no amendment existed
    for r in result:
        if r["original_accession"] is None:
            r["original_accession"] = r["accession"]

    if not result:
        raise ValueError("No 10-K filings found for this company.")
    return result


# ---------------------------------------------------------------------------
# XBRL helpers
# ---------------------------------------------------------------------------

def get_company_facts(cik: str) -> dict:
    url = f"https://data.sec.gov/api/xbrl/companyfacts/CIK{cik}.json"
    r = requests.get(url, headers=HEADERS, timeout=20)
    if r.status_code == 404:
        return {}
    r.raise_for_status()
    return r.json()


def _annual_entries(facts: dict, concept: str, namespace: str = "us-gaap") -> list[dict]:
    """All annual 10-K/10-K/A entries for a concept, sorted newest-first."""
    try:
        units = facts["facts"][namespace][concept]["units"]
        entries = units.get("USD", units.get("pure", []))
        annual = [e for e in entries
                  if e.get("form") in ("10-K", "10-K/A") and e.get("val") is not None]
        annual.sort(key=lambda e: e["end"], reverse=True)
        return annual
    except (KeyError, TypeError):
        return []


def value_for_year(facts: dict, concepts: list[str], fiscal_year: str,
                   namespace: str = "us-gaap") -> Optional[float]:
    """
    Return the value for one of the given XBRL concepts whose fiscal-year-end
    matches `fiscal_year` (a 4-digit string like "2023").
    Returns None if no concept has data for that exact year.
    """
    for concept in concepts:
        for e in _annual_entries(facts, concept, namespace):
            if e["end"].startswith(fiscal_year):
                return e["val"]
    return None


def _build_series(facts: dict, concepts: list[str]) -> dict[str, Optional[float]]:
    """
    Return {fiscal_year_str: value} covering all years present in XBRL.
    First concept that has data for a year wins.
    """
    series: dict[str, Optional[float]] = {}
    for concept in concepts:
        for e in _annual_entries(facts, concept):
            yr = e["end"][:4]
            if yr not in series:
                series[yr] = e["val"]
    return series


def _build_ebitda_series(facts: dict) -> dict[str, Optional[float]]:
    """
    EBITDA = Operating/pre-tax income + D&A, matched by fiscal year.
    """
    income_concepts = [
        "OperatingIncomeLoss",
        "IncomeLossFromContinuingOperationsBeforeIncomeTaxesExtraordinaryItemsNoncontrollingInterest",
        "IncomeLossFromContinuingOperationsBeforeIncomeTaxesMinorityInterestAndIncomeLossFromEquityMethodInvestments",
    ]
    da_concepts = [
        "DepreciationDepletionAndAmortization",
        "DepreciationAndAmortization",
        "Depreciation",
    ]
    income_series = _build_series(facts, income_concepts)
    da_series     = _build_series(facts, da_concepts)
    common = set(income_series) & set(da_series)
    return {yr: income_series[yr] + da_series[yr] for yr in common}  # type: ignore[operator]


# ---------------------------------------------------------------------------
# Unit-scale detection & normalisation
# ---------------------------------------------------------------------------

def detect_xbrl_scale(facts: dict) -> int:
    """
    Determine the unit multiplier from XBRL data.

    EDGAR XBRL values are always in *base units* (dollars, not thousands/millions).
    So the scale is always 1 — but we keep this function to document that assumption
    and to handle any future edge case.
    """
    return 1  # XBRL values from data.sec.gov are always in whole dollars


def detect_html_scale(html_text: str) -> int:
    """
    Return the multiplier to convert the filing's reported numbers to whole dollars.
    Scans for the disclosure phrase that governs dollar amounts (not shares).
    """
    low = html_text.lower()
    # Match "in millions" referring to dollar values (most large-cap filings)
    if re.search(r"(?:stated|expressed|reported|presented|denominated|"
                 r"amounts? are stated|values?.{0,30})\s*in millions", low):
        return 1_000_000
    # "in thousands" for dollar amounts (exclude "(shares in thousands)" etc.)
    if re.search(r"(?:stated|expressed|reported|presented|denominated|"
                 r"amounts? are stated|values?.{0,30})\s*in thousands", low):
        return 1_000
    # Bare "in thousands" / "in millions" — accept only if share context is absent nearby
    for m in re.finditer(r"\bin thousands\b", low):
        ctx = low[max(0, m.start() - 60): m.start() + 40]
        if "share" not in ctx and "stock" not in ctx and "unit" not in ctx:
            return 1_000
    for m in re.finditer(r"\bin millions\b", low):
        ctx = low[max(0, m.start() - 60): m.start() + 40]
        if "share" not in ctx and "stock" not in ctx and "unit" not in ctx:
            return 1_000_000
    return 1_000_000   # safe default for large-cap filings


def to_thousands(value: Optional[float]) -> Optional[float]:
    """Convert a whole-dollar XBRL value to thousands (our standard display unit)."""
    if value is None:
        return None
    return value / 1_000


# ---------------------------------------------------------------------------
# HTML filing fetch
# ---------------------------------------------------------------------------

def _resolve_doc_url(href: str) -> str:
    if href.startswith("/ix?doc="):
        href = href[len("/ix?doc="):]
    return href if href.startswith("http") else f"https://www.sec.gov{href}"


def fetch_10k_html(cik: str, accession: str) -> str:
    acc_nodash = accession.replace("-", "")
    index_url = (f"https://www.sec.gov/Archives/edgar/data/{int(cik)}"
                 f"/{acc_nodash}/{accession}-index.htm")
    r = requests.get(index_url, headers=HTML_HEADERS, timeout=15)
    soup = BeautifulSoup(r.text, "lxml")

    primary = None
    for row in soup.select("table tr"):
        cells = row.find_all("td")
        if len(cells) >= 4 and "10-K" in cells[3].get_text():
            link = cells[2].find("a")
            if link:
                href = link["href"]
                if "ix?doc=" in href or href.endswith((".htm", ".html")):
                    primary = _resolve_doc_url(href)
                    break
    if not primary:
        for a in soup.select("table a[href]"):
            href = a["href"]
            if "ix?doc=" in href or href.endswith((".htm", ".html")):
                primary = _resolve_doc_url(href)
                break
    if not primary:
        return ""

    time.sleep(0.3)
    r2 = requests.get(primary, headers=HTML_HEADERS, timeout=60)
    return r2.text


# ---------------------------------------------------------------------------
# HTML table scraping helpers (for PV-10 and reserves)
# ---------------------------------------------------------------------------

def _parse_num(s: str) -> Optional[float]:
    s = s.strip().replace(",", "").replace("$", "")
    negative = s.startswith("(") and s.endswith(")")
    s = s.strip("()")
    try:
        val = float(s)
        return -val if negative else val
    except ValueError:
        return None


def scrape_table_value(soup: BeautifulSoup, label_patterns: list[str]) -> Optional[float]:
    """Search all table cells for a label match, return the first numeric sibling."""
    for tag in soup.find_all(["td", "th"]):
        cell_text = tag.get_text(" ", strip=True).lower()
        if not any(re.search(p, cell_text) for p in label_patterns):
            continue
        row = tag.find_parent("tr")
        if not row:
            continue
        cells = list(row.find_all(["td", "th"]))
        try:
            idx = cells.index(tag)
        except ValueError:
            continue
        for td in cells[idx + 1:]:
            v = _parse_num(td.get_text(" ", strip=True))
            if v is not None and v != 0:
                return v
    return None


def scrape_number_from_text(text: str, patterns: list[str], window: int = 400) -> Optional[float]:
    low = text.lower()
    for pat in patterns:
        m = re.search(pat, low)
        if m:
            snippet = low[m.start(): m.start() + window]
            num_m = re.search(r"\(?([\d,]{3,}(?:\.\d+)?)\)?", snippet[len(m.group(0)):])
            if num_m:
                try:
                    val = float(num_m.group(1).replace(",", ""))
                    return -val if num_m.group(0).startswith("(") else val
                except ValueError:
                    continue
    return None


# ---------------------------------------------------------------------------
# Oil & gas HTML extraction (PV-10 + proved reserves)
# ---------------------------------------------------------------------------

def _reserves_year_end(plain_text: str, year: int) -> Optional[float]:
    """
    Find total proved reserves for `year` from the rollforward table.
    The last substantive number after 'December 31, {year}' is the BOE total column.
    """
    lower = plain_text.lower()
    start = lower.find("proved developed and undeveloped reserves")
    if start < 0:
        return None
    section = plain_text[start: start + 6_000]
    m = re.search(rf"december 31, {year}", section.lower())
    if not m:
        return None
    row_text = section[m.start(): m.start() + 300]
    nums = []
    for n in re.findall(r"[\d,]+(?:\.\d+)?", row_text):
        try:
            v = float(n.replace(",", ""))
            if v != year and v > 10:
                nums.append(v)
        except ValueError:
            pass
    return nums[-1] if nums else None


def scrape_pv10_and_reserves(html_text: str) -> tuple[Optional[float], Optional[float]]:
    """
    Return (pv10_raw, proved_reserves_raw) in the filing's own reported units.
    The caller is responsible for applying the HTML unit scale.
    """
    soup = BeautifulSoup(html_text, "html.parser")
    plain = soup.get_text(" ")

    pv10 = scrape_table_value(soup, [
        r"standardized measure of discounted future net cash flows",
        r"standardized measure of discounted",
        r"pv[\s\-]?10",
    ])
    if pv10 is None:
        pv10 = scrape_number_from_text(plain, [
            r"standardized measure of discounted future net cash flows",
            r"pv[\s\-]?10",
        ])

    current_year = datetime.now().year
    reserves = None
    for yr in (current_year - 1, current_year - 2):
        reserves = _reserves_year_end(plain, yr)
        if reserves is not None:
            break
    if reserves is None:
        reserves = scrape_table_value(soup, [r"total proved reserves", r"^total proved$"])
    if reserves is None:
        reserves = scrape_number_from_text(plain, [r"total proved reserves"])

    return pv10, reserves


# ---------------------------------------------------------------------------
# Per-year data extraction
# ---------------------------------------------------------------------------

def extract_year(
    facts: dict,
    fiscal_year: str,
    ebitda_series: dict[str, Optional[float]],
    is_oil_gas: bool,
    cik: str,
    accession: str,
    original_accession: str,
    scrape_html: bool = False,   # caller controls whether to fetch HTML for this year
) -> dict[str, Optional[float]]:
    """
    Pull all metrics for one fiscal year. All dollar values returned in $thousands.
    proved_reserves is returned as a raw float in the filing's own units (MMBoe/Bcfe).
    """

    def v(concepts: list[str]) -> Optional[float]:
        raw = value_for_year(facts, concepts, fiscal_year)
        return to_thousands(raw)

    revenue = v([
        "RevenueFromContractWithCustomerExcludingAssessedTax",
        "Revenues",
        "SalesRevenueNet",
        "RevenueFromContractWithCustomerIncludingAssessedTax",
        "SalesRevenueGoodsNet",
    ])
    net_income = v([
        "NetIncomeLoss",
        "NetIncomeLossAvailableToCommonStockholdersBasic",
        "ProfitLoss",
    ])
    total_assets = v(["Assets"])
    cash = v([
        "CashAndCashEquivalentsAtCarryingValue",
        "CashCashEquivalentsAndShortTermInvestments",
        "Cash",
    ])

    # Debt: try an explicit total first, then build from components
    total_debt_raw = value_for_year(facts, [
        "LongTermDebtAndCapitalLeaseObligations",   # XOM and others use this as full total
        "DebtAndCapitalLeaseObligations",
        "LongTermDebt",                             # some filers include current portion here
    ], fiscal_year)

    if total_debt_raw is not None:
        # May already include current; check if we should add separate current portion
        current_raw = value_for_year(facts, ["DebtCurrent", "LongTermDebtCurrent",
                                             "LongTermDebtAndCapitalLeaseObligationsCurrent"],
                                     fiscal_year)
        # Only add current if the total concept name suggests it's non-current only
        # Use heuristic: if current > 10% of total, it's likely already included
        if current_raw is not None and current_raw < total_debt_raw * 0.5:
            # Assume current not included; most single-concept filers include both
            # but LongTermDebt (noncurrent) filers need current added
            noncurrent_only = value_for_year(facts, ["LongTermDebtNoncurrent"], fiscal_year)
            if noncurrent_only is not None:
                total_debt_raw = noncurrent_only + current_raw
        total_debt = to_thousands(total_debt_raw)
    else:
        # Build from noncurrent + current components
        lt_raw = value_for_year(facts, ["LongTermDebtNoncurrent"], fiscal_year)
        st_raw = value_for_year(facts, ["ShortTermBorrowings", "DebtCurrent"], fiscal_year)
        if lt_raw is not None and st_raw is not None:
            total_debt = to_thousands(lt_raw + st_raw)
        elif lt_raw is not None:
            total_debt = to_thousands(lt_raw)
        elif st_raw is not None:
            total_debt = to_thousands(st_raw)
        else:
            total_debt = None

    # Equity: direct concept, or Assets - Liabilities
    equity = v([
        "StockholdersEquity",
        "StockholdersEquityIncludingPortionAttributableToNoncontrollingInterest",
    ])
    if equity is None and total_assets is not None:
        liabilities_raw = value_for_year(facts, ["Liabilities"], fiscal_year)
        if liabilities_raw is not None:
            # total_assets and liabilities_raw are both in whole dollars here
            assets_raw = value_for_year(facts, ["Assets"], fiscal_year)
            if assets_raw is not None:
                equity = to_thousands(assets_raw - liabilities_raw)

    # EBITDA from pre-built year-matched series
    ebitda_raw = ebitda_series.get(fiscal_year)
    ebitda = to_thousands(ebitda_raw)

    result: dict[str, Optional[float]] = {
        "revenue":         revenue,
        "ebitda":          ebitda,
        "net_income":      net_income,
        "total_assets":    total_assets,
        "total_debt":      total_debt,
        "cash":            cash,
        "equity":          equity,
        "pv10":            None,
        "proved_reserves": None,
    }

    if is_oil_gas:
        # PV-10: try XBRL concepts in order of specificity
        pv10_xbrl_concepts = [
            "StandardizedMeasureOfDiscountedFutureNetCashFlows",
            "StandardizedMeasureOfDiscountedFutureNetCashFlowsRelatingToProvedOilAndGasReserves",
            "DiscountedFutureNetCashFlowsRelatingToProvedOilAndGasReservesStandardizedMeasure",
            "FutureNetCashFlowsRelatingToProvedOilAndGasReservesTenPercentAnnualDiscountForEstimatedTimingOfCashFlows",
            "DiscountedFutureNetCashFlowsRelatingToProvedOilAndGasReserves10PercentAnnualDiscountForEstimatedTimingOfCashFlows",
        ]
        pv10_raw = value_for_year(facts, pv10_xbrl_concepts, fiscal_year)

        # Proved reserves: XBRL (energy units, usually MMBoe or Bcfe)
        reserves_raw = value_for_year(facts, [
            "ProvedDevelopedAndUndevelopedReserveNetEnergy",
            "ProvedDevelopedReservesBOE1",
        ], fiscal_year)

        # HTML scrape only when explicitly requested (most recent year, caller decides)
        if scrape_html and (pv10_raw is None or reserves_raw is None):
            accs = [accession]
            if original_accession != accession:
                accs.append(original_accession)
            for acc in accs:
                html_text = fetch_10k_html(cik, acc)
                if not html_text:
                    continue
                scale = detect_html_scale(html_text)
                pv10_html, res_html = scrape_pv10_and_reserves(html_text)
                if pv10_raw is None and pv10_html is not None:
                    pv10_raw = pv10_html * scale   # convert to whole dollars
                if reserves_raw is None:
                    reserves_raw = res_html
                # Stop if we found what we needed; otherwise try next accession
                if pv10_raw is not None and reserves_raw is not None:
                    break

        result["pv10"]            = to_thousands(pv10_raw)
        result["proved_reserves"] = reserves_raw    # NOT in $thousands — raw MMBoe/Bcfe

    return result


# ---------------------------------------------------------------------------
# Formatting helpers
# ---------------------------------------------------------------------------

def fmt_thousands(val: Optional[float], is_reserves: bool = False) -> str:
    """Format a value stored in $thousands for display."""
    if val is None:
        return "N/A"
    if is_reserves:
        return f"{val:,.1f}"
    # Display in $M (divide thousands by 1,000)
    m = val / 1_000
    return f"${m:,.1f}M"


def fmt_excel(val: Optional[float]) -> object:
    """Raw value for Excel (already in $thousands, or raw for reserves)."""
    return val if val is not None else ""


# ---------------------------------------------------------------------------
# Console output
# ---------------------------------------------------------------------------

def print_company_table(
    ticker: str,
    company_name: str,
    is_oil_gas: bool,
    years_data: list[tuple[str, dict]],   # [(fiscal_year, metrics), ...]
) -> None:
    year_labels = [f"FY{fy}" for fy, _ in years_data]
    metrics = [k for k in METRIC_KEYS if is_oil_gas or k not in ("pv10", "proved_reserves")]

    header = ["Metric (USD Millions)"] + year_labels
    rows = []
    for key in metrics:
        label = METRIC_LABELS[key]
        is_res = (key == "proved_reserves")
        row = [label] + [fmt_thousands(data.get(key), is_reserves=is_res)
                         for _, data in years_data]
        rows.append(row)

    width = 57 + 14 * len(years_data)
    print(f"\n{'=' * width}")
    print(f"  {ticker.upper()} — {company_name}")
    print(f"{'=' * width}")
    print(tabulate(rows, headers=header, tablefmt="rounded_outline"))
    print("\n  * EBITDA = Operating Income + D&A (same fiscal year)")
    if is_oil_gas:
        print("  * PV-10 = Standardized Measure of Discounted Future Net Cash Flows")
        print("  * Proved Reserves in filing units (MMBoe or Bcfe)")
    print()


# ---------------------------------------------------------------------------
# Excel export
# ---------------------------------------------------------------------------

def export_excel(
    all_results: list[tuple[str, str, bool, list[tuple[str, dict]]]],
    path: str,
) -> None:
    """
    Write one sheet per ticker. Each sheet has metrics as rows, fiscal years as columns.
    All dollar values are in $thousands; the sheet header row says so.
    """
    try:
        import openpyxl
        from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
        from openpyxl.utils import get_column_letter
    except ImportError:
        print("openpyxl not installed — run: pip3 install openpyxl", file=sys.stderr)
        return

    wb = openpyxl.Workbook()
    wb.remove(wb.active)  # remove default empty sheet

    # Colour palette
    HEADER_BG   = "1F3864"   # dark navy
    HEADER_FG   = "FFFFFF"
    SUBHDR_BG   = "2E75B6"   # medium blue
    SUBHDR_FG   = "FFFFFF"
    ALT_BG      = "EAF0FB"   # light blue alternating row
    OG_LABEL_BG = "FFF2CC"   # light yellow for oil/gas-only rows
    BORDER_CLR  = "BDD7EE"

    thin = Side(style="thin", color=BORDER_CLR)
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    def style_cell(cell, bold=False, bg=None, fg="000000", align="left",
                   num_fmt=None, wrap=False):
        cell.font      = Font(bold=bold, color=fg, name="Calibri", size=10)
        cell.alignment = Alignment(horizontal=align, vertical="center", wrap_text=wrap)
        cell.border    = border
        if bg:
            cell.fill  = PatternFill("solid", fgColor=bg)
        if num_fmt:
            cell.number_format = num_fmt

    for ticker, company_name, is_oil_gas, years_data in all_results:
        ws = wb.create_sheet(title=ticker.upper()[:31])
        ws.sheet_view.showGridLines = False

        fiscal_years = [fy for fy, _ in years_data]
        metrics = [k for k in METRIC_KEYS
                   if is_oil_gas or k not in ("pv10", "proved_reserves")]

        # --- Row 1: company banner ---
        ws.merge_cells(start_row=1, start_column=1,
                       end_row=1, end_column=1 + len(fiscal_years))
        banner = ws.cell(1, 1, f"{ticker.upper()} — {company_name}")
        style_cell(banner, bold=True, bg=HEADER_BG, fg=HEADER_FG, align="center")
        ws.row_dimensions[1].height = 20

        # --- Row 2: column headers (Metric | FY20XX ...) ---
        ws.cell(2, 1, "Metric")
        style_cell(ws.cell(2, 1), bold=True, bg=SUBHDR_BG, fg=SUBHDR_FG, align="center")
        for col_i, fy in enumerate(fiscal_years, start=2):
            c = ws.cell(2, col_i, f"FY{fy}")
            style_cell(c, bold=True, bg=SUBHDR_BG, fg=SUBHDR_FG, align="center")
        ws.row_dimensions[2].height = 16

        # --- Row 3: unit note ---
        ws.merge_cells(start_row=3, start_column=1,
                       end_row=3, end_column=1 + len(fiscal_years))
        unit_note = ws.cell(3, 1, "All dollar values in USD thousands ($000s), except Proved Reserves (MMBoe/Bcfe)")
        style_cell(unit_note, bg="F2F2F2", fg="595959", align="center", wrap=True)
        ws.row_dimensions[3].height = 14

        # --- Data rows ---
        for row_i, key in enumerate(metrics, start=4):
            is_res     = (key == "proved_reserves")
            is_og_only = key in ("pv10", "proved_reserves")
            label_bg   = OG_LABEL_BG if is_og_only else (ALT_BG if row_i % 2 == 0 else None)

            label_cell = ws.cell(row_i, 1, METRIC_LABELS[key])
            style_cell(label_cell, bold=True, bg=label_bg, align="left")

            for col_i, (fy, data) in enumerate(years_data, start=2):
                raw = data.get(key)
                cell = ws.cell(row_i, col_i)

                if raw is None:
                    cell.value = None
                elif is_res:
                    cell.value = raw           # MMBoe, no dollar format
                    cell.number_format = '#,##0.0'
                else:
                    cell.value = raw           # already in $thousands
                    cell.number_format = '#,##0'

                style_cell(cell, bg=label_bg, align="right")

            ws.row_dimensions[row_i].height = 15

        # --- Column widths ---
        ws.column_dimensions[get_column_letter(1)].width = 26
        for col_i in range(2, 2 + len(fiscal_years)):
            ws.column_dimensions[get_column_letter(col_i)].width = 18

        # --- Notes section ---
        note_row = 4 + len(metrics) + 1
        ws.merge_cells(start_row=note_row, start_column=1,
                       end_row=note_row, end_column=1 + len(fiscal_years))
        notes = [
            "Notes:",
            "• EBITDA estimated as Operating Income + D&A, matched within the same fiscal year.",
            "• Source: SEC EDGAR XBRL data (data.sec.gov); oil & gas supplemental via HTML parsing.",
        ]
        if is_oil_gas:
            notes.append("• PV-10 = Standardized Measure of Discounted Future Net Cash Flows.")
            notes.append("• Proved Reserves in filing units (MMBoe or Bcfe — verify per filing).")
        note_cell = ws.cell(note_row, 1, "  ".join(notes))
        style_cell(note_cell, fg="595959", bg="F9F9F9", wrap=True)
        ws.row_dimensions[note_row].height = 60

    wb.save(path)
    print(f"Excel file saved: {path}")


# ---------------------------------------------------------------------------
# Main per-ticker extraction
# ---------------------------------------------------------------------------

def extract_ticker(ticker: str, num_years: int) -> tuple[str, str, bool, list[tuple[str, dict]]]:
    """
    Fetch and extract data for one ticker.
    Returns (ticker, company_name, is_oil_gas, [(fiscal_year, metrics), ...]).
    """
    print(f"\n{'─'*55}")
    print(f"  {ticker.upper()}")
    print(f"{'─'*55}")

    cik, company_name = get_cik(ticker)
    print(f"  Company : {company_name}")
    print(f"  CIK     : {int(cik)}")

    info    = get_company_info(cik)
    sic     = str(info.get("sic", ""))
    is_oil_gas = sic in OIL_GAS_SIC
    if is_oil_gas:
        print(f"  SIC {sic} — Oil & Gas company detected.")

    filings = get_10k_filings(info, n=num_years)
    print(f"  Found {len(filings)} filing(s): " +
          ", ".join(f"FY{f['fiscal_year']} ({f['date']})" for f in filings))

    print("  Fetching XBRL facts...")
    facts = get_company_facts(cik)
    ebitda_series = _build_ebitda_series(facts)

    years_data: list[tuple[str, dict]] = []
    for i, filing in enumerate(filings):
        fy   = filing["fiscal_year"]
        acc  = filing["accession"]
        orig = filing["original_accession"]
        # HTML scrape only for the most recent year — it's slow and historical
        # XBRL data covers most prior years for major filers
        scrape_html = is_oil_gas and (i == 0)
        if scrape_html:
            print(f"  Fetching 10-K HTML for FY{fy} (PV-10 / reserves)...")

        metrics = extract_year(
            facts             = facts,
            fiscal_year       = fy,
            ebitda_series     = ebitda_series,
            is_oil_gas        = is_oil_gas,
            cik               = cik,
            accession         = acc,
            original_accession = orig,
            scrape_html       = scrape_html,
        )
        years_data.append((fy, metrics))

    return ticker, company_name, is_oil_gas, years_data


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------

def main() -> None:
    parser = argparse.ArgumentParser(
        description="Extract key financials from SEC EDGAR 10-K filings.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=(
            "Examples:\n"
            "  python3 sec_10k_extractor.py AAPL\n"
            "  python3 sec_10k_extractor.py AAPL,MSFT,GOOG --years 3\n"
            "  python3 sec_10k_extractor.py XOM,CVX --export oilgas.xlsx\n"
        )
    )
    parser.add_argument(
        "tickers",
        help="One or more ticker symbols, comma-separated (e.g. AAPL,MSFT,XOM)",
    )
    parser.add_argument(
        "--years", "-y",
        type=int, default=3,
        help="Number of fiscal years to extract (default: 3)",
    )
    parser.add_argument(
        "--export", "-e",
        metavar="FILE.xlsx",
        help="Export results to an Excel file",
    )
    args = parser.parse_args()

    tickers = [t.strip().upper() for t in args.tickers.split(",") if t.strip()]
    all_results = []
    errors = []

    for ticker in tickers:
        try:
            result = extract_ticker(ticker, args.years)
            all_results.append(result)
            ticker, company_name, is_oil_gas, years_data = result
            print_company_table(ticker, company_name, is_oil_gas, years_data)
        except Exception as exc:
            errors.append((ticker, str(exc)))
            print(f"\n  [ERROR] {ticker}: {exc}", file=sys.stderr)

    if args.export and all_results:
        export_excel(all_results, args.export)

    if errors:
        print("\nFailed tickers:", ", ".join(t for t, _ in errors), file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()
