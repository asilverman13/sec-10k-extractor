"""
SEC EDGAR 10-K Financial Data Extractor — Streamlit UI
Wraps sec_10k_extractor.py and exposes a browser-based interface.
"""

from __future__ import annotations

import io
import traceback
import tempfile
import os

import pandas as pd
import streamlit as st

from sec_10k_extractor import (
    extract_ticker,
    export_excel,
    METRIC_KEYS,
    METRIC_LABELS,
    TRUST_METRIC_LABELS,
    fmt_thousands,
)

# ---------------------------------------------------------------------------
# Page config
# ---------------------------------------------------------------------------

st.set_page_config(
    page_title="SEC 10-K Extractor",
    page_icon="📊",
    layout="wide",
)

st.title("📊 SEC EDGAR 10-K Financial Extractor")
st.caption(
    "Pulls revenue, EBITDA, net income, assets, debt, cash, equity — "
    "plus PV-10 and proved reserves for oil & gas companies — directly from SEC filings."
)

# ---------------------------------------------------------------------------
# Sidebar inputs
# ---------------------------------------------------------------------------

with st.sidebar:
    st.header("Settings")
    tickers_raw = st.text_input(
        "Ticker(s)",
        placeholder="e.g.  AAPL, XOM, SBR",
        help="Enter one or more ticker symbols separated by commas.",
    )
    num_years = st.slider("Years of history", min_value=1, max_value=5, value=3)
    run_btn = st.button("🔍 Extract Data", type="primary", use_container_width=True)
    st.divider()
    st.markdown(
        "**Data source:** [SEC EDGAR](https://www.sec.gov/)\n\n"
        "Values in **USD millions** on screen; "
        "Excel export stores raw **USD thousands** for precision."
    )

# ---------------------------------------------------------------------------
# Helper: build a DataFrame for one ticker result
# ---------------------------------------------------------------------------

def result_to_df(
    company_name: str,
    is_oil_gas: bool,
    years_data: list[tuple[str, dict]],
    is_trust: bool,
) -> pd.DataFrame:
    labels = TRUST_METRIC_LABELS if is_trust else METRIC_LABELS
    show_og = is_oil_gas or is_trust
    metrics = [
        k for k in METRIC_KEYS
        if show_og or k not in ("pv10", "proved_reserves")
    ]
    if is_trust:
        metrics = [k for k in metrics if k != "ebitda"]

    rows = []
    for key in metrics:
        label = labels[key]
        is_res = key == "proved_reserves"
        row = {"Metric": label}
        for fy, data in years_data:
            val = data.get(key)
            if val is None:
                row[f"FY{fy}"] = "N/A"
            elif is_res:
                row[f"FY{fy}"] = f"{val:,.1f}"
            else:
                m = val / 1_000
                row[f"FY{fy}"] = f"${m:,.1f}M"
        rows.append(row)

    return pd.DataFrame(rows)


# ---------------------------------------------------------------------------
# Main — run extraction when button is pressed
# ---------------------------------------------------------------------------

if run_btn:
    tickers = [t.strip().upper() for t in tickers_raw.split(",") if t.strip()]
    if not tickers:
        st.warning("Please enter at least one ticker symbol.")
        st.stop()

    all_results = []
    errors = []

    progress = st.progress(0, text="Starting…")

    for idx, ticker in enumerate(tickers):
        progress.progress(
            (idx) / len(tickers),
            text=f"Fetching {ticker} ({idx + 1} of {len(tickers)})…",
        )
        try:
            result = extract_ticker(ticker, num_years)
            all_results.append(result)
        except Exception as exc:
            errors.append((ticker, str(exc), traceback.format_exc()))

    progress.progress(1.0, text="Done.")

    # -----------------------------------------------------------------------
    # Display results
    # -----------------------------------------------------------------------

    for ticker, company_name, is_oil_gas, years_data, is_trust in all_results:
        trust_tag = " · Royalty Trust" if is_trust else ""
        og_tag    = " · Oil & Gas" if is_oil_gas and not is_trust else ""
        st.subheader(f"{ticker}  —  {company_name}{trust_tag}{og_tag}")

        df = result_to_df(company_name, is_oil_gas, years_data, is_trust)
        st.dataframe(df.set_index("Metric"), use_container_width=True)

        if is_trust:
            st.caption(
                "Modified cash-basis filer — EBITDA not applicable; "
                "Equity = Trust Corpus end-of-year. Data scraped from 10-K HTML."
            )
        if is_oil_gas or is_trust:
            st.caption(
                "PV-10 = Standardized Measure of Discounted Future Net Cash Flows. "
                "Proved Reserves in filing units (MMBoe or Bcfe)."
            )
        st.divider()

    # -----------------------------------------------------------------------
    # Error panel
    # -----------------------------------------------------------------------

    if errors:
        with st.expander(f"⚠️ {len(errors)} ticker(s) failed", expanded=True):
            for ticker, msg, tb in errors:
                st.error(f"**{ticker}**: {msg}")
                with st.expander("Stack trace"):
                    st.code(tb)

    # -----------------------------------------------------------------------
    # Excel download
    # -----------------------------------------------------------------------

    if all_results:
        with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
            tmp_path = tmp.name

        try:
            export_excel(all_results, tmp_path)
            with open(tmp_path, "rb") as f:
                excel_bytes = f.read()

            tickers_slug = "_".join(t for t, *_ in all_results)
            st.download_button(
                label="⬇️ Download Excel",
                data=excel_bytes,
                file_name=f"{tickers_slug}_10K.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
            )
        finally:
            os.unlink(tmp_path)

# ---------------------------------------------------------------------------
# Idle state — show instructions
# ---------------------------------------------------------------------------

else:
    st.info(
        "Enter one or more ticker symbols in the sidebar (e.g. **AAPL, XOM, SBR**) "
        "and click **Extract Data** to pull the latest 10-K financials from SEC EDGAR."
    )

    with st.expander("What data does this tool pull?"):
        st.markdown("""
| Metric | Source |
|--------|--------|
| Revenue | XBRL / 10-K HTML |
| EBITDA (est.) | Operating Income + D&A, year-matched |
| Net Income | XBRL |
| Total Assets | XBRL |
| Total Debt | XBRL (noncurrent + current) |
| Cash | XBRL |
| Equity | XBRL (or Assets − Liabilities) |
| **PV-10** *(oil & gas only)* | XBRL or HTML scrape |
| **Proved Reserves** *(oil & gas only)* | XBRL or HTML scrape |

**Royalty trusts** (e.g. SBR, MTR, BPT) use a dedicated HTML scraper because they
file under modified cash basis with no structured XBRL data.
        """)

    with st.expander("Supported tickers / known royalty trusts"):
        st.markdown("""
The tool works for **any publicly traded company** with SEC EDGAR filings.
Royalty trusts that have been specifically tested and verified:

`MTR` · `SBR` · `TIRTZ` · `MARPS` · `NRT` · `BPT` · `DMLP` · `SDT` · `CHKR`
        """)
