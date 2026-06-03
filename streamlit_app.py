"""
SEC EDGAR Financial Extractor — Streamlit UI
"""

import io
import warnings
warnings.filterwarnings("ignore")

import streamlit as st
import openpyxl

from sec_edgar_extractor import (
    get_cik, get_company_name, get_facts,
    fetch_row, compute_ebitdax,
    INCOME_CONCEPTS, BALANCE_CONCEPTS,
    write_income_sheet, write_balance_sheet,
)

# ── Page config ────────────────────────────────────────────────────────────────

st.set_page_config(
    page_title="SEC EDGAR Extractor",
    page_icon="📊",
    layout="wide",
)

st.title("📊 SEC EDGAR 10-K Financial Extractor")
st.caption("Pulls income statement and balance sheet data via SEC EDGAR XBRL API and exports to Excel.")

# ── Inputs ─────────────────────────────────────────────────────────────────────

col1, col2, col3 = st.columns([2, 1, 4])
with col1:
    ticker = st.text_input("Ticker Symbol", placeholder="e.g. AAPL, XOM, MSFT").strip().upper()
with col2:
    num_years = st.selectbox("Years", [3, 4, 5], index=0)
with col3:
    st.write("")
    st.write("")
    run = st.button("Run Extraction", type="primary", disabled=not ticker)

# ── Helper: build preview table ────────────────────────────────────────────────

def build_preview(data: dict, years: list, base_key: str) -> list[dict]:
    base = {y: data.get(base_key, {}).get(y) for y in years}
    rows = []
    for label, val_dict in data.items():
        row = {"Line Item": label}
        for y in years:
            v = val_dict.get(y)
            b = base.get(y)
            row[str(y)] = f"{v / 1_000_000:,.0f}" if v is not None else "—"
            row[f"{y} %"] = (
                f"{v / b * 100:.1f}%" if (v is not None and b) else "—"
            )
        rows.append(row)
    return rows


BOLD_INCOME  = {"Income / Revenue", "Operating Expenses", "EBITDAX", "Net Income"}
BOLD_BALANCE = {
    "Total Current Assets", "Total Assets",
    "Total Current Liabilities", "Total Liabilities",
    "Total Equity", "Total Liabilities and Equity",
}


def show_table(rows: list[dict], bold_set: set, years: list):
    col_order = ["Line Item"] + [c for y in years for c in (str(y), f"{y} %")]

    header_cols = st.columns([3] + [1, 0.7] * len(years))
    header_cols[0].markdown("**Line Item**")
    for i, y in enumerate(years):
        header_cols[1 + i * 2].markdown(f"**{y}**")
        header_cols[2 + i * 2].markdown(f"**%**")

    st.divider()

    for row in rows:
        is_bold = row["Line Item"] in bold_set
        cols = st.columns([3] + [1, 0.7] * len(years))
        label = f"**{row['Line Item']}**" if is_bold else row["Line Item"]
        cols[0].markdown(label)
        for i, y in enumerate(years):
            dollar = f"**{row[str(y)]}**" if is_bold else row[str(y)]
            pct    = f"**{row[f'{y} %']}**" if is_bold else row[f"{y} %"]
            cols[1 + i * 2].markdown(dollar)
            cols[2 + i * 2].markdown(pct)


# ── Main extraction ────────────────────────────────────────────────────────────

if run and ticker:
    try:
        # 1. Resolve ticker → CIK
        with st.spinner(f"Looking up {ticker}…"):
            cik = get_cik(ticker)
            company_name = get_company_name(cik)
        st.success(f"**{company_name}** (CIK {int(cik)})")

        # 2. Fetch XBRL facts
        with st.spinner("Fetching XBRL data from SEC EDGAR…"):
            facts = get_facts(cik)

        # 3. Determine fiscal years
        probe = fetch_row(facts, INCOME_CONCEPTS["Net Income"], list(range(2010, 2026)))
        available = sorted([y for y, v in probe.items() if v is not None], reverse=True)
        if not available:
            st.error("No annual Net Income data found for this ticker.")
            st.stop()
        years = sorted(available[:num_years])
        st.info(f"Fiscal years: {', '.join(str(y) for y in years)}")

        # 4. Fetch income statement
        with st.spinner("Fetching income statement…"):
            income_data: dict = {}
            for label, concepts in INCOME_CONCEPTS.items():
                income_data[label] = fetch_row(facts, concepts, years)
            income_data["EBITDAX"] = compute_ebitdax(income_data, years)

        # 5. Fetch balance sheet
        with st.spinner("Fetching balance sheet…"):
            bs_data: dict = {}
            for label, concepts in BALANCE_CONCEPTS.items():
                bs_data[label] = fetch_row(facts, concepts, years)

        # 6. Build Excel in memory
        wb = openpyxl.Workbook()
        write_income_sheet(wb, company_name, years, income_data)
        write_balance_sheet(wb, company_name, years, bs_data)
        buf = io.BytesIO()
        wb.save(buf)
        buf.seek(0)

        # ── Download button ────────────────────────────────────────────────────
        st.download_button(
            label="⬇️  Download Excel File",
            data=buf,
            file_name=f"{ticker}_financials.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
        )

        st.divider()

        # ── Income Statement preview ───────────────────────────────────────────
        st.subheader("Income Statement")
        st.caption("Values in thousands (USD)")

        income_display = {
            "Income / Revenue":                        income_data.get("Revenue", {}),
            "Operating Expenses":                      income_data.get("Total Expenses", {}),
            "EBITDAX":                                 income_data.get("EBITDAX", {}),
            "Depreciation / Depletion / Amortization": income_data.get("Depreciation / Depletion / Amortization", {}),
            "Other Income / Expense":                  income_data.get("Other Income / Expense", {}),
            "Net Income":                              income_data.get("Net Income", {}),
            "Distributions":                           income_data.get("Distributions", {}),
        }

        show_table(build_preview(income_display, years, "Revenue"), BOLD_INCOME, years)

        st.divider()

        # ── Balance Sheet preview ──────────────────────────────────────────────
        st.subheader("Balance Sheet")
        st.caption("Values in thousands (USD)")

        bs_display = {
            "Cash and Cash Equivalents":                      bs_data.get("Cash and Cash Equivalents", {}),
            "Other Current Assets":                           bs_data.get("Other Current Assets", {}),
            "Total Current Assets":                           bs_data.get("Total Current Assets", {}),
            "Fixed Assets, Net":                              bs_data.get("Fixed Assets, Net", {}),
            "Other Non-Current Assets":                       bs_data.get("Other Non-Current Assets", {}),
            "Total Assets":                                   bs_data.get("Total Assets", {}),
            "Notes Payable":                                  bs_data.get("Notes Payable", {}),
            "Accounts Payable and Other Current Liabilities": bs_data.get("Accounts Payable and Other Current Liabilities", {}),
            "Total Current Liabilities":                      bs_data.get("Total Current Liabilities", {}),
            "Long-Term Debt":                                 bs_data.get("Long-Term Debt", {}),
            "Other Liabilities":                              bs_data.get("Other Liabilities", {}),
            "Total Liabilities":                              bs_data.get("Total Liabilities", {}),
            "Total Equity":                                   bs_data.get("Total Equity", {}),
            "Total Liabilities and Equity":                   bs_data.get("Total Liabilities and Equity", {}),
        }

        show_table(build_preview(bs_display, years, "Total Assets"), BOLD_BALANCE, years)

        st.caption("Source: SEC EDGAR XBRL API")

    except ValueError as e:
        st.error(str(e))
    except Exception as e:
        st.error(f"Unexpected error: {e}")
