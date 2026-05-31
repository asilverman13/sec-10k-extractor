# SEC 10-K Financial Data Extractor

Automatically extracts key financial data from SEC EDGAR 10-K filings for any public company.

## Features
- Pulls data directly from SEC EDGAR — no manual input required
- Supports any public company by ticker symbol
- Extracts 3 years of historical data
- Detects and standardizes units (thousands vs. millions) automatically
- For oil & gas companies: extracts PV-10 and proved reserves
- Exports results to Excel with professional formatting

## Usage
```bash
# Single ticker
python3 sec_10k_extractor.py AAPL

# Multiple tickers
python3 sec_10k_extractor.py AAPL,MSFT,GOOG

# Export to Excel with 3 years of history
python3 sec_10k_extractor.py AAPL,MSFT,XOM --years 3 --export results.xlsx
```

## Fields Extracted
- Revenue, Net Income, EBITDA
- Total Assets, Total Debt, Cash, Equity
- PV-10 and Proved Reserves (oil & gas companies only)

## Built With
- Python
- SEC EDGAR XBRL API
- pandas, openpyxl
