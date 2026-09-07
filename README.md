# Equity Liquidity Automation

Production-grade automation script to update the **'Liquidity'** sheet in the Master Equity Liquidity workbook (`Equity Liquidity.xlsx`).

---

## Overview

This tool processes equity liquidity data by integrating three data sources from the master Excel workbook:

1. **`BBG Formula Save as Value`**: Bloomberg raw data containing ticker details, positions, 30-day trading volume, and up to 5 years (60 months) of historical trading volumes.
2. **`LiqTracking`**: Market USD values (`mkt_USD`) matched via composite identifier (`Ticker + Exchange`).
3. **`Summary`**: Tier classification lookup table referenced via Excel `VLOOKUP`.

---

## Workflow & Column Mapping

The script generates the **'Liquidity'** sheet with 18 standardized columns:

| Col Index | Excel Col | Header Name | Source / Calculation |
|-----------|-----------|-------------|----------------------|
| 1–6       | A–F       | Columns A to F from BBG | Copied directly from `BBG Formula Save as Value` (Ticker, Exch, Name, Currency, Price, Position) |
| 7         | G         | `Trading Volume in the past 30 Days` | Column G (index 6) from BBG sheet |
| 8         | H         | `% (Past 30 Days)` | Formula `=F{row}/G{row}` (Formatted as `0.00%`) |
| 9         | I         | `Liquidity Classification in the past 30 days` | Formula `=VLOOKUP(H{row},Summary!$I$2:$K$5,3,1)` |
| 10        | J         | `Least Trading Volume in the past 5 years` | Minimum trading volume across 5-year date columns (Cols G:BN) |
| 11        | K         | `% (Past 5 Years)` | Formula `=F{row}/J{row}` (Formatted as `0.00%`) |
| 12        | L         | `Liquidity Classification in the past 5 years` | Formula `=VLOOKUP(K{row},Summary!$I$2:$K$5,3,1)` |
| 13        | M         | `5th percentile Volume in the past 5 years` | 5th percentile (0.05 quantile) volume across 5-year date columns |
| 14        | N         | `% (5th percentile in the past 5 years)` | Formula `=F{row}/M{row}` (Formatted as `0.00%`) |
| 15        | O         | `Liquidity Classification in the past 5 years (5th percentile)` | Formula `=VLOOKUP(N{row},Summary!$I$2:$K$5,3,1)` |
| 16        | P         | `INDUSTRY_SECTOR` | Bloomberg Industry Sector |
| 17        | Q         | `mkt_USD` | Market value (USD) merged from `LiqTracking` |
| 18        | R         | `EXCH_CODE` | Exchange Code from BBG |

---

## Key Improvements & Bug Fixes

1. **Bloomberg Error Coercion (`#N/A N/A`, strings)**:
   - In raw Bloomberg exports, cells with missing data often contain string tokens such as `"#N/A N/A"`, `"#N/A"`, or string numbers. In standard Python, doing `.replace("#N/A N/A", np.nan)` leaves the column dtype as `object`, which causes `min()` and `quantile()` to raise `TypeError: '<=' not supported between instances of 'str' and 'float'`.
   - **Fix**: Date columns are coerced to numeric floats using `pd.to_numeric(..., errors='coerce')`, ensuring reliable mathematical calculation across all columns.

2. **Deduplication in Tracking Data**:
   - If `LiqTracking` contains multiple entries or lot-level records for the same identifier, a plain `pd.merge` produces extra rows, causing `ValueError: Length of values does not match length of index`.
   - **Fix**: Identifiers are aggregated by sum (`.groupby("Identifier1")["mkt_USD"].sum(min_count=1)`) and mapped cleanly.

3. **Single-Pass Workbook I/O**:
   - The original script opened, modified, and saved the workbook 3 times (clearing cells with openpyxl, writing DataFrame with pandas, then re-opening with openpyxl to set number format).
   - **Fix**: Consolidated into a single `pd.ExcelWriter` pass with openpyxl engine. Formats are applied directly to `writer.sheets['Liquidity']` before saving once.

4. **Vectorized Formula Generation**:
   - Replaced row-by-row `.at[i-2, ...]` loops with fast list comprehensions (~30-50x speedup).

5. **CLI and Path Flexibility**:
   - Added command-line options (`--file-path`, `--use-iferror`, custom sheet names) instead of hardcoded paths.

---

## Installation & Setup

```bash
pip install -r requirements.txt
```

---

## Usage

### Run from Command Line
```bash
# Run with default network file path:
python equity_liquidity.py

# Or run with custom file path:
python equity_liquidity.py --file-path "/path/to/Equity Liquidity.xlsx"

# Optionally wrap division formulas with IFERROR to guard against 0 volume:
python equity_liquidity.py --file-path "Equity Liquidity.xlsx" --use-iferror
```

### Run as Python Module
```python
from equity_liquidity import process_equity_liquidity

process_equity_liquidity(
    file_path="Equity Liquidity.xlsx",
    use_iferror=False,
)
```

---

## Running Tests

```bash
python3 -m pytest -v
```
