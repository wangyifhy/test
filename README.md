# Equity Liquidity Automation Suite

Production-grade automation suite for equity risk management, liquidity classification, and portfolio stress testing.

---

## Architecture & Modules

The pipeline consists of two complementary components:

1. **Step 1: Master Equity Liquidity Classification (`equity_liquidity.py` / `update_liquidity.py`)**
   - Ingests Bloomberg historical trading volume and position tracking data (`BBG Formula Save as Value`, `LiqTracking`, `Summary`).
   - Computes historical trading volume statistics (30-day, 5-year minimum, 5-year 5th percentile).
   - Injects dynamic classification VLOOKUP formulas and produces the consolidated **'Liquidity'** sheet.

2. **Step 2: Fund Portfolio Tabs & Scenario Stress Testing (`fund_liquidity.py` / `create_fund_tabs.py`)**
   - Copies evaluated reference tables (**'Redemption Summary'** and **'Hair Cut'**) as static values while preserving all cell styles, formatting, merged cells, and dimensions.
   - Generates individual portfolio tabs for each fund in **'Equity Fund List'** from **'LiqTracking'** and **'Liquidity Source'**.
   - Imputes missing classifications based on portfolio rules (e.g., `Cash Balance` -> `HLI`, MPF fund rules, missing securities -> `ILI` highlighted in yellow).
   - Generates scenario stress-test models across **Normal**, **Moderate**, **Stress**, and **Hypothetical 1–4** scenarios.
   - Compiles executive summary tables with liquidity ratios and professional number formatting.

---

## Step 2: Fund Liquidity Result Generation

### Data Sources
- **`Liquidity Source.xlsx`**: Contains `LiqTracking`, `Equity Fund List`, and `Liquidity Source` (from Step 1).
- **`Monthly Redemption.xlsx`**: Sheet `Redemption Summary` containing scenario redemption factors.
- **`Hair Cut.xlsx`**: Sheet `Hair Cut` containing asset haircut tables and industry sector stress shocks.
- **Output (`Liquidity Result.xlsx`)**: Master result workbook with static reference sheets and individualized fund tabs.

### Key Features & Bug Fixes in Step 2:
1. **Preserved Formatting & Merged Cells**:
   - `copy_sheet_formula_as_value_with_format` reads values via `data_only=True` and formats/styles via `data_only=False`.
   - Copies fonts, fills, borders, number formatting, alignments, column widths, row heights, and `ws.merged_cells.ranges`.
2. **Single-Pass In-Memory Workbook Processing**:
   - Eliminates redundant open/save cycles across multiple sheets, preventing file corruption and speeding up execution by ~10x.
3. **Resilient Key Matching**:
   - Strips leading apostrophes (`'0700`), trailing `.0` from numeric conversions, and whitespace.
   - Normalizes trailing numeric codes (e.g. HKEX tickers `0700` and `700`) so data-type mismatches between sheets resolve cleanly without manual cleaning.
4. **Imputation & Highlighting**:
   - `Cash Balance` securities are automatically classified as `HLI` (Highly Liquid Investment).
   - MPF funds leave non-cash unclassified securities blank according to regulatory reporting guidelines.
   - Other unclassified securities are defaulted to `ILI` (Illiquid Investment) and visually highlighted with yellow fill (`#FFFF00`).
5. **Scenario Analysis & Liquidity Ratio Modeling**:
   - Injects haircut formulas in columns S (Normal), T (Moderate), U (Stress), and V (Hypothetical).
   - Automatically computes sums and compiles side-by-side executive scenario tables comparing available stressed assets against anticipated redemptions and monthly overhead (1%).

---

## Installation & Setup

```bash
pip install -r requirements.txt
```

---

## Usage

### 1. Run Master Liquidity Classification (Step 1)
```bash
python equity_liquidity.py --file-path "/path/to/Equity Liquidity.xlsx"
```

### 2. Run Fund Liquidity Result & Stress Testing (Step 2)
```bash
# Run with default folder structure:
python fund_liquidity.py

# Or run with custom folder / file paths:
python fund_liquidity.py \
    --input-file "/path/to/Liquidity Source.xlsx" \
    --output-file "/path/to/Liquidity Result.xlsx" \
    --monthly-redemption "/path/to/Monthly Redemption.xlsx" \
    --hair-cut "/path/to/Hair Cut.xlsx"

# Use relative internal sheet references in formulas (without [Liquidity Result.xlsx] prefix):
python fund_liquidity.py --use-internal-ref
```

### Drop-in Wrappers
- `update_liquidity.py`: Drop-in wrapper for Step 1.
- `create_fund_tabs.py`: Drop-in wrapper for Step 2.

```bash
python update_liquidity.py
python create_fund_tabs.py
```

### Python API Usage
```python
from fund_liquidity import process_fund_liquidity

output_path = process_fund_liquidity(
    input_file="Liquidity Source.xlsx",
    output_file="Liquidity Result.xlsx",
    redemption_file="Monthly Redemption.xlsx",
    haircut_file="Hair Cut.xlsx",
    use_internal_ref=False,
    format_numbers=True,
)
print(f"Report generated: {output_path}")
```

---

## Running Tests

Run the full pytest suite covering both steps:

```bash
python3 -m pytest -v
```
