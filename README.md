# Liquidity Automation

Python utilities for HaiTong monthly liquidity / redemption reporting.

## Redemption Summary updater

`redemption_summary.py` rebuilds the **Redemption Summary** sheet in
`Monthly Redemption.xlsx` from `Input File.xlsx`.

For each fund in Input File column D, it computes over the past 36 months
(ending at the current month in `Input File!B1`):

| Statistic | Source |
| --- | --- |
| Max / Max Month | Peak `Outflow Base$` in the 36-month window |
| 95 percentile / 95 percentile month | Excel `PERCENTILE.INC` over the same window |
| Current Month redemption | Value for the month in `B1` |
| Average past 1 year | Mean of the last 12 months |

Monthly amounts come from each `YYYY MMM` tab, column **Outflow Base$**
(column E). Missing months, or a fund that is not listed on a month tab,
are treated as **0**.

The written column layout matches the existing working file used by
Equity Fund Liquidity VLOOKUPs:

- **A** fund, **B** max month, **C** max, **D** 95pct, **E** 95pct month,
  **F** current, **G** avg 1y, **H** base currency
- **I:N** USD conversion formulas (`R10` HKD, `R11` CNY)
- Hypo / FX cells in **Q:T** are left untouched

Note: headers on B/C (and I/J) are historically swapped versus the stored
values. Downstream formulas depend on that layout, so this script keeps it.

## Installation

```bash
pip install -r requirements.txt
```

## Usage

Default folder is `d:\HaiTong Work\Liquidity Automation` (override with
`--folder` or `LIQUIDITY_FOLDER`).

```bash
# Default Windows work folder
python update_redemption.py

# Explicit paths
python redemption_summary.py \
    --input-file "/path/to/Input File.xlsx" \
    --redemption-file "/path/to/Monthly Redemption.xlsx"
```

### Python API

```python
from redemption_summary import process_redemption_summary

path, rows = process_redemption_summary(
    input_file="Input File.xlsx",
    redemption_file="Monthly Redemption.xlsx",
)
```

If `Monthly Redemption.xlsx` is open and cannot be overwritten, the script
writes `Monthly Redemption_updated.xlsx` next to it instead.

## Tests

```bash
python3 -m pytest -v
```
