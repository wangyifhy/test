#!/usr/bin/env python3
"""
Update 'Redemption Summary' in Monthly Redemption.xlsx from Input File.xlsx.

For each fund in Input File column D, compute over the past 36 months
(ending at the current month in Input File!B1):
  - Max / Max Month
  - 95 percentile / 95 percentile month
  - Current Month redemption
  - Average past 1 year

Monthly amounts come from each 'YYYY MMM' tab, column 'Outflow Base$'.
Missing months (or fund not listed) are treated as 0.
"""

from __future__ import annotations

import argparse
import calendar
import logging
import math
import os
from datetime import datetime
from typing import Dict, Iterable, List, Optional, Sequence, Tuple

from openpyxl import load_workbook
from openpyxl.workbook.workbook import Workbook
from openpyxl.worksheet.worksheet import Worksheet

DEFAULT_FOLDER = os.environ.get(
    "LIQUIDITY_FOLDER",
    r"d:\HaiTong Work\Liquidity Automation",
)
DEFAULT_INPUT_FILE = "Input File.xlsx"
DEFAULT_REDEMPTION_FILE = "Monthly Redemption.xlsx"

LOOKBACK_MONTHS = 36  # past three years inclusive of current month
AVG_MONTHS = 12
OUTFLOW_BASE_COL = 5  # E: Outflow Base$
PORTFOLIO_COL = 1  # A: Portfolio
FUND_LIST_COL = 4  # D: Fund List
CURRENCY_COL = 6  # F: Base Currency

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
)
logger = logging.getLogger(__name__)


def add_months(dt: datetime, n: int) -> datetime:
    y = dt.year + (dt.month - 1 + n) // 12
    m = (dt.month - 1 + n) % 12 + 1
    return datetime(y, m, 1)


def month_label(dt: datetime) -> str:
    return f"{dt.year} {calendar.month_abbr[dt.month]}"


def parse_month(value) -> datetime:
    """Parse Input File!B1 (datetime, YYYY-MM-DD, YYYY-MM, or 'YYYY MMM')."""
    if isinstance(value, datetime):
        return datetime(value.year, value.month, 1)
    if value is None or str(value).strip() == "":
        raise ValueError("Current month (Input File!B1) is empty")
    text = str(value).strip()
    candidates = (
        (text[:19], "%Y-%m-%d %H:%M:%S"),
        (text[:10], "%Y-%m-%d"),
        (text[:7], "%Y-%m"),
        (text, "%Y %b"),
        (text, "%Y %B"),
    )
    for candidate, fmt in candidates:
        try:
            parsed = datetime.strptime(candidate, fmt)
            return datetime(parsed.year, parsed.month, 1)
        except ValueError:
            continue
    raise ValueError(f"Unrecognized current month value: {value!r}")


def parse_number(value) -> Optional[float]:
    if value is None:
        return None
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).strip().replace(",", "")
    if not text or text in {"-", "—", "–"} or set(text) <= {"-", " "}:
        return None
    try:
        return float(text)
    except ValueError:
        return None


def percentile_inc(values: Sequence[float], p: float) -> float:
    """Excel PERCENTILE / PERCENTILE.INC."""
    xs = sorted(values)
    n = len(xs)
    if n == 0:
        return 0.0
    if n == 1:
        return xs[0]
    k = 1 + (n - 1) * p
    f = int(math.floor(k))
    c = int(math.ceil(k))
    if f == c:
        return xs[f - 1]
    return xs[f - 1] + (xs[c - 1] - xs[f - 1]) * (k - f)


def percentile_month(series: Sequence[Tuple[str, float]], p: float) -> str:
    """
    Month corresponding to the upper order statistic used by PERCENTILE.INC.
    series: list of (month_label, value) in chronological order.
    """
    n = len(series)
    if n == 0:
        return ""
    k = 1 + (n - 1) * p
    rank = int(math.ceil(k))  # 1-based
    # Stable sort keeps chronological order on ties (matches existing file)
    ordered = sorted(series, key=lambda x: x[1])
    return ordered[rank - 1][0]


def month_window(end_month: datetime, lookback: int = LOOKBACK_MONTHS) -> List[str]:
    return [month_label(add_months(end_month, -i)) for i in range(lookback - 1, -1, -1)]


def build_month_index(
    wb: Workbook,
    end_month: datetime,
    lookback: int = LOOKBACK_MONTHS,
) -> Tuple[List[str], Dict[str, Dict[str, float]]]:
    """Preload Outflow Base$ by fund for the lookback window."""
    labels = month_window(end_month, lookback)
    by_fund: Dict[str, Dict[str, float]] = {}

    for label in labels:
        if label not in wb.sheetnames:
            continue
        ws = wb[label]
        for row in range(2, ws.max_row + 1):
            fund = ws.cell(row=row, column=PORTFOLIO_COL).value
            if fund is None:
                continue
            fund = str(fund).strip()
            if not fund:
                continue
            amount = parse_number(ws.cell(row=row, column=OUTFLOW_BASE_COL).value)
            if amount is None:
                amount = 0.0
            by_fund.setdefault(fund, {})[label] = amount

    return labels, by_fund


def fund_series(
    fund: str,
    labels: Sequence[str],
    by_fund: Dict[str, Dict[str, float]],
) -> List[Tuple[str, float]]:
    fund_map = by_fund.get(fund, {})
    # Missing month / fund not on sheet -> 0 (matches existing summary logic)
    return [(label, float(fund_map.get(label, 0.0))) for label in labels]


def summarize_fund(
    fund: str,
    labels: Sequence[str],
    by_fund: Dict[str, Dict[str, float]],
    avg_months: int = AVG_MONTHS,
) -> dict:
    series = fund_series(fund, labels, by_fund)
    values = [v for _, v in series]

    max_val = max(values) if values else 0.0
    # On ties, keep the earliest month (chronological first match)
    max_month = next((m for m, v in series if v == max_val), labels[0] if labels else "")

    p95 = percentile_inc(values, 0.95)
    p95_month = percentile_month(series, 0.95)

    current_val = series[-1][1] if series else 0.0
    avg_1y = sum(values[-avg_months:]) / avg_months if values else 0.0

    return {
        "fund": fund,
        # Existing sheet stores month in B and value in C (headers B/C are swapped vs labels)
        "max_month": max_month,
        "max": max_val,
        "p95": p95,
        "p95_month": p95_month,
        "current": current_val,
        "avg_1y": avg_1y,
    }


def read_input_funds(path: str) -> Tuple[datetime, List[Tuple[str, str]]]:
    wb = load_workbook(path, data_only=True)
    ws = wb.active
    current_month = parse_month(ws["B1"].value)

    funds: List[Tuple[str, str]] = []
    for row in range(2, ws.max_row + 1):
        fund = ws.cell(row=row, column=FUND_LIST_COL).value  # D: Fund List
        if fund is None or str(fund).strip() == "":
            continue
        currency = ws.cell(row=row, column=CURRENCY_COL).value  # F: Base Currency
        funds.append((str(fund).strip(), str(currency).strip() if currency else ""))
    return current_month, funds


def clear_summary_data_rows(ws: Worksheet, first_data_row: int = 2) -> None:
    """Clear prior fund rows but keep hypo / FX cells below the table."""
    # Preserve Q1:T2, Q9:R11, etc. Only clear A:N for old fund rows.
    max_row = ws.max_row
    for row in range(first_data_row, max_row + 1):
        # Stop once we leave the fund block (no fund name and no max-month text pattern)
        a_val = ws.cell(row=row, column=1).value
        if a_val is None:
            # still clear stray formula rows that only have I:N
            has_i = ws.cell(row=row, column=9).value is not None
            if not has_i:
                continue
        for col in range(1, 15):  # A:N
            ws.cell(row=row, column=col).value = None


def write_summary(ws: Worksheet, rows: Iterable[dict]) -> None:
    """
    Write summary rows.
    Column layout matches the existing working file (used by Equity Fund Liquidity):
      A fund, B max month, C max, D 95pct, E 95pct month, F current, G avg1y, H ccy
      I:N USD conversion formulas
    """
    # Ensure headers
    headers = [
        "Summary",
        "Max",
        "Max Month",
        "95 percentile",
        "95 percentile month",
        "Current Month",
        "Average past 1 year",
        "Base Currency",
        "Max",
        "Max Month",
        "95 percentile",
        "95 percentile month",
        "Current Month",
        "Average past 1 year",
    ]
    for i, h in enumerate(headers, 1):
        if ws.cell(1, i).value is None:
            ws.cell(1, i).value = h

    clear_summary_data_rows(ws)

    for i, row in enumerate(rows):
        r = i + 2
        ws.cell(r, 1).value = row["fund"]
        # NOTE: existing file stores Max Month in B and Max value in C
        # (headers say the opposite; I:N formulas and liquidity VLOOKUPs rely on this)
        ws.cell(r, 2).value = row["max_month"]
        ws.cell(r, 3).value = row["max"]
        ws.cell(r, 4).value = row["p95"]
        ws.cell(r, 5).value = row["p95_month"]
        ws.cell(r, 6).value = row["current"]
        ws.cell(r, 7).value = row["avg_1y"]
        ws.cell(r, 8).value = row["currency"]

        ws.cell(r, 9).value = f"=B{r}"
        ws.cell(r, 10).value = (
            f'=IF($H{r}<>"USD",IF($H{r}="CNY",C{r}/$R$11,C{r}/$R$10),C{r})'
        )
        ws.cell(r, 11).value = (
            f'=IF($H{r}<>"USD",IF($H{r}="CNY",D{r}/$R$11,D{r}/$R$10),D{r})'
        )
        ws.cell(r, 12).value = f"=E{r}"
        ws.cell(r, 13).value = (
            f'=IF($H{r}<>"USD",IF($H{r}="CNY",F{r}/$R$11,F{r}/$R$10),F{r})'
        )
        ws.cell(r, 14).value = (
            f'=IF($H{r}<>"USD",IF($H{r}="CNY",G{r}/$R$11,G{r}/$R$10),G{r})'
        )


def resolve_path(explicit: Optional[str], folder: str, filename: str) -> str:
    if explicit:
        return explicit
    return os.path.join(folder, filename)


def process_redemption_summary(
    input_file: Optional[str] = None,
    redemption_file: Optional[str] = None,
    folder: str = DEFAULT_FOLDER,
) -> Tuple[str, List[dict]]:
    """
    Rebuild Redemption Summary from Input File funds and monthly outflow tabs.

    Returns
    -------
    (output_path, rows)
        Path actually written and the computed summary rows.
    """
    input_path = resolve_path(input_file, folder, DEFAULT_INPUT_FILE)
    redemption_path = resolve_path(redemption_file, folder, DEFAULT_REDEMPTION_FILE)

    if not os.path.exists(input_path):
        raise FileNotFoundError(f"Input file not found: {input_path}")
    if not os.path.exists(redemption_path):
        raise FileNotFoundError(f"Monthly Redemption workbook not found: {redemption_path}")

    current_month, funds = read_input_funds(input_path)
    logger.info("Current month: %s", month_label(current_month))
    logger.info("Funds from Input File: %s", len(funds))

    wb = load_workbook(redemption_path)
    if "Redemption Summary" not in wb.sheetnames:
        raise ValueError("Missing 'Redemption Summary' sheet")

    labels, by_fund = build_month_index(wb, current_month)
    logger.info("Window: %s -> %s (%s months)", labels[0], labels[-1], len(labels))

    rows: List[dict] = []
    for fund, currency in funds:
        stats = summarize_fund(fund, labels, by_fund)
        stats["currency"] = currency
        rows.append(stats)
        logger.info(
            "  %s: max=%.2f (%s), p95=%.2f (%s), current=%.2f",
            fund,
            stats["max"],
            stats["max_month"],
            stats["p95"],
            stats["p95_month"],
            stats["current"],
        )

    write_summary(wb["Redemption Summary"], rows)
    try:
        wb.save(redemption_path)
        logger.info("Updated: %s", redemption_path)
        return redemption_path, rows
    except PermissionError:
        alt = os.path.join(os.path.dirname(redemption_path), "Monthly Redemption_updated.xlsx")
        wb.save(alt)
        logger.warning(
            "Could not overwrite %s (file is open). Wrote: %s",
            redemption_path,
            alt,
        )
        return alt, rows


def parse_args(argv: Optional[Sequence[str]] = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Update Monthly Redemption.xlsx Redemption Summary from Input File.xlsx",
    )
    parser.add_argument(
        "--folder",
        default=DEFAULT_FOLDER,
        help=f"Folder containing the workbooks (default: {DEFAULT_FOLDER})",
    )
    parser.add_argument(
        "--input-file",
        default=None,
        help="Path to Input File.xlsx (default: <folder>/Input File.xlsx)",
    )
    parser.add_argument(
        "--redemption-file",
        default=None,
        help="Path to Monthly Redemption.xlsx (default: <folder>/Monthly Redemption.xlsx)",
    )
    return parser.parse_args(argv)


def main(argv: Optional[Sequence[str]] = None) -> None:
    args = parse_args(argv)
    process_redemption_summary(
        input_file=args.input_file,
        redemption_file=args.redemption_file,
        folder=args.folder,
    )


if __name__ == "__main__":
    main()
