"""
Unit and integration tests for equity_liquidity module.
"""

import os
import openpyxl
import pandas as pd
import numpy as np
import pytest

from equity_liquidity import (
    read_bbg_data,
    read_liq_tracking_data,
    build_liquidity_dataframe,
    update_liquidity_sheet,
    process_equity_liquidity,
)


@pytest.fixture
def mock_excel_file(tmp_path):
    """Create a temporary Excel file matching the Bloomberg Master Liquidity format."""
    file_path = str(tmp_path / "Equity Liquidity Test.xlsx")

    wb = openpyxl.Workbook()

    # Sheet 1: Summary (for VLOOKUP table)
    ws_summary = wb.active
    ws_summary.title = "Summary"
    # Populate Summary table at I2:K5
    # Thresholds: 0.0 -> Liquid, 0.2 -> Semi-Liquid, 0.5 -> Illiquid, 1.0 -> Highly Illiquid
    table_data = [
        (0.0, "Tier 1", "Liquid"),
        (0.2, "Tier 2", "Semi-Liquid"),
        (0.5, "Tier 3", "Illiquid"),
        (1.0, "Tier 4", "Highly Illiquid"),
    ]
    for row_idx, (thresh, tier, desc) in enumerate(table_data, start=2):
        ws_summary[f"I{row_idx}"] = thresh
        ws_summary[f"J{row_idx}"] = tier
        ws_summary[f"K{row_idx}"] = desc

    # Sheet 2: BBG Formula Save as Value
    ws_bbg = wb.create_sheet(title="BBG Formula Save as Value")
    # Columns A-F (indices 0-5): Ticker, Exch, Name, Currency, Price, Position
    # Column G (index 6): 30-day Volume
    # Columns H-BN (indices 7-65): 59 historical date columns (total 60 volume columns)
    # Extra columns: INDUSTRY_SECTOR, EXCH_CODE
    headers = [
        "ID_BB_GLOBAL", "SECURITY_DES", "NAME", "CRNCY", "PX_LAST", "POSITION",
        "VOLUME_30D",
    ]
    # Add 59 monthly volume columns
    for m in range(1, 60):
        headers.append(f"VOL_M{m}")
    headers.extend(["INDUSTRY_SECTOR", "EXCH_CODE"])
    ws_bbg.append(headers)

    # Row 1: Normal data
    row1 = ["AAPL", "US", "Apple Inc", "USD", 180.0, 50000.0, 1000000.0]
    row1.extend([800000.0 + i * 10000 for i in range(59)])
    row1.extend(["Technology", "US"])
    ws_bbg.append(row1)

    # Row 2: Bloomberg with '#N/A N/A' and string errors
    row2 = ["MSFT", "US", "Microsoft Corp", "USD", 350.0, 25000.0, 500000.0]
    # Mix numbers with '#N/A N/A'
    row2.extend(["#N/A N/A" if i % 5 == 0 else 400000.0 + i * 5000 for i in range(59)])
    row2.extend(["Technology", "US"])
    ws_bbg.append(row2)

    # Row 3: All zeros or missing
    row3 = ["0700", "HK", "Tencent Holdings", "HKD", 380.0, 10000.0, 200000.0]
    row3.extend([150000.0] * 59)
    row3.extend(["Communications", "HK"])
    ws_bbg.append(row3)

    # Sheet 3: LiqTracking
    ws_track = wb.create_sheet(title="LiqTracking")
    # First 6 rows are header/metadata, data starts on row 7 (index 6 with 0-indexing)
    for _ in range(6):
        ws_track.append(["Metadata"] * 12)

    # Columns: 0=Ticker, 1=Exch (often with leading quote), ..., 11=mkt_USD
    def make_tracking_row(ticker, exch, mkt_usd):
        r = [""] * 12
        r[0] = ticker
        r[1] = f"'{exch}"  # Leading apostrophe as in Bloomberg Excel
        r[11] = mkt_usd
        return r

    ws_track.append(make_tracking_row("AAPL", "US", 9000000.0))
    ws_track.append(make_tracking_row("MSFT", "US", 8750000.0))
    # Add a duplicate for Tencent to test duplicate aggregation
    ws_track.append(make_tracking_row("0700", "HK", 2000000.0))
    ws_track.append(make_tracking_row("0700", "HK", 1800000.0))

    # Sheet 4: Liquidity (initially empty or old)
    ws_liq = wb.create_sheet(title="Liquidity")
    ws_liq.append(["Old Data", "To Be Replaced"])

    wb.save(file_path)
    return file_path


def test_read_bbg_data(mock_excel_file):
    df = read_bbg_data(mock_excel_file)
    assert len(df) == 3
    assert "ID_BB_GLOBAL" in df.columns
    assert "INDUSTRY_SECTOR" in df.columns
    assert "EXCH_CODE" in df.columns


def test_read_liq_tracking_data(mock_excel_file):
    df = read_liq_tracking_data(mock_excel_file)
    # Check that AAPLUS and MSFTUS exist, and 0700HK was aggregated (2000000 + 1800000 = 3800000)
    assert len(df) == 3
    mkt_map = df.set_index("Identifier1")["mkt_USD"].to_dict()
    assert mkt_map["AAPLUS"] == 9000000.0
    assert mkt_map["MSFTUS"] == 8750000.0
    assert mkt_map["0700HK"] == 3800000.0


def test_build_liquidity_dataframe(mock_excel_file):
    bbg_df = read_bbg_data(mock_excel_file)
    tracking_df = read_liq_tracking_data(mock_excel_file)

    combined = build_liquidity_dataframe(bbg_df, tracking_df)

    # Check dimensions
    assert len(combined) == 3
    assert len(combined.columns) == 18

    # Verify column names and ordering
    expected_columns = [
        "ID_BB_GLOBAL", "SECURITY_DES", "NAME", "CRNCY", "PX_LAST", "POSITION",
        "Trading Volume in the past 30 Days",
        "% (Past 30 Days)",
        "Liquidity Classification in the past 30 days",
        "Least Trading Volume in the past 5 years",
        "% (Past 5 Years)",
        "Liquidity Classification in the past 5 years",
        "5th percentile Volume in the past 5 years",
        "% (5th percentile in the past 5 years)",
        "Liquidity Classification in the past 5 years (5th percentile)",
        "INDUSTRY_SECTOR",
        "mkt_USD",
        "EXCH_CODE",
    ]
    assert list(combined.columns) == expected_columns

    # Verify formula syntax for row 2 (index 0) and row 3 (index 1)
    assert combined.loc[0, "% (Past 30 Days)"] == "=F2/G2"
    assert combined.loc[0, "Liquidity Classification in the past 30 days"] == "=VLOOKUP(H2,Summary!$I$2:$K$5,3,1)"
    assert combined.loc[0, "% (Past 5 Years)"] == "=F2/J2"
    assert combined.loc[0, "Liquidity Classification in the past 5 years"] == "=VLOOKUP(K2,Summary!$I$2:$K$5,3,1)"
    assert combined.loc[0, "% (5th percentile in the past 5 years)"] == "=F2/M2"
    assert combined.loc[0, "Liquidity Classification in the past 5 years (5th percentile)"] == "=VLOOKUP(N2,Summary!$I$2:$K$5,3,1)"

    assert combined.loc[1, "% (Past 30 Days)"] == "=F3/G3"

    # Verify MSFT row handled '#N/A N/A' without crashing, and computed least volume
    assert isinstance(combined.loc[1, "Least Trading Volume in the past 5 years"], (int, float))
    assert combined.loc[1, "Least Trading Volume in the past 5 years"] > 0

    # Verify mkt_USD merged values
    assert combined.loc[0, "mkt_USD"] == 9000000.0
    assert combined.loc[1, "mkt_USD"] == 8750000.0
    assert combined.loc[2, "mkt_USD"] == 3800000.0


def test_build_liquidity_dataframe_with_iferror(mock_excel_file):
    bbg_df = read_bbg_data(mock_excel_file)
    tracking_df = read_liq_tracking_data(mock_excel_file)

    combined = build_liquidity_dataframe(bbg_df, tracking_df, use_iferror=True)
    assert combined.loc[0, "% (Past 30 Days)"] == '=IFERROR(F2/G2, "")'
    assert combined.loc[0, "% (Past 5 Years)"] == '=IFERROR(F2/J2, "")'
    assert combined.loc[0, "% (5th percentile in the past 5 years)"] == '=IFERROR(F2/M2, "")'


def test_update_liquidity_sheet_and_formatting(mock_excel_file):
    bbg_df = read_bbg_data(mock_excel_file)
    tracking_df = read_liq_tracking_data(mock_excel_file)
    combined = build_liquidity_dataframe(bbg_df, tracking_df)

    update_liquidity_sheet(mock_excel_file, combined, sheet_name="Liquidity")

    wb = openpyxl.load_workbook(mock_excel_file)
    assert "Liquidity" in wb.sheetnames
    assert "Summary" in wb.sheetnames
    assert "BBG Formula Save as Value" in wb.sheetnames
    assert "LiqTracking" in wb.sheetnames

    ws = wb["Liquidity"]
    assert ws.max_row == 4  # 1 header + 3 data rows
    assert ws.max_column == 18

    # Check cell formula
    assert ws.cell(row=2, column=8).value == "=F2/G2"

    # Check percentage formatting on columns H (8), K (11), N (14)
    for col in [8, 11, 14]:
        for r in range(2, 5):
            assert ws.cell(row=r, column=col).number_format == "0.00%"

    # Check non-percentage column formatting
    assert ws.cell(row=2, column=7).number_format == "General"


def test_process_equity_liquidity_e2e(mock_excel_file):
    result_df = process_equity_liquidity(mock_excel_file)
    assert len(result_df) == 3

    wb = openpyxl.load_workbook(mock_excel_file)
    ws = wb["Liquidity"]
    assert ws["H2"].value == "=F2/G2"
    assert ws["I2"].value == "=VLOOKUP(H2,Summary!$I$2:$K$5,3,1)"
    assert ws["H2"].number_format == "0.00%"
    assert ws["K2"].number_format == "0.00%"
    assert ws["N2"].number_format == "0.00%"


def test_missing_file_raises_error():
    with pytest.raises(FileNotFoundError):
        read_bbg_data("non_existent_file.xlsx")
