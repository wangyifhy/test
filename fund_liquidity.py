#!/usr/bin/env python3
"""
Equity Fund Liquidity Result Generator

Processes equity liquidity data across individual fund portfolios:
1. Copies 'Redemption Summary' and 'Hair Cut' sheets as values with formatting into 'Liquidity Result.xlsx'.
2. Generates individual fund tabs from 'LiqTracking', enriched with classifications from 'Liquidity Source'.
3. Imputes missing classifications based on security type (e.g. Cash Balance -> HLI, MPF rules, ILI highlight).
4. Generates scenario stress-test formulas (Normal, Moderate, Stress, Hypothetical 1-4).
5. Compiles executive scenario summary and liquidity ratio tables for each fund.

Optimizations & Features:
- Preserves all cell formatting, styles, column widths, row heights, and merged cell ranges.
- Avoids multiple re-save cycles by performing in-memory batch operations.
- Clean vectorization and prevention of pandas SettingWithCopy warnings.
- Robust handling of leading quotes, whitespace, and type coercion in identifiers.
- Dynamic column letter resolution with fallback to standard L, O, P, Q, R, S, T, U, V layout.
- Professional number and percentage formatting for executive summary tables.
- Full CLI support and modular API for automated scheduled pipelines.
"""

import argparse
from copy import copy
import logging
import os
import shutil
import sys
from typing import Dict, List, Optional, Tuple, Union

import openpyxl
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet
import pandas as pd

# Default paths matching the equity liquidity master folder structure
DEFAULT_FOLDER = r"Q:\Risk Management\Liquidity\Master File\Equity"
DEFAULT_INPUT_FILE = "Liquidity Source.xlsx"
DEFAULT_OUTPUT_FILE = "Liquidity Result.xlsx"
DEFAULT_REDEMPTION_FILE = "Monthly Redemption.xlsx"
DEFAULT_HAIRCUT_FILE = "Hair Cut.xlsx"

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
)
logger = logging.getLogger(__name__)

# Yellow highlight fill for imputed ILI classifications
YELLOW_FILL = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")


def copy_sheet_formula_as_value_with_format(
    src_path_or_wb: Union[str, Workbook],
    sheet_name: str,
    dest_path_or_wb: Union[str, Workbook],
    save_dest: bool = True,
) -> Workbook:
    """
    Copy a worksheet from a source workbook to a destination workbook, converting formulas
    to static values while preserving all visual formatting, styles, column/row dimensions,
    and merged cell ranges.

    Parameters
    ----------
    src_path_or_wb : str or openpyxl.Workbook
        Path to source workbook or open workbook instance.
    sheet_name : str
        Name of sheet to copy.
    dest_path_or_wb : str or openpyxl.Workbook
        Path to destination workbook or open workbook instance.
    save_dest : bool, default True
        Whether to save the destination workbook if dest_path_or_wb is a filepath.

    Returns
    -------
    openpyxl.Workbook
        The modified destination workbook instance.
    """
    logger.info("Copying sheet '%s' as values with formatting...", sheet_name)

    # Load source values (evaluated) and styles/dimensions
    if isinstance(src_path_or_wb, str):
        if not os.path.exists(src_path_or_wb):
            raise FileNotFoundError(f"Source file not found: {src_path_or_wb}")
        src_wb_val = load_workbook(src_path_or_wb, data_only=True)
        src_wb_style = load_workbook(src_path_or_wb, data_only=False)
    else:
        src_wb_val = src_path_or_wb
        src_wb_style = src_path_or_wb

    if sheet_name not in src_wb_val.sheetnames:
        raise ValueError(f"Sheet '{sheet_name}' not found in source workbook.")

    # Load destination workbook
    if isinstance(dest_path_or_wb, str):
        if os.path.exists(dest_path_or_wb):
            dest_wb = load_workbook(dest_path_or_wb)
        else:
            dest_wb = Workbook()
            # Remove default active sheet if empty
            if "Sheet" in dest_wb.sheetnames and len(dest_wb.sheetnames) == 1:
                pass
    else:
        dest_wb = dest_path_or_wb

    # Remove existing sheet with same name in destination
    if sheet_name in dest_wb.sheetnames:
        del dest_wb[sheet_name]

    src_ws_val = src_wb_val[sheet_name]
    src_ws_style = src_wb_style[sheet_name]
    dest_ws = dest_wb.create_sheet(sheet_name)

    # Copy cell values and styles
    for i, (row_val, row_style) in enumerate(zip(src_ws_val.iter_rows(), src_ws_style.iter_rows()), start=1):
        for j, (cell_val, cell_style) in enumerate(zip(row_val, row_style), start=1):
            new_cell = dest_ws.cell(row=i, column=j, value=cell_val.value)
            if cell_style.has_style:
                if cell_style.font:
                    new_cell.font = copy(cell_style.font)
                if cell_style.border:
                    new_cell.border = copy(cell_style.border)
                if cell_style.fill:
                    new_cell.fill = copy(cell_style.fill)
                if cell_style.number_format:
                    new_cell.number_format = copy(cell_style.number_format)
                if cell_style.protection:
                    new_cell.protection = copy(cell_style.protection)
                if cell_style.alignment:
                    new_cell.alignment = copy(cell_style.alignment)

    # Copy merged cell ranges
    for merge_range in src_ws_style.merged_cells.ranges:
        dest_ws.merge_cells(str(merge_range))

    # Copy column dimensions
    for col_letter, col_dim in src_ws_style.column_dimensions.items():
        if col_dim.width is not None:
            dest_ws.column_dimensions[col_letter].width = col_dim.width
        if col_dim.hidden:
            dest_ws.column_dimensions[col_letter].hidden = True

    # Copy row dimensions
    for row_num, row_dim in src_ws_style.row_dimensions.items():
        if row_dim.height is not None:
            dest_ws.row_dimensions[row_num].height = row_dim.height
        if row_dim.hidden:
            dest_ws.row_dimensions[row_num].hidden = True

    # Save destination if filepath provided and requested
    if isinstance(dest_path_or_wb, str) and save_dest:
        dest_wb.save(dest_path_or_wb)
        logger.info("Saved copied sheet '%s' to '%s'.", sheet_name, dest_path_or_wb)

    return dest_wb


import re


def normalize_identifier(id_str: object) -> str:
    """
    Normalize an identifier string (e.g. 'Fund A0700' and 'Fund A700') so that
    trailing numeric codes with differing leading zeros still match.
    """
    s = str(id_str).strip().lstrip("'").rstrip()
    if s.endswith(".0"):
        s = s[:-2]
    # If ends with digits, normalize leading zeros while keeping at least one digit
    m = re.match(r"^(.*?)0*(\d+)$", s)
    if m:
        prefix, digits = m.groups()
        return f"{prefix}{digits}"
    return s


def clean_security_id(series: pd.Series) -> pd.Series:
    """Normalize Security ID strings by stripping quotes, trailing .0, and whitespace."""
    return (
        series.astype(str)
        .str.lstrip("'")
        .str.strip()
        .apply(lambda s: s[:-2] if s.endswith(".0") else s)
    )


def prepare_fund_data(
    fund_name: str,
    liq_tracking_df: pd.DataFrame,
    liquidity_source_df: pd.DataFrame,
    fund_col: str = "Fund Name",
    sec_id_col: str = "Security ID",
    sec_type_col: str = "Sec Type",
) -> pd.DataFrame:
    """
    Filter LiqTracking records for a specific fund, sort Sec Type descending,
    and merge the 4 classification and sector columns from Liquidity Source.

    Parameters
    ----------
    fund_name : str
        Target fund identifier.
    liq_tracking_df : pd.DataFrame
        DataFrame containing LiqTracking portfolio positions.
    liquidity_source_df : pd.DataFrame
        DataFrame containing liquidity source classifications.

    Returns
    -------
    pd.DataFrame
        Filtered and merged portfolio positions ready for Excel output.
    """
    # 1. Filter rows for target fund
    mask = liq_tracking_df[fund_col].astype(str).str.strip() == str(fund_name).strip()
    fund_data = liq_tracking_df[mask].copy()

    if fund_data.empty:
        logger.warning("No records found in LiqTracking for fund '%s'.", fund_name)
        return fund_data

    # 2. Sort Sec Type descending
    if sec_type_col in fund_data.columns:
        fund_data.sort_values(by=sec_type_col, ascending=False, inplace=True)
    else:
        logger.warning("Column '%s' not found for sorting fund '%s'.", sec_type_col, fund_name)

    # 3. Create cleaned matching identifiers
    fund_clean_id = clean_security_id(fund_data[sec_id_col])
    fund_data["Identifier 1"] = fund_data[fund_col].astype(str).str.strip() + fund_clean_id
    fund_data["_match_key"] = fund_data["Identifier 1"].apply(normalize_identifier)

    # Prepare liquidity source columns
    source_cols = [
        "Liquidity Classification in the past 30 days",
        "Liquidity Classification in the past 5 years (5th percentile)",
        "Liquidity Classification in the past 5 years",
    ]

    # Ensure Identifier 2 exists on liquidity_source_df
    if "Identifier 2" in liquidity_source_df.columns:
        source_cols.insert(0, "Identifier 2")
    else:
        fund_col_src = fund_col if fund_col in liquidity_source_df.columns else liquidity_source_df.columns[0]
        sec_id_col_src = sec_id_col if sec_id_col in liquidity_source_df.columns else liquidity_source_df.columns[1]
        liquidity_source_df["Identifier 2"] = (
            liquidity_source_df[fund_col_src].astype(str).str.strip()
            + clean_security_id(liquidity_source_df[sec_id_col_src])
        )
        source_cols.insert(0, "Identifier 2")

    # Check for industry sector column casing
    if "INDUSTRY_SECTOR" in liquidity_source_df.columns:
        source_cols.append("INDUSTRY_SECTOR")
        rename_sector = True
    elif "Industry_Sector" in liquidity_source_df.columns:
        source_cols.append("Industry_Sector")
        rename_sector = False
    else:
        # Create empty placeholder if missing
        liquidity_source_df["Industry_Sector"] = ""
        source_cols.append("Industry_Sector")
        rename_sector = False

    source_subset = liquidity_source_df[source_cols].copy()
    source_subset["_match_key"] = source_subset["Identifier 2"].apply(normalize_identifier)

    # Deduplicate liquidity source on _match_key
    source_subset = source_subset.drop_duplicates(subset=["_match_key"])

    # 4. Merge classifications
    merged = pd.merge(
        fund_data,
        source_subset,
        on="_match_key",
        how="left",
    )

    # 5. Clean helper columns and normalize Industry_Sector
    cols_to_drop = ["_match_key"]
    if "Identifier 1" in merged.columns:
        cols_to_drop.append("Identifier 1")
    if "Identifier 2" in merged.columns:
        cols_to_drop.append("Identifier 2")
    merged.drop(columns=[c for c in cols_to_drop if c in merged.columns], inplace=True)

    if rename_sector:
        merged.rename(columns={"INDUSTRY_SECTOR": "Industry_Sector"}, inplace=True)

    # Ensure Industry_Sector is placed as the last column of the merged block
    cols = [c for c in merged.columns if c != "Industry_Sector"] + ["Industry_Sector"]
    merged = merged[cols]

    return merged


def populate_fund_worksheet(
    worksheet: Worksheet,
    fund_name: str,
    fund_data: pd.DataFrame,
    haircut_sheet_name: str = "Hair Cut",
    redemption_sheet_name: str = "Redemption Summary",
    workbook_name_ref: Optional[str] = "[Liquidity Result.xlsx]",
    format_numbers: bool = True,
) -> None:
    """
    Populate an openpyxl Worksheet for a single fund with data, missing value imputation,
    stress-test scenario formulas (Normal, Moderate, Stress, Hypothetical), and summary tables.

    Parameters
    ----------
    worksheet : openpyxl.worksheet.worksheet.Worksheet
        The worksheet to populate.
    fund_name : str
        Name of the equity fund.
    fund_data : pd.DataFrame
        DataFrame of fund positions and merged liquidity classifications.
    haircut_sheet_name : str, default 'Hair Cut'
        Worksheet name holding haircut tables.
    redemption_sheet_name : str, default 'Redemption Summary'
        Worksheet name holding redemption summary tables.
    workbook_name_ref : Optional[str], default '[Liquidity Result.xlsx]'
        Workbook prefix in formula links (e.g. '[Liquidity Result.xlsx]').
        Set to None or empty string to use relative internal sheet references without workbook name.
    format_numbers : bool, default True
        Whether to apply professional currency and percentage number formats.
    """
    # Write DataFrame to worksheet starting at A1
    # 1. Write header row
    headers = list(fund_data.columns)
    for col_idx, header in enumerate(headers, start=1):
        worksheet.cell(row=1, column=col_idx, value=header)

    # 2. Write data rows
    for row_idx, row_values in enumerate(fund_data.itertuples(index=False), start=2):
        for col_idx, val in enumerate(row_values, start=1):
            # Write NaN/None as None
            cell_val = None if pd.isna(val) else val
            worksheet.cell(row=row_idx, column=col_idx, value=cell_val)

    # Determine column positions dynamically
    header_map = {col_name: idx + 1 for idx, col_name in enumerate(headers)}

    col_30d = header_map.get("Liquidity Classification in the past 30 days", 15)
    col_5p = header_map.get("Liquidity Classification in the past 5 years (5th percentile)", 16)
    col_5y = header_map.get("Liquidity Classification in the past 5 years", 17)
    col_sec_type = header_map.get("Sec Type")
    col_industry = header_map.get("Industry_Sector", 18)

    # Find Market Value column (typically column L / 12)
    col_mkt_val = 12
    for candidate in ["Mkt Value L$", "Mkt Value USD", "Mkt Value", "Market Value", "mkt_USD"]:
        if candidate in header_map:
            col_mkt_val = header_map[candidate]
            break

    # Impute missing classifications
    # Rule 1: Cash Balance -> HLI (no fill)
    # Rule 2: MPF funds and not Cash Balance -> Leave blank
    # Rule 3: Other -> ILI and highlight yellow
    is_mpf = fund_name.startswith("MPF")
    max_row = worksheet.max_row

    if max_row >= 2 and col_sec_type is not None:
        classification_cols = [col_30d, col_5p, col_5y]
        for row in range(2, max_row + 1):
            sec_type_raw = worksheet.cell(row=row, column=col_sec_type).value
            sec_type = str(sec_type_raw).strip() if sec_type_raw is not None else ""

            for col in classification_cols:
                cell = worksheet.cell(row=row, column=col)
                if cell.value in (None, "", "nan", "NaN") or pd.isna(cell.value):
                    if sec_type == "Cash Balance":
                        cell.value = "HLI"
                    elif is_mpf and sec_type != "Cash Balance":
                        continue
                    else:
                        cell.value = "ILI"
                        cell.fill = YELLOW_FILL

    # Freeze header row
    worksheet.freeze_panes = "A2"

    # Define haircut table reference string
    hc_prefix = f"'{workbook_name_ref}{haircut_sheet_name}'!" if workbook_name_ref else f"'{haircut_sheet_name}'!"
    hc_internal_prefix = f"'{haircut_sheet_name}'!"

    # Resolve column letters
    l_let = get_column_letter(col_mkt_val)
    o_let = get_column_letter(col_30d)
    p_let = get_column_letter(col_5p)
    q_let = get_column_letter(col_5y)
    r_let = get_column_letter(col_industry)

    # Output scenario columns: S (19), T (20), U (21), V (22)
    s_let = "S"
    t_let = "T"
    u_let = "U"
    v_let = "V"

    # Add scenario headers
    worksheet[f"{s_let}1"] = "Normal"
    worksheet[f"{t_let}1"] = "Moderate"
    worksheet[f"{u_let}1"] = "Stress"
    worksheet[f"{v_let}1"] = "Hypothetical"

    # Fill formulas in columns S, T, U, V
    if max_row >= 2:
        for row in range(2, max_row + 1):
            o_cell = f"{o_let}{row}"
            p_cell = f"{p_let}{row}"
            q_cell = f"{q_let}{row}"
            r_cell = f"{r_let}{row}"
            l_cell = f"{l_let}{row}"

            # S (Normal Scenario Haircut)
            worksheet[f"{s_let}{row}"] = f"=(1-VLOOKUP({o_cell},{hc_prefix}$A$1:$D$5,2,0))*{l_cell}"
            # T (Moderate Scenario Haircut)
            worksheet[f"{t_let}{row}"] = f"=(1-VLOOKUP({p_cell},{hc_prefix}$A$1:$D$5,3,0))*{l_cell}"
            # U (Stress Scenario Haircut)
            worksheet[f"{u_let}{row}"] = f"=(1-VLOOKUP({q_cell},{hc_prefix}$A$1:$D$5,4,0))*{l_cell}"
            # V (Hypothetical Industry Sector Haircut)
            worksheet[f"{v_let}{row}"] = (
                f"=(1+IF(ISNA(VLOOKUP({r_cell},{hc_internal_prefix}$A$9:$B$10,2,0)),"
                f"MIN({hc_internal_prefix}$B$9:$B$10),"
                f"VLOOKUP({r_cell},{hc_internal_prefix}$A$9:$B$10,2,0)))*{l_cell}"
            )

        sum_row = max_row + 1
        worksheet[f"{s_let}{sum_row}"] = f"=SUM({s_let}2:{s_let}{max_row})"
        worksheet[f"{t_let}{sum_row}"] = f"=SUM({t_let}2:{t_let}{max_row})"
        worksheet[f"{u_let}{sum_row}"] = f"=SUM({u_let}2:{u_let}{max_row})"
        worksheet[f"{v_let}{sum_row}"] = f"=SUM({v_let}2:{v_let}{max_row})"
        worksheet[f"{l_let}{sum_row}"] = f"=SUM({l_let}2:{l_let}{max_row})"
    else:
        # Edge case: No data rows
        sum_row = 2
        worksheet[f"{s_let}{sum_row}"] = 0
        worksheet[f"{t_let}{sum_row}"] = 0
        worksheet[f"{u_let}{sum_row}"] = 0
        worksheet[f"{v_let}{sum_row}"] = 0
        worksheet[f"{l_let}{sum_row}"] = 0

    # Add Hypothetical table in columns E, F, G, H
    hypo_cols = ["E", "F", "G", "H"]
    hypo_names = ["Hypothetical 1", "Hypothetical 2", "Hypothetical 3", "Hypothetical 4"]
    redemption_cols = ["Q", "R", "S", "T"]

    for idx, col in enumerate(hypo_cols):
        r_col = redemption_cols[idx]
        # Two rows below the summation formula: Header
        worksheet[f"{col}{sum_row+2}"] = hypo_names[idx]
        # Three rows below: Asset value (=V sum)
        worksheet[f"{col}{sum_row+3}"] = f"={v_let}{sum_row}"
        # Four rows below: Redemption amount (=Total Mkt Val * Redemption Factor)
        worksheet[f"{col}{sum_row+4}"] = f"=${l_let}${sum_row}*'{redemption_sheet_name}'!{r_col}$2"
        # Five rows below: Monthly overhead (= 1% of asset)
        worksheet[f"{col}{sum_row+5}"] = f"=0.01*{col}{sum_row+3}"
        # Six rows below: Liquidity Ratio
        worksheet[f"{col}{sum_row+6}"] = f"={col}{sum_row+3}/SUM({col}{sum_row+4}:{col}{sum_row+5})"

    # Add Scenario Summary Table in columns A, B, C, D
    table_row = sum_row + 2

    # Column A (Row Labels)
    worksheet.cell(row=table_row, column=1, value=fund_name)
    worksheet.cell(row=table_row + 1, column=1, value="Asset")
    worksheet.cell(row=table_row + 2, column=1, value="Redemption")
    worksheet.cell(row=table_row + 3, column=1, value="Monthly overhead")
    worksheet.cell(row=table_row + 4, column=1, value="Liquidity Ratio")

    # Column B (Normal Scenario)
    worksheet.cell(row=table_row, column=2, value="Normal Scenario")
    worksheet.cell(row=table_row + 1, column=2, value=f"={s_let}{sum_row}")
    worksheet.cell(row=table_row + 2, column=2, value=f"=VLOOKUP(A2,'{redemption_sheet_name}'!A:N,14,0)")
    worksheet.cell(row=table_row + 3, column=2, value=f"=0.01*B{table_row+1}")
    worksheet.cell(row=table_row + 4, column=2, value=f"=B{table_row+1}/SUM(B{table_row+2}:B{table_row+3})")

    # Column C (Moderate Scenario)
    worksheet.cell(row=table_row, column=3, value="Moderate Scenario")
    worksheet.cell(row=table_row + 1, column=3, value=f"={t_let}{sum_row}")
    worksheet.cell(row=table_row + 2, column=3, value=f"=VLOOKUP(A2,'{redemption_sheet_name}'!A:N,11,0)")
    worksheet.cell(row=table_row + 3, column=3, value=f"=0.01*C{table_row+1}")
    worksheet.cell(row=table_row + 4, column=3, value=f"=C{table_row+1}/SUM(C{table_row+2}:C{table_row+3})")

    # Column D (Severe Scenario)
    worksheet.cell(row=table_row, column=4, value="Severe Scenario")
    worksheet.cell(row=table_row + 1, column=4, value=f"={u_let}{sum_row}")
    worksheet.cell(row=table_row + 2, column=4, value=f"=VLOOKUP(A2,'{redemption_sheet_name}'!A:N,10,0)")
    worksheet.cell(row=table_row + 3, column=4, value=f"=0.01*D{table_row+1}")
    worksheet.cell(row=table_row + 4, column=4, value=f"=D{table_row+1}/SUM(D{table_row+2}:D{table_row+3})")

    # Optional number formatting
    if format_numbers:
        # Currency formatting for summation row
        currency_fmt = "#,##0.00"
        for col_name in [l_let, s_let, t_let, u_let, v_let]:
            worksheet[f"{col_name}{sum_row}"].number_format = currency_fmt

        # Formatting for Asset, Redemption, and Overhead rows across cols B to H
        for col_char in ["B", "C", "D", "E", "F", "G", "H"]:
            worksheet[f"{col_char}{table_row+1}"].number_format = currency_fmt
            worksheet[f"{col_char}{table_row+2}"].number_format = currency_fmt
            worksheet[f"{col_char}{table_row+3}"].number_format = currency_fmt
            # Liquidity Ratio row
            worksheet[f"{col_char}{table_row+4}"].number_format = "0.00%"


def process_fund_liquidity(
    folder: Optional[str] = None,
    input_file: Optional[str] = None,
    output_file: Optional[str] = None,
    redemption_file: Optional[str] = None,
    haircut_file: Optional[str] = None,
    use_internal_ref: bool = False,
    format_numbers: bool = True,
) -> str:
    """
    Execute the complete fund liquidity workflow:
    1. Resolve file paths and initialize output workbook.
    2. Copy 'Redemption Summary' and 'Hair Cut' sheets with evaluated values and formatting.
    3. Read 'LiqTracking', 'Equity Fund List', and 'Liquidity Source' sheets.
    4. For each fund, generate individual portfolio sheets with stress scenarios and summary tables.
    5. Save the output workbook in a single pass.

    Returns
    -------
    str
        Path to the generated output workbook.
    """
    base_folder = folder if folder else (DEFAULT_FOLDER if os.path.exists(DEFAULT_FOLDER) else ".")

    input_path = input_file if input_file else os.path.join(base_folder, DEFAULT_INPUT_FILE)
    output_path = output_file if output_file else os.path.join(base_folder, DEFAULT_OUTPUT_FILE)
    monthly_redemption_path = redemption_file if redemption_file else os.path.join(base_folder, DEFAULT_REDEMPTION_FILE)
    hair_cut_path = haircut_file if haircut_file else os.path.join(base_folder, DEFAULT_HAIRCUT_FILE)

    logger.info("Initializing Equity Fund Liquidity generation...")
    logger.info("Input file: %s", input_path)
    logger.info("Output file: %s", output_path)
    logger.info("Monthly Redemption file: %s", monthly_redemption_path)
    logger.info("Hair Cut file: %s", hair_cut_path)

    if not os.path.exists(input_path):
        raise FileNotFoundError(f"Input workbook not found: {input_path}")
    if not os.path.exists(monthly_redemption_path):
        raise FileNotFoundError(f"Monthly Redemption workbook not found: {monthly_redemption_path}")
    if not os.path.exists(hair_cut_path):
        raise FileNotFoundError(f"Hair Cut workbook not found: {hair_cut_path}")

    # Ensure output file directory exists
    out_dir = os.path.dirname(output_path)
    if out_dir and not os.path.exists(out_dir):
        os.makedirs(out_dir, exist_ok=True)

    # 1. Initialize output workbook
    if not os.path.exists(output_path):
        logger.info("Creating output workbook by copying '%s' to '%s'...", input_path, output_path)
        shutil.copy(input_path, output_path)

    # Open destination workbook in memory for single-pass processing
    dest_wb = load_workbook(output_path)

    # 2. Copy 'Redemption Summary' and 'Hair Cut' sheets into destination workbook
    copy_sheet_formula_as_value_with_format(
        monthly_redemption_path,
        "Redemption Summary",
        dest_wb,
        save_dest=False,
    )
    copy_sheet_formula_as_value_with_format(
        hair_cut_path,
        "Hair Cut",
        dest_wb,
        save_dest=False,
    )

    # 3. Read input datasets
    logger.info("Loading input datasets from '%s'...", input_path)
    liq_tracking_df = pd.read_excel(input_path, sheet_name="LiqTracking")
    equity_fund_list_df = pd.read_excel(input_path, sheet_name="Equity Fund List")
    liquidity_source_df = pd.read_excel(input_path, sheet_name="Liquidity Source")

    # Clean the second column of LiqTracking (Security ID) by removing leading single quotes
    if liq_tracking_df.shape[1] > 1:
        liq_tracking_df.iloc[:, 1] = clean_security_id(liq_tracking_df.iloc[:, 1])

    # Clean Security ID in Liquidity Source if present
    if "Security ID" in liquidity_source_df.columns:
        liquidity_source_df["Security ID"] = clean_security_id(liquidity_source_df["Security ID"])
    elif liquidity_source_df.shape[1] > 1:
        liquidity_source_df.iloc[:, 1] = clean_security_id(liquidity_source_df.iloc[:, 1])

    # Precompute Identifier 2 on liquidity_source_df if not already present
    fund_col_source = "Fund Name" if "Fund Name" in liquidity_source_df.columns else liquidity_source_df.columns[0]
    sec_id_col_source = "Security ID" if "Security ID" in liquidity_source_df.columns else liquidity_source_df.columns[1]
    if "Identifier 2" not in liquidity_source_df.columns:
        liquidity_source_df["Identifier 2"] = (
            liquidity_source_df[fund_col_source].astype(str).str.strip()
            + clean_security_id(liquidity_source_df[sec_id_col_source])
        )

    # Extract target funds
    equity_funds = equity_fund_list_df.iloc[:, 0].dropna().astype(str).str.strip().tolist()
    logger.info("Found %d equity funds to process: %s", len(equity_funds), equity_funds)

    # Workbook name reference for formulas
    wb_name_ref = "" if use_internal_ref else f"[{os.path.basename(output_path)}]"

    # 4. Generate fund sheets
    for fund in equity_funds:
        logger.info("Processing fund: %s", fund)
        fund_data = prepare_fund_data(fund, liq_tracking_df, liquidity_source_df)

        # Get or create sheet in destination workbook
        if fund in dest_wb.sheetnames:
            del dest_wb[fund]
        ws = dest_wb.create_sheet(title=fund)

        populate_fund_worksheet(
            worksheet=ws,
            fund_name=fund,
            fund_data=fund_data,
            haircut_sheet_name="Hair Cut",
            redemption_sheet_name="Redemption Summary",
            workbook_name_ref=wb_name_ref,
            format_numbers=format_numbers,
        )

    # 5. Save destination workbook in a single pass
    logger.info("Saving consolidated output workbook to '%s'...", output_path)
    dest_wb.save(output_path)
    logger.info("Successfully completed fund liquidity processing.")
    return output_path


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Generate 'Liquidity Result.xlsx' with fund portfolio tabs and scenario stress testing."
    )
    parser.add_argument(
        "--folder",
        "-d",
        default=None,
        help=f"Working directory folder (default: '{DEFAULT_FOLDER}' if exists, else current directory)",
    )
    parser.add_argument(
        "--input-file",
        "-i",
        default=None,
        help="Path to input 'Liquidity Source.xlsx'",
    )
    parser.add_argument(
        "--output-file",
        "-o",
        default=None,
        help="Path to output 'Liquidity Result.xlsx'",
    )
    parser.add_argument(
        "--monthly-redemption",
        "-r",
        default=None,
        help="Path to 'Monthly Redemption.xlsx'",
    )
    parser.add_argument(
        "--hair-cut",
        "-c",
        default=None,
        help="Path to 'Hair Cut.xlsx'",
    )
    parser.add_argument(
        "--use-internal-ref",
        action="store_true",
        help="Use internal relative sheet references without workbook filename prefix in formulas",
    )
    parser.add_argument(
        "--no-format-numbers",
        action="store_true",
        help="Disable currency and percentage formatting on summary tables",
    )
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    try:
        process_fund_liquidity(
            folder=args.folder,
            input_file=args.input_file,
            output_file=args.output_file,
            redemption_file=args.monthly_redemption,
            haircut_file=args.hair_cut,
            use_internal_ref=args.use_internal_ref,
            format_numbers=not args.no_format_numbers,
        )
    except Exception as exc:
        logger.error("Execution failed: %s", exc, exc_info=True)
        sys.exit(1)


if __name__ == "__main__":
    main()
