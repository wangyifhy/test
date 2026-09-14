import os
from datetime import datetime

import pandas as pd

# -------------------------- 日期统一处理 --------------------------
today_now = datetime.today()
date_string = today_now.strftime("%d%m%Y")
date_string_1 = today_now.strftime("%Y%m%d")
today_ts = pd.Timestamp(date_string_1)  # 统一pandas时间戳，消除日期相减警告

# -------------------------- 路径（原始字符串杜绝转义警告） --------------------------
directory_source = r"Q:\Risk Management\每日 04 Stop_Loss\Daily Output"
directory_destination = r"Q:\Risk Management\每日 04 Stop_Loss\Daily Notification"
directory_breach = r"Q:\Risk Management\每日 04 Stop_Loss"

# daily_output now uses: daily_output_<ddmmyyyy> test.xlsx
# and includes YTD Drawdown, Holding Period Drawdown, High Last Year, Drawdown from High.
source_file_name = f"daily_output_{date_string} test.xlsx"
legacy_source_file_name = f"daily_output_{date_string}.xlsx"
destination_file_name = f"Stop Loss {date_string_1}.xlsx"
notification_file_name = f"Stop Loss Notification {date_string_1}.xlsx"

source_file_path = os.path.join(directory_source, source_file_name)
legacy_source_file_path = os.path.join(directory_source, legacy_source_file_name)
destination_file_path = os.path.join(directory_destination, destination_file_name)
notification_file_path = os.path.join(directory_destination, notification_file_name)

# 历史违约日志文件
breach_list_name = "stop loss log.xlsx"
breach_list_path = os.path.join(directory_breach, breach_list_name)

NOTICE_SHEETS = ("Equity", "High Yield", "Investment Grade")

# Keep new drawdown columns next to the breach fields when present.
PREFERRED_COLUMNS = [
    "Fund Name",
    "Security ID",
    "Security Desc",
    "Issuer",
    "Sec Type",
    "Sec Curr",
    "Mkt Price",
    "Average Cost",
    "Price Benchmark",
    "Initial Mkt Price",
    "High Last Year",
    "YTD Drawdown",
    "Holding Period Drawdown",
    "Drawdown from High",
    "Grade",
    "Limit 1",
    "Limit 2",
    "Breach",
    "Severity",
    "Last Breach Date",
]


def resolve_source_file_path(primary_path=None, legacy_path=None):
    primary = primary_path if primary_path is not None else source_file_path
    legacy = legacy_path if legacy_path is not None else legacy_source_file_path
    candidates = [path for path in (primary, legacy) if path and os.path.exists(path)]
    if not candidates:
        return primary
    return max(candidates, key=os.path.getmtime)


def drop_unnamed_columns(df):
    drop_cols = [col for col in df.columns if str(col).startswith("Unnamed")]
    return df.drop(columns=drop_cols, errors="ignore")


def order_output_columns(df):
    if df is None or df.empty:
        return df
    df = drop_unnamed_columns(df)
    preferred = [col for col in PREFERRED_COLUMNS if col in df.columns]
    rest = [col for col in df.columns if col not in preferred]
    return df[preferred + rest]


def _clean_name_part(value):
    text = str(value).replace("\xa0", " ").strip()
    if text.lower() in ("nan", "none", "<na>", "nat"):
        return ""
    return " ".join(text.split())


def name_key(fund_name, security_desc):
    return _clean_name_part(fund_name) + " | " + _clean_name_part(security_desc)


def is_active_breach(value):
    text = _clean_name_part(value)
    return text not in ("", "No Breach")


def history_matches(history_df, combined_text):
    if history_df is None or history_df.empty or "Fund Name Security Desc" not in history_df.columns:
        return pd.DataFrame()
    keys = history_df["Fund Name Security Desc"].map(_clean_name_part)
    return history_df[keys == _clean_name_part(combined_text)]


def load_breach_sheet(sheet_name: str, log_path=None) -> pd.DataFrame:
    path = log_path if log_path is not None else breach_list_path
    df_raw = pd.read_excel(path, sheet_name=sheet_name)
    df_raw["Fund Name"] = df_raw["Fund Name"].fillna("")
    df_raw["Security Desc"] = df_raw["Security Desc"].fillna("")
    df_out = pd.DataFrame({
        "Date": pd.to_datetime(df_raw["Date"], errors="coerce").ffill(),
        "Fund Name Security Desc": [
            name_key(fund, desc)
            for fund, desc in zip(df_raw["Fund Name"], df_raw["Security Desc"])
        ],
        "Severity": pd.to_numeric(df_raw["Severity"], errors="coerce"),
    })
    df_out["Date"] = pd.to_datetime(df_out["Date"], errors="coerce").dt.strftime("%Y-%m-%d")
    df_out["Severity"] = df_out["Severity"].fillna(0).astype(int)
    return df_out


def filter_today_breaches(source_path, dest_path, sheet_names=NOTICE_SHEETS):
    """Copy today's breached rows, keeping the updated drawdown columns."""
    xl = pd.ExcelFile(source_path)
    os.makedirs(os.path.dirname(dest_path) or ".", exist_ok=True)
    with pd.ExcelWriter(dest_path, engine="openpyxl") as writer:
        for sheet_name in sheet_names:
            if sheet_name not in xl.sheet_names:
                pd.DataFrame().to_excel(writer, sheet_name=sheet_name, index=False)
                continue
            df = pd.read_excel(source_path, sheet_name=sheet_name)
            df = drop_unnamed_columns(df)
            if "Breach" in df.columns:
                df = df[df["Breach"].map(is_active_breach)].copy()
            else:
                df = df.iloc[0:0].copy()
            df = order_output_columns(df)
            df.to_excel(writer, sheet_name=sheet_name, index=False)


def get_new_breach_df(sheet_name: str, history_df: pd.DataFrame, dest_path=None, as_of=None) -> pd.DataFrame:
    path = dest_path if dest_path is not None else destination_file_path
    as_of_ts = pd.Timestamp(as_of) if as_of is not None else today_ts
    df_today = pd.read_excel(path, sheet_name=sheet_name)
    df_today = drop_unnamed_columns(df_today)
    if df_today.empty or "Fund Name" not in df_today.columns:
        return pd.DataFrame()

    df_today["Fund Name"] = df_today["Fund Name"].fillna("")
    if "Security Desc" in df_today.columns:
        df_today["Security Desc"] = df_today["Security Desc"].fillna("")
    else:
        df_today["Security Desc"] = ""
    new_rows = []

    for _, row in df_today.iterrows():
        combined_text = name_key(row["Fund Name"], row["Security Desc"])
        severity_today = pd.to_numeric(row.get("Severity"), errors="coerce")
        if pd.isna(severity_today):
            severity_today = 0
        last_breach_info = "New Breach"
        last_breach_date = None
        should_add = False

        match_rows = history_matches(history_df, combined_text)
        if match_rows.empty:
            # Not in the log at all — always notify.
            should_add = True
        else:
            max_sev = pd.to_numeric(match_rows["Severity"], errors="coerce").max()
            last_dt_str = match_rows["Date"].max()
            last_breach_date = pd.to_datetime(last_dt_str, errors="coerce")
            if pd.isna(max_sev) or pd.isna(last_breach_date):
                should_add = True
            elif severity_today > max_sev:
                should_add = True
            elif (as_of_ts.normalize() - pd.Timestamp(last_breach_date).normalize()).days > 90:
                should_add = True

        if should_add:
            row_dict = row.to_dict()
            if last_breach_date is not None and not pd.isna(last_breach_date):
                row_dict["Last Breach Date"] = pd.Timestamp(last_breach_date).strftime("%Y-%m-%d")
            else:
                row_dict["Last Breach Date"] = last_breach_info
            new_rows.append(row_dict)
    return order_output_columns(pd.DataFrame(new_rows))


def write_notification_workbook(notification_path, sheet_frames):
    os.makedirs(os.path.dirname(notification_path) or ".", exist_ok=True)
    with pd.ExcelWriter(notification_path, engine="openpyxl") as writer:
        for sheet_name, df in sheet_frames.items():
            out = order_output_columns(df) if df is not None else pd.DataFrame()
            if out is None:
                out = pd.DataFrame()
            out.to_excel(writer, sheet_name=sheet_name, index=False)


def main():
    source_path = resolve_source_file_path()
    filter_today_breaches(source_path, destination_file_path)

    equity_list = load_breach_sheet("Equity")
    high_yield_list = load_breach_sheet("High Yield")
    investment_grade_list = load_breach_sheet("Investment Grade")

    df_equity_notice = get_new_breach_df("Equity", equity_list)
    df_hy_notice = get_new_breach_df("High Yield", high_yield_list)
    df_ig_notice = get_new_breach_df("Investment Grade", investment_grade_list)

    write_notification_workbook(
        notification_file_path,
        {
            "Equity": df_equity_notice,
            "High Yield": df_hy_notice,
            "Investment Grade": df_ig_notice,
        },
    )

    print("Source file:", source_path)
    print("Destination file:", destination_file_path)
    print("Notification file:", notification_file_path)
    print("===== Equity New Breach =====")
    print(df_equity_notice)
    print("===== High Yield New Breach =====")
    print(df_hy_notice)
    print("===== Investment Grade New Breach =====")
    print(df_ig_notice)


if __name__ == "__main__":
    main()
