from cmath import nan
from sqlite3 import Row
import pandas as pd
import numpy as np
#import pyfolio
import openpyxl
import csv
import sys
from pandas.tseries.offsets import BDay
import datetime as dt
import holidays
from datetime import datetime, timedelta
import re
import yfinance as yf

OUTPUT_DIR = r"Q:\\Risk Management\\每日 04 Stop_Loss\\Daily Output\\"

# ---------------------------------------------------------------------------
# Stop-loss limits — edit these values only when policy changes.
# Limit 1 is the first (milder) breach; Limit 2 is the stricter breach.
# ---------------------------------------------------------------------------
# Default equity limits: used for any Sec Curr not listed in EQUITY_LIMITS_BY_CURRENCY.
EQUITY_LIMIT_1 = -0.20
EQUITY_LIMIT_2 = -0.30
MUTUAL_FUND_LIMIT_1 = EQUITY_LIMIT_1
MUTUAL_FUND_LIMIT_2 = EQUITY_LIMIT_2
HY_LIMIT_1 = -0.15
HY_LIMIT_2 = -0.25
IG_LIMIT_1 = -0.08
IG_LIMIT_2 = -0.15

# Excluded funds — add or remove fund codes here only.
# These names are never flagged as a stop-loss breach (any asset class).
EXCLUDED_FUNDS = (
    "DCFH2024",
    "TBHTHYEF",
)

# Equity limits by Portia column "Sec Curr".
# Each entry is (Limit 1, Limit 2). Edit one currency without changing the others.
# Unlisted currencies fall back to EQUITY_LIMIT_1 / EQUITY_LIMIT_2.
EQUITY_LIMITS_BY_CURRENCY = {
    "HKD": (-0.20, -0.30),
    "SGD": (-0.20, -0.30),
    "KRW": (-0.20, -0.30),
    "USD": (-0.20, -0.30),
    "EUR": (-0.20, -0.30),
    "JPY": (-0.20, -0.30),
    "CAD": (-0.20, -0.30),
    "GBP": (-0.20, -0.30),
    "CNH": (-0.20, -0.30),
    "AUD": (-0.20, -0.30),
}

BLOOMBERG_EXCHANGE_TO_YAHOO = {
    "US": "",
    "UN": "",
    "UQ": "",
    "UW": "",
    "UA": "",
    "UP": "",
    "UF": "",
    "UV": "",
    "HK": ".HK",
    "TT": ".TW",
    "TW": ".TW",
    "KS": ".KS",
    "KQ": ".KQ",
    "KN": ".KQ",
    "JP": ".T",
    "JT": ".T",
    "LN": ".L",
    "LI": ".L",
    "FP": ".PA",
    "GY": ".DE",
    "GR": ".DE",
    "GF": ".DE",
    "SW": ".SW",
    "AU": ".AX",
    "AT": ".AX",
    "SP": ".SI",
    "CG": ".SS",
    "C1": ".SS",
    "CS": ".SZ",
    "C2": ".SZ",
}


def output_excel_path(today):
    return OUTPUT_DIR + "daily_output_" + today + " test.xlsx"


def _as_frame(df):
    """Return an independent DataFrame so later writes are not chained assignment.

    Pandas 3 Copy-on-Write raises ChainedAssignmentError for patterns such as
    ``df["Limit 1"][i] = -0.20`` or assigning into a filtered view. Always copy
    first, then set columns with ``.loc[:, col] = value``.
    """
    if df is None:
        return df
    return df.copy()


def _to_numeric_columns(df, columns):
    df = _as_frame(df)
    updates = {}
    for column in columns:
        if column in df.columns:
            updates[column] = pd.to_numeric(df[column], errors="coerce")
    return df.assign(**updates) if updates else df


def to_yahoo_ticker(security_id, sec_curr=None):
    """Convert a Portia / Bloomberg-style security ID to a Yahoo Finance ticker."""
    if security_id is None:
        return None
    if isinstance(security_id, float) and np.isnan(security_id):
        return None

    raw = str(security_id).replace("\ufeff", "").strip().strip("'")
    if not raw or raw.lower() in ("nan", "none", "nat"):
        return None

    raw_upper = raw.upper()
    if "." in raw_upper and " " not in raw_upper:
        return raw_upper

    raw_upper = re.sub(r"\s+EQUITY$", "", raw_upper).strip()
    parts = raw_upper.split()
    exch = parts[-1] if len(parts) >= 2 else None
    ticker_body = " ".join(parts[:-1]) if exch is not None else raw_upper

    if exch in BLOOMBERG_EXCHANGE_TO_YAHOO or exch in ("CH", "CN"):
        if exch in ("CH", "CN"):
            digits = re.sub(r"\D", "", ticker_body)
            if not digits:
                return None
            padded = digits.zfill(6)
            if padded.startswith("6") or padded.startswith("9"):
                return padded + ".SS"
            return padded + ".SZ"
        if exch == "HK":
            digits = re.sub(r"\D", "", ticker_body)
            if digits:
                return digits.zfill(4) + ".HK"
            return None
        suffix = BLOOMBERG_EXCHANGE_TO_YAHOO[exch]
        yahoo_body = ticker_body.replace("/", "-")
        if suffix == "":
            yahoo_body = yahoo_body.replace(".", "-")
        return yahoo_body + suffix

    curr = str(sec_curr).upper().strip() if sec_curr is not None and not (isinstance(sec_curr, float) and np.isnan(sec_curr)) else ""
    digits = re.sub(r"\D", "", raw_upper)
    purely_numeric = bool(digits) and re.sub(r"[\s]", "", raw_upper) == digits
    if purely_numeric:
        if curr == "HKD":
            return digits.zfill(4) + ".HK"
        if curr in ("CNY", "CNH"):
            padded = digits.zfill(6)
            if padded.startswith("6") or padded.startswith("9"):
                return padded + ".SS"
            return padded + ".SZ"
        if curr == "JPY":
            return digits + ".T"
        if curr == "KRW":
            return digits + ".KS"
        if curr == "TWD":
            return digits + ".TW"

    return raw_upper.replace("/", "-")


def _safe_max_high(series):
    if series is None:
        return np.nan
    try:
        value = series.max()
    except Exception:
        return np.nan
    return float(value) if pd.notna(value) else np.nan


def _max_high_from_history(hist):
    if hist is None or getattr(hist, "empty", True):
        return np.nan
    if "High" not in getattr(hist, "columns", []):
        return np.nan
    return _safe_max_high(hist["High"])


def _max_high_from_download(data, ticker, single=False):
    if data is None or getattr(data, "empty", True):
        return np.nan

    if single and not isinstance(data.columns, pd.MultiIndex) and "High" in data.columns:
        return _max_high_from_history(data)

    if isinstance(data.columns, pd.MultiIndex):
        try:
            return _safe_max_high(data[ticker]["High"])
        except Exception:
            pass
        try:
            return _safe_max_high(data["High"][ticker])
        except Exception:
            pass
        ticker_upper = str(ticker).upper()
        for col in data.columns:
            parts = [str(x).upper() for x in (col if isinstance(col, tuple) else (col,))]
            if ticker_upper in parts and "HIGH" in parts:
                return _safe_max_high(data[col])
        return np.nan

    if "High" in data.columns:
        return _max_high_from_history(data)
    return np.nan


LONDON_YAHOO_SUFFIXES = (".L", ".IL")
PENCE_CURRENCIES = {"GBp", "GBX", "GBx"}


def is_pence_currency(currency):
    """Yahoo uses GBp/GBX for sterling pence; GBP would mean pounds."""
    if currency is None:
        return False
    text = str(currency).strip()
    return text in PENCE_CURRENCIES or text.upper() == "GBX"


def fetch_yahoo_currency(yahoo_ticker):
    try:
        return yf.Ticker(yahoo_ticker).fast_info.get("currency")
    except Exception:
        return None


def yahoo_price_scale(yahoo_ticker, currency=None):
    """Convert Yahoo minor units to the major currency (GBp pence -> GBP pounds)."""
    if is_pence_currency(currency):
        return 0.01
    ticker_upper = str(yahoo_ticker or "").upper()
    if currency is None and ticker_upper.endswith(LONDON_YAHOO_SUFFIXES):
        return 0.01
    return 1.0


def _maybe_fetch_currency(yahoo_ticker):
    ticker_upper = str(yahoo_ticker or "").upper()
    if ticker_upper.endswith(LONDON_YAHOO_SUFFIXES):
        return fetch_yahoo_currency(yahoo_ticker)
    return None


def apply_yahoo_price_scale(highs):
    scaled = {}
    for ticker, high in highs.items():
        if high is None or (isinstance(high, float) and np.isnan(high)) or pd.isna(high):
            scaled[ticker] = high
            continue
        currency = _maybe_fetch_currency(ticker)
        scaled[ticker] = float(high) * yahoo_price_scale(ticker, currency)
    return scaled


def fetch_last_year_highs(yahoo_tickers):
    """Return a dict of Yahoo ticker -> raw (unadjusted) max High over the trailing year.

    Prices are not adjusted for stock splits or dividends. London quotes in GBp
    (pence) are converted to GBP (pounds), e.g. AZN.L 15732 GBp -> 157.32 GBP.
    """
    highs = {ticker: np.nan for ticker in yahoo_tickers}
    tickers = [ticker for ticker in yahoo_tickers if ticker]
    if not tickers:
        return highs

    try:
        data = yf.download(
            tickers=tickers,
            period="1y",
            interval="1d",
            auto_adjust=False,
            progress=False,
            threads=True,
            group_by="ticker",
        )
        single = len(tickers) == 1
        for ticker in tickers:
            highs[ticker] = _max_high_from_download(data, ticker, single=single)
    except Exception as exc:
        print("Yahoo Finance batch download failed:", exc)

    missing = [ticker for ticker in tickers if pd.isna(highs.get(ticker))]
    for ticker in missing:
        try:
            hist = yf.Ticker(ticker).history(period="1y", auto_adjust=False, interval="1d")
            highs[ticker] = _max_high_from_history(hist)
        except Exception as exc:
            print("Yahoo Finance history failed for " + str(ticker) + ":", exc)
    return apply_yahoo_price_scale(highs)


def add_high_last_year(df):
    """Add High Last Year from Yahoo Finance for each equity ticker."""
    df = _as_frame(df)
    if df.empty:
        df.loc[:, "High Last Year"] = np.nan
        return df

    if "Sec Curr" in df.columns:
        sec_curr = df["Sec Curr"]
    else:
        sec_curr = pd.Series([None] * len(df), index=df.index)

    yahoo_tickers = [
        to_yahoo_ticker(security_id, currency)
        for security_id, currency in zip(df["Security ID"], sec_curr)
    ]
    unique_tickers = sorted({ticker for ticker in yahoo_tickers if ticker})
    print("Fetching last-year highs from Yahoo Finance for " + str(len(unique_tickers)) + " unique ticker(s)...")
    high_map = fetch_last_year_highs(unique_tickers)
    df.loc[:, "High Last Year"] = [
        high_map.get(ticker, np.nan) if ticker else np.nan
        for ticker in yahoo_tickers
    ]
    missing = sorted({ticker for ticker, high in zip(yahoo_tickers, df["High Last Year"]) if ticker and pd.isna(high)})
    if missing:
        preview = ", ".join(missing[:30])
        if len(missing) > 30:
            preview += ", ..."
        print("Could not fetch last-year high for:", preview)
    return df


def add_ytd_drawdown(df):
    df = _as_frame(df)
    df.loc[:, "YTD Drawdown"] = (df["Mkt Price"] - df["Price Benchmark"]) / df["Price Benchmark"]
    return df


def add_holding_period_drawdown(df):
    df = _as_frame(df)
    average_cost = df["Average Cost"].replace(0, np.nan)
    df.loc[:, "Holding Period Drawdown"] = (df["Mkt Price"] - df["Average Cost"]) / average_cost
    return df


def add_drawdown_from_high(df):
    df = _as_frame(df)
    high_last_year = df["High Last Year"].replace(0, np.nan)
    df.loc[:, "Drawdown from High"] = (df["Mkt Price"] - df["High Last Year"]) / high_last_year
    return df


BREACH_DRAWDOWN_COLUMN = "Holding Period Drawdown"


def classify_two_limit_breach(drawdown, limit_1, limit_2):
    """Classify Limit 1 / Limit 2 using holding-period drawdown vs the given thresholds."""
    try:
        dd = float(drawdown)
    except (TypeError, ValueError):
        return "No Breach"
    if pd.isna(dd):
        return "No Breach"
    if dd <= limit_1 and dd > limit_2:
        return "Breach Limit 1"
    if dd <= limit_2:
        return "Breach Limit 2"
    return "No Breach"


def apply_excluded_fund_breaches(df):
    """Force No Breach for any fund in EXCLUDED_FUNDS, regardless of asset class."""
    if df is None or df.empty or "Fund Name" not in df.columns or "Breach" not in df.columns:
        return df
    df = _as_frame(df)
    excluded = df["Fund Name"].astype(str).isin(EXCLUDED_FUNDS)
    df.loc[excluded, "Breach"] = "No Breach"
    return df


def _holding_period_drawdown_series(df):
    if "Holding Period Drawdown" not in df.columns:
        raise KeyError("Holding Period Drawdown is required to assign Breach; YTD Drawdown is not used")
    return pd.to_numeric(df["Holding Period Drawdown"], errors="coerce")


def equity_limits_for_currency(sec_curr):
    """Return (Limit 1, Limit 2) for a Portia Sec Curr value."""
    if sec_curr is None or (isinstance(sec_curr, float) and np.isnan(sec_curr)):
        return (EQUITY_LIMIT_1, EQUITY_LIMIT_2)
    curr = str(sec_curr).replace("\xa0", " ").strip().upper()
    if not curr or curr in ("NAN", "NONE", "<NA>"):
        return (EQUITY_LIMIT_1, EQUITY_LIMIT_2)
    return EQUITY_LIMITS_BY_CURRENCY.get(curr, (EQUITY_LIMIT_1, EQUITY_LIMIT_2))


def assign_two_limit_breaches(df, limit_1, limit_2, extra_limit_2_mask=None):
    """Assign Limit 1/2 and Breach from Holding Period Drawdown only. Never uses YTD Drawdown."""
    df = _as_frame(df)
    dd = _holding_period_drawdown_series(df)
    df.loc[:, "Limit 1"] = limit_1
    df.loc[:, "Limit 2"] = limit_2
    breach = pd.Series("No Breach", index=df.index)
    limit_2_mask = dd <= limit_2
    if extra_limit_2_mask is not None:
        extra = extra_limit_2_mask.reindex(df.index)
        extra = extra.fillna(False).astype(bool)
        limit_2_mask = limit_2_mask & extra
    breach = breach.mask((dd <= limit_1) & (dd > limit_2), "Breach Limit 1")
    breach = breach.mask(limit_2_mask, "Breach Limit 2")
    df.loc[:, "Breach"] = breach
    return df


def assign_equity_breaches(df):
    """Assign equity Limit 1/2 and Breach using Sec Curr-specific thresholds."""
    df = _as_frame(df)
    dd = _holding_period_drawdown_series(df)
    if "Sec Curr" in df.columns:
        limits = df["Sec Curr"].map(equity_limits_for_currency)
    else:
        limits = pd.Series([(EQUITY_LIMIT_1, EQUITY_LIMIT_2)] * len(df), index=df.index)
    df.loc[:, "Limit 1"] = limits.map(lambda pair: pair[0])
    df.loc[:, "Limit 2"] = limits.map(lambda pair: pair[1])
    breach = pd.Series("No Breach", index=df.index)
    breach = breach.mask((dd <= df["Limit 1"]) & (dd > df["Limit 2"]), "Breach Limit 1")
    breach = breach.mask(dd <= df["Limit 2"], "Breach Limit 2")
    df.loc[:, "Breach"] = breach
    return apply_excluded_fund_breaches(df)


def assign_fi_breaches(df):
    df = _as_frame(df)
    dd = _holding_period_drawdown_series(df)
    hy = df["Grade"].astype(str) == "HY"
    ig = df["Grade"].astype(str) == "IG"

    df.loc[:, "Limit 1"] = np.where(hy, HY_LIMIT_1, IG_LIMIT_1)
    df.loc[:, "Limit 2"] = np.where(hy, HY_LIMIT_2, IG_LIMIT_2)

    breach = pd.Series("No Breach", index=df.index)
    breach = breach.mask(hy & (dd <= HY_LIMIT_1) & (dd > HY_LIMIT_2), "Breach Limit 1")
    breach = breach.mask(hy & (dd <= HY_LIMIT_2), "Breach Limit 2")
    breach = breach.mask(ig & (dd <= IG_LIMIT_1) & (dd > IG_LIMIT_2), "Breach Limit 1")
    breach = breach.mask(ig & (dd <= IG_LIMIT_2), "Breach Limit 2")
    df.loc[:, "Breach"] = breach
    return apply_excluded_fund_breaches(df)


def assign_mutualfund_breaches(df):
    df = _as_frame(df)
    extra_limit_2 = pd.to_numeric(df["Average Cost"], errors="coerce") > pd.to_numeric(df["Mkt Price"], errors="coerce")
    df = assign_two_limit_breaches(df, MUTUAL_FUND_LIMIT_1, MUTUAL_FUND_LIMIT_2, extra_limit_2_mask=extra_limit_2)
    return apply_excluded_fund_breaches(df)


def assign_severity(df):
    """Set Severity from Breach text: 'Breach Limit 1' -> 1, 'Breach Limit 2' -> 2."""
    df = _as_frame(df)
    severity = df["Breach"].astype(str).str.extract(r"(\d+)\s*$", expand=False)
    df.loc[:, "Severity"] = severity.fillna("")
    return df


def add_equity_drawdown_columns(df):
    df = add_high_last_year(df)
    df = add_ytd_drawdown(df)
    df = add_holding_period_drawdown(df)
    df = add_drawdown_from_high(df)
    return df


def add_non_equity_drawdown_columns(df):
    df = add_ytd_drawdown(df)
    df = add_holding_period_drawdown(df)
    return df


def clean_portia_columns(portia_data):
    portia_data = _as_frame(portia_data)
    portia_data.columns = [str(column).replace("\ufeff", "").replace("\xef\xbb\xbf", "").strip() for column in portia_data.columns]
    if "Security ID" in portia_data.columns:
        stripped_ids = portia_data["Security ID"].map(
            lambda x: x[1:] if isinstance(x, str) and x.startswith("'") else x
        )
        portia_data.loc[:, "Security ID"] = stripped_ids
    return portia_data


def _filter_sec_type_equals(df, value):
    return df.loc[df["Sec Type"] == value].copy()


def _filter_sec_type_contains(df, pattern):
    mask = df["Sec Type"].str.contains(pattern, case=False).fillna(False)
    return df.loc[mask].copy()


def _initial_price_chunk(df):
    return df.loc[:, ["Fund Name", "Security ID", "Mkt Price"]].rename(
        columns={"Mkt Price": "Initial Mkt Price"}
    )


def _add_price_benchmark(merged):
    merged = _as_frame(merged)
    benchmark = merged["Initial Mkt Price"].fillna(merged["Average Cost"])
    merged = merged.assign(**{"Price Benchmark": benchmark})
    return _to_numeric_columns(merged, ["Price Benchmark", "Mkt Price", "Average Cost"])


############################## Data Cleaning ############################################# add DC on 2024 Dec to fulfill SFC requirement
def data_cleaning(portia_data_initial, portia_data_last):
    # Filter for Equity. .copy() is required under pandas Copy-on-Write.
    portia_data_initial_equity = _filter_sec_type_equals(portia_data_initial, "Common Stock")
    portia_data_last_equity = _filter_sec_type_equals(portia_data_last, "Common Stock")

    # Filter for Fixed Income
    portia_data_initial_fi = _filter_sec_type_contains(portia_data_initial, "bond")
    portia_data_last_fi = _filter_sec_type_contains(portia_data_last, "bond")
    #print(portia_data_last_fi)

    # Filter for Mutual Fund
    portia_data_initial_mutualfund = _filter_sec_type_equals(portia_data_initial, "Mutual Fund")
    portia_data_last_mutualfund = _filter_sec_type_equals(portia_data_last, "Mutual Fund")
    #print (portia_data_last_mutualfund)


    # Need to filter 20220419 Fund Name, Security ID, Market Price
    portia_data_initial_equity_chunk = _initial_price_chunk(portia_data_initial_equity)
    portia_data_initial_fi_chunk = _initial_price_chunk(portia_data_initial_fi)
    portia_data_initial_mutualfund_chunk = _initial_price_chunk(portia_data_initial_mutualfund)

    ########################## left join
    portia_data_equity_merge = portia_data_last_equity.merge(portia_data_initial_equity_chunk, how = "left", on = ["Fund Name", "Security ID"])
    portia_data_fi_merge = portia_data_last_fi.merge(portia_data_initial_fi_chunk, how = "left", on = ["Fund Name", "Security ID"])
    portia_data_mutualfund_merge = portia_data_last_mutualfund.merge(portia_data_initial_mutualfund_chunk, how = "left", on = ["Fund Name", "Security ID"])
    #print(portia_data_equity_merge.head(6))

    
########### Replace Initial Mkt Price with Average Cost, changed on 2024/12/11 due to update of Operation Procedure
    portia_data_equity_merge = _add_price_benchmark(portia_data_equity_merge)
    portia_data_fi_merge = _add_price_benchmark(portia_data_fi_merge)
    portia_data_mutualfund_merge = _add_price_benchmark(portia_data_mutualfund_merge)

    return portia_data_equity_merge, portia_data_fi_merge,portia_data_mutualfund_merge


def main():
    # today_temp = datetime.today()
    # today = datetime.today().strftime("%d%m%Y")

    # last_business_day_temp = today_temp - BDay(1)
    # last_business_day = last_business_day_temp.strftime("%d%m%Y")

    current_year = dt.date.today().year

    #set the start date to the first day of the current year
    start_date=dt.date(current_year,1,1)

    end_date = dt.date.today()

    # Generate the list of Hong Kong holidays
    hk_holidays = holidays.HK(years=[start_date.year, end_date.year])

    # Generate the list of working days
    working_days = [day for day in (start_date + dt.timedelta(days=i) for i in range((end_date - start_date).days + 1)) if day.weekday() < 5 and day not in hk_holidays]
    #print (working_days)
    # Find the last working day before today
    last_business_day = max(day for day in working_days if day < dt.date.today())
    last_business_day = (last_business_day).strftime("%Y%m%d")
    today = dt.date.today().strftime("%d%m%Y")

    # The benchmark date
    initial_date = "02012026"
    #initial_date = "31122021"

    # Check the date
    #print("Date Today:", today, ". Last Bussiness Day:", last_business_day, ". Initial Date:", initial_date)

    portia_report_path = "Q:\\Risk Management\\Daily Portia Report\\"
    portia_file_name_initial = "LiqTrackingRpt2_" + initial_date
    portia_file_initial = portia_report_path + portia_file_name_initial + ".csv"

    # Define path of Portia report of last buz day
    portia_file_last = portia_report_path + "LiqTrackingRpt_" + last_business_day + "_wind.csv"

    # Check the path of initial Portia file
    #print(portia_file_initial)
    #print(portia_file_last)

    # Load the initial Portia file
    portia_data_initial = pd.read_csv(portia_file_initial, encoding = 'unicode_escape', skiprows = 3)
    portia_data_initial = clean_portia_columns(portia_data_initial)
    #print(portia_data_initial.head(5))

    # Load the T-1 data
    # New T-1 format already starts from the header row, so do not skip the first three rows.
    portia_data_last = pd.read_csv(portia_file_last, encoding = 'unicode_escape')
    portia_data_last = clean_portia_columns(portia_data_last)
    #print(portia_data_last.head(5))

    #print (portia_data_last)

    portia_data_initial_mutualfund = _filter_sec_type_equals(portia_data_initial, "Mutual Fund")
    portia_data_last_mutualfund = _filter_sec_type_equals(portia_data_last, "Mutual Fund")

    #print (portia_data_initial_mutualfund)


    # Load Fund Manager Data, PAY ATTENTION to the fund manager list! The format may be changed
    # fund_manager_data_path = r"Q:\Risk Management\Fund Manager List\List of Primary and Secondary Fund Manager 20210628 onwards.xlsb"
    # fund_manager_data = pd.read_excel(fund_manager_data_path, sheet_name = "20220511", skiprows = 1).iloc[np.arange(51), [0, 2, 3]].rename(columns = {"Fund Code": "Fund Name"})

    #print (data_cleaning(portia_data_initial, portia_data_last)[2])


    ##################### Start to Calculate Equity Drawdowns #######################################
    portia_data_equity_merge = data_cleaning(portia_data_initial, portia_data_last)[0]
    portia_data_equity_merge = add_equity_drawdown_columns(portia_data_equity_merge)

    #print(portia_data_equity_merge.head(6))

    ##################### Start to Calculate Fixed Income Drawdowns #################################
    portia_data_fi_merge = data_cleaning(portia_data_initial, portia_data_last)[1]
    portia_data_fi_merge = add_non_equity_drawdown_columns(portia_data_fi_merge)
    portia_data_fi_merge.insert(loc=3, column='Issuer', value=portia_data_fi_merge['Security Desc'].str.split(' ', n=1, expand=True)[0])


    ##################### Identify IG/HY ############################################################
    bond_grades = pd.read_excel(r"Q:\Risk Management\每日 04 Stop_Loss\IG List.xlsx", sheet_name = "HY")
    bond_grades = bond_grades[["Original ticker", "IG or HY"]].rename(columns = {"Original ticker": "Security ID", "IG or HY": "Grade"})

    print (bond_grades)

    portia_data_fi_merge = portia_data_fi_merge.merge(bond_grades, how = "left", on = "Security ID")


    # Here need to be extremely CAREFUL, since the default is IG, it's more conservative to assmue bond without rating is high yield, which has strict limit
    portia_data_fi_merge.loc[:, "Grade"] = portia_data_fi_merge["Grade"].fillna("IG")

    #print(portia_data_fi_merge)

    ##################### Start to Calculate Mutual Fund Drawdowns #######################################
    portia_data_mutualfund_merge = data_cleaning(portia_data_initial, portia_data_last)[2]
    portia_data_mutualfund_merge = add_non_equity_drawdown_columns(portia_data_mutualfund_merge)

    ##################### Set Limits ################################################################
    # Breach / Severity use Holding Period Drawdown only. YTD Drawdown is kept as a
    # reporting column and is never used to decide Limit 1 / Limit 2.
    portia_data_equity_merge = assign_equity_breaches(portia_data_equity_merge)
              
    ################################# IF MARKET SITUATION CHANGE, UPDATE HERE########################################          
      # if portia_data_equity_merge["Sec Curr"][i]=="KRW": #or portia_data_equity_merge["Sec"][i] == "CNY":
      #      portia_data_equity_merge["Limit 1"][i]=-0.28
      #      portia_data_equity_merge["Limit 2"][i]=-0.38  
      #      if portia_data_equity_merge["Holding Period Drawdown"][i] <= -0.28 and portia_data_equity_merge["Holding Period Drawdown"][i] > -0.38:
      #            portia_data_equity_merge.loc[i, "Breach"] = "Breach Limit 1"
      #      elif portia_data_equity_merge["Holding Period Drawdown"][i] <= -0.38:
      #            portia_data_equity_merge.loc[i, "Breach"] = "Breach Limit 2"
      #      else:
      #            portia_data_equity_merge.loc[i, "Breach"] = "No Breach"         
          
    portia_data_equity_merge = portia_data_equity_merge.sort_values("Breach")

    #print(portia_data_equity_merge.head(6))

    # For Fixed Income (Holding Period Drawdown; EXCLUDED_FUNDS skip breach)
    portia_data_fi_merge = assign_fi_breaches(portia_data_fi_merge)
    portia_data_equity_merge = assign_severity(portia_data_equity_merge)
    portia_data_fi_merge = assign_severity(portia_data_fi_merge)

    ##portia_data_equity_merge['Severity'] = portia_data_equity_merge['Severity'].astype(int)

    # Sort by Grade
    portia_data_fi_merge = portia_data_fi_merge.sort_values("Grade")
    # Seperate by investment grade
    portia_data_hy_merge = portia_data_fi_merge.loc[portia_data_fi_merge["Grade"] == "HY"].sort_values("Breach").copy()
    portia_data_ig_merge = portia_data_fi_merge.loc[portia_data_fi_merge["Grade"] == "IG"].sort_values("Breach").copy()

    ######################################## For Mutual Fund################################################################
    # Breach uses Holding Period Drawdown, not YTD Drawdown.
    portia_data_mutualfund_merge = assign_mutualfund_breaches(portia_data_mutualfund_merge)





    #################### Merge with Fund Manager#########################

    # portia_data_equity_merge = portia_data_equity_merge.merge(fund_manager_data, how = "left", on = ["Fund Name"])
    # portia_data_hy_merge = portia_data_hy_merge.merge(fund_manager_data, how = "left", on = ["Fund Name"])
    # portia_data_ig_merge = portia_data_ig_merge.merge(fund_manager_data, how = "left", on = ["Fund Name"])

    # Output Excel



    output_path = output_excel_path(today)
    with pd.ExcelWriter(output_path) as writer:
        portia_data_equity_merge.to_excel(writer, sheet_name='Equity', index=False)
        portia_data_hy_merge.to_excel(writer, sheet_name='High Yield', index=False)
        portia_data_ig_merge.to_excel(writer, sheet_name='Investment Grade', index=False)
        portia_data_mutualfund_merge.to_excel(writer, sheet_name='Mutual Fund', index=False)
    print("Wrote", output_path)


if __name__ == "__main__":
    main()
