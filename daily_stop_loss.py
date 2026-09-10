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


def fetch_last_year_highs(yahoo_tickers):
    """Return a dict of Yahoo ticker -> max High over the trailing year."""
    highs = {ticker: np.nan for ticker in yahoo_tickers}
    tickers = [ticker for ticker in yahoo_tickers if ticker]
    if not tickers:
        return highs

    try:
        data = yf.download(
            tickers=tickers,
            period="1y",
            interval="1d",
            auto_adjust=True,
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
            hist = yf.Ticker(ticker).history(period="1y", auto_adjust=True, interval="1d")
            highs[ticker] = _max_high_from_history(hist)
        except Exception as exc:
            print("Yahoo Finance history failed for " + str(ticker) + ":", exc)
    return highs


def add_high_last_year(df):
    """Add High Last Year from Yahoo Finance for each equity ticker."""
    if df.empty:
        df["High Last Year"] = pd.Series(dtype=float)
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
    df["High Last Year"] = [
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
    df["YTD Drawdown"] = (df["Mkt Price"] - df["Price Benchmark"]) / df["Price Benchmark"]
    return df


def add_holding_period_drawdown(df):
    average_cost = df["Average Cost"].replace(0, np.nan)
    df["Holding Period Drawdown"] = (df["Mkt Price"] - df["Average Cost"]) / average_cost
    return df


def add_drawdown_from_high(df):
    high_last_year = df["High Last Year"].replace(0, np.nan)
    df["Drawdown from High"] = (df["Mkt Price"] - df["High Last Year"]) / high_last_year
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
    portia_data.columns = [str(column).replace("\ufeff", "").replace("\xef\xbb\xbf", "").strip() for column in portia_data.columns]
    if "Security ID" in portia_data.columns:
        portia_data["Security ID"] = portia_data["Security ID"].apply(
            lambda x: x[1:] if isinstance(x, str) and x.startswith("'") else x
        )
    return portia_data


############################## Data Cleaning ############################################# add DC on 2024 Dec to fulfill SFC requirement
def data_cleaning(portia_data_initial, portia_data_last):
    # Filter for Equity
    portia_data_initial_equity = portia_data_initial[portia_data_initial["Sec Type"] == "Common Stock"]
    portia_data_last_equity = portia_data_last[portia_data_last["Sec Type"] == "Common Stock"]

    # Filter for Fixed Income
    portia_data_initial_fi = portia_data_initial.loc[portia_data_initial["Sec Type"].str.contains("bond", case=False).fillna(False)]
    portia_data_last_fi = portia_data_last.loc[portia_data_last["Sec Type"].str.contains("bond", case=False).fillna(False)]
    #print(portia_data_last_fi)

    # Filter for Mutual Fund
    portia_data_initial_mutualfund = portia_data_initial[portia_data_initial["Sec Type"] == "Mutual Fund"]
    portia_data_last_mutualfund = portia_data_last[portia_data_last["Sec Type"] == "Mutual Fund"]
    #print (portia_data_last_mutualfund)


    # Need to filter 20220419 Fund Name, Security ID, Market Price
    portia_data_initial_equity_chunk = portia_data_initial_equity[["Fund Name", "Security ID", "Mkt Price"]].rename(columns={"Mkt Price": "Initial Mkt Price"})
    portia_data_initial_fi_chunk = portia_data_initial_fi[["Fund Name", "Security ID", "Mkt Price"]].rename(columns={"Mkt Price": "Initial Mkt Price"})
    portia_data_initial_mutualfund_chunk = portia_data_initial_mutualfund[["Fund Name", "Security ID", "Mkt Price"]].rename(columns={"Mkt Price": "Initial Mkt Price"})

    ########################## left join
    portia_data_equity_merge = portia_data_last_equity.merge(portia_data_initial_equity_chunk, how = "left", on = ["Fund Name", "Security ID"])
    portia_data_fi_merge = portia_data_last_fi.merge(portia_data_initial_fi_chunk, how = "left", on = ["Fund Name", "Security ID"])
    portia_data_mutualfund_merge = portia_data_last_mutualfund.merge(portia_data_initial_mutualfund_chunk, how = "left", on = ["Fund Name", "Security ID"])
    #print(portia_data_equity_merge.head(6))

    
########### Replace Initial Mkt Price with Average Cost, changed on 2024/12/11 due to update of Operation Procedure
    portia_data_equity_merge["Price Benchmark"] = portia_data_equity_merge["Initial Mkt Price"].fillna(portia_data_equity_merge["Average Cost"])
    #portia_data_equity_merge["Price Benchmark"] = portia_data_equity_merge["Average Cost"]    
    portia_data_fi_merge["Price Benchmark"] = portia_data_fi_merge["Initial Mkt Price"].fillna(portia_data_fi_merge["Average Cost"])
    #portia_data_fi_merge["Price Benchmark"] = portia_data_fi_merge["Average Cost"]
    portia_data_mutualfund_merge["Price Benchmark"] = portia_data_mutualfund_merge["Initial Mkt Price"].fillna(portia_data_mutualfund_merge["Average Cost"])
    #portia_data_mutualfund_merge["Price Benchmark"] = portia_data_mutualfund_merge["Average Cost"]
    #print(portia_data_equity_merge.head(6))

    #print(type(portia_data_equity_merge["Price Benchmark"][2]))

    ##### Change numbers needed into NUMBERS
    portia_data_equity_merge["Price Benchmark"] = portia_data_equity_merge["Price Benchmark"].apply(lambda x: float(x))
    portia_data_equity_merge["Mkt Price"] = portia_data_equity_merge["Mkt Price"].apply(lambda x: float(x))
    portia_data_equity_merge["Average Cost"] = portia_data_equity_merge["Average Cost"].apply(lambda x: float(x))

    portia_data_fi_merge["Price Benchmark"] = portia_data_fi_merge["Price Benchmark"].apply(lambda x: float(x))
    portia_data_fi_merge["Mkt Price"] = portia_data_fi_merge["Mkt Price"].apply(lambda x: float(x))
    portia_data_fi_merge["Average Cost"] = portia_data_fi_merge["Average Cost"].apply(lambda x: float(x))
    
    portia_data_mutualfund_merge["Price Benchmark"] = portia_data_mutualfund_merge["Price Benchmark"].apply(lambda x: float(x))
    portia_data_mutualfund_merge["Mkt Price"] = portia_data_mutualfund_merge["Mkt Price"].apply(lambda x: float(x))
    portia_data_mutualfund_merge["Average Cost"] = portia_data_mutualfund_merge["Average Cost"].apply(lambda x: float(x))
           
    
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

    portia_data_initial_mutualfund = portia_data_initial[portia_data_initial["Sec Type"] == "Mutual Fund"]
    portia_data_last_mutualfund = portia_data_last[portia_data_last["Sec Type"] == "Mutual Fund"]

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
    portia_data_fi_merge = portia_data_fi_merge[portia_data_fi_merge["Fund Name"] != "DCFH2024"].reset_index(drop=True)
    portia_data_fi_merge = add_non_equity_drawdown_columns(portia_data_fi_merge)
    portia_data_fi_merge.insert(loc=3, column='Issuer', value=portia_data_fi_merge['Security Desc'].str.split(' ', n=1, expand=True)[0])


    ##################### Identify IG/HY ############################################################
    bond_grades = pd.read_excel(r"Q:\Risk Management\每日 04 Stop_Loss\IG List.xlsx", sheet_name = "HY")
    bond_grades = bond_grades[["Original ticker", "IG or HY"]].rename(columns = {"Original ticker": "Security ID", "IG or HY": "Grade"})

    print (bond_grades)

    portia_data_fi_merge = portia_data_fi_merge.merge(bond_grades, how = "left", on = "Security ID")


    # Here need to be extremely CAREFUL, since the default is IG, it's more conservative to assmue bond without rating is high yield, which has strict limit
    portia_data_fi_merge["Grade"] = portia_data_fi_merge["Grade"].fillna("IG")

    #print(portia_data_fi_merge)

    ##################### Start to Calculate Mutual Fund Drawdowns #######################################
    portia_data_mutualfund_merge = data_cleaning(portia_data_initial, portia_data_last)[2]
    portia_data_mutualfund_merge = add_non_equity_drawdown_columns(portia_data_mutualfund_merge)

    ##################### Set Limits ################################################################
    # For Equity

    portia_data_equity_merge[["Limit 1", "Limit 2"]] = [-0.20, -0.30]
    # Pay Attention to the deault value
    portia_data_equity_merge["Breach"] = "No Breach"

    #for i in range(0,len(portia_data_equity_merge.index)):
        
    # if portia_data_equity_merge["Sec"][i]=="CNY" or portia_data_equity_merge["Sec"][i]=="HKD":
    #  portia_data_equity_merge[["Limit 1"]][i] = [-0.2]
    #  portia_data_equity_merge[["Limit 2"]][i] = [-0.3]

    for i in range(0,len(portia_data_equity_merge.index)):
      # if portia_data_equity_merge["Sec Curr"][i]!="KRW": #or portia_data_equity_merge["Sec"][i] != "CNY":  IF MARKET SITUATION CHANGE, UPDATE HERE！！！########
      #     portia_data_equity_merge["Limit 1"][i]=-0.20
      #     portia_data_equity_merge["Limit 2"][i]=-0.30
          if portia_data_equity_merge["YTD Drawdown"][i] <= -0.20 and portia_data_equity_merge["YTD Drawdown"][i] > -0.30:
              portia_data_equity_merge.loc[i, "Breach"] = "Breach Limit 1"
          elif portia_data_equity_merge["YTD Drawdown"][i] <= -0.30: #and portia_data_equity_merge["Average Cost"][i] > portia_data_equity_merge["Mkt Price"][i]:  # add another condition to exclude the stock split 
              portia_data_equity_merge.loc[i, "Breach"] = "Breach Limit 2"
          else:
              portia_data_equity_merge.loc[i, "Breach"] = "No Breach"
              
    ################################# IF MARKET SITUATION CHANGE, UPDATE HERE########################################          
      # if portia_data_equity_merge["Sec Curr"][i]=="KRW": #or portia_data_equity_merge["Sec"][i] == "CNY":
      #      portia_data_equity_merge["Limit 1"][i]=-0.28
      #      portia_data_equity_merge["Limit 2"][i]=-0.38  
      #      if portia_data_equity_merge["YTD Drawdown"][i] <= -0.28 and portia_data_equity_merge["YTD Drawdown"][i] > -0.38:
      #            portia_data_equity_merge.loc[i, "Breach"] = "Breach Limit 1"
      #      elif portia_data_equity_merge["YTD Drawdown"][i] <= -0.38:
      #            portia_data_equity_merge.loc[i, "Breach"] = "Breach Limit 2"
      #      else:
      #            portia_data_equity_merge.loc[i, "Breach"] = "No Breach"         
          
    portia_data_equity_merge = portia_data_equity_merge.sort_values("Breach")

    #print(portia_data_equity_merge.head(6))

    # For Fixed Income
    # Set different Limits
    for i in range(0,len(portia_data_fi_merge)):
        if portia_data_fi_merge["Grade"][i] == "HY":
            portia_data_fi_merge.loc[i, "Limit 1"] = -0.15
            portia_data_fi_merge.loc[i, "Limit 2"] = -0.25
        else:
            portia_data_fi_merge.loc[i, "Limit 1"] = -0.08
            portia_data_fi_merge.loc[i, "Limit 2"] = -0.15

    # exclude TBHTHYEF
    for i in range(0,len(portia_data_fi_merge)):
        if portia_data_fi_merge["Grade"][i] == "HY" and portia_data_fi_merge["Fund Name"][i] != "TBHTHYEF" and portia_data_fi_merge["YTD Drawdown"][i] <= -0.15 and portia_data_fi_merge["YTD Drawdown"][i] > -0.25: #and portia_data_fi_merge["Average Cost"][i] > portia_data_fi_merge["Mkt Price"][i]:
            portia_data_fi_merge.loc[i, "Breach"] = "Breach Limit 1"
        elif portia_data_fi_merge["Grade"][i] == "HY" and portia_data_fi_merge["Fund Name"][i] != "TBHTHYEF" and  portia_data_fi_merge["YTD Drawdown"][i] <= -0.25: #and portia_data_fi_merge["Average Cost"][i] > portia_data_fi_merge["Mkt Price"][i]:
            portia_data_fi_merge.loc[i, "Breach"] = "Breach Limit 2"
        elif portia_data_fi_merge["Grade"][i] == "HY":
            portia_data_fi_merge.loc[i, "Breach"] = "No Breach"
        elif portia_data_fi_merge["Grade"][i] == "IG" and portia_data_fi_merge["YTD Drawdown"][i] <= -0.08 and portia_data_fi_merge["YTD Drawdown"][i] > -0.15: #and portia_data_fi_merge["Average Cost"][i] > portia_data_fi_merge["Mkt Price"][i]:
            portia_data_fi_merge.loc[i, "Breach"] = "Breach Limit 1"
        elif portia_data_fi_merge["Grade"][i] == "IG" and portia_data_fi_merge["YTD Drawdown"][i] <= -0.15: #and portia_data_fi_merge["Average Cost"][i] > portia_data_fi_merge["Mkt Price"][i]:
            portia_data_fi_merge.loc[i, "Breach"] = "Breach Limit 2"
        elif portia_data_fi_merge["Grade"][i] == "IG":
            portia_data_fi_merge.loc[i, "Breach"] = "No Breach"
            

    #### add a new column called Severity #########
    portia_data_equity_merge['Severity'] = ''
    for index, row in portia_data_equity_merge.iterrows():
        if row['Breach'] != 'No Breach':
            first_string_from_right = row['Breach'].rsplit(maxsplit=1)[-1]
            first_string_from_right = int(first_string_from_right)
            portia_data_equity_merge.at[index, 'Severity'] =str(first_string_from_right)
            

    #### add a new column called Severity #########
    portia_data_fi_merge['Severity'] = ''
    for index, row in portia_data_fi_merge.iterrows():
        if row['Breach'] != 'No Breach':
            first_string_from_right = row['Breach'].rsplit(maxsplit=1)[-1]
            first_string_from_right = int(first_string_from_right)
            portia_data_fi_merge.at[index, 'Severity'] = str(first_string_from_right)

    ##portia_data_equity_merge['Severity'] = portia_data_equity_merge['Severity'].astype(int)

    # Sort by Grade
    portia_data_fi_merge = portia_data_fi_merge.sort_values("Grade")
    # Seperate by investment grade
    portia_data_hy_merge = portia_data_fi_merge[portia_data_fi_merge["Grade"] == "HY"].sort_values("Breach")
    portia_data_ig_merge = portia_data_fi_merge[portia_data_fi_merge["Grade"] == "IG"].sort_values("Breach")

    ######################################## For Mutual Fund################################################################
    portia_data_mutualfund_merge[["Limit 1", "Limit 2"]] = [-0.20, -0.30]
    # Pay Attention to the deault value
    portia_data_mutualfund_merge["Breach"] = "No Breach"

    #for i in range(0,len(portia_data_equity_merge.index)):
        
    # if portia_data_equity_merge["Sec"][i]=="CNY" or portia_data_equity_merge["Sec"][i]=="HKD":
    #  portia_data_equity_merge[["Limit 1"]][i] = [-0.2]
    #  portia_data_equity_merge[["Limit 2"]][i] = [-0.3]

    for i in range(0,len(portia_data_mutualfund_merge.index)):
      ## if portia_data_equity_merge["Sec"][i]!="HKD": #or portia_data_equity_merge["Sec"][i] != "CNY": 
          portia_data_mutualfund_merge["Limit 1"][i]=-0.20
          portia_data_mutualfund_merge["Limit 2"][i]=-0.30
          if portia_data_mutualfund_merge["YTD Drawdown"][i] <= -0.20 and portia_data_mutualfund_merge["YTD Drawdown"][i] > -0.30 :
              portia_data_mutualfund_merge.loc[i, "Breach"] = "Breach Limit 1"
          elif portia_data_mutualfund_merge["YTD Drawdown"][i] <= -0.30 and portia_data_mutualfund_merge["Average Cost"][i] > portia_data_mutualfund_merge["Mkt Price"][i]:  # add another condition to exclude the stock split 
              portia_data_mutualfund_merge.loc[i, "Breach"] = "Breach Limit 2"
          else:
              portia_data_mutualfund_merge.loc[i, "Breach"] = "No Breach"





    #################### Merge with Fund Manager#########################

    # portia_data_equity_merge = portia_data_equity_merge.merge(fund_manager_data, how = "left", on = ["Fund Name"])
    # portia_data_hy_merge = portia_data_hy_merge.merge(fund_manager_data, how = "left", on = ["Fund Name"])
    # portia_data_ig_merge = portia_data_ig_merge.merge(fund_manager_data, how = "left", on = ["Fund Name"])

    # Output Excel



    output_path = output_excel_path(today)
    with pd.ExcelWriter(output_path) as writer:
        portia_data_equity_merge.to_excel(writer, sheet_name='Equity')
        portia_data_hy_merge.to_excel(writer, sheet_name='High Yield')
        portia_data_ig_merge.to_excel(writer, sheet_name='Investment Grade')
        portia_data_mutualfund_merge.to_excel(writer, sheet_name='Mutual Fund')
    print("Wrote", output_path)


if __name__ == "__main__":
    main()
