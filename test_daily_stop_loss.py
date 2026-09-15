import unittest
from unittest.mock import patch

import numpy as np
import pandas as pd

import daily_stop_loss as sl


class TickerConversionTests(unittest.TestCase):
    def test_bloomberg_hk_pads_to_four_digits(self):
        self.assertEqual(sl.to_yahoo_ticker("700 HK"), "0700.HK")
        self.assertEqual(sl.to_yahoo_ticker("5 HK Equity"), "0005.HK")
        self.assertEqual(sl.to_yahoo_ticker("9988 HK"), "9988.HK")

    def test_us_and_china_and_existing_yahoo(self):
        self.assertEqual(sl.to_yahoo_ticker("AAPL US"), "AAPL")
        self.assertEqual(sl.to_yahoo_ticker("BRK/B US"), "BRK-B")
        self.assertEqual(sl.to_yahoo_ticker("600519 CH"), "600519.SS")
        self.assertEqual(sl.to_yahoo_ticker("000858 CH"), "000858.SZ")
        self.assertEqual(sl.to_yahoo_ticker("0700.HK"), "0700.HK")
        self.assertEqual(sl.to_yahoo_ticker("7203 JT"), "7203.T")
        self.assertEqual(sl.to_yahoo_ticker("005930 KS"), "005930.KS")

    def test_numeric_id_uses_currency(self):
        self.assertEqual(sl.to_yahoo_ticker("5", "HKD"), "0005.HK")
        self.assertEqual(sl.to_yahoo_ticker("600519", "CNY"), "600519.SS")
        self.assertEqual(sl.to_yahoo_ticker("000858", "CNY"), "000858.SZ")


class DrawdownColumnTests(unittest.TestCase):
    def test_ytd_holding_and_from_high_formulas_and_order(self):
        df = pd.DataFrame(
            {
                "Fund Name": ["F1"],
                "Security ID": ["AAPL US"],
                "Sec Curr": ["USD"],
                "Mkt Price": [80.0],
                "Average Cost": [100.0],
                "Price Benchmark": [100.0],
            }
        )
        with patch.object(sl, "fetch_last_year_highs", return_value={"AAPL": 200.0}):
            result = sl.add_equity_drawdown_columns(df.copy())

        self.assertAlmostEqual(result.loc[0, "YTD Drawdown"], -0.20)
        self.assertAlmostEqual(result.loc[0, "Holding Period Drawdown"], -0.20)
        self.assertAlmostEqual(result.loc[0, "High Last Year"], 200.0)
        self.assertAlmostEqual(result.loc[0, "Drawdown from High"], -0.60)

        cols = list(result.columns)
        self.assertEqual(cols[cols.index("YTD Drawdown") - 1], "High Last Year")
        self.assertEqual(cols[cols.index("YTD Drawdown") + 1], "Holding Period Drawdown")
        self.assertEqual(cols[cols.index("Holding Period Drawdown") + 1], "Drawdown from High")

    def test_non_equity_has_holding_period_but_not_high_columns(self):
        df = pd.DataFrame(
            {
                "Mkt Price": [90.0],
                "Average Cost": [100.0],
                "Price Benchmark": [120.0],
            }
        )
        result = sl.add_non_equity_drawdown_columns(df.copy())
        self.assertAlmostEqual(result.loc[0, "YTD Drawdown"], (90.0 - 120.0) / 120.0)
        self.assertAlmostEqual(result.loc[0, "Holding Period Drawdown"], -0.10)
        self.assertNotIn("High Last Year", result.columns)
        self.assertNotIn("Drawdown from High", result.columns)
        cols = list(result.columns)
        self.assertEqual(cols[cols.index("YTD Drawdown") + 1], "Holding Period Drawdown")

    def test_output_filename_includes_test_suffix(self):
        path = sl.output_excel_path("10092026")
        self.assertTrue(path.endswith("daily_output_10092026 test.xlsx"))
        self.assertIn("Daily Output", path)

    def test_breach_uses_holding_period_drawdown_not_ytd(self):
        self.assertEqual(sl.BREACH_DRAWDOWN_COLUMN, "Holding Period Drawdown")
        l1, l2 = sl.EQUITY_LIMIT_1, sl.EQUITY_LIMIT_2
        self.assertEqual(sl.classify_two_limit_breach(l1 + 0.10, l1, l2), "No Breach")
        self.assertEqual(sl.classify_two_limit_breach(l1, l1, l2), "Breach Limit 1")
        self.assertEqual(sl.classify_two_limit_breach((l1 + l2) / 2, l1, l2), "Breach Limit 1")
        self.assertEqual(sl.classify_two_limit_breach(l2, l1, l2), "Breach Limit 2")
        self.assertEqual(sl.classify_two_limit_breach(l2 - 0.05, l1, l2), "Breach Limit 2")
        ytd = l2 - 0.05
        holding = l1 + 0.10
        self.assertEqual(sl.classify_two_limit_breach(ytd, l1, l2), "Breach Limit 2")
        self.assertEqual(sl.classify_two_limit_breach(holding, l1, l2), "No Breach")

    def test_assign_equity_breaches_ignores_ytd_and_offset_index(self):
        l1, l2 = sl.EQUITY_LIMIT_1, sl.EQUITY_LIMIT_2
        df = pd.DataFrame(
            {
                "YTD Drawdown": [l2 - 0.10, l2 - 0.10, -0.01],
                "Holding Period Drawdown": [l1 + 0.15, (l1 + l2) / 2, l2 - 0.05],
            },
            index=[10, 20, 30],
        )
        result = sl.assign_equity_breaches(df)
        self.assertEqual(list(result["Breach"]), ["No Breach", "Breach Limit 1", "Breach Limit 2"])
        self.assertTrue((result["Limit 1"] == l1).all())
        self.assertTrue((result["Limit 2"] == l2).all())

    def test_assign_fi_breaches_uses_holding_period(self):
        hy1, hy2 = sl.HY_LIMIT_1, sl.HY_LIMIT_2
        ig1, ig2 = sl.IG_LIMIT_1, sl.IG_LIMIT_2
        excluded = sl.EXCLUDED_FUNDS[0]
        df = pd.DataFrame(
            {
                "Fund Name": ["A", "A", excluded, "B"],
                "Grade": ["HY", "HY", "HY", "IG"],
                "YTD Drawdown": [hy2 - 0.15, hy2 - 0.15, hy2 - 0.15, hy2 - 0.15],
                "Holding Period Drawdown": [hy1 + 0.05, (hy1 + hy2) / 2, hy2 - 0.15, (ig1 + ig2) / 2],
            }
        )
        result = sl.assign_fi_breaches(df)
        self.assertEqual(
            list(result["Breach"]),
            ["No Breach", "Breach Limit 1", "No Breach", "Breach Limit 1"],
        )
        self.assertEqual(list(result["Limit 1"]), [hy1, hy1, hy1, ig1])
        self.assertEqual(list(result["Limit 2"]), [hy2, hy2, hy2, ig2])

    def test_excluded_fund_lists_are_defined_at_top(self):
        self.assertIn("DCFH2024", sl.EXCLUDED_FUNDS)
        self.assertIn("TBHTHYEF", sl.EXCLUDED_FUNDS)

    def test_excluded_funds_skip_breach_on_equity(self):
        l1, l2 = sl.EQUITY_LIMIT_1, sl.EQUITY_LIMIT_2
        excluded = sl.EXCLUDED_FUNDS[0]
        df = pd.DataFrame(
            {
                "Fund Name": ["NORMAL", excluded],
                "Sec Curr": ["USD", "USD"],
                "Holding Period Drawdown": [l2 - 0.05, l2 - 0.05],
            }
        )
        result = sl.assign_equity_breaches(df)
        self.assertEqual(list(result["Breach"]), ["Breach Limit 2", "No Breach"])

    def test_mutual_fund_limits_follow_equity_constants(self):
        self.assertEqual(sl.MUTUAL_FUND_LIMIT_1, sl.EQUITY_LIMIT_1)
        self.assertEqual(sl.MUTUAL_FUND_LIMIT_2, sl.EQUITY_LIMIT_2)

    def test_equity_limits_listed_by_currency(self):
        expected = ["HKD", "SGD", "KRW", "USD", "EUR", "JPY", "CAD", "GBP", "CNH", "AUD"]
        self.assertEqual(list(sl.EQUITY_LIMITS_BY_CURRENCY.keys()), expected)
        self.assertEqual(sl.equity_limits_for_currency("hkd"), sl.EQUITY_LIMITS_BY_CURRENCY["HKD"])
        self.assertEqual(sl.equity_limits_for_currency("TWD"), (sl.EQUITY_LIMIT_1, sl.EQUITY_LIMIT_2))

    def test_assign_equity_breaches_uses_sec_curr_limits(self):
        original = dict(sl.EQUITY_LIMITS_BY_CURRENCY)
        try:
            sl.EQUITY_LIMITS_BY_CURRENCY["KRW"] = (-0.28, -0.38)
            sl.EQUITY_LIMITS_BY_CURRENCY["USD"] = (-0.20, -0.30)
            df = pd.DataFrame(
                {
                    "Sec Curr": ["KRW", "KRW", "USD"],
                    "YTD Drawdown": [0.0, 0.0, 0.0],
                    "Holding Period Drawdown": [-0.30, -0.40, -0.30],
                }
            )
            result = sl.assign_equity_breaches(df)
            self.assertEqual(list(result["Limit 1"]), [-0.28, -0.28, -0.20])
            self.assertEqual(list(result["Limit 2"]), [-0.38, -0.38, -0.30])
            self.assertEqual(list(result["Breach"]), ["Breach Limit 1", "Breach Limit 2", "Breach Limit 2"])
        finally:
            sl.EQUITY_LIMITS_BY_CURRENCY.clear()
            sl.EQUITY_LIMITS_BY_CURRENCY.update(original)


class YahooHighExtractionTests(unittest.TestCase):
    def test_max_high_from_single_ticker_frame(self):
        hist = pd.DataFrame({"High": [10.0, 12.5, 11.0]})
        self.assertAlmostEqual(sl._max_high_from_history(hist), 12.5)

    def test_max_high_from_multiindex_download(self):
        arrays = [["AAPL", "AAPL"], ["Open", "High"]]
        columns = pd.MultiIndex.from_arrays(arrays)
        data = pd.DataFrame([[1.0, 10.0], [1.0, 15.0]], columns=columns)
        self.assertAlmostEqual(sl._max_high_from_download(data, "AAPL"), 15.0)

    def test_fetch_last_year_highs_uses_download(self):
        hist = pd.DataFrame({"High": [10.0, 22.0, 18.0]})
        with patch.object(sl.yf, "download", return_value=hist) as mock_download:
            highs = sl.fetch_last_year_highs(["AAPL"])
        self.assertAlmostEqual(highs["AAPL"], 22.0)
        self.assertEqual(mock_download.call_args.kwargs["auto_adjust"], False)

    def test_fetch_last_year_highs_fallback_is_unadjusted(self):
        hist = pd.DataFrame({"High": [10.0, 22.0, 18.0]})

        class FakeTicker:
            def history(self, **kwargs):
                self.kwargs = kwargs
                return hist

        fake = FakeTicker()
        with patch.object(sl.yf, "download", side_effect=RuntimeError("batch failed")):
            with patch.object(sl.yf, "Ticker", return_value=fake):
                highs = sl.fetch_last_year_highs(["AAPL"])
        self.assertAlmostEqual(highs["AAPL"], 22.0)
        self.assertEqual(fake.kwargs["auto_adjust"], False)

    def test_pence_scale_converts_gbpence_to_pounds(self):
        self.assertEqual(sl.yahoo_price_scale("AZN.L", "GBp"), 0.01)
        self.assertEqual(sl.yahoo_price_scale("AZN.L", "GBX"), 0.01)
        self.assertEqual(sl.yahoo_price_scale("AZN.L", None), 0.01)
        self.assertEqual(sl.yahoo_price_scale("AAPL", "USD"), 1.0)
        self.assertEqual(sl.yahoo_price_scale("AAPL", "GBP"), 1.0)

    def test_london_high_is_converted_from_pence_to_pounds(self):
        hist = pd.DataFrame({"High": [15732.0, 11708.0]})

        class FakeTicker:
            fast_info = {"currency": "GBp"}

            def history(self, **kwargs):
                return hist

        with patch.object(sl.yf, "download", return_value=hist):
            with patch.object(sl.yf, "Ticker", return_value=FakeTicker()):
                highs = sl.fetch_last_year_highs(["AZN.L"])
        self.assertAlmostEqual(highs["AZN.L"], 157.32)


class CopyOnWriteAssignmentTests(unittest.TestCase):
    """Pandas 3 Copy-on-Write forbids chained assignment such as df[col][i] = value."""

    def test_data_cleaning_writes_benchmark_on_filtered_rows(self):
        initial = pd.DataFrame(
            {
                "Fund Name": ["F1", "F2", "F3"],
                "Security ID": ["A", "B", "C"],
                "Sec Type": ["Common Stock", "Corporate Bond", "Mutual Fund"],
                "Mkt Price": ["100", "90", "80"],
                "Average Cost": ["110", "95", "85"],
            }
        )
        last = pd.DataFrame(
            {
                "Fund Name": ["F1", "F2", "F3"],
                "Security ID": ["A", "B", "C"],
                "Sec Type": ["Common Stock", "Corporate Bond", "Mutual Fund"],
                "Mkt Price": ["80", "70", "60"],
                "Average Cost": ["110", "95", "85"],
                "Security Desc": ["Apple Inc", "FOO 5 1/2", "Fund X"],
            }
        )
        equity, fi, mf = sl.data_cleaning(initial, last)
        self.assertAlmostEqual(equity.iloc[0]["Price Benchmark"], 100.0)
        self.assertAlmostEqual(equity.iloc[0]["Mkt Price"], 80.0)
        self.assertAlmostEqual(fi.iloc[0]["Price Benchmark"], 90.0)
        self.assertAlmostEqual(mf.iloc[0]["Price Benchmark"], 80.0)
        self.assertEqual(list(equity.columns).count("Price Benchmark"), 1)

    def test_assign_equity_breaches_on_boolean_slice(self):
        parent = pd.DataFrame(
            {
                "Fund Name": ["A", "B", "C"],
                "Sec Curr": ["USD", "USD", "USD"],
                "Sec Type": ["Common Stock", "Common Stock", "Bond"],
                "Holding Period Drawdown": [-0.35, -0.10, -0.50],
            }
        )
        slice_df = parent[parent["Sec Type"] == "Common Stock"]
        result = sl.assign_equity_breaches(slice_df)
        self.assertEqual(list(result["Breach"]), ["Breach Limit 2", "No Breach"])
        self.assertTrue((result["Limit 1"] == sl.EQUITY_LIMIT_1).all())

    def test_drawdown_helpers_on_boolean_slice(self):
        parent = pd.DataFrame(
            {
                "Security ID": ["AAPL US", "MSFT US", "BOND"],
                "Sec Curr": ["USD", "USD", "USD"],
                "Sec Type": ["Common Stock", "Common Stock", "Bond"],
                "Mkt Price": [80.0, 90.0, 70.0],
                "Average Cost": [100.0, 100.0, 100.0],
                "Price Benchmark": [100.0, 100.0, 100.0],
            }
        )
        slice_df = parent[parent["Sec Type"] == "Common Stock"]
        with patch.object(sl, "fetch_last_year_highs", return_value={"AAPL": 200.0, "MSFT": 180.0}):
            result = sl.add_equity_drawdown_columns(slice_df)
        self.assertAlmostEqual(result.iloc[0]["YTD Drawdown"], -0.20)
        self.assertAlmostEqual(result.iloc[0]["Holding Period Drawdown"], -0.20)
        self.assertAlmostEqual(result.iloc[0]["High Last Year"], 200.0)


if __name__ == "__main__":
    unittest.main()
