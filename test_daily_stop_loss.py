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
        self.assertEqual(sl.classify_two_limit_breach(-0.10, -0.20, -0.30), "No Breach")
        self.assertEqual(sl.classify_two_limit_breach(-0.20, -0.20, -0.30), "Breach Limit 1")
        self.assertEqual(sl.classify_two_limit_breach(-0.25, -0.20, -0.30), "Breach Limit 1")
        self.assertEqual(sl.classify_two_limit_breach(-0.30, -0.20, -0.30), "Breach Limit 2")
        self.assertEqual(sl.classify_two_limit_breach(-0.35, -0.20, -0.30), "Breach Limit 2")
        # YTD of -35% would have been Limit 2; holding-period -10% is not a breach.
        ytd = -0.35
        holding = -0.10
        self.assertEqual(sl.classify_two_limit_breach(ytd, -0.20, -0.30), "Breach Limit 2")
        self.assertEqual(sl.classify_two_limit_breach(holding, -0.20, -0.30), "No Breach")

    def test_assign_equity_breaches_ignores_ytd_and_offset_index(self):
        df = pd.DataFrame(
            {
                "YTD Drawdown": [-0.40, -0.40, -0.01],
                "Holding Period Drawdown": [-0.05, -0.22, -0.35],
            },
            index=[10, 20, 30],
        )
        result = sl.assign_equity_breaches(df)
        self.assertEqual(list(result["Breach"]), ["No Breach", "Breach Limit 1", "Breach Limit 2"])

    def test_assign_fi_breaches_uses_holding_period(self):
        df = pd.DataFrame(
            {
                "Fund Name": ["A", "A", "TBHTHYEF", "B"],
                "Grade": ["HY", "HY", "HY", "IG"],
                "YTD Drawdown": [-0.40, -0.40, -0.40, -0.40],
                "Holding Period Drawdown": [-0.10, -0.20, -0.40, -0.10],
            }
        )
        result = sl.assign_fi_breaches(df)
        self.assertEqual(
            list(result["Breach"]),
            ["No Breach", "Breach Limit 1", "No Breach", "Breach Limit 1"],
        )


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


if __name__ == "__main__":
    unittest.main()
