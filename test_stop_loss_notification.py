import os
import tempfile
import unittest
from datetime import datetime

import pandas as pd

import stop_loss_notification as notice


class NotificationColumnTests(unittest.TestCase):
    def test_order_keeps_new_drawdown_columns_before_breach(self):
        df = pd.DataFrame(
            {
                "Fund Name": ["F1"],
                "Breach": ["Breach Limit 1"],
                "Unnamed: 0": [0],
                "YTD Drawdown": [-0.10],
                "Holding Period Drawdown": [-0.22],
                "High Last Year": [200.0],
                "Drawdown from High": [-0.40],
                "Severity": [1],
                "Security Desc": ["APPLE"],
            }
        )
        ordered = notice.order_output_columns(df)
        cols = list(ordered.columns)
        self.assertNotIn("Unnamed: 0", cols)
        self.assertLess(cols.index("High Last Year"), cols.index("YTD Drawdown"))
        self.assertLess(cols.index("YTD Drawdown"), cols.index("Holding Period Drawdown"))
        self.assertLess(cols.index("Holding Period Drawdown"), cols.index("Drawdown from High"))
        self.assertLess(cols.index("Drawdown from High"), cols.index("Breach"))

    def test_filter_today_breaches_keeps_updated_columns(self):
        equity = pd.DataFrame(
            {
                "Fund Name": ["F1", "F2"],
                "Security Desc": ["AAA", "BBB"],
                "YTD Drawdown": [-0.05, -0.40],
                "Holding Period Drawdown": [-0.22, -0.05],
                "High Last Year": [10.0, 20.0],
                "Drawdown from High": [-0.30, -0.10],
                "Breach": ["Breach Limit 1", "No Breach"],
                "Severity": [1, ""],
            }
        )
        with tempfile.TemporaryDirectory() as tmp:
            source = os.path.join(tmp, "daily_output_14092026 test.xlsx")
            dest = os.path.join(tmp, "Stop Loss 20260914.xlsx")
            with pd.ExcelWriter(source) as writer:
                equity.to_excel(writer, sheet_name="Equity", index=False)
                equity.iloc[0:0].to_excel(writer, sheet_name="High Yield", index=False)
                equity.iloc[0:0].to_excel(writer, sheet_name="Investment Grade", index=False)
            notice.filter_today_breaches(source, dest)
            out = pd.read_excel(dest, sheet_name="Equity")
            self.assertEqual(list(out["Fund Name"]), ["F1"])
            self.assertIn("Holding Period Drawdown", out.columns)
            self.assertIn("Drawdown from High", out.columns)
            self.assertIn("YTD Drawdown", out.columns)
            self.assertIn("High Last Year", out.columns)
            self.assertAlmostEqual(out.loc[0, "Holding Period Drawdown"], -0.22)

    def test_new_breach_when_not_in_history(self):
        today = pd.DataFrame(
            {
                "Fund Name": ["F1"],
                "Security Desc": ["AAA"],
                "Holding Period Drawdown": [-0.22],
                "Drawdown from High": [-0.30],
                "YTD Drawdown": [-0.05],
                "Breach": ["Breach Limit 1"],
                "Severity": [1],
            }
        )
        history = pd.DataFrame(
            {
                "Date": ["2026-01-01"],
                "Fund Name Security Desc": ["OTHER | ZZZ"],
                "Severity": [2],
            }
        )
        with tempfile.TemporaryDirectory() as tmp:
            dest = os.path.join(tmp, "dest.xlsx")
            with pd.ExcelWriter(dest) as writer:
                today.to_excel(writer, sheet_name="Equity", index=False)
            result = notice.get_new_breach_df(
                "Equity",
                history,
                dest_path=dest,
                as_of=pd.Timestamp("2026-09-14"),
            )
            self.assertEqual(len(result), 1)
            self.assertEqual(result.loc[0, "Last Breach Date"], "New Breach")
            self.assertIn("Holding Period Drawdown", result.columns)
            self.assertIn("Drawdown from High", result.columns)

    def test_resolve_prefers_updated_test_filename(self):
        with tempfile.TemporaryDirectory() as tmp:
            primary = os.path.join(tmp, "daily_output_14092026 test.xlsx")
            legacy = os.path.join(tmp, "daily_output_14092026.xlsx")
            open(primary, "w").close()
            open(legacy, "w").close()
            self.assertEqual(notice.resolve_source_file_path(primary, legacy), primary)
            os.remove(primary)
            self.assertEqual(notice.resolve_source_file_path(primary, legacy), legacy)


if __name__ == "__main__":
    unittest.main()
