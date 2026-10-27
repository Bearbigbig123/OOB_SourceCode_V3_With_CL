import contextlib
import io
import unittest

import numpy as np
import pandas as pd
from matplotlib.figure import Figure

import oob_module_NGK_nostatic as oob


class OobChartStatisticsTests(unittest.TestCase):
    def test_resolution_decimal_places(self):
        cases = {
            1: 0,
            0.1: 1,
            0.01: 2,
            0.001: 3,
            0.5: 1,
            0.25: 2,
        }
        for resolution, expected in cases.items():
            with self.subTest(resolution=resolution):
                self.assertEqual(
                    oob.resolution_decimal_places(resolution), expected
                )

    def test_resolution_rounding_uses_half_up_for_positive_and_negative_values(self):
        raw = pd.DataFrame({"point_val": [1.234, 1.235, -1.234, -1.235]})

        rounded, decimal_places = oob.round_point_values_by_resolution(raw, 0.01)

        self.assertEqual(decimal_places, 2)
        self.assertEqual(
            rounded["point_val"].tolist(), [1.23, 1.24, -1.23, -1.24]
        )

    def test_preprocess_resolution_rounding_is_optional_and_auditable(self):
        raw = pd.DataFrame({
            "point_time": ["2026/10/20 10:00", "2026/10/21 10:00"],
            "point_val": [1.235, 2.344],
        })
        chart = pd.Series({
            "Characteristics": "Nominal", "USL": 10, "LSL": 0,
            "UCL": 9, "LCL": 1, "Target": 5, "Resolution": 0.01,
        })

        ok_plain, plain, plain_info = oob.preprocess_data(chart, raw, False)
        ok_rounded, rounded, rounded_info = oob.preprocess_data(chart, raw, True)

        self.assertTrue(ok_plain)
        self.assertTrue(ok_rounded)
        self.assertEqual(plain["point_val"].tolist(), [1.235, 2.344])
        self.assertEqual(rounded["point_val"].tolist(), [1.24, 2.34])
        self.assertFalse(plain_info["raw_data_rounding_enabled"])
        self.assertTrue(rounded_info["raw_data_rounding_enabled"])
        self.assertEqual(rounded_info["raw_data_rounding_decimal_places"], 2)

    def test_invalid_resolution_skips_rounding_without_failing_chart(self):
        raw = pd.DataFrame({"point_val": [1.235]})

        for resolution in (None, np.nan, 0, -0.01, "invalid"):
            with self.subTest(resolution=resolution):
                rounded, decimal_places = oob.round_point_values_by_resolution(
                    raw, resolution
                )
                self.assertIsNone(decimal_places)
                self.assertEqual(rounded["point_val"].tolist(), [1.235])

    def test_weekly_oos_counts_both_spec_sides(self):
        weekly = pd.DataFrame({"point_val": [4.0, 5.0, 7.0, 10.0, 11.0]})

        result = oob.calculate_oos_statistics(
            weekly,
            {"Characteristics": "Nominal", "USL": 10.0, "LSL": 5.0},
        )

        self.assertEqual(result["oos_cnt"], 2)
        self.assertEqual(result["oos_high_cnt"], 1)
        self.assertEqual(result["oos_low_cnt"], 1)
        self.assertEqual(result["oos_rate"], 0.4)

    def test_weekly_oos_counts_only_active_one_sided_spec(self):
        weekly = pd.DataFrame({"point_val": [4.0, 7.0, 11.0]})

        smaller = oob.calculate_oos_statistics(
            weekly,
            {"Characteristics": "Smaller", "USL": 10.0, "LSL": 5.0},
        )
        bigger = oob.calculate_oos_statistics(
            weekly,
            {"Characteristics": "Bigger", "USL": 10.0, "LSL": 5.0},
        )

        self.assertEqual(smaller["oos_cnt"], 1)
        self.assertEqual(smaller["oos_high_cnt"], 1)
        self.assertEqual(smaller["oos_low_cnt"], 0)
        self.assertEqual(bigger["oos_cnt"], 1)
        self.assertEqual(bigger["oos_high_cnt"], 0)
        self.assertEqual(bigger["oos_low_cnt"], 1)

    def test_chart_data_filters_historical_oos_but_keeps_weekly_oos(self):
        raw = pd.DataFrame({
            "point_time": pd.to_datetime([
                "2026-10-19", "2026-10-20", "2026-10-21", "2026-10-22"
            ]),
            "point_val": [20.0, 7.0, 20.0, 7.0],
        })
        chart_info = {"Characteristics": "Nominal", "USL": 10.0, "LSL": 5.0}

        result = oob.prepare_oob_chart_data(
            raw, chart_info, pd.Timestamp("2026-10-21"), pd.Timestamp("2026-10-27")
        )

        self.assertEqual(result["point_val"].tolist(), [7.0, 20.0, 7.0])

    def test_continuous_analysis_filters_baseline_and_counts_weekly_oos(self):
        weekly_start = pd.Timestamp("2026-07-20")
        weekly_end = pd.Timestamp("2026-07-26 23:59:59")
        baseline_end = weekly_start - pd.Timedelta(seconds=1)
        baseline_start = baseline_end - pd.Timedelta(days=365)
        raw = pd.DataFrame({
            "point_time": list(pd.date_range("2026-01-01", periods=10, freq="10D"))
            + [pd.Timestamp("2026-05-01"), pd.Timestamp("2026-07-21"), pd.Timestamp("2026-07-22")],
            "point_val": list(np.linspace(5.0, 6.0, 10)) + [15.0, 7.0, 12.0],
        })
        chart = {
            "Characteristics": "Smaller", "UCL": 20.0, "LCL": np.nan,
            "USL": 10.0, "LSL": np.nan, "Target": 6.0, "Resolution": 0.1,
        }

        with contextlib.redirect_stdout(io.StringIO()):
            result = oob.process_single_chart(
                chart, raw, baseline_start, baseline_end, weekly_start, weekly_end
            )

        self.assertEqual(result["baseline_n"], 10)
        self.assertEqual(result["weekly_n"], 2)
        self.assertEqual(result["oos_cnt"], 1)
        self.assertEqual(result["oos_high_cnt"], 1)
        self.assertEqual(result["oos_low_cnt"], 0)
        self.assertEqual(result["oos_rate"], 0.5)

    def test_audit_statistics_include_periods_and_distribution(self):
        kshift = pd.Series({
            "k_source": "Weekly",
            "k_p95": 1.25,
            "k_p95_abs": 1.25,
        })
        result = oob.build_oob_audit_statistics(
            [9.0, 10.0, 11.0],
            [8.0, 9.0, 10.0, 11.0, 12.0],
            kshift,
            "2026-10-21",
            "2026-10-27",
            "2025-10-20",
            "2026-10-20",
        )

        self.assertEqual(result["weekly_n"], 3)
        self.assertEqual(result["baseline_n"], 5)
        self.assertEqual(result["weekly_p50"], 10.0)
        self.assertAlmostEqual(result["baseline_p95"], 11.8)
        self.assertEqual(result["k_source"], "Weekly")
        self.assertEqual(result["k_p95"], 1.25)
        self.assertEqual(result["weekly_start"], pd.Timestamp("2026-10-21"))

    def test_single_weekly_point_exposes_rolling_k_as_display_value(self):
        baseline = {
            "values": np.linspace(8.0, 12.0, 40),
            "cnt": 40,
            "mean": 10.0,
            "sigma": 1.0,
        }
        weekly = {
            "values": np.array([11.5]),
            "cnt": 1,
            "mean": 11.5,
            "sigma": 0.0,
        }
        with contextlib.redirect_stdout(io.StringIO()):
            result = oob.kshift_sigma_ratio_calculator(
                baseline, weekly, "Nominal", 0.01, 12.0, 8.0
            )

        self.assertEqual(result["k_source"], "Rolling")
        self.assertTrue(np.isfinite(result["weekly_k_p95"]))
        self.assertTrue(np.isfinite(result["rolling_k_p95"]))
        self.assertEqual(result["k_p95"], result["rolling_k_p95"])

    def test_header_renders_matrix_and_highlight_color(self):
        figure = Figure(figsize=(13, 3.5))
        summary = {
            "weekly_n": 42,
            "Cpk": 1.36,
            "ooc_cnt": 2,
            "ooc_rate": 2 / 42,
            "baseline_p95": 1.234,
            "baseline_p50": 1.012,
            "baseline_p05": 0.845,
            "weekly_p95": 1.468,
            "weekly_p50": 1.205,
            "weekly_p05": 0.901,
            "k_source": "Rolling",
            "k_p95": 1.24,
            "k_p50": 0.82,
            "k_p05": 0.31,
            "HL_P95_shift": "HIGHLIGHT",
            "HL_P50_shift": "NO_HIGHLIGHT",
            "HL_P05_shift": "NO_HIGHLIGHT",
        }

        oob.add_oob_statistics_header(figure, summary)

        texts = {item.get_text(): item for item in figure.texts}
        self.assertNotIn("N=42  |  Cpk=1.36  |  OOC=2/42 (4.8%)", texts)
        self.assertIn("Baseline", texts)
        self.assertIn("Weekly", texts)
        self.assertIn("K Rolling", texts)
        self.assertIn("1.23", texts)
        self.assertNotIn("1.234", texts)
        self.assertEqual(texts["Baseline"].get_ha(), "right")
        self.assertEqual(texts["Weekly"].get_ha(), "right")
        self.assertEqual(texts["P95"].get_ha(), "center")
        self.assertEqual(texts["P50"].get_ha(), "center")
        self.assertEqual(texts["P05"].get_ha(), "center")
        self.assertAlmostEqual(
            texts["P50"].get_position()[0] - texts["P95"].get_position()[0],
            0.12,
        )
        self.assertAlmostEqual(
            texts["P05"].get_position()[0] - texts["P50"].get_position()[0],
            0.12,
        )
        self.assertEqual(texts["P95"].get_fontsize(), 10)
        self.assertEqual(texts["+1.24"].get_color(), "#DC2626")
        self.assertEqual(texts["+0.82"].get_color(), "#334155")


if __name__ == "__main__":
    unittest.main()
