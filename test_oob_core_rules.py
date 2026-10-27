import ast
import contextlib
import io
import logging
import traceback
import unittest
from pathlib import Path

import numpy as np
import pandas as pd


SOURCE_PATH = Path(__file__).with_name("oob_module_NGK_nostatic.py")
FUNCTIONS = {
    "is_valid_number",
    "unavailable_kshift_result",
    "get_percentiles",
    "summarize_oob_values",
    "build_oob_audit_statistics",
    "rolling_calculation",
    "safe_division",
    "review_kshift_results",
    "kshift_sigma_ratio_calculator",
    "record_high_low_calculator",
    "optional_record_high_low_calculator",
    "normalize_characteristic",
    "format_and_clean_data",
    "resolution_decimal_places",
    "round_point_values_by_resolution",
    "update_chart_limits",
    "exclude_oos_data",
    "calculate_oos_statistics",
    "prepare_oob_chart_data",
    "preprocess_data",
    "ooc_calculator",
    "review_ooc_results",
    "process_single_chart",
    "_process_discrete_chart",
}


def load_core_namespace():
    tree = ast.parse(SOURCE_PATH.read_text(encoding="utf-8"))
    selected = [
        node
        for node in tree.body
        if isinstance(node, ast.FunctionDef) and node.name in FUNCTIONS
    ]
    for node in tree.body:
        if isinstance(node, ast.ClassDef):
            selected.extend(
                child
                for child in node.body
                if isinstance(child, ast.FunctionDef) and child.name in FUNCTIONS
            )
    module = ast.Module(body=selected, type_ignores=[])
    namespace = {
        "pd": pd,
        "np": np,
        "logger": logging.getLogger("test_oob_core"),
        "traceback": traceback,
        "OOB_CALCULATED": "CALCULATED",
        "OOB_ERROR": "ERROR",
        "OOB_INSUFFICIENT_DATA": "INSUFFICIENT_DATA",
        "OOB_KEYS": [
            "HL_P95_shift",
            "HL_P50_shift",
            "HL_P05_shift",
            "HL_sticking_shift",
            "HL_trending",
            "HL_high_OOC",
            "HL_record_high_low",
            "HL_category_LT_shift",
        ],
        "KSHIFT_AUDIT_KEYS": tuple(
            key
            for percentile in ("p95", "p50", "p05")
            for key in (
                f"weekly_k_{percentile}", f"weekly_k_{percentile}_abs",
                f"rolling_k_{percentile}", f"rolling_k_{percentile}_abs",
                f"k_{percentile}", f"k_{percentile}_abs",
            )
        ) + ("k_source",),
    }
    exec(compile(module, str(SOURCE_PATH), "exec"), namespace)
    return namespace


CORE = load_core_namespace()


class OobCoreRuleTests(unittest.TestCase):
    def test_record_high_low_is_disabled_by_default(self):
        result = CORE["optional_record_high_low_calculator"](
            np.array([11.0]), np.arange(1.0, 11.0)
        )

        self.assertEqual(result["highlight_status"], "NO_HIGHLIGHT")
        self.assertFalse(result["record_high"])
        self.assertFalse(result["record_low"])

    def test_record_high_low_can_be_enabled(self):
        result = CORE["optional_record_high_low_calculator"](
            np.array([11.0]), np.arange(1.0, 11.0), enabled=True
        )

        self.assertEqual(result["highlight_status"], "HIGHLIGHT")
        self.assertTrue(result["record_high"])
        self.assertFalse(result["record_low"])

    def _assert_single_sided_kshift_denominator(self, ucl, lcl):
        base_values = np.arange(-10.0, 11.0)
        data_values = np.array([1.0, 2.0, 3.0])
        base_percentiles = CORE["get_percentiles"](base_values)
        data_percentiles = CORE["get_percentiles"](data_values)
        if not pd.isna(ucl):
            single_side_deno = (ucl - base_percentiles["P50"]) / 6
        else:
            single_side_deno = (base_percentiles["P50"] - lcl) / 6

        percentile_denominators = {
            "P95": (base_percentiles["P99.865"] - base_percentiles["P50"]) / 3,
            "P50": (base_percentiles["P99.865"] - base_percentiles["P0.135"]) / 6,
            "P05": (base_percentiles["P50"] - base_percentiles["P0.135"]) / 3,
        }

        with contextlib.redirect_stdout(io.StringIO()):
            result = CORE["kshift_sigma_ratio_calculator"](
                {"values": base_values},
                {"values": data_values},
                "Nominal",
                0.1,
                ucl,
                lcl,
            )

        for percentile in ("P95", "P50", "P05"):
            expected_deno = round(
                max(percentile_denominators[percentile], single_side_deno), 8
            )
            expected_k = round(
                abs(data_percentiles[percentile] - base_percentiles[percentile]), 8
            ) / expected_deno
            self.assertAlmostEqual(result[f"{percentile}_k"], expected_k)

    def test_ucl_only_is_used_by_all_kshift_denominators(self):
        self._assert_single_sided_kshift_denominator(60.0, np.nan)

    def test_lcl_only_is_used_by_all_kshift_denominators(self):
        self._assert_single_sided_kshift_denominator(np.nan, -60.0)

    def test_invalid_point_values_are_removed(self):
        raw = pd.DataFrame(
            {
                "point_val": ["1.0", "", "bad", np.nan, np.inf, "4.5"],
                "point_time": ["2026/07/20 00:00"] * 6,
            }
        )

        cleaned = CORE["format_and_clean_data"](raw, {})

        self.assertEqual(cleaned["point_val"].tolist(), [1.0, 4.5])

    def test_smaller_uses_only_upper_control_limit(self):
        data = pd.DataFrame({"point_val": [8, 11, 12]})

        result = CORE["ooc_calculator"](data, 10, np.nan, "Smaller")

        self.assertEqual(result, (3, 2, 2 / 3))

    def test_bigger_uses_only_lower_control_limit(self):
        data = pd.DataFrame({"point_val": [4, 5, 7]})

        result = CORE["ooc_calculator"](data, np.nan, 6, "Bigger")

        self.assertEqual(result, (3, 2, 2 / 3))

    def test_nominal_still_requires_both_control_limits(self):
        data = pd.DataFrame({"point_val": [4, 6, 11]})

        with self.assertRaises(ValueError):
            CORE["ooc_calculator"](data, 10, np.nan, "Nominal")

    def test_single_sided_spec_filter_ignores_unused_side(self):
        raw = pd.DataFrame(
            {
                "point_val": [1.0, 5.0, 9.0],
                "usl_val": [8.0] * 3,
                "lsl_val": [np.nan] * 3,
            }
        )

        filtered = CORE["exclude_oos_data"](raw, {"Characteristics": "Smaller"})

        self.assertEqual(filtered["point_val"].tolist(), [1.0, 5.0])

    def test_unavailable_kshift_is_not_normal(self):
        result = CORE["unavailable_kshift_result"]("ERROR")

        self.assertEqual(result["P95_shift"], "ERROR")
        self.assertEqual(result["P50_shift"], "ERROR")
        self.assertEqual(result["P05_shift"], "ERROR")

    def test_insufficient_baseline_does_not_disable_ooc(self):
        weekly_start = pd.Timestamp("2026-07-20")
        weekly_end = pd.Timestamp("2026-07-26 23:59:59")
        baseline_end = weekly_start - pd.Timedelta(seconds=1)
        baseline_start = baseline_end - pd.Timedelta(days=365)
        raw = pd.DataFrame(
            {
                "point_time": [
                    pd.Timestamp("2026-01-01"),
                    pd.Timestamp("2026-02-01"),
                    pd.Timestamp("2026-03-01"),
                    pd.Timestamp("2026-04-01"),
                    pd.Timestamp("2026-05-01"),
                    pd.Timestamp("2026-07-21"),
                    pd.Timestamp("2026-07-22"),
                ],
                "point_val": [5, 5, 5, 5, 5, 11, 12],
            }
        )
        chart = {
            "Characteristics": "Smaller",
            "UCL": 10,
            "LCL": np.nan,
            "USL": 20,
            "LSL": np.nan,
            "Target": 5,
            "Resolution": 1,
        }

        with contextlib.redirect_stdout(io.StringIO()):
            result = CORE["process_single_chart"](
                chart,
                raw,
                baseline_start,
                baseline_end,
                weekly_start,
                weekly_end,
            )

        self.assertEqual(result["OOB_Status"], "INSUFFICIENT_DATA")
        self.assertEqual(result["HL_high_OOC"], "HIGHLIGHT")
        self.assertEqual(result["HL_P50_shift"], "INSUFFICIENT_DATA")

    def test_discrete_insufficient_baseline_still_calculates_ooc(self):
        weekly_start = pd.Timestamp("2026-07-20")
        weekly_end = pd.Timestamp("2026-07-26 23:59:59")
        baseline_end = weekly_start - pd.Timedelta(seconds=1)
        baseline_start = baseline_end - pd.Timedelta(days=365)
        raw = pd.DataFrame(
            {
                "point_time": [
                    pd.Timestamp("2026-01-01"),
                    pd.Timestamp("2026-02-01"),
                    pd.Timestamp("2026-03-01"),
                    pd.Timestamp("2026-07-21"),
                    pd.Timestamp("2026-07-22"),
                ],
                "point_val": [5, 5, 5, 11, 12],
            }
        )
        chart = {
            "Characteristics": "Smaller",
            "UCL": 10,
            "LCL": np.nan,
            "USL": 20,
            "LSL": np.nan,
            "Target": 5,
            "Resolution": 1,
        }

        with contextlib.redirect_stdout(io.StringIO()):
            result = CORE["_process_discrete_chart"](
                None,
                raw,
                chart,
                weekly_start,
                weekly_end,
                baseline_start,
                baseline_end,
            )

        self.assertEqual(result["OOB_Status"], "INSUFFICIENT_DATA")
        self.assertEqual(result["HL_high_OOC"], "HIGHLIGHT")
        self.assertEqual(result["ooc_cnt"], 2)


if __name__ == "__main__":
    unittest.main()
