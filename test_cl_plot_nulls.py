import contextlib
import io
import tempfile
import unittest
from pathlib import Path

import numpy as np
import pandas as pd

from CL_limit_class import CLTightenCalculator


class PlotNullValueTests(unittest.TestCase):
    def setUp(self):
        self.calculator = CLTightenCalculator()
        self.chart_info = pd.Series(
            {
                'GroupName': 'NullTest',
                'ChartName': 'Plot',
                'Characteristics': 'Nominal',
                'Target': 10.0,
                'UCL': 11.0,
                'LCL': 9.0,
            }
        )

    def plot(self, chart_data, output_dir):
        log = io.StringIO()
        with contextlib.redirect_stdout(log):
            result = self.calculator.plot_control_chart(
                chart_data=chart_data,
                chart_info=self.chart_info,
                suggest_ucl=10.5,
                suggest_lcl=9.5,
                static_ucl=10.5,
                static_lcl=9.5,
                cl_center=10.0,
                pattern='Normal',
                total_data_count=3,
                used_data_count=3,
                output_dir=output_dir,
            )
        return result, log.getvalue()

    def test_nulls_at_start_middle_and_end_are_removed(self):
        chart_data = pd.DataFrame(
            {
                'date': pd.date_range('2026-01-01', periods=8),
                'value': [None, 10.0, '', 10.2, 'not-a-number', 9.8, np.nan, None],
            }
        )

        with tempfile.TemporaryDirectory() as output_dir:
            result, log = self.plot(chart_data, output_dir)

            self.assertIsNotNone(result)
            self.assertTrue(Path(result).exists())
            self.assertIn('原始=8', log)
            self.assertIn('無效數值=5', log)
            self.assertIn('最終繪圖=3', log)
            self.assertNotIn('Axis limits cannot be NaN or Inf', log)

    def test_invalid_dates_are_removed_with_their_values(self):
        chart_data = pd.DataFrame(
            {
                'date': ['2026-01-01', 'invalid', '2026-01-03', None],
                'value': [10.0, 10.1, 10.2, 10.3],
            }
        )

        with tempfile.TemporaryDirectory() as output_dir:
            result, log = self.plot(chart_data, output_dir)

            self.assertIsNotNone(result)
            self.assertTrue(Path(result).exists())
            self.assertIn('無效日期=2', log)
            self.assertIn('最終繪圖=2', log)

    def test_all_invalid_values_return_none_without_axis_error(self):
        chart_data = pd.DataFrame(
            {
                'date': pd.date_range('2026-01-01', periods=4),
                'value': [None, '', 'not-a-number', np.nan],
            }
        )

        with tempfile.TemporaryDirectory() as output_dir:
            result, log = self.plot(chart_data, output_dir)

            self.assertIsNone(result)
            self.assertIn('最終繪圖=0', log)
            self.assertIn('清理後沒有有效數據點', log)
            self.assertNotIn('Axis limits cannot be NaN or Inf', log)
            self.assertEqual(list(Path(output_dir).iterdir()), [])

    def test_no_null_data_still_plots(self):
        chart_data = pd.DataFrame(
            {
                'date': pd.date_range('2026-01-01', periods=4),
                'value': [9.9, 10.0, 10.1, 10.2],
            }
        )

        with tempfile.TemporaryDirectory() as output_dir:
            result, log = self.plot(chart_data, output_dir)

            self.assertIsNotNone(result)
            self.assertTrue(Path(result).exists())
            self.assertIn('排除=0', log)
            self.assertIn('最終繪圖=4', log)


if __name__ == '__main__':
    unittest.main()
