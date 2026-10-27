import contextlib
import io
import unittest
from unittest.mock import patch

import numpy as np
import pandas as pd

from CL_limit_class import CLTightenCalculator


class FixedDataCalculator(CLTightenCalculator):
    def __init__(self, values, resolution):
        super().__init__()
        self._values = np.asarray(values, dtype=float)
        self._resolution = resolution

    def data_integrity(self, df, date_col, value_col, oos_col):
        return self._values, self._resolution


class ControlMultiplierTests(unittest.TestCase):
    def setUp(self):
        self.calculator = CLTightenCalculator()
        self.today = pd.Timestamp.today().normalize()

    def get_k(self, count, create_time, kurtosis=10.0):
        return self.calculator.get_k_value(
            count,
            'Nominal',
            pattern='Normal',
            kurtosis_value=kurtosis,
            chart_create_time=create_time,
        )

    def test_4_to_15_points_use_chart_age(self):
        young = self.today - pd.DateOffset(months=6)
        exactly_one_year = self.today - pd.DateOffset(years=1)

        for count in (4, 15):
            self.assertEqual(self.get_k(count, young), 8.0)
            self.assertEqual(self.get_k(count, exactly_one_year), 5.0)
            self.assertEqual(self.get_k(count, pd.NaT), 5.0)
            self.assertEqual(self.get_k(count, 'invalid'), 5.0)

    def test_16_to_29_points_always_use_5_sigma(self):
        young = self.today - pd.DateOffset(days=1)
        for count in (16, 29):
            self.assertEqual(self.get_k(count, young), 5.0)
            self.assertEqual(self.get_k(count, pd.NaT), 5.0)

    def test_30_or_more_keeps_existing_kurtosis_rule(self):
        self.assertEqual(self.get_k(30, pd.NaT, kurtosis=0.0), 3.0)
        self.assertEqual(self.get_k(30, pd.NaT, kurtosis=1.1), 4.0)


class HardRuleResolutionTests(unittest.TestCase):
    def run_hard_rule(
        self,
        values,
        resolution,
        create_time,
        characteristic='Nominal',
        original_ucl=100.0,
        original_lcl=-100.0,
    ):
        calculator = FixedDataCalculator(values, resolution)
        df = pd.DataFrame(
            {
                'value': values,
                'date': [pd.Timestamp.today()] * len(values),
                'oos_flag': [False] * len(values),
                'DetectionLimit': [np.nan] * len(values),
                'Target': [0.0] * len(values),
                'UCL': [original_ucl] * len(values),
                'LCL': [original_lcl] * len(values),
                'Resolution': [resolution] * len(values),
            }
        )
        with contextlib.redirect_stdout(io.StringIO()):
            return calculator.process_chart(
                df,
                value_col='value',
                date_col='date',
                oos_col='oos_flag',
                characteristic=characteristic,
                chart_create_time=create_time,
            )

    def test_hard_rule_1_uses_one_or_two_resolutions(self):
        today = pd.Timestamp.today().normalize()
        old = today - pd.DateOffset(years=2)
        young = today - pd.DateOffset(months=6)

        old_result = self.run_hard_rule([10] * 5, 1.0, old)
        self.assertEqual(old_result['Suggest UCL'], 11.0)
        self.assertEqual(old_result['Suggest LCL'], 9.0)

        young_result = self.run_hard_rule([10] * 5, 1.0, young)
        self.assertEqual(young_result['Suggest UCL'], 12.0)
        self.assertEqual(young_result['Suggest LCL'], 8.0)

    def test_constant_data_falls_back_to_configured_resolution(self):
        calculator = CLTightenCalculator()
        old = pd.Timestamp.today().normalize() - pd.DateOffset(years=2)
        values = [10.0] * 5
        df = pd.DataFrame(
            {
                'value': values,
                'date': [pd.Timestamp.today()] * len(values),
                'oos_flag': [False] * len(values),
                'DetectionLimit': [np.nan] * len(values),
                'Target': [0.0] * len(values),
                'UCL': [100.0] * len(values),
                'LCL': [-100.0] * len(values),
                'Resolution': [0.5] * len(values),
            }
        )

        with contextlib.redirect_stdout(io.StringIO()):
            result = calculator.process_chart(
                df,
                value_col='value',
                date_col='date',
                oos_col='oos_flag',
                characteristic='Nominal',
                chart_create_time=old,
            )

        self.assertEqual(result['Resolution_Estimated'], 0.5)
        self.assertEqual(result['Suggest UCL'], 10.5)
        self.assertEqual(result['Suggest LCL'], 9.5)

    def test_hard_rules_2_and_3_use_one_resolution_at_16_points(self):
        old = pd.Timestamp.today().normalize() - pd.DateOffset(years=2)

        rule_2 = self.run_hard_rule([10, 11] * 8, 1.0, old)
        self.assertEqual(rule_2['HardRule'], 'Hard Rule 2: Two Categories')
        self.assertEqual(rule_2['Suggest UCL'], 12.0)
        self.assertEqual(rule_2['Suggest LCL'], 9.0)

        rule_3_values = [10, 11, 12, 10, 11, 12, 10, 11] * 2
        rule_3 = self.run_hard_rule(rule_3_values, 1.0, old)
        self.assertEqual(
            rule_3['HardRule'],
            'Hard Rule 3: Three Categories Spaced by Resolution',
        )
        self.assertEqual(rule_3['Suggest UCL'], 13.0)
        self.assertEqual(rule_3['Suggest LCL'], 9.0)

    def test_one_sided_characteristics_only_adjust_active_limit(self):
        old = pd.Timestamp.today().normalize() - pd.DateOffset(years=2)
        values = [10, 11] * 8

        smaller = self.run_hard_rule(
            values, 1.0, old, characteristic='Smaller'
        )
        self.assertEqual(smaller['Suggest UCL'], 12.0)
        self.assertEqual(smaller['Suggest LCL'], -100.0)

        bigger = self.run_hard_rule(
            values, 1.0, old, characteristic='Bigger'
        )
        self.assertEqual(bigger['Suggest UCL'], 100.0)
        self.assertEqual(bigger['Suggest LCL'], 9.0)

    def test_hard_rule_ooc_ignores_inactive_limit(self):
        old = pd.Timestamp.today().normalize() - pd.DateOffset(years=2)
        values = [10, 11] * 8

        smaller = self.run_hard_rule(
            values,
            1.0,
            old,
            characteristic='Smaller',
            original_ucl=100.0,
            original_lcl=10.5,
        )
        self.assertEqual(smaller['Ori_OOC_Count'], 0)
        self.assertEqual(smaller['Final_OOC_Count'], 0)

        bigger = self.run_hard_rule(
            values,
            1.0,
            old,
            characteristic='Bigger',
            original_ucl=10.5,
            original_lcl=-100.0,
        )
        self.assertEqual(bigger['Ori_OOC_Count'], 0)
        self.assertEqual(bigger['Final_OOC_Count'], 0)

    def test_clamp_recalculates_final_values_and_tighten_status(self):
        old = pd.Timestamp.today().normalize() - pd.DateOffset(years=2)
        result = self.run_hard_rule(
            [10, 11] * 8,
            1.0,
            old,
            original_ucl=11.5,
            original_lcl=10.5,
        )

        self.assertEqual(result['Suggest UCL'], 11.5)
        self.assertEqual(result['Suggest LCL'], 10.5)
        self.assertFalse(result['TightenNeeded'])
        self.assertEqual(result['Static_OOC_Count'], 8)

    def test_30_points_do_not_expand_hard_rule_limits(self):
        old = pd.Timestamp.today().normalize() - pd.DateOffset(years=2)
        result = self.run_hard_rule([10, 11] * 15, 1.0, old)

        self.assertEqual(result['Suggest UCL'], 11.0)
        self.assertEqual(result['Suggest LCL'], 10.0)


class OOCRetreatLimitTests(unittest.TestCase):
    def setUp(self):
        self.calculator = CLTightenCalculator()

    def adjust(self, pattern, resolution, sigma, max_adj_units=2):
        with contextlib.redirect_stdout(io.StringIO()):
            return self.calculator.adjust_CL_based_on_OOC(
                values=np.array([10.0, 10.0, 10.0, 10.0]),
                UCL=0.0,
                LCL=-100.0,
                pattern=pattern,
                resolution=resolution,
                sigma_est_u=sigma,
                sigma_est_l=sigma,
                max_adj_units=max_adj_units,
                characteristic='Smaller',
            )

    def test_discrete_retreat_uses_resolution_not_sigma(self):
        ucl, _, _, _, units = self.adjust(
            pattern='Attribute', resolution=1.0, sigma=0.3
        )
        self.assertEqual(ucl, 2.0)
        self.assertEqual(units, 2)

    def test_max_adj_units_controls_discrete_retreat_cap(self):
        ucl, _, _, _, units = self.adjust(
            pattern='Attribute', resolution=1.0, sigma=0.3, max_adj_units=1
        )
        self.assertEqual(ucl, 1.0)
        self.assertEqual(units, 1)

    def test_continuous_retreat_remains_capped_by_sigma(self):
        ucl, _, _, _, units = self.adjust(
            pattern='Normal', resolution=None, sigma=2.0
        )
        self.assertEqual(ucl, 4.0)
        self.assertEqual(units, 2.0)

    def test_ooc_count_uses_only_active_one_sided_limit(self):
        values = np.array([-10.0, 0.0, 5.0, 20.0])
        self.assertEqual(
            self.calculator.count_ooc(
                values, ucl=10.0, lcl=-5.0, characteristic='Smaller'
            ),
            1,
        )
        self.assertEqual(
            self.calculator.count_ooc(
                values, ucl=10.0, lcl=-5.0, characteristic='Bigger'
            ),
            1,
        )
        self.assertEqual(
            self.calculator.count_ooc(
                values, ucl=10.0, lcl=-5.0, characteristic='Nominal'
            ),
            2,
        )

    def test_one_sided_count_does_not_require_inactive_limit(self):
        values = np.array([-10.0, 0.0, 5.0, 20.0])
        self.assertEqual(
            self.calculator.count_ooc(
                values, ucl=10.0, lcl=np.nan, characteristic='Smaller'
            ),
            1,
        )
        self.assertEqual(
            self.calculator.count_ooc(
                values, ucl=np.nan, lcl=-5.0, characteristic='Bigger'
            ),
            1,
        )

    def test_counts_are_recomputed_after_final_precision_lock(self):
        values = np.array([0.0, 1.02, 2.0, 3.0])
        calculator = FixedDataCalculator(values, resolution=0.1)
        df = pd.DataFrame({
            'value': values,
            'date': pd.date_range('2026-08-01', periods=4, freq='D'),
            'oos_flag': [False] * 4,
            'DetectionLimit': [np.nan] * 4,
            'UCL': [np.nan] * 4,
            'LCL': [np.nan] * 4,
        })
        prepared = (values, values, 'Normal', 0.0, 0.0)
        calculated_limits = (
            1.04, -10.0, 1.0, 1.0, np.nan, np.nan, np.nan, np.nan
        )

        with patch.object(
            calculator, 'prepare_pattern_and_outliers', return_value=prepared
        ), patch.object(
            calculator, 'calc_CL', return_value=calculated_limits
        ), patch.object(
            calculator, 'adjust_CL_based_on_OOC',
            return_value=(1.04, -10.0, 999, 0, 0.0)
        ), contextlib.redirect_stdout(io.StringIO()):
            result = calculator.process_chart(
                df, 'value', 'date', 'oos_flag', 'Nominal'
            )

        self.assertEqual(result['Static UCL'], 1.0)
        self.assertEqual(result['Suggest UCL'], 1.0)
        self.assertEqual(result['Static_OOC_Count'], 3)
        self.assertEqual(result['Final_OOC_Count'], 3)

    def test_nonfinite_values_are_removed_before_resolution_and_analysis(self):
        calculator = CLTightenCalculator()
        df = pd.DataFrame({
            'value': [1.0, np.inf, 2.0, -np.inf, 3.0],
            'date': [pd.Timestamp.today()] * 5,
            'oos_flag': [False] * 5,
        })

        with contextlib.redirect_stdout(io.StringIO()):
            values, resolution = calculator.data_integrity(
                df, 'date', 'value', 'oos_flag'
            )

        np.testing.assert_array_equal(values, np.array([1.0, 2.0, 3.0]))
        self.assertEqual(resolution, 1.0)

    def test_fewer_than_four_points_after_filter_returns_cleanly(self):
        values = np.arange(10, dtype=float)
        calculator = FixedDataCalculator(values, resolution=1.0)
        df = pd.DataFrame({
            'value': values,
            'date': [pd.Timestamp.today()] * len(values),
            'oos_flag': [False] * len(values),
            'DetectionLimit': [np.nan] * len(values),
            'UCL': [20.0] * len(values),
            'LCL': [-20.0] * len(values),
        })
        filtered = np.array([1.0, 2.0, 3.0])

        with patch.object(
            calculator,
            'prepare_pattern_and_outliers',
            return_value=(filtered, filtered, 'Insufficient Data', np.nan, np.nan),
        ), contextlib.redirect_stdout(io.StringIO()):
            result = calculator.process_chart(
                df, 'value', 'date', 'oos_flag', 'Nominal'
            )

        self.assertEqual(result['Pattern'], 'Insufficient Data After Filter')
        self.assertEqual(result['DataCountUsed'], 3)
        self.assertTrue(np.isnan(result['Suggest UCL']))

    def test_final_precision_recomputes_tighten_tolerance_and_k_metrics(self):
        values = np.array([0.0, 1.0, 2.0, 3.0])
        calculator = FixedDataCalculator(values, resolution=1.0)
        df = pd.DataFrame({
            'value': values,
            'date': [pd.Timestamp.today()] * 4,
            'oos_flag': [False] * 4,
            'DetectionLimit': [np.nan] * 4,
            'Target': [1.5] * 4,
            'UCL': [4.0] * 4,
            'LCL': [-2.0] * 4,
        })
        calculated_limits = (
            2.6, -0.6, 1.0, 1.0, np.nan, np.nan, np.nan, np.nan
        )

        with patch.object(
            calculator,
            'prepare_pattern_and_outliers',
            return_value=(values, values, 'Normal', 0.0, 0.0),
        ), patch.object(
            calculator, 'calc_CL', return_value=calculated_limits
        ), patch.object(
            calculator,
            'adjust_CL_based_on_OOC',
            return_value=(2.6, -0.6, 0, 0, 0.0),
        ), contextlib.redirect_stdout(io.StringIO()):
            result = calculator.process_chart(
                df, 'value', 'date', 'oos_flag', 'Nominal'
            )

        self.assertEqual(result['Suggest UCL'], 3)
        self.assertEqual(result['Suggest LCL'], -1)
        self.assertEqual(result['CL_Center'], 2)
        self.assertEqual(result['New_Tolerance'], 4.0)
        self.assertAlmostEqual(result['Diff_Ratio_%'], 100 / 3)
        self.assertFalse(result['TightenNeeded'])
        sigma = np.std(values, ddof=1)
        self.assertAlmostEqual(result['Suggest_UCL_K_Set'], 1 / sigma)
        self.assertAlmostEqual(result['Suggest_LCL_K_Set'], 3 / sigma)
        self.assertAlmostEqual(result['Sug_K_Set'], 3 / sigma)


class OutputFieldNamingTests(unittest.TestCase):
    def test_original_limits_preserve_chart_settings(self):
        calculator = CLTightenCalculator()
        chart_info = pd.Series({
            'GroupName': 'G1',
            'ChartName': 'C1',
            'UCL': 12.3456,
            'LCL': 3.2109,
            'USL': np.nan,
            'LSL': np.nan,
            'Target': 8.0,
            'Characteristics': 'Nominal',
        })
        raw_data = pd.DataFrame({
            'point_time': pd.date_range('2026-08-01', periods=4, freq='D'),
            'point_val': [7.0, 8.0, 9.0, 10.0],
        })
        calculated = {
            'Suggest UCL': 11.0,
            'Suggest LCL': 5.0,
            'Static UCL': 10.0,
            'Static LCL': 6.0,
            'CL_Center': 8.0,
            'Pattern': 'Normal',
            'Resolution_Estimated': None,
            'TotalDataCount': 4,
            'DataCountUsed': 4,
        }

        with patch.object(
            calculator, 'process_chart', return_value=calculated
        ), patch.object(
            calculator, 'plot_control_chart', return_value=None
        ), contextlib.redirect_stdout(io.StringIO()):
            output = calculator.process_single_chart_data(chart_info, raw_data)

        self.assertEqual(output['Original UCL'], 12.3456)
        self.assertEqual(output['Original LCL'], 3.2109)
        self.assertEqual(output['Suggest UCL'], 11.0)
        self.assertEqual(output['Suggest LCL'], 5.0)


if __name__ == '__main__':
    unittest.main()
