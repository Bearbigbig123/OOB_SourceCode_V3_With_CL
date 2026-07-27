import contextlib
import io
import unittest

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


if __name__ == '__main__':
    unittest.main()
