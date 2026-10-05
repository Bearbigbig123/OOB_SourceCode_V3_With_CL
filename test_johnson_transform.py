import numpy as np
import unittest
import contextlib
import io
from unittest.mock import patch
from scipy.stats import kurtosis, norm, skew

from CL_limit_class import (
    CLTightenCalculator,
    estimate_johnson_slifker_shapiro_parameters,
    transform_johnson_slifker_shapiro_full,
)


def test_su_parameters_match_sas_formula():
    m, n, p = 3.0, 4.0, 2.0
    x1z, xm1z, zval = 2.0, 0.0, 0.524
    family, parms = estimate_johnson_slifker_shapiro_parameters(
        m, n, p, x1z, xm1z, z_val=zval
    )

    temp = 0.5 * (m / p + n / p)
    eta = 2 * zval / np.log(temp + np.sqrt(temp * temp - 1))
    temp = (n / p - m / p) / (2 * np.sqrt(m * n / p**2 - 1))
    gamma = eta * np.log(temp + np.sqrt(temp * temp + 1))
    lam = 2 * p * np.sqrt(m * n / p**2 - 1) / (
        (m / p + n / p - 2) * np.sqrt(m / p + n / p + 2)
    )
    epsilon = (x1z + xm1z) / 2 + p * (n / p - m / p) / (
        2 * (m / p + n / p - 2)
    )

    assert family == "SU"
    assert np.isclose(parms["eta"], eta)
    assert np.isclose(parms["gamma"], gamma)
    assert np.isclose(parms["lambda"], lam)
    assert np.isclose(parms["epsilon"], epsilon)


def test_sb_parameters_match_sas_formula():
    m, n, p = 1.5, 1.2, 2.0
    x1z, xm1z, zval = 2.0, 0.0, 0.524
    family, parms = estimate_johnson_slifker_shapiro_parameters(
        m, n, p, x1z, xm1z, z_val=zval
    )

    product = (1 + p / m) * (1 + p / n)
    temp = 0.5 * np.sqrt(product)
    eta = zval / np.log(temp + np.sqrt(temp * temp - 1))
    temp = (p / n - p / m) * np.sqrt(product - 4) / (
        2 * (p * p / (m * n) - 1)
    )
    gamma = eta * np.log(temp + np.sqrt(temp * temp + 1))
    lam = p * np.sqrt((product - 2) ** 2 - 4) / (p * p / (m * n) - 1)
    epsilon = (x1z + xm1z) / 2 - lam / 2 + p * (p / n - p / m) / (
        2 * (p * p / (m * n) - 1)
    )

    assert family == "SB"
    assert np.isclose(parms["eta"], eta)
    assert np.isclose(parms["gamma"], gamma)
    assert np.isclose(parms["lambda"], lam)
    assert np.isclose(parms["epsilon"], epsilon)


def test_sl_parameters_match_sas_formula():
    m, n, p = 2.02, 1.98, 2.0
    x1z, xm1z, zval = 2.0, 0.0, 0.524
    family, parms = estimate_johnson_slifker_shapiro_parameters(
        m, n, p, x1z, xm1z, z_val=zval
    )

    eta = 2 * zval / np.log(m / p)
    gamma = eta * np.log((m / p - 1) / (p * np.sqrt(m / p)))
    epsilon = (x1z + xm1z) / 2 - (p / 2) * (m / p + 1) / (m / p - 1)

    assert family == "SL"
    assert np.isclose(parms["eta"], eta)
    assert np.isclose(parms["gamma"], gamma)
    assert parms["lambda"] is None
    assert np.isclose(parms["epsilon"], epsilon)


def test_each_johnson_family_transforms_back_to_normal_scores():
    scores = norm.ppf((np.arange(1, 5001) - 0.5) / 5000)
    samples = {
        "SU": 1.2 + 2.5 * np.sinh((scores - 0.3) / 1.4),
        "SB": -1.0 + 5.0 / (1 + np.exp(-(scores + 0.2) / 1.1)),
        "SL": 0.5 + np.exp((scores - 0.1) / 0.9),
    }

    for expected_family, sample in samples.items():
        transformed, family = transform_johnson_slifker_shapiro_full(sample)
        assert family == expected_family
        assert np.corrcoef(scores, transformed)[0, 1] > 0.999


class JohnsonTransformTests(unittest.TestCase):
    def test_su_parameters(self):
        test_su_parameters_match_sas_formula()

    def test_sb_parameters(self):
        test_sb_parameters_match_sas_formula()

    def test_sl_parameters(self):
        test_sl_parameters_match_sas_formula()

    def test_family_transformations(self):
        test_each_johnson_family_transforms_back_to_normal_scores()

    def test_johnson_failure_raises_instead_of_rank_int(self):
        with self.assertRaises(ValueError):
            transform_johnson_slifker_shapiro_full(np.ones(20))
        with self.assertRaises(ValueError):
            transform_johnson_slifker_shapiro_full(np.arange(9))

    def test_failure_fallback_filters_all_original_values_above_robust_z_4_5(self):
        calculator = CLTightenCalculator()
        values = np.concatenate([np.linspace(-1, 1, 20), [50.0, 100.0]])
        expected = calculator.filter_original_by_robust_z(values, threshold=4.5)
        self.assertNotIn(50.0, expected)
        self.assertNotIn(100.0, expected)
        self.assertEqual(len(expected), 20)

        with patch(
            "CL_limit_class.transform_johnson_slifker_shapiro_full",
            side_effect=ValueError("forced Johnson failure"),
        ):
            with contextlib.redirect_stdout(io.StringIO()):
                pattern_result = calculator.data_prep_for_pattern(values)
                outlier_result = calculator.outlier_filter(values, "Other")

        np.testing.assert_array_equal(pattern_result, expected)
        np.testing.assert_array_equal(outlier_result, expected)

    def test_discrete_data_skips_johnson_and_uses_original_robust_z_4_5(self):
        calculator = CLTightenCalculator()
        values = np.array([0.0] * 10 + [1.0] * 10 + [2.0] * 9 + [100.0])

        self.assertTrue(calculator.is_discrete_data(values))
        expected = calculator.filter_original_by_robust_z(values, threshold=4.5)

        with patch(
            "CL_limit_class.transform_johnson_slifker_shapiro_full",
            side_effect=AssertionError("Johnson must not run for discrete data"),
        ):
            with contextlib.redirect_stdout(io.StringIO()):
                pattern_values, filtered, pattern, _, _ = (
                    calculator.prepare_pattern_and_outliers(values, resolution=1.0)
                )

        np.testing.assert_array_equal(pattern_values, expected)
        np.testing.assert_array_equal(filtered, expected)
        self.assertNotIn(100.0, filtered)
        self.assertEqual(pattern, "Attribute")

    def test_continuous_data_is_not_misclassified_as_discrete(self):
        calculator = CLTightenCalculator()
        self.assertFalse(calculator.is_discrete_data(np.linspace(0.0, 1.0, 30)))

    def test_bimodality_coefficient_uses_finite_sample_correction(self):
        calculator = CLTightenCalculator()
        values = np.array([
            0.1, 0.4, 0.9, 1.1, 1.8, 2.0, 2.7, 3.2, 3.8, 4.0,
            4.9, 5.1, 5.8, 6.3, 7.0, 7.4, 8.2, 8.9, 9.1, 10.0,
        ])
        sample_skew = skew(values, bias=False)
        sample_excess_kurtosis = kurtosis(values, fisher=True, bias=False)
        sample_size = len(values)
        correction = (
            3 * (sample_size - 1) ** 2
            / ((sample_size - 2) * (sample_size - 3))
        )
        expected = (sample_skew ** 2 + 1) / (
            sample_excess_kurtosis + correction
        )

        self.assertAlmostEqual(calculator.compute_CB(values), expected, places=12)
        legacy = (sample_skew ** 2 + 1) / (sample_excess_kurtosis + 3)
        self.assertNotAlmostEqual(calculator.compute_CB(values), legacy, places=6)
