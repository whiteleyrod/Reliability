from __future__ import annotations

import unittest
from unittest.mock import patch

import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
from matplotlib.collections import PolyCollection

from app import (
    app,
    build_bland_altman_plot,
    build_html_report,
    build_regression_confidence_band,
    build_scatter_plot,
    build_typical_error_table,
    detect_reliability_pairs,
    extract_ci_bounds,
)


class ReliabilityAppTests(unittest.TestCase):
    def test_detect_reliability_pairs_finds_and_sorts_pairs(self) -> None:
        columns = [
            "Knee Flexion Test 2",
            "Ankle ROM Test 1",
            "Knee Flexion Test 1",
            "Ignore Me",
            "Ankle ROM Test 2",
        ]

        pairs = detect_reliability_pairs(columns)

        self.assertEqual(len(pairs), 2)
        self.assertEqual(pairs[0]["label"], "Ankle ROM")
        self.assertEqual(pairs[1]["label"], "Knee Flexion")
        self.assertEqual(pairs[1]["test_1"], "Knee Flexion Test 1")
        self.assertEqual(pairs[1]["test_2"], "Knee Flexion Test 2")

    def test_extract_ci_bounds_handles_string_and_sequence_inputs(self) -> None:
        self.assertEqual(extract_ci_bounds("[0.12, 0.98]"), (0.12, 0.98))
        self.assertEqual(extract_ci_bounds((0.33, 0.77)), (0.33, 0.77))
        self.assertEqual(extract_ci_bounds("invalid"), (None, None))

    def test_regression_confidence_band_requires_non_constant_x(self) -> None:
        x_values = np.array([5.0, 5.0, 5.0])
        y_values = np.array([4.0, 5.0, 6.0])

        band = build_regression_confidence_band(x_values, y_values)

        self.assertIsNone(band)

    def test_scatter_plot_is_square_and_includes_confidence_band(self) -> None:
        dataframe = pd.DataFrame(
            {
                "Test 1": [10.0, 11.5, 13.0, 14.0, 15.5],
                "Test 2": [10.3, 11.2, 13.4, 14.6, 15.2],
            }
        )

        figure = build_scatter_plot(dataframe, "Test 1", "Test 2")
        self.addCleanup(plt.close, figure)
        axis = figure.axes[0]

        self.assertEqual(axis.get_aspect(), 1.0)
        self.assertTrue(any(isinstance(collection, PolyCollection) for collection in axis.collections))
        self.assertEqual(len(axis.lines), 2)

    def test_bland_altman_plot_uses_symmetric_y_limits(self) -> None:
        dataframe = pd.DataFrame(
            {
                "Test 1": [8.0, 9.5, 11.0, 12.5, 14.0],
                "Test 2": [8.4, 9.2, 10.8, 12.9, 14.3],
            }
        )

        figure = build_bland_altman_plot(dataframe, "Test 1", "Test 2")
        self.addCleanup(plt.close, figure)
        axis = figure.axes[0]
        lower, upper = axis.get_ylim()

        self.assertAlmostEqual(abs(lower), abs(upper), places=6)

    def test_typical_error_table_includes_mdc_95(self) -> None:
        dataframe = pd.DataFrame(
            {
                "Test 1": [10.0, 12.0, 14.0, 15.0],
                "Test 2": [11.0, 13.5, 13.5, 16.0],
            }
        )

        metrics = build_typical_error_table(dataframe)

        self.assertEqual(len(metrics), 1)
        expected_sd = dataframe["Test 2"].sub(dataframe["Test 1"]).std(ddof=1)
        expected_te = round(float(expected_sd / np.sqrt(2)), 4)
        expected_mdc = round(float((expected_sd / np.sqrt(2)) * 1.96 * np.sqrt(2)), 4)
        self.assertEqual(metrics[0]["typical_error"], expected_te)
        self.assertEqual(metrics[0]["minimum_detectable_change_95"], expected_mdc)

    def test_build_html_report_includes_metrics_and_figures(self) -> None:
        upload_record = {
            "id": "upload-test-html",
            "original_filename": "test.csv",
        }
        analysis_record = {
            "id": "analysis-test-html",
            "config": {
                "upload_id": "upload-test-html",
                "selected_sheet": "Sheet1",
                "subject_column": "Subject",
                "selected_pair_labels": ["Test 1 vs Test 2"],
            },
            "recommendation": {
                "design_label": "Two-way random",
                "agreement_label": "Absolute agreement",
                "measurement_label": "Single measurement",
                "rationale": "Example rationale",
            },
            "dataset_summary": {"pair_count": 1, "source_rows": 3},
            "pair_results": [
                {
                    "pair_key": "pair-1",
                    "pair_label": "Test 1 vs Test 2",
                    "primary_x_column": "Test 1",
                    "primary_y_column": "Test 2",
                    "icc_result": {
                        "model": "ICC2",
                        "estimate": 0.95,
                        "ci_lower": 0.8,
                        "ci_upper": 0.99,
                        "f_value": 10.0,
                        "p_value": 0.01,
                        "description": "Example",
                    },
                    "dataset_summary": {"observations": 3, "dropped_rows": 0},
                    "overall_summary": {"count": 6, "mean": 10.5, "sd": 1.0, "median": 10.5, "iqr": 1.0, "min": 9.0, "max": 12.0, "residual_mse": 0.25},
                    "column_summaries": [{"name": "Test 1", "count": 3, "mean": 10.0, "sd": 1.0, "median": 10.0, "iqr": 1.0, "min": 9.0, "max": 11.0}],
                    "observation_summaries": [{"observation": "A", "mean": 10.0, "sd": 0.5, "median": 10.0, "iqr": 0.5, "min": 9.5, "max": 10.5}],
                    "pair_metrics": [{"pair": "Test 1 vs Test 2", "bias": 0.1, "typical_error": 0.2, "minimum_detectable_change_95": 0.5544, "loa_lower": -0.5, "loa_upper": 0.7}],
                    "typical_error_formula": "Typical error = SD(differences) / √2",
                    "minimum_detectable_change_formula": "Minimum detectable change (95%) = Typical error × 1.96 × √2",
                }
            ],
        }

        source_frame = pd.DataFrame({"Subject": ["A"], "Test 1": [10.0], "Test 2": [10.1]})

        with patch("app.load_upload", return_value=upload_record), patch("app.build_source_data_export_frame", return_value=source_frame), patch("app.build_source_data_frame", return_value=source_frame):
            report = build_html_report(analysis_record, "http://127.0.0.1:5000")

        self.assertIn("<!doctype html>", report)
        self.assertIn("Minimum detectable change formula", report)
        self.assertIn("/plots/analysis-test-html/pair-1/scatter.svg", report)
        self.assertIn("reliability Analysis Report".lower(), report.lower())

    def test_health_endpoint_returns_ok(self) -> None:
        app.testing = True

        with app.test_client() as client:
            response = client.get("/healthz")

        self.assertEqual(response.status_code, 200)
        self.assertEqual(response.get_json(), {"status": "ok", "service": "reliability-web-tool"})


if __name__ == "__main__":
    unittest.main()
