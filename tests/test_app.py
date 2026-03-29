from __future__ import annotations

import unittest

import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
from matplotlib.collections import PolyCollection

from app import (
    app,
    build_bland_altman_plot,
    build_regression_confidence_band,
    build_scatter_plot,
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

    def test_health_endpoint_returns_ok(self) -> None:
        app.testing = True

        with app.test_client() as client:
            response = client.get("/healthz")

        self.assertEqual(response.status_code, 200)
        self.assertEqual(response.get_json(), {"status": "ok", "service": "reliability-web-tool"})


if __name__ == "__main__":
    unittest.main()
