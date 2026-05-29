# Correlation Fragility

A Streamlit web application for testing the fragility of Pearson correlations by identifying which data points, when replaced with the dataset median, most damage (or destroy) a statistically significant correlation.

## Background

This tool implements a brute-force approach to correlation fragility analysis. Given a dataset with a statistically significant Pearson correlation, it exhaustively tests every possible combination of `n` data point replacements (substituting selected points with the median X/Y values) to find the combination that produces the highest (least significant) p-value. The result reveals how robust — or fragile — the original correlation truly is.

## Features

- Upload any two-column CSV file (no headers required)
- Select the number of replacements `n` to test
- Visual output: fragile r and p-value line charts across all combinations
- Visual output: scatter plot highlighting the most fragility-inducing data points
- Downloadable PDF exports of both charts
- Summary statistics: original r/p, total combinations tested, and worst-case replacement

## Requirements

### Python

Python 3.9 or later is recommended.

### Dependencies

Install all dependencies with:

```bash
pip install -r requirements.txt
```

**`requirements.txt`**

```
streamlit
numpy
scipy
matplotlib
```

### Input Data Format

- CSV file with **two columns** (X values and Y values)
- **No header row**
- Comma-separated
- Compatible with the default export format from [WebPlotDigitizer](https://automeris.io/WebPlotDigitizer/)

Example:
```
1.2,3.4
2.5,5.1
3.8,6.9
```

## Usage

```bash
streamlit run app.py
```

Then open the URL shown in the terminal (typically `http://localhost:8501`).

1. Upload your CSV file using the file uploader in the sidebar
2. Set the number of replacements `n` (minimum 1)
3. Click **Run Analysis**
4. Review the charts and summary, and download PDFs if needed

## Performance Note

The number of combinations tested is $\binom{N}{n}$ where $N$ is the number of data points. This grows rapidly:

| Dataset size | n=1 | n=2 | n=3 |
|---|---|---|---|
| 20 points | 20 | 190 | 1,140 |
| 50 points | 50 | 1,225 | 19,600 |
| 100 points | 100 | 4,950 | 161,700 |

For large datasets with `n > 2`, computation may take significant time. The app will warn you before running if the combination count exceeds a configurable threshold.

## Output

- **Line chart**: Pearson r and p-value across all tested combinations, with the original values shown as dotted reference lines
- **Scatter plot**: Original data in blue, the most fragility-inducing points highlighted in red, and the median replacement point in orange
- **Summary**: The maximum achievable p-value, the corresponding r, and which data points to replace
