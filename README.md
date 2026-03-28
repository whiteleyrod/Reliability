# Reliability Web Tool

Flask-based starter app for inter-rater reliability analysis.

Current build includes:

- CSV and XLSX upload
- worksheet scanning and column heading discovery
- wide-format column selection for reliability analysis
- ICC suggestion and calculation
- 95% confidence intervals
- descriptive summaries for the full analysed group
- per-observation summaries
- typical error, bias, and limits of agreement
- square scatter plots with a `y = x` reference line
- Bland-Altman plots centered symmetrically around 0 on the y-axis
- SVG preview and SVG/PDF download routes
- optional Markdown report export including analysed source data and method summary
- CSV export of the analysed source data used in each reliability analysis

Sample workbook for testing:

- [SampleData/All Metrics Cleaned.xlsx](SampleData/All%20Metrics%20Cleaned.xlsx)
- primary sample sheet: `Variables`
- the app now auto-detects `Test 1` / `Test 2` column pairs from this sheet

## Activate the virtual environment

### PowerShell

```powershell
.\.venv\Scripts\Activate.ps1
```

### Command Prompt

```bat
.\.venv\Scripts\activate.bat
```

## Run the app

```powershell
python app.py
```

Then open http://127.0.0.1:5000

## Workflow

1. Upload a `.csv` or `.xlsx` file.
2. Review detected worksheets and column headings.
3. Choose the worksheet to analyse.
4. Select an observation ID column if one exists.
5. Select two or more numeric measurement columns.
6. Choose the primary pair for plots.
7. Confirm the ICC design settings and figure action.
8. Run the analysis.

## Installed packages

- Flask
- pandas
- openpyxl
- pingouin
- matplotlib
- seaborn

## Deployment note

This remains a Flask web app and now includes [wsgi.py](wsgi.py) for future deployment to a personal website or other WSGI-compatible hosting.
