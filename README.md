# Reliability Web Tool

Flask-based starter app for inter-rater reliability analysis.

Current build includes:

- CSV and XLSX upload
- worksheet scanning and column heading discovery
- wide-format column selection for reliability analysis
- multiple guessed Test 1 / Test 2 pair selection in one run
- optional pre-upload user guide on the landing page
- ICC suggestion and calculation
- 95% confidence intervals
- descriptive summaries for the full analysed group
- per-observation summaries
- typical error, bias, and limits of agreement
- square scatter plots with a `y = x` reference line
- Bland-Altman plots centered symmetrically around 0 on the y-axis
- SVG preview and SVG/PDF download routes
- single PDF report export with analysed data first, analysis description, package list, commands used, results, and figures
- CSV export of the analysed source data used in each reliability analysis
- deployment-ready WSGI runner with `waitress`
- Docker, Compose, and Procfile deployment support

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

## Production-style local run

```powershell
python run_web.py
```

Then open http://127.0.0.1:8000

## Workflow

1. Upload a `.csv` or `.xlsx` file.
2. Review detected worksheets and column headings.
3. Choose the worksheet to analyse.
4. Select an observation ID column if one exists.
5. Select one or more guessed reliability pairs, or use the manual X/Y pair fallback.
6. Review the selected columns and primary pair settings.
7. Confirm the ICC design settings and figure action.
8. Run the analysis.
9. Choose whether to keep viewing results in the browser or download a single PDF report.

## Installed packages

- Flask
- pandas
- openpyxl
- pingouin
- matplotlib
- seaborn

## Deployment note

This remains a Flask web app and now includes [wsgi.py](wsgi.py), [run_web.py](run_web.py), [Dockerfile](Dockerfile), [compose.yaml](compose.yaml), and [Procfile](Procfile) for deployment.

## Recommended deployment hosts

Best open-source-friendly options for this app:

1. **Coolify**
	- open-source, self-hostable platform
	- connect the GitHub repo and deploy from the included [Dockerfile](Dockerfile)
	- best fit if using a VPS and wanting a modern web UI

2. **Dokku**
	- open-source Heroku-style self-hosted platform
	- can deploy from Git or via the included [Procfile](Procfile)
	- lightweight option for a small VPS

3. **Any Docker host**
	- deploy with the included [Dockerfile](Dockerfile)
	- examples: self-hosted VPS, Coolify, CapRover, or Kubernetes-based hosting

Managed fallback options if wanted later:

- Render
- Railway
- Azure App Service

## Docker test

```powershell
docker compose up --build
```

Then open http://127.0.0.1:8000
