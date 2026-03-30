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
- square scatter plots with a `y = x` reference line plus shaded 95% confidence intervals around the line of best fit
- Bland-Altman plots centered symmetrically around 0 on the y-axis
- SVG preview and SVG/PDF download routes
- single PDF report export with analysed data first, analysis description, package list, commands used, results, and figures
- DOCX report export with figures embedded as SVG image parts plus Word-compatible PNG fallbacks
- CSV export of the analysed source data used in each reliability analysis
- health endpoint for container and platform checks at `/healthz`
- basic automated CI coverage for plotting and core helper functions
- reverse-proxy-safe configuration for subdomain deployment
- Windows standalone packaging support with PyInstaller
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
9. Choose whether to keep viewing results in the browser or download a PDF or DOCX report.

## Installed packages

- Flask
- pandas
- openpyxl
- pingouin
- matplotlib
- seaborn
- python-docx

## Automated checks

The repo now includes a GitHub Actions workflow that installs dependencies and runs:

```powershell
python -m unittest discover -s tests -p "test_*.py"
```

Local smoke-check before pushing:

```powershell
python -m unittest discover -s tests -p "test_*.py"
```

## Build a Windows standalone app

This repo can also be packaged into a standalone Windows `.exe` that starts the app locally and opens it in the user's browser.

Build steps:

```powershell
.\build_windows_exe.ps1
```

Build output:

- [dist/Reliability.exe](dist/Reliability.exe) after a successful local build

Packaging notes:

- the executable bundles Python, templates, static assets, and the app code
- uploaded files and analysis outputs are written to the user's local app data folder when running as a packaged app
- the packaged app opens a local browser session automatically

## Deployment note

This remains a Flask web app and now includes [wsgi.py](wsgi.py), [run_web.py](run_web.py), [Dockerfile](Dockerfile), [compose.yaml](compose.yaml), and [Procfile](Procfile) for deployment.

For a production subdomain deployment, the repo also includes:

- [deploy/compose.production.yaml](deploy/compose.production.yaml)
- [deploy/Caddyfile.example](deploy/Caddyfile.example)

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

## Subdomain deployment: reliability.whiteley.work

Recommended production shape:

- app container runs locally on the server at `127.0.0.1:8000`
- Caddy terminates HTTPS for `reliability.whiteley.work`
- Caddy reverse-proxies traffic to the Flask app

### 1. DNS

Create a DNS record for the subdomain:

- `A` record: `reliability.whiteley.work` → your server IPv4 address
- optional `AAAA` record: `reliability.whiteley.work` → your server IPv6 address

### 2. Start the app on the server

From the server checkout of this repo:

```powershell
docker compose -f deploy/compose.production.yaml up -d --build
```

This production compose file:

- binds the app to `127.0.0.1:8000` only
- enables proxy-aware handling with `TRUST_PROXY=1`
- sets `PREFERRED_URL_SCHEME=https`

### 3. Configure Caddy

Copy [deploy/Caddyfile.example](deploy/Caddyfile.example) into your live Caddy config and reload Caddy.

Example site block:

```caddyfile
reliability.whiteley.work {
	encode zstd gzip
	reverse_proxy 127.0.0.1:8000
}
```

### 4. Open ports

Ensure the server allows inbound:

- `80/tcp`
- `443/tcp`

### 5. Verify the deployment

Check these URLs after DNS and proxy setup:

- `https://reliability.whiteley.work/`
- `https://reliability.whiteley.work/healthz`

Expected health response:

```json
{"service":"reliability-web-tool","status":"ok"}
```

### 6. Update the app after future pushes

On the server:

```powershell
git pull
docker compose -f deploy/compose.production.yaml up -d --build
```

### What is still manual

This repo is now prepared for subdomain deployment, but these steps still have to be performed on the actual host or DNS provider:

- create the `reliability.whiteley.work` DNS record
- install or configure Caddy on the server
- reload the reverse proxy after adding the site block
- run the production compose stack on the server

## Docker test

```powershell
docker compose up --build
```

Then open http://127.0.0.1:8000

Container health can be checked at http://127.0.0.1:8000/healthz
