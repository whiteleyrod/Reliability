# Project Plan: Web-Based Inter-Rater Reliability Tool

## Implementation Status Snapshot

This plan now also serves as an implementation record.

### Implemented in the current build

- Flask web application scaffolded and running locally
- Virtual environment and Python dependency management established
- CSV and XLSX upload support
- Worksheet scanning for Excel workbooks
- Uploaded data preview in the browser
- Automatic detection of likely `Test 1` / `Test 2` reliability pairs
- Clean pair-selection interface with support for multiple selected pairs in one run
- Manual X/Y fallback when no detected pair is selected
- Observation identifier selection
- ICC design, agreement, and measurement-unit selection
- ICC recommendation rationale shown in plain language
- Tooltip guidance for study design, agreement target, and measurement unit choices
- ICC calculation using `pingouin`
- 95% confidence intervals displayed correctly
- Descriptive summaries for the analysed sample
- Descriptive summaries for each selected series
- Observation-level descriptive summaries
- Median and IQR included in descriptive outputs
- Residual mean square error included at the analysis-level descriptive summary
- Typical error, bias, and limits of agreement calculations
- Square scatter plots with `y = x`
- Bland-Altman plots centred symmetrically around `0`
- SVG preview and SVG/PDF figure download routes
- PDF report export including data, methods, results, and figures
- DOCX report export including embedded SVG figure parts with PNG fallbacks for Word compatibility
- CSV export of analysed source data
- Post-analysis export chooser in the UI
- Optional landing-page user guide
- Health-check endpoint for deployment verification
- Basic automated CI workflow for helper and plotting smoke tests
- Reverse-proxy-safe app configuration for subdomain deployment
- Windows standalone packaging path via PyInstaller
- Deployment support via `waitress`, `wsgi.py`, `run_web.py`, `Dockerfile`, `compose.yaml`, and `Procfile`
- GitHub repository setup and pushed updates

### Partially implemented or still open

- Long-format data mapping is not yet implemented as a full workflow
- Missing-data handling is currently drop-complete-cases per selected pair, with reporting but limited user options
- Automated test coverage is still limited, though a basic CI workflow now runs helper and plotting smoke tests
- Interpretation guidance is present in a limited form and could be expanded

## 1. Goal

Build a web-based application that allows users to upload `.xlsx` or `.csv` files, select the relevant columns for analysis, confirm the structure of raters and observations, receive guidance on the appropriate intraclass correlation coefficient (ICC) approach, and generate reliability outputs including:

- ICC estimate
- 95% confidence intervals
- descriptive summaries including mean, SD, median, IQR, and residual mean square error
- Scatter plots
- Bland-Altman plots
- PDF and DOCX report exports

## 2. Core User Workflow

1. User uploads a data file (`.csv` or `.xlsx`).
2. Application scans the available worksheet or table structure and reads column headings.
3. Application previews the dataset.
4. User selects the relevant worksheet when needed.
5. Application presents a clean interface for selecting reliability data pairs or mapped rater/value columns.
6. Application asks the user to confirm:
   - which column identifies observations/subjects
   - which columns or rows represent raters
   - whether raters are fixed or random
   - whether the goal is absolute agreement or consistency
   - whether the result should be for single measures or average measures
7. Application suggests the most appropriate ICC model.
8. User reviews and confirms the suggested model.
9. Application calculates:
   - ICC value
   - 95% confidence interval
  - descriptive summaries for the entire group
  - descriptive summaries for each selected series
  - descriptive summaries for each set of observations
  - residual mean square error for the overall selected analysis
  - typical error metrics
   - summary statistics
10. Application generates visual outputs:
   - scatter plots between raters
   - Bland-Altman plots
11. User downloads or saves results as CSV, PDF, or DOCX.

## 3. Functional Requirements

### 3.1 File Upload and Parsing

- Support `.csv` uploads
- Support `.xlsx` uploads
- For Excel files, scan each worksheet and extract column headings for selection
- Present worksheet names clearly before data mapping begins
- Detect parsing errors and show user-friendly messages
- Preview uploaded data in a table
- Handle missing values with current pairwise complete-case filtering and clear reporting of dropped rows

### 3.2 Data Selection and Mapping

- Let user choose relevant columns for analysis
- Provide a clean interface for selecting reliability pairs of data
- Allow the user to select multiple detected pairs in one run
- Allow the user to select from detected worksheets and scanned column headings
- Support both common layouts:
  - wide format: one column per rater
  - long format: subject column, rater column, score column
- Allow the user to map:
  - subject/observation identifier
  - rater identifier
  - score/value field
- Support explicit pairing of measurement columns when pairwise reliability analysis is required
- Validate that the chosen structure is compatible with ICC calculation

Implementation status:
- Wide-format workflow: implemented
- Auto-detected `Test 1` / `Test 2` pair workflow: implemented
- Multiple-pair workflow: implemented
- Long-format workflow: planned, not yet implemented

### 3.3 ICC Recommendation Engine

- Ask structured questions to identify the correct ICC family
- Suggest the ICC form based on study design:
  - one-way random
  - two-way random
  - two-way mixed
- Distinguish between:
  - absolute agreement
  - consistency
- Distinguish between:
  - single measurement
  - average measurement
- Present the suggested ICC in plain language and technical notation
- Provide hover help describing when each choice is appropriate

### 3.4 Reliability Analysis

- Compute ICC point estimate
- Compute 95% confidence interval
- Compute descriptive statistics for the full analysed sample
- Compute descriptive statistics for each observation set or selected rating series
- Include median and IQR in descriptive summaries
- Include residual mean square error in the overall descriptive summary for the selected analysis
- Compute typical error metrics and define clearly how they are calculated
- Report sample size, number of raters, and missing-data handling
- Display assumptions and caveats
- Return informative errors when the dataset is not valid for the selected ICC

### 3.5 Visualisation

- Scatter plots for rater-pair comparisons
- Scatter plots must use a square aspect ratio
- Scatter plots must include a reference identity line (`y = x`)
- Bland-Altman plots for pairwise agreement
- Bland-Altman plots must be symmetric around 0 on the y-axis
- Bland-Altman plots must show the mean bias line
- Bland-Altman plots must show upper and lower limits of agreement
- Clear labels, legends, and axis titles
- By default, figure outputs should be generated as SVG and PDF
- User should be able to choose to view figures, save figures, or do both
- Figures should also be embedded in exported reports

### 3.6 Reporting

- Display a clear results summary
- Include:
  - ICC model used
  - ICC estimate
  - 95% confidence interval
  - descriptive statistics for the full group
  - descriptive statistics for each selected observation set
  - median and IQR
  - residual mean square error in the analysis-level summary
  - typical error metrics
  - interpretation guidance
- Provide figure output actions for view, save, or both
- Export analysed source data to CSV
- Export report to PDF
- Export report to DOCX

## 4. Non-Functional Requirements

- Simple browser-based workflow
- Clear and defensible statistical language
- Transparent user confirmation before running analysis
- Good error handling for malformed files and invalid selections
- Reproducible analysis logic
- Maintainable project structure for future expansion
- Deployable on a personal server or container host

## 5. Proposed Technical Stack

Given the current project setup, the initial implementation can use:

- Backend: Flask
- Data handling: pandas
- Excel support: openpyxl
- Statistics: pingouin and/or statsmodels
- Plotting: matplotlib and seaborn
- Frontend: Jinja templates, HTML, CSS, light JavaScript as needed
- Reporting: ReportLab and python-docx
- Serving: waitress
- Deployment: Docker / Compose / Procfile / WSGI

## 6. Proposed Application Modules

### 6.1 Upload Module

Responsibilities:
- upload file
- detect format
- scan worksheets for Excel uploads
- extract and store column headings by worksheet
- parse into a DataFrame
- store temporary analysis session data

Status: implemented

### 6.2 Data Mapping Module

Responsibilities:
- preview worksheets and columns
- provide a clean pair-selection interface
- let user assign subject, rater, and score fields
- let user select paired reliability columns where applicable
- reshape data when needed
- validate completeness

Status:
- worksheet preview and pair selection: implemented
- subject-column selection: implemented
- multi-pair selection: implemented
- full long-format mapping: not yet implemented

### 6.3 ICC Decision Module

Responsibilities:
- ask decision questions
- infer appropriate ICC type
- explain rationale
- request final user confirmation

Status: implemented, including hover guidance text for user choices

### 6.4 Analysis Module

Responsibilities:
- prepare clean dataset
- run ICC calculation
- compute 95% confidence intervals
- compute whole-group descriptive statistics
- compute per-observation-set descriptive statistics
- compute typical error metrics
- summarise key metadata

Status: implemented

### 6.5 Visualisation Module

Responsibilities:
- build scatter plots
- build Bland-Altman plots
- enforce square scatter plot dimensions and draw the `y = x` reference line
- enforce symmetric Bland-Altman y-axis scaling around 0
- annotate bias and limits of agreement on Bland-Altman plots
- render default figure outputs as SVG and PDF
- support user choice to view, save, or both
- return rendered images or embeddable outputs

Status: implemented

### 6.6 Results Module

Responsibilities:
- display numeric outputs
- present plot outputs
- provide export options later

Status: implemented with CSV, PDF, and DOCX export options

## 7. Data and Statistical Considerations

- Confirm whether every subject is rated by every rater
- Detect and report incomplete or unbalanced rating structures
- Define a clear policy for missing values
- Define exactly which descriptive statistics are shown at group and observation-set level
- Define the chosen typical error formula and ensure it is reported transparently
- Ensure the ICC recommendation logic matches standard statistical definitions
- Document exactly which ICC notation is used, for example:
  - ICC(1,1)
  - ICC(2,1)
  - ICC(2,k)
  - ICC(3,1)
  - ICC(3,k)
- Provide interpretation guidance carefully and avoid oversimplifying clinical or research conclusions

Current implementation notes:
- Missing values are handled by dropping incomplete rows for each selected pair
- ICC notation is aligned to the labels returned by `pingouin`, for example `ICC(1,1)`, `ICC(A,1)`, and `ICC(C,1)`
- Descriptives currently include mean, SD, median, IQR, min, max, and residual mean square error at the analysis level

## 8. Development Phases

### Phase 1: Foundation

- Finalise project structure
- Add dependencies for file parsing and plotting
- Build upload form
- Parse CSV and XLSX files
- Scan worksheet names and column headings
- Display data preview

Status: completed

### Phase 2: Data Mapping UI

- Build column selection workflow
- Build reliability-pair selection workflow
- Support wide and long data formats
- Add validation messages
- Confirm observations and raters

Status: partially completed
- wide-format and pair-selection workflows completed
- long-format workflow still open

### Phase 3: ICC Recommendation Logic

- Implement questionnaire for study design
- Map responses to ICC type suggestion
- Add explanatory help text

Status: completed

### Phase 4: Statistical Analysis

- Implement ICC calculation with 95% confidence intervals
- Implement descriptive summaries for the full group and each observation set
- Implement typical error calculations
- Validate against known examples
- Add summary output tables

Status: completed for the current wide-format pairwise workflow

### Phase 5: Visualisation

- Add scatter plot generation
- Add Bland-Altman plot generation
- Embed plots in results page
- Add SVG and PDF figure export flow
- Add user controls for view, save, or both

Status: completed

### Phase 6: Refinement

- Improve UX and error handling
- Add exports
- Add test coverage
- Prepare deployment configuration

Status: mostly completed
- UX improved with guide, chooser, navigator, and tooltips
- exports added for CSV, PDF, and DOCX
- deployment configuration added
- broader automated test coverage remains open

## 9. Key Risks and Open Questions

- Choosing a statistics library that provides trustworthy ICC estimates and confidence intervals
- Handling partially missing or unbalanced rater data
- Supporting both long and wide dataset layouts cleanly
- Determining how much automation should be used before requiring explicit user confirmation
- Deciding whether Bland-Altman plots should be limited to two raters or extended pairwise for multiple raters
- Ensuring SVG preview and PDF export both work reliably in the browser workflow
- Ensuring DOCX SVG rendering behaves consistently across Word versions and viewers

## 10. Recommended Immediate Next Steps

1. Add full long-format data mapping workflow.
2. Expand automated test coverage for upload, analysis, plotting, PDF export, and DOCX export.
3. Improve interpretation guidance and reporting notes.
4. Add stronger missing-data options if users need alternatives to complete-case filtering.
5. Verify DOCX SVG behavior in Microsoft Word across target environments.
6. Prepare production deployment target and live hosting workflow.

## 11. Definition of Done for Initial Release

The first usable release should allow a user to:

- upload CSV or XLSX data
- scan worksheets and column headings
- preview and map relevant columns
- select reliability pairs through a clean interface
- confirm subjects and raters
- receive an ICC recommendation
- calculate ICC with 95% confidence intervals
- view descriptive summaries for the full group and each observation set
- view median, IQR, and residual mean square error in the descriptive outputs
- view typical error metrics
- view square scatter plots with a `y = x` reference line
- view Bland-Altman plots centered symmetrically around 0 and showing bias and limits of agreement
- receive figure outputs in SVG and PDF by default
- choose whether to view figures, save figures, or both
- understand what model was used and why

Current release status: this definition is met for the implemented wide-format, pair-based workflow, with additional CSV, PDF, and DOCX export support.
