# Project Plan: Web-Based Inter-Rater Reliability Tool

## 1. Goal

Build a web-based application that allows users to upload `.xlsx` or `.csv` files, select the relevant columns for analysis, confirm the structure of raters and observations, receive guidance on the appropriate intraclass correlation coefficient (ICC) approach, and generate reliability outputs including:

- ICC estimate
- 95% confidence intervals
- Scatter plots
- Bland-Altman plots

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
  - descriptive summaries for each set of observations
  - typical error metrics
   - summary statistics
10. Application generates visual outputs:
   - scatter plots between raters
   - Bland-Altman plots
11. User downloads or saves results.

## 3. Functional Requirements

### 3.1 File Upload and Parsing

- Support `.csv` uploads
- Support `.xlsx` uploads
- For Excel files, scan each worksheet and extract column headings for selection
- Present worksheet names clearly before data mapping begins
- Detect parsing errors and show user-friendly messages
- Preview uploaded data in a table
- Handle missing values with clear user options or warnings

### 3.2 Data Selection and Mapping

- Let user choose relevant columns for analysis
- Provide a clean interface for selecting reliability pairs of data
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

### 3.4 Reliability Analysis

- Compute ICC point estimate
- Compute 95% confidence interval
- Compute descriptive statistics for the full analysed sample
- Compute descriptive statistics for each observation set or selected rating series
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

### 3.6 Reporting

- Display a clear results summary
- Include:
  - ICC model used
  - ICC estimate
  - 95% confidence interval
  - descriptive statistics for the full group
  - descriptive statistics for each selected observation set
  - typical error metrics
  - interpretation guidance
- Provide figure output actions for view, save, or both
- Optional export of summary results to CSV or PDF in a later phase

## 4. Non-Functional Requirements

- Simple browser-based workflow
- Clear and defensible statistical language
- Transparent user confirmation before running analysis
- Good error handling for malformed files and invalid selections
- Reproducible analysis logic
- Maintainable project structure for future expansion

## 5. Proposed Technical Stack

Given the current project setup, the initial implementation can use:

- Backend: Flask
- Data handling: pandas
- Excel support: openpyxl
- Statistics: pingouin and/or statsmodels
- Plotting: matplotlib and seaborn
- Frontend: Jinja templates, HTML, CSS, light JavaScript as needed

## 6. Proposed Application Modules

### 6.1 Upload Module

Responsibilities:
- upload file
- detect format
- scan worksheets for Excel uploads
- extract and store column headings by worksheet
- parse into a DataFrame
- store temporary analysis session data

### 6.2 Data Mapping Module

Responsibilities:
- preview worksheets and columns
- provide a clean pair-selection interface
- let user assign subject, rater, and score fields
- let user select paired reliability columns where applicable
- reshape data when needed
- validate completeness

### 6.3 ICC Decision Module

Responsibilities:
- ask decision questions
- infer appropriate ICC type
- explain rationale
- request final user confirmation

### 6.4 Analysis Module

Responsibilities:
- prepare clean dataset
- run ICC calculation
- compute 95% confidence intervals
- compute whole-group descriptive statistics
- compute per-observation-set descriptive statistics
- compute typical error metrics
- summarise key metadata

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

### 6.6 Results Module

Responsibilities:
- display numeric outputs
- present plot outputs
- provide export options later

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

## 8. Development Phases

### Phase 1: Foundation

- Finalise project structure
- Add dependencies for file parsing and plotting
- Build upload form
- Parse CSV and XLSX files
- Scan worksheet names and column headings
- Display data preview

### Phase 2: Data Mapping UI

- Build column selection workflow
- Build reliability-pair selection workflow
- Support wide and long data formats
- Add validation messages
- Confirm observations and raters

### Phase 3: ICC Recommendation Logic

- Implement questionnaire for study design
- Map responses to ICC type suggestion
- Add explanatory help text

### Phase 4: Statistical Analysis

- Implement ICC calculation with 95% confidence intervals
- Implement descriptive summaries for the full group and each observation set
- Implement typical error calculations
- Validate against known examples
- Add summary output tables

### Phase 5: Visualisation

- Add scatter plot generation
- Add Bland-Altman plot generation
- Embed plots in results page
- Add SVG and PDF figure export flow
- Add user controls for view, save, or both

### Phase 6: Refinement

- Improve UX and error handling
- Add exports
- Add test coverage
- Prepare deployment configuration

## 9. Key Risks and Open Questions

- Choosing a statistics library that provides trustworthy ICC estimates and confidence intervals
- Handling partially missing or unbalanced rater data
- Supporting both long and wide dataset layouts cleanly
- Determining how much automation should be used before requiring explicit user confirmation
- Deciding whether Bland-Altman plots should be limited to two raters or extended pairwise for multiple raters
- Ensuring SVG preview and PDF export both work reliably in the browser workflow

## 10. Recommended Immediate Next Steps

1. Add core analysis dependencies:
   - `pandas`
   - `openpyxl`
   - `pingouin`
   - `matplotlib`
   - `seaborn`
2. Create upload, worksheet scan, and dataset preview flow in the Flask app.
3. Define the internal canonical data structure for subject, rater, score, and paired comparison data.
4. Design the UI for worksheet, column, and reliability-pair selection.
5. Implement the ICC decision questionnaire.
6. Validate ICC and typical error outputs against reference examples before exposing final results.

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
- view typical error metrics
- view square scatter plots with a `y = x` reference line
- view Bland-Altman plots centered symmetrically around 0 and showing bias and limits of agreement
- receive figure outputs in SVG and PDF by default
- choose whether to view figures, save figures, or both
- understand what model was used and why
