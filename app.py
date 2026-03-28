from __future__ import annotations

import io
import json
import os
import re
from datetime import datetime
from itertools import combinations
from pathlib import Path
from uuid import uuid4

import matplotlib

matplotlib.use("Agg")

import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
import pingouin as pg
from flask import Flask, abort, redirect, render_template, request, send_file, url_for
from werkzeug.utils import secure_filename

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 16 * 1024 * 1024
app.config["SECRET_KEY"] = os.environ.get("FLASK_SECRET_KEY", "dev-secret-key")

BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "instance" / "data"
UPLOAD_DIR = DATA_DIR / "uploads"
ANALYSIS_DIR = DATA_DIR / "analyses"
ALLOWED_EXTENSIONS = {"csv", "xlsx"}
PAIR_PATTERN = re.compile(r"^(?P<label>.+?)\s+Test\s+(?P<test>[12])(?:\.\d+)?\s*$", re.IGNORECASE)


class AnalysisError(ValueError):
    pass


def ensure_storage() -> None:
    for directory in (UPLOAD_DIR, ANALYSIS_DIR):
        directory.mkdir(parents=True, exist_ok=True)


def allowed_file(filename: str) -> bool:
    return "." in filename and filename.rsplit(".", 1)[1].lower() in ALLOWED_EXTENSIONS


def save_json(file_path: Path, payload: dict) -> None:
    file_path.write_text(json.dumps(payload, indent=2), encoding="utf-8")


def load_json(file_path: Path) -> dict:
    return json.loads(file_path.read_text(encoding="utf-8"))


def make_unique_columns(columns: list[str]) -> list[str]:
    seen: dict[str, int] = {}
    unique_columns: list[str] = []

    for index, column in enumerate(columns, start=1):
        base_name = str(column).strip() or f"Column {index}"
        count = seen.get(base_name, 0)
        seen[base_name] = count + 1
        unique_columns.append(base_name if count == 0 else f"{base_name} ({count + 1})")

    return unique_columns


def normalise_dataframe(dataframe: pd.DataFrame) -> pd.DataFrame:
    cleaned = dataframe.copy()
    cleaned = cleaned.dropna(axis=0, how="all").dropna(axis=1, how="all")
    cleaned.columns = make_unique_columns([str(column) for column in cleaned.columns])
    return cleaned


def read_csv_file(file_path: Path) -> pd.DataFrame:
    encodings = ("utf-8-sig", "utf-8", "cp1252", "latin-1")
    last_error: UnicodeDecodeError | None = None

    for encoding in encodings:
        try:
            return pd.read_csv(file_path, encoding=encoding)
        except UnicodeDecodeError as exc:
            last_error = exc

    if last_error is not None:
        raise AnalysisError("The CSV file could not be decoded using common encodings.") from last_error

    raise AnalysisError("The CSV file could not be read.")


def read_dataset(file_path: Path, file_type: str, sheet_name: str | None = None) -> pd.DataFrame:
    if file_type == "csv":
        return normalise_dataframe(read_csv_file(file_path))

    if not sheet_name:
        raise AnalysisError("Select a worksheet before continuing.")

    dataframe = pd.read_excel(file_path, sheet_name=sheet_name)
    return normalise_dataframe(dataframe)


def scan_sheet(name: str, dataframe: pd.DataFrame) -> dict:
    numeric_columns = [
        column
        for column in dataframe.columns
        if pd.api.types.is_numeric_dtype(dataframe[column])
    ]
    detected_pairs = detect_reliability_pairs(numeric_columns)

    return {
        "name": name,
        "row_count": int(len(dataframe)),
        "column_count": int(len(dataframe.columns)),
        "columns": list(dataframe.columns),
        "numeric_columns": numeric_columns,
        "detected_pairs": detected_pairs,
    }


def detect_reliability_pairs(columns: list[str]) -> list[dict]:
    pair_map: dict[str, dict] = {}

    for column in columns:
        match = PAIR_PATTERN.match(str(column).strip())
        if not match:
            continue

        label = re.sub(r"\s+", " ", match.group("label")).strip()
        key = label.casefold()
        pair_entry = pair_map.setdefault(
            key,
            {
                "key": key,
                "label": label,
                "test_1": None,
                "test_2": None,
            },
        )
        pair_entry[f"test_{match.group('test')}"] = column

    detected_pairs = [
        pair for pair in pair_map.values() if pair["test_1"] is not None and pair["test_2"] is not None
    ]
    detected_pairs.sort(key=lambda pair: pair["label"])
    return detected_pairs


def scan_upload(file_path: Path, file_type: str) -> list[dict]:
    if file_type == "csv":
        dataframe = read_dataset(file_path, file_type)
        return [scan_sheet("CSV data", dataframe)]

    workbook = pd.ExcelFile(file_path)
    scanned_sheets: list[dict] = []

    for sheet_name in workbook.sheet_names:
        dataframe = normalise_dataframe(pd.read_excel(file_path, sheet_name=sheet_name, nrows=200))
        scanned_sheets.append(scan_sheet(sheet_name, dataframe))

    if not scanned_sheets:
        raise AnalysisError("No worksheets were found in the uploaded workbook.")

    return scanned_sheets


def save_uploaded_file(file_storage) -> dict:
    filename = secure_filename(file_storage.filename or "")
    if not filename or not allowed_file(filename):
        raise AnalysisError("Upload a CSV or XLSX file.")

    file_id = uuid4().hex
    extension = filename.rsplit(".", 1)[1].lower()
    stored_filename = f"{file_id}.{extension}"
    stored_path = UPLOAD_DIR / stored_filename
    file_storage.save(stored_path)

    sheets = scan_upload(stored_path, extension)
    upload_record = {
        "id": file_id,
        "original_filename": filename,
        "stored_filename": stored_filename,
        "file_type": extension,
        "sheets": sheets,
    }
    save_json(UPLOAD_DIR / f"{file_id}.json", upload_record)
    return upload_record


def load_upload(upload_id: str | None) -> dict | None:
    if not upload_id:
        return None

    record_path = UPLOAD_DIR / f"{upload_id}.json"
    if not record_path.exists():
        return None

    return load_json(record_path)


def load_analysis(analysis_id: str | None) -> dict | None:
    if not analysis_id:
        return None

    record_path = ANALYSIS_DIR / f"{analysis_id}.json"
    if not record_path.exists():
        return None

    return load_json(record_path)


def get_upload_path(upload_record: dict) -> Path:
    return UPLOAD_DIR / upload_record["stored_filename"]


def get_sheet_options(upload_record: dict) -> list[str]:
    return [sheet["name"] for sheet in upload_record["sheets"]]


def get_active_sheet(upload_record: dict, requested_sheet: str | None) -> str:
    sheet_names = get_sheet_options(upload_record)
    if requested_sheet in sheet_names:
        return requested_sheet
    return sheet_names[0]


def get_sheet_meta(upload_record: dict, sheet_name: str) -> dict:
    for sheet in upload_record["sheets"]:
        if sheet["name"] == sheet_name:
            return sheet
    raise AnalysisError("The selected worksheet could not be found.")


def build_preview_rows(dataframe: pd.DataFrame, limit: int = 12) -> list[dict]:
    preview = dataframe.head(limit).replace({np.nan: None})
    return preview.to_dict(orient="records")


def find_pair_by_key(sheet_meta: dict, pair_key: str | None) -> dict | None:
    if not pair_key:
        return None

    for pair in sheet_meta.get("detected_pairs", []):
        if pair["key"] == pair_key:
            return pair

    return None


def to_float(value: object, digits: int = 4) -> float | None:
    if value is None or pd.isna(value):
        return None
    return round(float(value), digits)


def summarise_series(series: pd.Series) -> dict:
    cleaned = pd.to_numeric(series, errors="coerce").dropna()

    if cleaned.empty:
        return {
            "count": 0,
            "mean": None,
            "sd": None,
            "median": None,
            "min": None,
            "max": None,
        }

    return {
        "count": int(cleaned.count()),
        "mean": to_float(cleaned.mean()),
        "sd": to_float(cleaned.std(ddof=1)) if cleaned.count() > 1 else 0.0,
        "median": to_float(cleaned.median()),
        "min": to_float(cleaned.min()),
        "max": to_float(cleaned.max()),
    }


def build_column_summaries(dataframe: pd.DataFrame, measurement_columns: list[str]) -> list[dict]:
    summaries: list[dict] = []

    for column in measurement_columns:
        summary = summarise_series(dataframe[column])
        summary["name"] = column
        summaries.append(summary)

    return summaries


def build_observation_summaries(
    wide_frame: pd.DataFrame,
    subject_labels: list[str],
    limit: int = 25,
) -> tuple[list[dict], int]:
    summary_frame = pd.DataFrame(
        {
            "observation": subject_labels,
            "mean": wide_frame.mean(axis=1),
            "sd": wide_frame.std(axis=1, ddof=1).fillna(0.0),
            "min": wide_frame.min(axis=1),
            "max": wide_frame.max(axis=1),
        }
    )
    total_count = int(len(summary_frame))
    summary_rows = summary_frame.head(limit).to_dict(orient="records")

    cleaned_rows: list[dict] = []
    for row in summary_rows:
        cleaned_rows.append(
            {
                "observation": str(row["observation"]),
                "mean": to_float(row["mean"]),
                "sd": to_float(row["sd"]),
                "min": to_float(row["min"]),
                "max": to_float(row["max"]),
            }
        )

    return cleaned_rows, total_count


def choose_icc_code(
    study_design: str,
    agreement_definition: str,
    measurement_unit: str,
) -> str:
    lookup = {
        ("one_way_random", "absolute", "single"): "ICC(1,1)",
        ("one_way_random", "absolute", "average"): "ICC(1,k)",
        ("one_way_random", "consistency", "single"): "ICC(1,1)",
        ("one_way_random", "consistency", "average"): "ICC(1,k)",
        ("two_way_random", "absolute", "single"): "ICC(A,1)",
        ("two_way_random", "absolute", "average"): "ICC(A,k)",
        ("two_way_random", "consistency", "single"): "ICC(C,1)",
        ("two_way_random", "consistency", "average"): "ICC(C,k)",
        ("two_way_mixed", "absolute", "single"): "ICC(A,1)",
        ("two_way_mixed", "absolute", "average"): "ICC(A,k)",
        ("two_way_mixed", "consistency", "single"): "ICC(C,1)",
        ("two_way_mixed", "consistency", "average"): "ICC(C,k)",
    }

    icc_code = lookup.get((study_design, agreement_definition, measurement_unit))
    if icc_code is None:
        raise AnalysisError("The ICC configuration is not supported.")
    return icc_code


def build_icc_recommendation(
    study_design: str,
    agreement_definition: str,
    measurement_unit: str,
) -> dict:
    icc_code = choose_icc_code(study_design, agreement_definition, measurement_unit)

    design_labels = {
        "one_way_random": "One-way random effects",
        "two_way_random": "Two-way random effects",
        "two_way_mixed": "Two-way mixed effects",
    }
    agreement_labels = {
        "absolute": "absolute agreement",
        "consistency": "consistency",
    }
    measure_labels = {
        "single": "single measurement",
        "average": "average measurement",
    }

    rationale = (
        f"Suggested model: {icc_code} based on a {design_labels[study_design].lower()} design, "
        f"with {agreement_labels[agreement_definition]} and a {measure_labels[measurement_unit]} target."
    )

    if study_design == "two_way_mixed" and agreement_definition == "absolute":
        rationale += " This starter build maps fixed-rater agreement requests to the ICC(3) family for reporting transparency."

    return {
        "icc_code": icc_code,
        "design_label": design_labels[study_design],
        "agreement_label": agreement_labels[agreement_definition].title(),
        "measurement_label": measure_labels[measurement_unit].title(),
        "rationale": rationale,
    }


def prepare_analysis_frame(
    dataframe: pd.DataFrame,
    subject_column: str | None,
    measurement_columns: list[str],
) -> dict:
    required_columns = [*measurement_columns]
    if subject_column:
        required_columns.insert(0, subject_column)

    analysis_frame = dataframe[required_columns].copy()
    for column in measurement_columns:
        analysis_frame[column] = pd.to_numeric(analysis_frame[column], errors="coerce")

    source_row_count = int(len(analysis_frame))
    complete_frame = analysis_frame.dropna(subset=measurement_columns).copy()
    dropped_rows = source_row_count - int(len(complete_frame))

    if len(measurement_columns) < 2:
        raise AnalysisError("Select at least two measurement columns.")

    if complete_frame.shape[0] < 2:
        raise AnalysisError("At least two complete observations are required after removing missing values.")

    if subject_column:
        subject_labels = complete_frame[subject_column].astype(str).tolist()
    else:
        subject_labels = [f"Observation {index + 1}" for index in range(len(complete_frame))]

    wide_frame = complete_frame[measurement_columns].astype(float)
    return {
        "wide_frame": wide_frame,
        "subject_labels": subject_labels,
        "source_row_count": source_row_count,
        "complete_row_count": int(len(complete_frame)),
        "dropped_rows": dropped_rows,
    }


def build_typical_error_table(wide_frame: pd.DataFrame) -> list[dict]:
    metrics: list[dict] = []

    for first_column, second_column in combinations(wide_frame.columns, 2):
        first_series = wide_frame[first_column]
        second_series = wide_frame[second_column]
        difference = second_series - first_series
        bias = float(difference.mean())
        sd_difference = float(difference.std(ddof=1)) if len(difference) > 1 else 0.0
        typical_error = sd_difference / np.sqrt(2)
        loa_upper = bias + 1.96 * sd_difference
        loa_lower = bias - 1.96 * sd_difference

        metrics.append(
            {
                "pair": f"{first_column} vs {second_column}",
                "first_column": first_column,
                "second_column": second_column,
                "bias": to_float(bias),
                "typical_error": to_float(typical_error),
                "loa_lower": to_float(loa_lower),
                "loa_upper": to_float(loa_upper),
            }
        )

    return metrics


def extract_ci_bounds(ci_value: object) -> tuple[float | None, float | None]:
    if isinstance(ci_value, str):
        stripped = ci_value.strip("[]()")
        parts = [part.strip() for part in stripped.split(",") if part.strip()]
        if len(parts) == 2:
            return to_float(float(parts[0])), to_float(float(parts[1]))
        return None, None

    if isinstance(ci_value, (list, tuple, np.ndarray, pd.Series)) and len(ci_value) >= 2:
        return to_float(ci_value[0]), to_float(ci_value[1])

    return None, None


def analyse_wide_dataset(
    upload_record: dict,
    sheet_name: str,
    subject_column: str | None,
    measurement_columns: list[str],
    primary_x_column: str,
    primary_y_column: str,
    selected_pair_key: str,
    study_design: str,
    agreement_definition: str,
    measurement_unit: str,
    figure_action: str,
) -> dict:
    dataframe = read_dataset(get_upload_path(upload_record), upload_record["file_type"], sheet_name)
    prepared = prepare_analysis_frame(dataframe, subject_column, measurement_columns)
    wide_frame: pd.DataFrame = prepared["wide_frame"]
    subject_labels: list[str] = prepared["subject_labels"]

    if primary_x_column not in measurement_columns or primary_y_column not in measurement_columns:
        raise AnalysisError("Select plot columns from the chosen measurement columns.")

    if primary_x_column == primary_y_column:
        raise AnalysisError("Choose two different columns for the primary comparison plots.")

    sheet_meta = get_sheet_meta(upload_record, sheet_name)
    selected_pair = find_pair_by_key(sheet_meta, selected_pair_key)

    long_frame = wide_frame.copy()
    long_frame.insert(0, "subject", subject_labels)
    melted = long_frame.melt(id_vars="subject", var_name="rater", value_name="score")

    recommendation = build_icc_recommendation(study_design, agreement_definition, measurement_unit)
    icc_table = pg.intraclass_corr(data=melted, targets="subject", raters="rater", ratings="score")
    icc_row = icc_table.loc[icc_table["Type"] == recommendation["icc_code"]]

    if icc_row.empty:
        raise AnalysisError("The selected ICC model could not be calculated for this dataset.")

    selected_row = icc_row.iloc[0]
    ci_lower, ci_upper = extract_ci_bounds(selected_row.get("CI95%"))
    overall_summary = summarise_series(wide_frame.stack())
    column_summaries = build_column_summaries(wide_frame, measurement_columns)
    observation_summaries, observation_total = build_observation_summaries(wide_frame, subject_labels)
    pair_metrics = build_typical_error_table(wide_frame)

    return {
        "config": {
            "upload_id": upload_record["id"],
            "selected_sheet": sheet_name,
            "subject_column": subject_column,
            "measurement_columns": measurement_columns,
            "primary_x_column": primary_x_column,
            "primary_y_column": primary_y_column,
            "selected_pair_key": selected_pair_key,
            "selected_pair_label": selected_pair["label"] if selected_pair else None,
            "study_design": study_design,
            "agreement_definition": agreement_definition,
            "measurement_unit": measurement_unit,
            "figure_action": figure_action,
        },
        "recommendation": recommendation,
        "dataset_summary": {
            "observations": prepared["complete_row_count"],
            "raters": len(measurement_columns),
            "source_rows": prepared["source_row_count"],
            "dropped_rows": prepared["dropped_rows"],
            "subject_label": subject_column or "Generated row labels",
        },
        "icc_result": {
            "model": recommendation["icc_code"],
            "estimate": to_float(selected_row["ICC"]),
            "ci_lower": ci_lower,
            "ci_upper": ci_upper,
            "f_value": to_float(selected_row.get("F")),
            "p_value": to_float(selected_row.get("pval")),
            "description": str(selected_row.get("Description", "")),
        },
        "overall_summary": overall_summary,
        "column_summaries": column_summaries,
        "observation_summaries": observation_summaries,
        "observation_total": observation_total,
        "pair_metrics": pair_metrics,
        "typical_error_formula": "Typical error = SD(differences) / √2",
    }


def save_analysis_record(analysis_result: dict) -> str:
    analysis_id = uuid4().hex
    payload = {"id": analysis_id, **analysis_result}
    save_json(ANALYSIS_DIR / f"{analysis_id}.json", payload)
    return analysis_id


def build_source_data_frame(upload_record: dict, analysis_record: dict) -> pd.DataFrame:
    config = analysis_record["config"]
    dataframe = read_dataset(
        get_upload_path(upload_record),
        upload_record["file_type"],
        config["selected_sheet"],
    )
    prepared = prepare_analysis_frame(
        dataframe,
        config.get("subject_column"),
        config["measurement_columns"],
    )
    wide_frame: pd.DataFrame = prepared["wide_frame"].copy()
    source_frame = wide_frame.copy()
    subject_column = config.get("subject_column")

    if subject_column:
        source_frame.insert(0, subject_column, prepared["subject_labels"])
    else:
        source_frame.insert(0, "Observation", prepared["subject_labels"])

    return source_frame


def markdown_table(dataframe: pd.DataFrame) -> str:
    formatted = dataframe.copy().replace({np.nan: ""})
    numeric_columns = formatted.select_dtypes(include=[np.number]).columns
    if len(numeric_columns) > 0:
        formatted[numeric_columns] = formatted[numeric_columns].round(4)
    return formatted.to_markdown(index=False)


def build_markdown_report(analysis_record: dict) -> str:
    config = analysis_record["config"]
    upload_record = load_upload(config["upload_id"])
    if upload_record is None:
        raise AnalysisError("The source upload for this analysis is no longer available.")

    source_frame = build_source_data_frame(upload_record, analysis_record)
    overall_summary_frame = pd.DataFrame([analysis_record["overall_summary"]])
    column_summary_frame = pd.DataFrame(analysis_record["column_summaries"])
    observation_summary_frame = pd.DataFrame(analysis_record["observation_summaries"])
    pair_metrics_frame = pd.DataFrame(analysis_record["pair_metrics"])

    selected_pair = config.get("selected_pair_label") or "Manual column selection"
    measurement_columns = ", ".join(config["measurement_columns"])

    lines = [
        "# Reliability Analysis Report",
        "",
        f"Generated: {datetime.now().isoformat(timespec='seconds')}",
        "",
        "## Source",
        "",
        f"- File: {upload_record['original_filename']}",
        f"- Worksheet: {config['selected_sheet']}",
        f"- Selected pair: {selected_pair}",
        f"- Measurement columns: {measurement_columns}",
        f"- Observation identifier: {config.get('subject_column') or 'Generated row labels'}",
        "",
        "## Analysis settings",
        "",
        f"- ICC model: {analysis_record['icc_result']['model']}",
        f"- Study design: {analysis_record['recommendation']['design_label']}",
        f"- Agreement target: {analysis_record['recommendation']['agreement_label']}",
        f"- Measurement unit: {analysis_record['recommendation']['measurement_label']}",
        f"- Rationale: {analysis_record['recommendation']['rationale']}",
        f"- Typical error formula: {analysis_record['typical_error_formula']}",
        "",
        "## Reliability result",
        "",
        f"- ICC estimate: {analysis_record['icc_result']['estimate']}",
        f"- 95% CI: {analysis_record['icc_result']['ci_lower']} to {analysis_record['icc_result']['ci_upper']}",
        f"- F value: {analysis_record['icc_result']['f_value']}",
        f"- P value: {analysis_record['icc_result']['p_value']}",
        f"- Description: {analysis_record['icc_result']['description']}",
        "",
        "## Dataset summary",
        "",
        f"- Complete observations analysed: {analysis_record['dataset_summary']['observations']}",
        f"- Raters / series: {analysis_record['dataset_summary']['raters']}",
        f"- Source rows: {analysis_record['dataset_summary']['source_rows']}",
        f"- Dropped rows due to missing values: {analysis_record['dataset_summary']['dropped_rows']}",
        "",
        "## Overall descriptive summary",
        "",
        markdown_table(overall_summary_frame),
        "",
        "## Series summaries",
        "",
        markdown_table(column_summary_frame),
        "",
        "## Observation summaries",
        "",
        markdown_table(observation_summary_frame),
        "",
        "## Typical error and limits of agreement",
        "",
        markdown_table(pair_metrics_frame),
        "",
        "## Source data used in the analysis",
        "",
        markdown_table(source_frame),
        "",
        "## Notes",
        "",
        "- This report includes the analysed source rows after missing-value filtering.",
        "- Figures are available separately in SVG and PDF from the web app.",
        "- Full implementation code is intentionally not embedded in the report by default.",
    ]
    return "\n".join(lines)


def build_scatter_plot(dataframe: pd.DataFrame, x_column: str, y_column: str) -> plt.Figure:
    x_values = dataframe[x_column].to_numpy(dtype=float)
    y_values = dataframe[y_column].to_numpy(dtype=float)
    lower = float(min(x_values.min(), y_values.min()))
    upper = float(max(x_values.max(), y_values.max()))
    padding = max((upper - lower) * 0.05, 0.1)
    axis_min = lower - padding
    axis_max = upper + padding

    figure, axis = plt.subplots(figsize=(6, 6))
    axis.scatter(x_values, y_values, color="#2563eb", edgecolors="white", linewidths=0.8, s=55)
    axis.plot([axis_min, axis_max], [axis_min, axis_max], linestyle="--", color="#dc2626", linewidth=1.5)
    axis.set_xlim(axis_min, axis_max)
    axis.set_ylim(axis_min, axis_max)
    axis.set_aspect("equal", adjustable="box")
    axis.set_xlabel(x_column)
    axis.set_ylabel(y_column)
    axis.set_title(f"Scatter plot: {x_column} vs {y_column}")
    axis.grid(alpha=0.25)
    figure.tight_layout()
    return figure


def build_bland_altman_plot(dataframe: pd.DataFrame, x_column: str, y_column: str) -> plt.Figure:
    x_values = dataframe[x_column].to_numpy(dtype=float)
    y_values = dataframe[y_column].to_numpy(dtype=float)
    means = (x_values + y_values) / 2
    differences = y_values - x_values
    bias = float(np.mean(differences))
    sd_difference = float(np.std(differences, ddof=1)) if len(differences) > 1 else 0.0
    loa_upper = bias + 1.96 * sd_difference
    loa_lower = bias - 1.96 * sd_difference
    max_extent = max(abs(bias), abs(loa_upper), abs(loa_lower), float(np.max(np.abs(differences))), 0.1)
    y_limit = max_extent * 1.1

    figure, axis = plt.subplots(figsize=(7, 5.5))
    axis.scatter(means, differences, color="#0f766e", edgecolors="white", linewidths=0.8, s=55)
    axis.axhline(0, color="#1e293b", linewidth=1.1)
    axis.axhline(bias, color="#2563eb", linestyle="-", linewidth=1.6, label=f"Bias = {bias:.3f}")
    axis.axhline(loa_upper, color="#dc2626", linestyle="--", linewidth=1.4, label=f"Upper LoA = {loa_upper:.3f}")
    axis.axhline(loa_lower, color="#dc2626", linestyle="--", linewidth=1.4, label=f"Lower LoA = {loa_lower:.3f}")
    axis.set_ylim(-y_limit, y_limit)
    axis.set_xlabel(f"Mean of {x_column} and {y_column}")
    axis.set_ylabel(f"Difference ({y_column} - {x_column})")
    axis.set_title(f"Bland-Altman plot: {x_column} vs {y_column}")
    axis.legend(loc="upper right")
    axis.grid(alpha=0.25)
    figure.tight_layout()
    return figure


def build_plot_response(analysis_record: dict, plot_kind: str, file_format: str, download: bool):
    if file_format not in {"svg", "pdf"}:
        abort(404)

    config = analysis_record["config"]
    upload_record = load_upload(config["upload_id"])
    if upload_record is None:
        abort(404)

    dataframe = read_dataset(
        get_upload_path(upload_record),
        upload_record["file_type"],
        config["selected_sheet"],
    )
    prepared = prepare_analysis_frame(
        dataframe,
        config.get("subject_column"),
        config["measurement_columns"],
    )
    wide_frame: pd.DataFrame = prepared["wide_frame"]
    x_column = config["primary_x_column"]
    y_column = config["primary_y_column"]

    if plot_kind == "scatter":
        figure = build_scatter_plot(wide_frame, x_column, y_column)
    elif plot_kind == "bland-altman":
        figure = build_bland_altman_plot(wide_frame, x_column, y_column)
    else:
        abort(404)

    buffer = io.BytesIO()
    figure.savefig(buffer, format=file_format, bbox_inches="tight")
    plt.close(figure)
    buffer.seek(0)

    download_name = f"{plot_kind}-{x_column}-vs-{y_column}.{file_format}".replace(" ", "-")
    mimetype = "image/svg+xml" if file_format == "svg" else "application/pdf"
    return send_file(buffer, mimetype=mimetype, as_attachment=download, download_name=download_name)


def default_form_state(sheet_meta: dict | None, analysis_record: dict | None = None) -> dict:
    all_columns = sheet_meta["columns"] if sheet_meta else []
    numeric_columns = sheet_meta["numeric_columns"] if sheet_meta else []
    detected_pairs = sheet_meta.get("detected_pairs", []) if sheet_meta else []
    default_pair = detected_pairs[0] if detected_pairs else None
    default_measurements = (
        [default_pair["test_1"], default_pair["test_2"]]
        if default_pair
        else numeric_columns[:2]
    )

    form_state = {
        "subject_column": "",
        "measurement_columns": default_measurements,
        "primary_x_column": default_measurements[0] if len(default_measurements) >= 1 else "",
        "primary_y_column": default_measurements[1] if len(default_measurements) >= 2 else "",
        "selected_pair_key": default_pair["key"] if default_pair else "",
        "study_design": "two_way_random",
        "agreement_definition": "absolute",
        "measurement_unit": "single",
        "figure_action": "both",
        "sheet_name": sheet_meta["name"] if sheet_meta else "",
        "available_columns": all_columns,
        "numeric_columns": numeric_columns,
        "detected_pairs": detected_pairs,
    }

    if analysis_record:
        config = analysis_record["config"]
        form_state.update(
            {
                "subject_column": config.get("subject_column") or "",
                "measurement_columns": config.get("measurement_columns", default_measurements),
                "primary_x_column": config.get("primary_x_column", ""),
                "primary_y_column": config.get("primary_y_column", ""),
                "selected_pair_key": config.get("selected_pair_key", form_state["selected_pair_key"]),
                "study_design": config.get("study_design", "two_way_random"),
                "agreement_definition": config.get("agreement_definition", "absolute"),
                "measurement_unit": config.get("measurement_unit", "single"),
                "figure_action": config.get("figure_action", "both"),
                "sheet_name": config.get("selected_sheet", form_state["sheet_name"]),
            }
        )

    return form_state


def form_state_from_request(sheet_meta: dict) -> dict:
    selected_pair_key = request.form.get("selected_pair_key", "")
    measurement_columns = request.form.getlist("measurement_columns")
    selected_pair = find_pair_by_key(sheet_meta, selected_pair_key)

    if selected_pair is not None:
        measurement_columns = [selected_pair["test_1"], selected_pair["test_2"]]
        primary_x_column = selected_pair["test_1"]
        primary_y_column = selected_pair["test_2"]
    else:
        primary_x_column = request.form.get("primary_x_column", "")
        primary_y_column = request.form.get("primary_y_column", "")

    return {
        "subject_column": request.form.get("subject_column", ""),
        "measurement_columns": measurement_columns,
        "primary_x_column": primary_x_column,
        "primary_y_column": primary_y_column,
        "selected_pair_key": selected_pair_key,
        "study_design": request.form.get("study_design", "two_way_random"),
        "agreement_definition": request.form.get("agreement_definition", "absolute"),
        "measurement_unit": request.form.get("measurement_unit", "single"),
        "figure_action": request.form.get("figure_action", "both"),
        "sheet_name": request.form.get("sheet_name", sheet_meta["name"]),
        "available_columns": sheet_meta["columns"],
        "numeric_columns": sheet_meta["numeric_columns"],
        "detected_pairs": sheet_meta.get("detected_pairs", []),
    }


@app.route("/", methods=["GET", "POST"])
def index() -> str:
    ensure_storage()
    error_message: str | None = None

    if request.method == "POST":
        action = request.form.get("action")

        if action == "upload":
            try:
                uploaded_file = request.files.get("dataset")
                if uploaded_file is None or not uploaded_file.filename:
                    raise AnalysisError("Choose a CSV or XLSX file to upload.")

                upload_record = save_uploaded_file(uploaded_file)
                return redirect(url_for("index", upload_id=upload_record["id"]))
            except AnalysisError as exc:
                error_message = str(exc)

        if action == "analyze":
            upload_id = request.form.get("upload_id", "")
            upload_record = load_upload(upload_id)

            if upload_record is None:
                error_message = "Upload a dataset before running the analysis."
            else:
                requested_sheet = request.form.get("sheet_name")
                active_sheet = get_active_sheet(upload_record, requested_sheet)
                sheet_meta = get_sheet_meta(upload_record, active_sheet)
                form_state = form_state_from_request(sheet_meta)

                try:
                    analysis_result = analyse_wide_dataset(
                        upload_record=upload_record,
                        sheet_name=active_sheet,
                        subject_column=form_state["subject_column"] or None,
                        measurement_columns=form_state["measurement_columns"],
                        primary_x_column=form_state["primary_x_column"],
                        primary_y_column=form_state["primary_y_column"],
                        selected_pair_key=form_state["selected_pair_key"],
                        study_design=form_state["study_design"],
                        agreement_definition=form_state["agreement_definition"],
                        measurement_unit=form_state["measurement_unit"],
                        figure_action=form_state["figure_action"],
                    )
                    analysis_id = save_analysis_record(analysis_result)
                    return redirect(
                        url_for(
                            "index",
                            upload_id=upload_id,
                            sheet=active_sheet,
                            analysis_id=analysis_id,
                        )
                    )
                except AnalysisError as exc:
                    error_message = str(exc)
                    preview_frame = read_dataset(get_upload_path(upload_record), upload_record["file_type"], active_sheet)
                    return render_template(
                        "index.html",
                        error=error_message,
                        upload=upload_record,
                        selected_sheet=active_sheet,
                        active_sheet_meta=sheet_meta,
                        preview_columns=list(preview_frame.columns),
                        preview_rows=build_preview_rows(preview_frame),
                        form_state=form_state,
                        analysis=None,
                    )

    upload_id = request.args.get("upload_id")
    analysis_id = request.args.get("analysis_id")
    analysis_record = load_analysis(analysis_id)

    if analysis_record and not upload_id:
        upload_id = analysis_record["config"]["upload_id"]

    upload_record = load_upload(upload_id)
    selected_sheet: str | None = request.args.get("sheet")
    active_sheet_meta = None
    preview_columns: list[str] = []
    preview_rows: list[dict] = []
    form_state = default_form_state(None, analysis_record)

    if upload_record:
        if analysis_record and not selected_sheet:
            selected_sheet = analysis_record["config"]["selected_sheet"]
        selected_sheet = get_active_sheet(upload_record, selected_sheet)
        active_sheet_meta = get_sheet_meta(upload_record, selected_sheet)
        preview_frame = read_dataset(get_upload_path(upload_record), upload_record["file_type"], selected_sheet)
        preview_columns = list(preview_frame.columns)
        preview_rows = build_preview_rows(preview_frame)
        form_state = default_form_state(active_sheet_meta, analysis_record)

    return render_template(
        "index.html",
        error=error_message,
        upload=upload_record,
        selected_sheet=selected_sheet,
        active_sheet_meta=active_sheet_meta,
        preview_columns=preview_columns,
        preview_rows=preview_rows,
        form_state=form_state,
        analysis=analysis_record,
    )


@app.get("/plots/<analysis_id>/<plot_kind>.<file_format>")
def plot_file(analysis_id: str, plot_kind: str, file_format: str):
    ensure_storage()
    analysis_record = load_analysis(analysis_id)
    if analysis_record is None:
        abort(404)
    download = request.args.get("download") == "1"
    return build_plot_response(analysis_record, plot_kind, file_format, download)


@app.get("/reports/<analysis_id>.md")
def markdown_report(analysis_id: str):
    ensure_storage()
    analysis_record = load_analysis(analysis_id)
    if analysis_record is None:
        abort(404)

    try:
        report = build_markdown_report(analysis_record)
    except AnalysisError:
        abort(404)

    buffer = io.BytesIO(report.encode("utf-8"))
    download_name = f"reliability-report-{analysis_id}.md"
    return send_file(
        buffer,
        mimetype="text/markdown; charset=utf-8",
        as_attachment=True,
        download_name=download_name,
    )


@app.get("/reports/<analysis_id>-source-data.csv")
def source_data_csv(analysis_id: str):
    ensure_storage()
    analysis_record = load_analysis(analysis_id)
    if analysis_record is None:
        abort(404)

    upload_record = load_upload(analysis_record["config"]["upload_id"])
    if upload_record is None:
        abort(404)

    try:
        source_frame = build_source_data_frame(upload_record, analysis_record)
    except AnalysisError:
        abort(404)

    buffer = io.BytesIO(source_frame.to_csv(index=False).encode("utf-8"))
    download_name = f"reliability-source-data-{analysis_id}.csv"
    return send_file(
        buffer,
        mimetype="text/csv; charset=utf-8",
        as_attachment=True,
        download_name=download_name,
    )


if __name__ == "__main__":
    ensure_storage()
    app.run(debug=True)
