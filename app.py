from __future__ import annotations

import io
import json
import os
import re
import zipfile
import xml.etree.ElementTree as ET
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
from docx import Document
from docx.shared import Inches
from flask import Flask, abort, redirect, render_template, request, send_file, url_for
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.units import inch
from reportlab.platypus import Image, PageBreak, Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle
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
DOCX_XML_NAMESPACES = {
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
    "ct": "http://schemas.openxmlformats.org/package/2006/content-types",
    "asvg": "http://schemas.microsoft.com/office/drawing/2016/SVG/main",
}
SVG_BLIP_EXTENSION_URI = "{96DAC541-7B7A-43D3-8B79-37D633B846F1}"
IMAGE_RELATIONSHIP_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"

for prefix, namespace in DOCX_XML_NAMESPACES.items():
    ET.register_namespace(prefix, namespace)


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
            "iqr": None,
            "min": None,
            "max": None,
        }

    q1 = cleaned.quantile(0.25)
    q3 = cleaned.quantile(0.75)

    return {
        "count": int(cleaned.count()),
        "mean": to_float(cleaned.mean()),
        "sd": to_float(cleaned.std(ddof=1)) if cleaned.count() > 1 else 0.0,
        "median": to_float(cleaned.median()),
        "iqr": to_float(q3 - q1),
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
            "median": wide_frame.median(axis=1),
            "iqr": wide_frame.quantile(0.75, axis=1) - wide_frame.quantile(0.25, axis=1),
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
                "median": to_float(row["median"]),
                "iqr": to_float(row["iqr"]),
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
        parts = [part.strip() for part in re.split(r"[,;\s]+", stripped) if part.strip()]
        if len(parts) == 2:
            return to_float(float(parts[0])), to_float(float(parts[1]))
        return None, None

    if isinstance(ci_value, (list, tuple, np.ndarray, pd.Series)) and len(ci_value) >= 2:
        return to_float(ci_value[0]), to_float(ci_value[1])

    return None, None


def approximate_t_critical_95(degrees_of_freedom: int) -> float:
    lookup = [
        (1, 12.706),
        (2, 4.303),
        (3, 3.182),
        (4, 2.776),
        (5, 2.571),
        (6, 2.447),
        (7, 2.365),
        (8, 2.306),
        (9, 2.262),
        (10, 2.228),
        (12, 2.179),
        (15, 2.131),
        (20, 2.086),
        (25, 2.060),
        (30, 2.042),
        (40, 2.021),
        (60, 2.000),
        (120, 1.980),
    ]

    for max_df, critical_value in lookup:
        if degrees_of_freedom <= max_df:
            return critical_value

    return 1.960


def build_regression_confidence_band(x_values: np.ndarray, y_values: np.ndarray) -> dict | None:
    if len(x_values) < 3 or np.allclose(x_values, x_values[0]):
        return None

    slope, intercept = np.polyfit(x_values, y_values, 1)
    x_grid = np.linspace(float(np.min(x_values)), float(np.max(x_values)), 200)
    fit_line = intercept + slope * x_grid
    fitted_values = intercept + slope * x_values
    residuals = y_values - fitted_values
    degrees_of_freedom = len(x_values) - 2

    if degrees_of_freedom <= 0:
        return None

    residual_standard_error = float(np.sqrt(np.sum(residuals**2) / degrees_of_freedom))
    x_mean = float(np.mean(x_values))
    sum_squared_x = float(np.sum((x_values - x_mean) ** 2))

    if np.isclose(sum_squared_x, 0.0):
        return None

    critical_value = approximate_t_critical_95(degrees_of_freedom)
    fit_standard_error = residual_standard_error * np.sqrt(
        (1 / len(x_values)) + ((x_grid - x_mean) ** 2 / sum_squared_x)
    )
    confidence_margin = critical_value * fit_standard_error

    return {
        "x_grid": x_grid,
        "fit_line": fit_line,
        "lower_band": fit_line - confidence_margin,
        "upper_band": fit_line + confidence_margin,
        "slope": float(slope),
        "intercept": float(intercept),
    }


def slugify(value: str) -> str:
    slug = re.sub(r"[^a-z0-9]+", "-", value.casefold()).strip("-")
    return slug or "pair"


def manual_pair_key(x_column: str, y_column: str) -> str:
    return f"manual-{slugify(x_column)}-vs-{slugify(y_column)}"


def build_selected_pair_definitions(
    sheet_meta: dict,
    selected_pair_keys: list[str],
    measurement_columns: list[str],
    primary_x_column: str,
    primary_y_column: str,
) -> list[dict]:
    selected_pairs: list[dict] = []

    for pair_key in selected_pair_keys:
        pair = find_pair_by_key(sheet_meta, pair_key)
        if pair is None:
            continue
        selected_pairs.append(
            {
                "pair_key": pair["key"],
                "pair_label": pair["label"],
                "measurement_columns": [pair["test_1"], pair["test_2"]],
                "primary_x_column": pair["test_1"],
                "primary_y_column": pair["test_2"],
            }
        )

    if selected_pairs:
        return selected_pairs

    if primary_x_column and primary_y_column and primary_x_column != primary_y_column:
        if primary_x_column not in measurement_columns or primary_y_column not in measurement_columns:
            raise AnalysisError("Select manual plot columns from the chosen measurement columns.")

        return [
            {
                "pair_key": manual_pair_key(primary_x_column, primary_y_column),
                "pair_label": f"{primary_x_column} vs {primary_y_column}",
                "measurement_columns": [primary_x_column, primary_y_column],
                "primary_x_column": primary_x_column,
                "primary_y_column": primary_y_column,
            }
        ]

    raise AnalysisError("Select at least one detected pair or choose a valid manual X/Y column combination.")


def analyse_pair_result(
    dataframe: pd.DataFrame,
    subject_column: str | None,
    pair_definition: dict,
    recommendation: dict,
) -> dict:
    measurement_columns = pair_definition["measurement_columns"]
    prepared = prepare_analysis_frame(dataframe, subject_column, measurement_columns)
    wide_frame: pd.DataFrame = prepared["wide_frame"]
    subject_labels: list[str] = prepared["subject_labels"]

    long_frame = wide_frame.copy()
    long_frame.insert(0, "subject", subject_labels)
    melted = long_frame.melt(id_vars="subject", var_name="rater", value_name="score")
    icc_table = pg.intraclass_corr(data=melted, targets="subject", raters="rater", ratings="score")
    icc_row = icc_table.loc[icc_table["Type"] == recommendation["icc_code"]]

    if icc_row.empty:
        raise AnalysisError(f"The selected ICC model could not be calculated for {pair_definition['pair_label']}.")

    selected_row = icc_row.iloc[0]
    ci_value = selected_row.get("CI95", selected_row.get("CI95%"))
    ci_lower, ci_upper = extract_ci_bounds(ci_value)
    overall_summary = summarise_series(wide_frame.stack())
    row_means = wide_frame.mean(axis=1)
    residual_mse = to_float(wide_frame.sub(row_means, axis=0).pow(2).to_numpy().mean())
    overall_summary["residual_mse"] = residual_mse
    column_summaries = build_column_summaries(wide_frame, measurement_columns)
    observation_summaries, observation_total = build_observation_summaries(wide_frame, subject_labels)
    pair_metrics = build_typical_error_table(wide_frame)

    return {
        "pair_key": pair_definition["pair_key"],
        "pair_label": pair_definition["pair_label"],
        "measurement_columns": measurement_columns,
        "primary_x_column": pair_definition["primary_x_column"],
        "primary_y_column": pair_definition["primary_y_column"],
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


def analyse_wide_dataset(
    upload_record: dict,
    sheet_name: str,
    subject_column: str | None,
    measurement_columns: list[str],
    primary_x_column: str,
    primary_y_column: str,
    selected_pair_keys: list[str],
    study_design: str,
    agreement_definition: str,
    measurement_unit: str,
    figure_action: str,
) -> dict:
    dataframe = read_dataset(get_upload_path(upload_record), upload_record["file_type"], sheet_name)
    sheet_meta = get_sheet_meta(upload_record, sheet_name)
    recommendation = build_icc_recommendation(study_design, agreement_definition, measurement_unit)
    pair_definitions = build_selected_pair_definitions(
        sheet_meta,
        selected_pair_keys,
        measurement_columns,
        primary_x_column,
        primary_y_column,
    )
    pair_results = [
        analyse_pair_result(dataframe, subject_column, pair_definition, recommendation)
        for pair_definition in pair_definitions
    ]

    return {
        "config": {
            "upload_id": upload_record["id"],
            "selected_sheet": sheet_name,
            "subject_column": subject_column,
            "measurement_columns": measurement_columns,
            "primary_x_column": primary_x_column,
            "primary_y_column": primary_y_column,
            "selected_pair_keys": [pair_result["pair_key"] for pair_result in pair_results],
            "selected_pair_labels": [pair_result["pair_label"] for pair_result in pair_results],
            "study_design": study_design,
            "agreement_definition": agreement_definition,
            "measurement_unit": measurement_unit,
            "figure_action": figure_action,
        },
        "recommendation": recommendation,
        "dataset_summary": {
            "pair_count": len(pair_results),
            "source_rows": int(len(dataframe)),
            "subject_label": subject_column or "Generated row labels",
        },
        "pair_results": pair_results,
    }


def save_analysis_record(analysis_result: dict) -> str:
    analysis_id = uuid4().hex
    payload = {"id": analysis_id, **analysis_result}
    save_json(ANALYSIS_DIR / f"{analysis_id}.json", payload)
    return analysis_id


def build_source_data_frame(upload_record: dict, analysis_record: dict, pair_result: dict) -> pd.DataFrame:
    config = analysis_record["config"]
    dataframe = read_dataset(
        get_upload_path(upload_record),
        upload_record["file_type"],
        config["selected_sheet"],
    )
    prepared = prepare_analysis_frame(
        dataframe,
        config.get("subject_column"),
        pair_result["measurement_columns"],
    )
    wide_frame: pd.DataFrame = prepared["wide_frame"].copy()
    source_frame = wide_frame.copy()
    subject_column = config.get("subject_column")

    if subject_column:
        source_frame.insert(0, subject_column, prepared["subject_labels"])
    else:
        source_frame.insert(0, "Observation", prepared["subject_labels"])

    return source_frame


def build_source_data_export_frame(upload_record: dict, analysis_record: dict) -> pd.DataFrame:
    export_frames: list[pd.DataFrame] = []

    for pair_result in analysis_record["pair_results"]:
        pair_frame = build_source_data_frame(upload_record, analysis_record, pair_result)
        subject_name = pair_frame.columns[0]
        export_frames.append(
            pd.DataFrame(
                {
                    "pair_key": pair_result["pair_key"],
                    "pair_label": pair_result["pair_label"],
                    "observation_id": pair_frame[subject_name],
                    "x_column": pair_result["primary_x_column"],
                    "x_value": pair_frame[pair_result["primary_x_column"]],
                    "y_column": pair_result["primary_y_column"],
                    "y_value": pair_frame[pair_result["primary_y_column"]],
                }
            )
        )

    if not export_frames:
        raise AnalysisError("No analysed pairs are available for export.")

    return pd.concat(export_frames, ignore_index=True)


def get_pair_result(analysis_record: dict, pair_key: str | None) -> dict:
    pair_results = analysis_record.get("pair_results", [])
    if not pair_results:
        raise AnalysisError("No pair results were found for this analysis.")

    if pair_key:
        for pair_result in pair_results:
            if pair_result["pair_key"] == pair_key:
                return pair_result

    return pair_results[0]


def markdown_table(dataframe: pd.DataFrame) -> str:
    formatted = dataframe.copy().replace({np.nan: ""})
    numeric_columns = formatted.select_dtypes(include=[np.number]).columns
    if len(numeric_columns) > 0:
        formatted[numeric_columns] = formatted[numeric_columns].round(4)
    return formatted.to_markdown(index=False)


def load_package_versions() -> list[str]:
    requirements_path = BASE_DIR / "requirements.txt"
    if not requirements_path.exists():
        return []

    packages: list[str] = []
    for line in requirements_path.read_text(encoding="utf-8").splitlines():
        stripped = line.strip()
        if stripped and not stripped.startswith("#"):
            packages.append(stripped)
    return packages


def dataframe_for_pdf(dataframe: pd.DataFrame, max_rows: int | None = None) -> list[list[str]]:
    frame = dataframe.copy().replace({np.nan: ""})
    if max_rows is not None:
        frame = frame.head(max_rows)

    table_rows: list[list[str]] = [list(frame.columns)]
    for _, row in frame.iterrows():
        formatted_row: list[str] = []
        for value in row.tolist():
            if isinstance(value, float):
                formatted_row.append(f"{value:.4f}")
            else:
                formatted_row.append(str(value))
        table_rows.append(formatted_row)
    return table_rows


def xml_tag(namespace_key: str, tag_name: str) -> str:
    return f"{{{DOCX_XML_NAMESPACES[namespace_key]}}}{tag_name}"


def add_docx_table(document: Document, dataframe: pd.DataFrame, max_rows: int | None = None) -> None:
    table_rows = dataframe_for_pdf(dataframe, max_rows=max_rows)
    if not table_rows or not table_rows[0]:
        document.add_paragraph("No rows available.")
        return

    table = document.add_table(rows=len(table_rows), cols=len(table_rows[0]))
    table.style = "Table Grid"

    for row_index, row_values in enumerate(table_rows):
        for column_index, value in enumerate(row_values):
            cell = table.cell(row_index, column_index)
            cell.text = value
            if row_index == 0:
                for run in cell.paragraphs[0].runs:
                    run.bold = True


def figure_to_bytes(figure: plt.Figure, file_format: str) -> bytes:
    buffer = io.BytesIO()
    save_kwargs = {"format": file_format, "bbox_inches": "tight"}
    if file_format == "png":
        save_kwargs["dpi"] = 220
    figure.savefig(buffer, **save_kwargs)
    buffer.seek(0)
    return buffer.getvalue()


def figure_to_docx_assets(figure: plt.Figure) -> tuple[bytes, bytes]:
    png_bytes = figure_to_bytes(figure, "png")
    svg_bytes = figure_to_bytes(figure, "svg")
    plt.close(figure)
    return png_bytes, svg_bytes


def embed_svgs_in_docx(docx_bytes: bytes, svg_images: list[bytes]) -> bytes:
    if not svg_images:
        return docx_bytes

    with zipfile.ZipFile(io.BytesIO(docx_bytes), "r") as source_zip:
        packaged_files = {name: source_zip.read(name) for name in source_zip.namelist()}

    document_xml_name = "word/document.xml"
    relationships_xml_name = "word/_rels/document.xml.rels"
    content_types_xml_name = "[Content_Types].xml"

    document_root = ET.fromstring(packaged_files[document_xml_name])
    relationships_root = ET.fromstring(packaged_files[relationships_xml_name])
    content_types_root = ET.fromstring(packaged_files[content_types_xml_name])

    blips = document_root.findall(".//a:blip", DOCX_XML_NAMESPACES)
    if len(blips) < len(svg_images):
        raise AnalysisError("The DOCX report could not be patched with SVG figures.")

    relationship_numbers = [
        int(match.group(1))
        for relationship in relationships_root.findall(xml_tag("rel", "Relationship"))
        if (match := re.fullmatch(r"rId(\d+)", relationship.get("Id", "")))
    ]
    next_relationship_number = max(relationship_numbers, default=0) + 1
    existing_media_names = {name.rsplit("/", 1)[-1] for name in packaged_files if name.startswith("word/media/")}

    for image_index, (blip, svg_bytes) in enumerate(zip(blips, svg_images), start=1):
        media_name = f"svg-image-{image_index}.svg"
        while media_name in existing_media_names:
            image_index += 1
            media_name = f"svg-image-{image_index}.svg"
        existing_media_names.add(media_name)

        relationship_id = f"rId{next_relationship_number}"
        next_relationship_number += 1

        ET.SubElement(
            relationships_root,
            xml_tag("rel", "Relationship"),
            {
                "Id": relationship_id,
                "Type": IMAGE_RELATIONSHIP_TYPE,
                "Target": f"media/{media_name}",
            },
        )

        extension_list = blip.find("a:extLst", DOCX_XML_NAMESPACES)
        if extension_list is None:
            extension_list = ET.SubElement(blip, xml_tag("a", "extLst"))

        svg_extension = ET.SubElement(
            extension_list,
            xml_tag("a", "ext"),
            {"uri": SVG_BLIP_EXTENSION_URI},
        )
        ET.SubElement(
            svg_extension,
            xml_tag("asvg", "svgBlip"),
            {xml_tag("r", "embed"): relationship_id},
        )

        packaged_files[f"word/media/{media_name}"] = svg_bytes

    has_svg_content_type = any(
        default.get("Extension") == "svg"
        for default in content_types_root.findall(xml_tag("ct", "Default"))
    )
    if not has_svg_content_type:
        ET.SubElement(
            content_types_root,
            xml_tag("ct", "Default"),
            {"Extension": "svg", "ContentType": "image/svg+xml"},
        )

    packaged_files[document_xml_name] = ET.tostring(document_root, encoding="utf-8", xml_declaration=True)
    packaged_files[relationships_xml_name] = ET.tostring(relationships_root, encoding="utf-8", xml_declaration=True)
    packaged_files[content_types_xml_name] = ET.tostring(content_types_root, encoding="utf-8", xml_declaration=True)

    output_buffer = io.BytesIO()
    with zipfile.ZipFile(output_buffer, "w", compression=zipfile.ZIP_DEFLATED) as target_zip:
        for file_name, file_bytes in packaged_files.items():
            target_zip.writestr(file_name, file_bytes)

    output_buffer.seek(0)
    return output_buffer.getvalue()


def pdf_table(dataframe: pd.DataFrame, max_rows: int | None = None) -> Table:
    table = Table(dataframe_for_pdf(dataframe, max_rows=max_rows), repeatRows=1)
    table.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#dbeafe")),
                ("TEXTCOLOR", (0, 0), (-1, 0), colors.HexColor("#0f172a")),
                ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                ("GRID", (0, 0), (-1, -1), 0.25, colors.HexColor("#94a3b8")),
                ("BACKGROUND", (0, 1), (-1, -1), colors.white),
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("FONTSIZE", (0, 0), (-1, -1), 8),
                ("LEADING", (0, 0), (-1, -1), 10),
            ]
        )
    )
    return table


def figure_to_pdf_image(figure: plt.Figure, width_inches: float) -> Image:
    buffer = io.BytesIO()
    figure.savefig(buffer, format="png", dpi=180, bbox_inches="tight")
    plt.close(figure)
    buffer.seek(0)
    image = Image(buffer)
    image.drawWidth = width_inches * inch
    aspect_ratio = image.imageHeight / image.imageWidth if image.imageWidth else 1
    image.drawHeight = image.drawWidth * aspect_ratio
    return image


def build_pdf_report(analysis_record: dict) -> bytes:
    config = analysis_record["config"]
    upload_record = load_upload(config["upload_id"])
    if upload_record is None:
        raise AnalysisError("The source upload for this analysis is no longer available.")

    styles = getSampleStyleSheet()
    story: list = []
    package_versions = load_package_versions()
    source_frame = build_source_data_export_frame(upload_record, analysis_record)
    selected_pairs = ", ".join(config.get("selected_pair_labels", [])) or "Manual column selection"

    story.extend(
        [
            Paragraph("Reliability Analysis Report", styles["Title"]),
            Spacer(1, 0.2 * inch),
            Paragraph(f"Generated: {datetime.now().isoformat(timespec='seconds')}", styles["Normal"]),
            Spacer(1, 0.2 * inch),
            Paragraph("1. Analysed source data", styles["Heading1"]),
            Paragraph(f"Source file: {upload_record['original_filename']}", styles["Normal"]),
            Paragraph(f"Worksheet: {config['selected_sheet']}", styles["Normal"]),
            Paragraph(f"Selected pairs: {selected_pairs}", styles["Normal"]),
            Paragraph(f"Observation identifier: {config.get('subject_column') or 'Generated row labels'}", styles["Normal"]),
            Spacer(1, 0.12 * inch),
            pdf_table(source_frame),
            Spacer(1, 0.24 * inch),
            Paragraph("2. Analysis description", styles["Heading1"]),
            Paragraph(f"Study design: {analysis_record['recommendation']['design_label']}", styles["Normal"]),
            Paragraph(f"Agreement target: {analysis_record['recommendation']['agreement_label']}", styles["Normal"]),
            Paragraph(f"Measurement unit: {analysis_record['recommendation']['measurement_label']}", styles["Normal"]),
            Paragraph(f"Rationale: {analysis_record['recommendation']['rationale']}", styles["Normal"]),
            Paragraph(f"Analysed pairs: {analysis_record['dataset_summary']['pair_count']}", styles["Normal"]),
            Paragraph(f"Source rows: {analysis_record['dataset_summary']['source_rows']}", styles["Normal"]),
            Spacer(1, 0.18 * inch),
            Paragraph("Python packages used", styles["Heading2"]),
        ]
    )

    for package in package_versions:
        story.append(Paragraph(f"• {package}", styles["Normal"]))

    story.extend(
        [
            Spacer(1, 0.16 * inch),
            Paragraph("Commands and analysis steps used", styles["Heading2"]),
            Paragraph("Run commands:", styles["Normal"]),
            Paragraph("python app.py", styles["Code"]),
            Paragraph("python run_web.py", styles["Code"]),
            Spacer(1, 0.08 * inch),
            Paragraph("Analysis workflow:", styles["Normal"]),
            Paragraph("1. Load the worksheet and selected columns.", styles["Normal"]),
            Paragraph("2. Drop rows with missing values for each selected pair.", styles["Normal"]),
            Paragraph("3. Reshape the pair data and run pingouin.intraclass_corr(...).", styles["Normal"]),
            Paragraph("4. Compute typical error as SD(y − x) / √2.", styles["Normal"]),
            Paragraph("5. Compute bias and limits of agreement as bias ± 1.96 × SD(y − x).", styles["Normal"]),
            Paragraph("6. Generate square scatter plots with a y = x line and Bland-Altman plots centered symmetrically around 0.", styles["Normal"]),
        ]
    )

    for index, pair_result in enumerate(analysis_record["pair_results"], start=1):
        pair_source_frame = build_source_data_frame(upload_record, analysis_record, pair_result)
        overall_summary_frame = pd.DataFrame([pair_result["overall_summary"]])
        column_summary_frame = pd.DataFrame(pair_result["column_summaries"])
        observation_summary_frame = pd.DataFrame(pair_result["observation_summaries"])
        pair_metrics_frame = pd.DataFrame(pair_result["pair_metrics"])

        pair_frame = pair_source_frame[[pair_result["primary_x_column"], pair_result["primary_y_column"]]].copy()
        scatter_image = figure_to_pdf_image(
            build_scatter_plot(pair_frame, pair_result["primary_x_column"], pair_result["primary_y_column"]),
            width_inches=5.8,
        )
        bland_image = figure_to_pdf_image(
            build_bland_altman_plot(pair_frame, pair_result["primary_x_column"], pair_result["primary_y_column"]),
            width_inches=5.8,
        )

        story.extend(
            [
                PageBreak() if index > 1 else Spacer(1, 0.2 * inch),
                Paragraph(f"3.{index} Results for {pair_result['pair_label']}", styles["Heading1"]),
                Paragraph(f"Columns: {pair_result['primary_x_column']} vs {pair_result['primary_y_column']}", styles["Normal"]),
                Paragraph(f"ICC model: {pair_result['icc_result']['model']}", styles["Normal"]),
                Paragraph(f"ICC estimate: {pair_result['icc_result']['estimate']}", styles["Normal"]),
                Paragraph(f"95% CI: {pair_result['icc_result']['ci_lower']} to {pair_result['icc_result']['ci_upper']}", styles["Normal"]),
                Paragraph(f"F value: {pair_result['icc_result']['f_value']}", styles["Normal"]),
                Paragraph(f"P value: {pair_result['icc_result']['p_value']}", styles["Normal"]),
                Paragraph(f"Description: {pair_result['icc_result']['description']}", styles["Normal"]),
                Paragraph(f"Complete observations analysed: {pair_result['dataset_summary']['observations']}", styles["Normal"]),
                Paragraph(f"Dropped rows due to missing values: {pair_result['dataset_summary']['dropped_rows']}", styles["Normal"]),
                Paragraph(f"Typical error formula: {pair_result['typical_error_formula']}", styles["Normal"]),
                Spacer(1, 0.12 * inch),
                Paragraph("Overall descriptive summary", styles["Heading2"]),
                pdf_table(overall_summary_frame),
                Spacer(1, 0.12 * inch),
                Paragraph("Series summaries", styles["Heading2"]),
                pdf_table(column_summary_frame),
                Spacer(1, 0.12 * inch),
                Paragraph("Observation summaries", styles["Heading2"]),
                pdf_table(observation_summary_frame),
                Spacer(1, 0.12 * inch),
                Paragraph("Typical error and limits of agreement", styles["Heading2"]),
                pdf_table(pair_metrics_frame),
                Spacer(1, 0.12 * inch),
                Paragraph("Source data used for this pair", styles["Heading2"]),
                pdf_table(pair_source_frame),
                Spacer(1, 0.14 * inch),
                Paragraph("Figures", styles["Heading2"]),
                scatter_image,
                Spacer(1, 0.12 * inch),
                bland_image,
            ]
        )

    buffer = io.BytesIO()
    document = SimpleDocTemplate(buffer, pagesize=A4, leftMargin=30, rightMargin=30, topMargin=30, bottomMargin=30)
    document.build(story)
    buffer.seek(0)
    return buffer.getvalue()


def build_docx_report(analysis_record: dict) -> bytes:
    config = analysis_record["config"]
    upload_record = load_upload(config["upload_id"])
    if upload_record is None:
        raise AnalysisError("The source upload for this analysis is no longer available.")

    document = Document()
    document.core_properties.title = "Reliability Analysis Report"
    document.add_heading("Reliability Analysis Report", level=0)
    document.add_paragraph(f"Generated: {datetime.now().isoformat(timespec='seconds')}")

    package_versions = load_package_versions()
    source_frame = build_source_data_export_frame(upload_record, analysis_record)
    selected_pairs = ", ".join(config.get("selected_pair_labels", [])) or "Manual column selection"
    svg_images: list[bytes] = []

    document.add_heading("1. Analysed source data", level=1)
    document.add_paragraph(f"Source file: {upload_record['original_filename']}")
    document.add_paragraph(f"Worksheet: {config['selected_sheet']}")
    document.add_paragraph(f"Selected pairs: {selected_pairs}")
    document.add_paragraph(f"Observation identifier: {config.get('subject_column') or 'Generated row labels'}")
    add_docx_table(document, source_frame)

    document.add_heading("2. Analysis description", level=1)
    document.add_paragraph(f"Study design: {analysis_record['recommendation']['design_label']}")
    document.add_paragraph(f"Agreement target: {analysis_record['recommendation']['agreement_label']}")
    document.add_paragraph(f"Measurement unit: {analysis_record['recommendation']['measurement_label']}")
    document.add_paragraph(f"Rationale: {analysis_record['recommendation']['rationale']}")
    document.add_paragraph(f"Analysed pairs: {analysis_record['dataset_summary']['pair_count']}")
    document.add_paragraph(f"Source rows: {analysis_record['dataset_summary']['source_rows']}")

    document.add_heading("Python packages used", level=2)
    for package in package_versions:
        document.add_paragraph(package, style="List Bullet")

    document.add_heading("Commands and analysis steps used", level=2)
    document.add_paragraph("Run commands:")
    document.add_paragraph("python app.py")
    document.add_paragraph("python run_web.py")
    document.add_paragraph("Analysis workflow:")
    for step in [
        "Load the worksheet and selected columns.",
        "Drop rows with missing values for each selected pair.",
        "Reshape the pair data and run pingouin.intraclass_corr(...).",
        "Compute typical error as SD(y − x) / √2.",
        "Compute bias and limits of agreement as bias ± 1.96 × SD(y − x).",
        "Generate square scatter plots with a y = x line and Bland-Altman plots centered symmetrically around 0.",
    ]:
        document.add_paragraph(step, style="List Number")

    for index, pair_result in enumerate(analysis_record["pair_results"], start=1):
        if index > 1:
            document.add_page_break()

        pair_source_frame = build_source_data_frame(upload_record, analysis_record, pair_result)
        overall_summary_frame = pd.DataFrame([pair_result["overall_summary"]])
        column_summary_frame = pd.DataFrame(pair_result["column_summaries"])
        observation_summary_frame = pd.DataFrame(pair_result["observation_summaries"])
        pair_metrics_frame = pd.DataFrame(pair_result["pair_metrics"])
        pair_frame = pair_source_frame[[pair_result["primary_x_column"], pair_result["primary_y_column"]]].copy()

        document.add_heading(f"3.{index} Results for {pair_result['pair_label']}", level=1)
        document.add_paragraph(f"Columns: {pair_result['primary_x_column']} vs {pair_result['primary_y_column']}")
        document.add_paragraph(f"ICC model: {pair_result['icc_result']['model']}")
        document.add_paragraph(f"ICC estimate: {pair_result['icc_result']['estimate']}")
        document.add_paragraph(f"95% CI: {pair_result['icc_result']['ci_lower']} to {pair_result['icc_result']['ci_upper']}")
        document.add_paragraph(f"F value: {pair_result['icc_result']['f_value']}")
        document.add_paragraph(f"P value: {pair_result['icc_result']['p_value']}")
        document.add_paragraph(f"Description: {pair_result['icc_result']['description']}")
        document.add_paragraph(f"Complete observations analysed: {pair_result['dataset_summary']['observations']}")
        document.add_paragraph(f"Dropped rows due to missing values: {pair_result['dataset_summary']['dropped_rows']}")
        document.add_paragraph(f"Typical error formula: {pair_result['typical_error_formula']}")

        document.add_heading("Overall descriptive summary", level=2)
        add_docx_table(document, overall_summary_frame)
        document.add_heading("Series summaries", level=2)
        add_docx_table(document, column_summary_frame)
        document.add_heading("Observation summaries", level=2)
        add_docx_table(document, observation_summary_frame)
        document.add_heading("Typical error and limits of agreement", level=2)
        add_docx_table(document, pair_metrics_frame)
        document.add_heading("Source data used for this pair", level=2)
        add_docx_table(document, pair_source_frame)

        document.add_heading("Figures", level=2)
        scatter_png, scatter_svg = figure_to_docx_assets(
            build_scatter_plot(pair_frame, pair_result["primary_x_column"], pair_result["primary_y_column"])
        )
        document.add_paragraph(f"Scatter plot: {pair_result['pair_label']}")
        document.add_picture(io.BytesIO(scatter_png), width=Inches(6.0))
        svg_images.append(scatter_svg)

        bland_png, bland_svg = figure_to_docx_assets(
            build_bland_altman_plot(pair_frame, pair_result["primary_x_column"], pair_result["primary_y_column"])
        )
        document.add_paragraph(f"Bland-Altman plot: {pair_result['pair_label']}")
        document.add_picture(io.BytesIO(bland_png), width=Inches(6.0))
        svg_images.append(bland_svg)

    buffer = io.BytesIO()
    document.save(buffer)
    return embed_svgs_in_docx(buffer.getvalue(), svg_images)


def build_markdown_report(analysis_record: dict, base_url: str | None = None) -> str:
    config = analysis_record["config"]
    upload_record = load_upload(config["upload_id"])
    if upload_record is None:
        raise AnalysisError("The source upload for this analysis is no longer available.")

    source_frame = build_source_data_export_frame(upload_record, analysis_record)
    selected_pairs = ", ".join(config.get("selected_pair_labels", [])) or "Manual column selection"
    base_url = (base_url or "").rstrip("/")
    package_versions = load_package_versions()

    lines = [
        "# Reliability Analysis Report",
        "",
        f"Generated: {datetime.now().isoformat(timespec='seconds')}",
        "",
        "## Analysed source data",
        "",
        f"Source file: {upload_record['original_filename']}",
        "",
        markdown_table(source_frame),
        "",
        "## Analysis description",
        "",
        f"- Worksheet: {config['selected_sheet']}",
        f"- Selected pairs: {selected_pairs}",
        f"- Observation identifier: {config.get('subject_column') or 'Generated row labels'}",
        f"- Study design: {analysis_record['recommendation']['design_label']}",
        f"- Agreement target: {analysis_record['recommendation']['agreement_label']}",
        f"- Measurement unit: {analysis_record['recommendation']['measurement_label']}",
        f"- Rationale: {analysis_record['recommendation']['rationale']}",
        f"- Analysed pairs: {analysis_record['dataset_summary']['pair_count']}",
        f"- Source rows: {analysis_record['dataset_summary']['source_rows']}",
        f"- Subject identifier: {analysis_record['dataset_summary']['subject_label']}",
        "",
        "### Python packages used",
        "",
    ]

    lines.extend([f"- {package}" for package in package_versions])
    lines.extend(
        [
            "",
            "### Commands and analysis steps used",
            "",
            "The web app was run with one of the following commands:",
            "",
            "```powershell",
            "python app.py",
            "# or",
            "python run_web.py",
            "```",
            "",
            "For each selected pair, the app executed the following analysis workflow:",
            "",
            "```python",
            "dataframe = read_dataset(upload_path, file_type, selected_sheet)",
            "pair_frame = dataframe[[subject_column, x_column, y_column]].copy()",
            "pair_frame = pair_frame.dropna(subset=[x_column, y_column])",
            "melted = pair_frame.melt(id_vars='subject', var_name='rater', value_name='score')",
            "icc_table = pingouin.intraclass_corr(data=melted, targets='subject', raters='rater', ratings='score')",
            "typical_error = SD(y - x) / sqrt(2)",
            "bias = mean(y - x)",
            "limits_of_agreement = bias ± 1.96 * SD(y - x)",
            "scatter_plot = square plot with y = x reference line",
            "bland_altman_plot = mean(x, y) vs (y - x), centred symmetrically around 0",
            "```",
            "",
            "### Notes on figures",
            "",
            "- The Markdown report is a single text file.",
            "- Figures are included below as links and Markdown image references to the live web app routes.",
            "- If the app is no longer running when the Markdown file is opened, the figure links may not render.",
        ]
    )

    for pair_result in analysis_record["pair_results"]:
        overall_summary_frame = pd.DataFrame([pair_result["overall_summary"]])
        column_summary_frame = pd.DataFrame(pair_result["column_summaries"])
        observation_summary_frame = pd.DataFrame(pair_result["observation_summaries"])
        pair_metrics_frame = pd.DataFrame(pair_result["pair_metrics"])
        pair_source_frame = build_source_data_frame(upload_record, analysis_record, pair_result)
        scatter_svg = (
            f"{base_url}/plots/{analysis_record['id']}/{pair_result['pair_key']}/scatter.svg"
            if base_url
            else None
        )
        scatter_pdf = (
            f"{base_url}/plots/{analysis_record['id']}/{pair_result['pair_key']}/scatter.pdf?download=1"
            if base_url
            else None
        )
        bland_svg = (
            f"{base_url}/plots/{analysis_record['id']}/{pair_result['pair_key']}/bland-altman.svg"
            if base_url
            else None
        )
        bland_pdf = (
            f"{base_url}/plots/{analysis_record['id']}/{pair_result['pair_key']}/bland-altman.pdf?download=1"
            if base_url
            else None
        )

        lines.extend(
            [
                "",
                f"## Results for pair: {pair_result['pair_label']}",
                "",
                f"- Columns: {pair_result['primary_x_column']} vs {pair_result['primary_y_column']}",
                f"- ICC model: {pair_result['icc_result']['model']}",
                f"- ICC estimate: {pair_result['icc_result']['estimate']}",
                f"- 95% CI: {pair_result['icc_result']['ci_lower']} to {pair_result['icc_result']['ci_upper']}",
                f"- F value: {pair_result['icc_result']['f_value']}",
                f"- P value: {pair_result['icc_result']['p_value']}",
                f"- Description: {pair_result['icc_result']['description']}",
                f"- Complete observations analysed: {pair_result['dataset_summary']['observations']}",
                f"- Dropped rows due to missing values: {pair_result['dataset_summary']['dropped_rows']}",
                f"- Typical error formula: {pair_result['typical_error_formula']}",
                "",
                "### Overall descriptive summary",
                "",
                markdown_table(overall_summary_frame),
                "",
                "### Series summaries",
                "",
                markdown_table(column_summary_frame),
                "",
                "### Observation summaries",
                "",
                markdown_table(observation_summary_frame),
                "",
                "### Typical error and limits of agreement",
                "",
                markdown_table(pair_metrics_frame),
                "",
                "### Source data used for this pair",
                "",
                markdown_table(pair_source_frame),
            ]
        )

        if scatter_svg and bland_svg:
            lines.extend(
                [
                    "",
                    "### Figures",
                    "",
                    f"- [Scatter SVG]({scatter_svg})",
                    f"- [Scatter PDF]({scatter_pdf})",
                    f"- [Bland-Altman SVG]({bland_svg})",
                    f"- [Bland-Altman PDF]({bland_pdf})",
                    "",
                    f"![Scatter plot for {pair_result['pair_label']}]({scatter_svg})",
                    "",
                    f"![Bland-Altman plot for {pair_result['pair_label']}]({bland_svg})",
                ]
            )

    lines.extend(
        [
            "## Notes",
            "",
            "- This report includes the analysed source rows after pair-specific missing-value filtering.",
            "- This Markdown file combines the analysed data, analysis description, results, and figure links in one report.",
            "- Full implementation source code is not embedded, but the commands and package set used by the app are listed above.",
        ]
    )
    return "\n".join(lines)


def build_scatter_plot(dataframe: pd.DataFrame, x_column: str, y_column: str) -> plt.Figure:
    x_values = dataframe[x_column].to_numpy(dtype=float)
    y_values = dataframe[y_column].to_numpy(dtype=float)
    regression_band = build_regression_confidence_band(x_values, y_values)
    lower = float(min(x_values.min(), y_values.min()))
    upper = float(max(x_values.max(), y_values.max()))
    padding = max((upper - lower) * 0.05, 0.1)
    axis_min = lower - padding
    axis_max = upper + padding

    figure, axis = plt.subplots(figsize=(6, 6))
    axis.scatter(x_values, y_values, color="#2563eb", edgecolors="white", linewidths=0.8, s=55)
    if regression_band:
        axis.fill_between(
            regression_band["x_grid"],
            regression_band["lower_band"],
            regression_band["upper_band"],
            color="#0f766e",
            alpha=0.16,
            label="95% CI of best-fit line",
            zorder=1,
        )
        axis.plot(
            regression_band["x_grid"],
            regression_band["fit_line"],
            color="#0f766e",
            linewidth=1.8,
            label=(
                f"Best fit: y = {regression_band['slope']:.3f}x "
                f"{regression_band['intercept']:+.3f}"
            ),
            zorder=2,
        )
    axis.plot(
        [axis_min, axis_max],
        [axis_min, axis_max],
        linestyle="--",
        color="#dc2626",
        linewidth=1.5,
        label="Identity line (y = x)" if regression_band else None,
    )
    axis.set_xlim(axis_min, axis_max)
    axis.set_ylim(axis_min, axis_max)
    axis.set_aspect("equal", adjustable="box")
    axis.set_xlabel(x_column)
    axis.set_ylabel(y_column)
    axis.set_title(f"Scatter plot: {x_column} vs {y_column}")
    axis.grid(alpha=0.25)
    if regression_band:
        axis.legend(loc="upper left")
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


def build_plot_response(analysis_record: dict, pair_key: str, plot_kind: str, file_format: str, download: bool):
    if file_format not in {"svg", "pdf"}:
        abort(404)

    config = analysis_record["config"]
    upload_record = load_upload(config["upload_id"])
    if upload_record is None:
        abort(404)
    try:
        pair_result = get_pair_result(analysis_record, pair_key)
    except AnalysisError:
        abort(404)

    pair_frame = build_source_data_frame(upload_record, analysis_record, pair_result)
    x_column = pair_result["primary_x_column"]
    y_column = pair_result["primary_y_column"]
    wide_frame = pair_frame[[x_column, y_column]].copy()

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

    download_name = f"{plot_kind}-{pair_result['pair_key']}.{file_format}".replace(" ", "-")
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
        "selected_pair_keys": [default_pair["key"]] if default_pair else [],
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
                "selected_pair_keys": config.get("selected_pair_keys", form_state["selected_pair_keys"]),
                "study_design": config.get("study_design", "two_way_random"),
                "agreement_definition": config.get("agreement_definition", "absolute"),
                "measurement_unit": config.get("measurement_unit", "single"),
                "figure_action": config.get("figure_action", "both"),
                "sheet_name": config.get("selected_sheet", form_state["sheet_name"]),
            }
        )

    return form_state


def form_state_from_request(sheet_meta: dict) -> dict:
    selected_pair_keys = request.form.getlist("selected_pair_keys")
    measurement_columns = request.form.getlist("measurement_columns")
    selected_pairs = [pair for pair_key in selected_pair_keys if (pair := find_pair_by_key(sheet_meta, pair_key)) is not None]

    if selected_pairs:
        measurement_columns = list(
            dict.fromkeys(
                column
                for pair in selected_pairs
                for column in (pair["test_1"], pair["test_2"])
            )
        )
        primary_x_column = selected_pairs[0]["test_1"]
        primary_y_column = selected_pairs[0]["test_2"]
    else:
        primary_x_column = request.form.get("primary_x_column", "")
        primary_y_column = request.form.get("primary_y_column", "")

    return {
        "subject_column": request.form.get("subject_column", ""),
        "measurement_columns": measurement_columns,
        "primary_x_column": primary_x_column,
        "primary_y_column": primary_y_column,
        "selected_pair_keys": selected_pair_keys,
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
                        selected_pair_keys=form_state["selected_pair_keys"],
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
                            prompt="1",
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
    show_export_prompt = request.args.get("prompt") == "1"

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
        show_export_prompt=show_export_prompt,
    )


@app.get("/healthz")
def healthcheck() -> tuple[dict[str, str], int]:
    ensure_storage()
    return {
        "status": "ok",
        "service": "reliability-web-tool",
    }, 200


@app.get("/plots/<analysis_id>/<pair_key>/<plot_kind>.<file_format>")
def plot_file(analysis_id: str, pair_key: str, plot_kind: str, file_format: str):
    ensure_storage()
    analysis_record = load_analysis(analysis_id)
    if analysis_record is None:
        abort(404)
    download = request.args.get("download") == "1"
    return build_plot_response(analysis_record, pair_key, plot_kind, file_format, download)


@app.get("/reports/<analysis_id>.md")
def markdown_report(analysis_id: str):
    ensure_storage()
    analysis_record = load_analysis(analysis_id)
    if analysis_record is None:
        abort(404)

    try:
        report = build_markdown_report(analysis_record, request.url_root)
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


@app.get("/reports/<analysis_id>.pdf")
def pdf_report(analysis_id: str):
    ensure_storage()
    analysis_record = load_analysis(analysis_id)
    if analysis_record is None:
        abort(404)

    try:
        pdf_bytes = build_pdf_report(analysis_record)
    except AnalysisError:
        abort(404)

    buffer = io.BytesIO(pdf_bytes)
    download_name = f"reliability-results-{analysis_id}.pdf"
    return send_file(
        buffer,
        mimetype="application/pdf",
        as_attachment=True,
        download_name=download_name,
    )


@app.get("/reports/<analysis_id>.docx")
def docx_report(analysis_id: str):
    ensure_storage()
    analysis_record = load_analysis(analysis_id)
    if analysis_record is None:
        abort(404)

    try:
        docx_bytes = build_docx_report(analysis_record)
    except AnalysisError:
        abort(404)

    buffer = io.BytesIO(docx_bytes)
    download_name = f"reliability-results-{analysis_id}.docx"
    return send_file(
        buffer,
        mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
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
        source_frame = build_source_data_export_frame(upload_record, analysis_record)
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
    app.run(
        host=os.environ.get("HOST", "127.0.0.1"),
        port=int(os.environ.get("PORT", "5000")),
        debug=os.environ.get("FLASK_DEBUG", "1") == "1",
    )
