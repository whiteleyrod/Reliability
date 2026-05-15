from __future__ import annotations

import hashlib
from pathlib import Path
from uuid import uuid4

import pandas as pd
import streamlit as st

from app import (
    ALLOWED_EXTENSIONS,
    AnalysisError,
    DEFAULT_FIGURE_PALETTE,
    UPLOAD_DIR,
    analyse_wide_dataset,
    build_bland_altman_plot,
    build_palette_preview_frame,
    build_docx_report,
    build_html_report,
    build_pdf_report,
    build_scatter_plot,
    build_source_data_export_frame,
    build_source_data_frame,
    ensure_storage,
    figure_palette_options,
    figure_to_bytes,
    get_sheet_meta,
    read_dataset,
    save_json,
    scan_upload,
)

st.set_page_config(
    page_title="Reliability Web Tool",
    layout="wide",
)

STUDY_DESIGN_OPTIONS = {
    "One-way random": "one_way_random",
    "Two-way random": "two_way_random",
    "Two-way mixed": "two_way_mixed",
}
AGREEMENT_OPTIONS = {
    "Absolute agreement": "absolute",
    "Consistency": "consistency",
}
MEASUREMENT_OPTIONS = {
    "Single measurement": "single",
    "Average measurement": "average",
}


def persist_uploaded_file(uploaded_file) -> dict:
    ensure_storage()
    suffix = Path(uploaded_file.name).suffix.lower()
    extension = suffix.lstrip(".")
    if extension not in ALLOWED_EXTENSIONS:
        raise AnalysisError("Upload a CSV or XLSX file.")

    file_bytes = uploaded_file.getvalue()
    digest = hashlib.md5(file_bytes).hexdigest()
    existing_record = st.session_state.get("upload_record")
    if existing_record and existing_record.get("digest") == digest:
        return existing_record

    stored_filename = f"{digest}.{extension}"
    stored_path = UPLOAD_DIR / stored_filename
    if not stored_path.exists():
        stored_path.write_bytes(file_bytes)

    upload_record = {
        "id": digest,
        "digest": digest,
        "original_filename": uploaded_file.name,
        "stored_filename": stored_filename,
        "file_type": extension,
        "sheets": scan_upload(stored_path, extension),
    }
    save_json(UPLOAD_DIR / f"{digest}.json", upload_record)
    st.session_state["upload_record"] = upload_record
    st.session_state.pop("analysis_record", None)
    return upload_record


def default_measurement_columns(sheet_meta: dict) -> list[str]:
    detected_pairs = sheet_meta.get("detected_pairs", [])
    if detected_pairs:
        first_pair = detected_pairs[0]
        return [first_pair["test_1"], first_pair["test_2"]]
    return sheet_meta.get("numeric_columns", [])[:2]


def render_pair_section(upload_record: dict, analysis_record: dict, pair_result: dict) -> None:
    st.subheader(pair_result["pair_label"])
    figure_palette = analysis_record["config"].get("figure_palette", DEFAULT_FIGURE_PALETTE)
    metric_columns = st.columns(6)
    metric_columns[0].metric("ICC", str(pair_result["icc_result"]["estimate"]))
    metric_columns[1].metric(
        "95% CI",
        f"{pair_result['icc_result']['ci_lower']} to {pair_result['icc_result']['ci_upper']}",
    )
    metric_columns[2].metric("Observations", str(pair_result["dataset_summary"]["observations"]))
    metric_columns[3].metric("Dropped rows", str(pair_result["dataset_summary"]["dropped_rows"]))
    metric_columns[4].metric(
        "Typical error",
        str(pair_result["pair_metrics"][0]["typical_error"] if pair_result["pair_metrics"] else "—"),
    )
    metric_columns[5].metric(
        "MDC 95%",
        str(
            pair_result["pair_metrics"][0]["minimum_detectable_change_95"]
            if pair_result["pair_metrics"]
            else "—"
        ),
    )

    st.caption(pair_result["typical_error_formula"])
    st.caption(pair_result["minimum_detectable_change_formula"])

    with st.expander("Series summaries", expanded=False):
        st.dataframe(pd.DataFrame(pair_result["column_summaries"]), use_container_width=True)

    with st.expander("Observation summaries", expanded=False):
        st.dataframe(pd.DataFrame(pair_result["observation_summaries"]), use_container_width=True)

    with st.expander("Typical error, minimum detectable change, and limits of agreement", expanded=True):
        st.dataframe(pd.DataFrame(pair_result["pair_metrics"]), use_container_width=True)

    pair_source_frame = build_source_data_frame(upload_record, analysis_record, pair_result)
    pair_frame = pair_source_frame[[pair_result["primary_x_column"], pair_result["primary_y_column"]]].copy()

    plot_column_1, plot_column_2 = st.columns(2)
    with plot_column_1:
        scatter_figure = build_scatter_plot(pair_frame, pair_result["primary_x_column"], pair_result["primary_y_column"], figure_palette)
        st.pyplot(scatter_figure, use_container_width=True)
        scatter_svg = figure_to_bytes(
            build_scatter_plot(pair_frame, pair_result["primary_x_column"], pair_result["primary_y_column"], figure_palette),
            "svg",
        )
        st.download_button(
            f"Download {pair_result['pair_label']} scatter SVG",
            data=scatter_svg,
            file_name=f"scatter-{pair_result['pair_key']}.svg",
            mime="image/svg+xml",
            key=f"scatter-{pair_result['pair_key']}",
        )

    with plot_column_2:
        bland_figure = build_bland_altman_plot(pair_frame, pair_result["primary_x_column"], pair_result["primary_y_column"], figure_palette)
        st.pyplot(bland_figure, use_container_width=True)
        bland_svg = figure_to_bytes(
            build_bland_altman_plot(pair_frame, pair_result["primary_x_column"], pair_result["primary_y_column"], figure_palette),
            "svg",
        )
        st.download_button(
            f"Download {pair_result['pair_label']} Bland-Altman SVG",
            data=bland_svg,
            file_name=f"bland-altman-{pair_result['pair_key']}.svg",
            mime="image/svg+xml",
            key=f"bland-{pair_result['pair_key']}",
        )

    with st.expander("Source data used for this pair", expanded=False):
        st.dataframe(pair_source_frame, use_container_width=True)


def main() -> None:
    ensure_storage()
    st.title("Inter-Rater Reliability Tool")
    st.caption(
        "Streamlit deployment branch for CSV/XLSX upload, ICC analysis, plots, and downloadable reports."
    )

    uploaded_file = st.file_uploader("Choose a CSV or XLSX file", type=["csv", "xlsx"])
    if not uploaded_file:
        st.info("Upload a dataset to begin.")
        return

    try:
        upload_record = persist_uploaded_file(uploaded_file)
    except AnalysisError as exc:
        st.error(str(exc))
        return

    sheet_names = [sheet["name"] for sheet in upload_record["sheets"]]
    selected_sheet = st.selectbox("Worksheet", options=sheet_names)
    sheet_meta = get_sheet_meta(upload_record, selected_sheet)
    preview_frame = read_dataset(UPLOAD_DIR / upload_record["stored_filename"], upload_record["file_type"], selected_sheet)

    with st.expander("Preview data", expanded=False):
        st.dataframe(preview_frame.head(12), use_container_width=True)

    numeric_columns = sheet_meta.get("numeric_columns", [])
    if len(numeric_columns) < 2:
        st.error("At least two numeric columns are required for reliability analysis.")
        return

    detected_pairs = sheet_meta.get("detected_pairs", [])
    pair_lookup = {pair["key"]: f"{pair['label']}: {pair['test_1']} ↔ {pair['test_2']}" for pair in detected_pairs}
    default_measurements = default_measurement_columns(sheet_meta)
    palette_options = figure_palette_options()
    palette_labels = [option["label"] for option in palette_options]
    palette_keys_by_label = {option["label"]: option["key"] for option in palette_options}
    default_palette_key = st.session_state.get("figure_palette", DEFAULT_FIGURE_PALETTE)
    default_palette_label = next(
        (option["label"] for option in palette_options if option["key"] == default_palette_key),
        palette_options[0]["label"],
    )

    selected_palette_label = st.selectbox(
        "Figure color palette",
        options=palette_labels,
        index=palette_labels.index(default_palette_label),
    )
    selected_palette_key = palette_keys_by_label[selected_palette_label]
    st.session_state["figure_palette"] = selected_palette_key

    preview_frame = build_palette_preview_frame()
    preview_column_1, preview_column_2 = st.columns(2)
    with preview_column_1:
        st.caption("Scatter preview")
        st.pyplot(build_scatter_plot(preview_frame, "Test 1", "Test 2", selected_palette_key), use_container_width=True)
    with preview_column_2:
        st.caption("Bland-Altman preview")
        st.pyplot(build_bland_altman_plot(preview_frame, "Test 1", "Test 2", selected_palette_key), use_container_width=True)

    with st.form("analysis-form"):
        subject_options = [""] + sheet_meta["columns"]
        subject_column = st.selectbox(
            "Observation ID column",
            options=subject_options,
            format_func=lambda value: "Use generated row labels" if value == "" else value,
        )
        selected_pair_keys = st.multiselect(
            "Detected reliability pairs",
            options=list(pair_lookup.keys()),
            default=list(pair_lookup.keys())[:1],
            format_func=lambda key: pair_lookup[key],
        )
        measurement_columns = st.multiselect(
            "Measurement columns",
            options=numeric_columns,
            default=default_measurements,
        )

        plot_column_1, plot_column_2 = st.columns(2)
        with plot_column_1:
            primary_x_column = st.selectbox(
                "Primary plot X column",
                options=numeric_columns,
                index=numeric_columns.index(default_measurements[0]) if default_measurements else 0,
            )
        with plot_column_2:
            default_y_index = 1 if len(default_measurements) > 1 else min(1, len(numeric_columns) - 1)
            primary_y_column = st.selectbox(
                "Primary plot Y column",
                options=numeric_columns,
                index=numeric_columns.index(default_measurements[1]) if len(default_measurements) > 1 else default_y_index,
            )

        setting_column_1, setting_column_2, setting_column_3 = st.columns(3)
        with setting_column_1:
            study_design_label = st.selectbox("Study design", options=list(STUDY_DESIGN_OPTIONS.keys()), index=1)
        with setting_column_2:
            agreement_label = st.selectbox("Agreement target", options=list(AGREEMENT_OPTIONS.keys()), index=0)
        with setting_column_3:
            measurement_label = st.selectbox("Measurement unit", options=list(MEASUREMENT_OPTIONS.keys()), index=0)

        submitted = st.form_submit_button("Run reliability analysis")

    if submitted:
        try:
            analysis_result = analyse_wide_dataset(
                upload_record=upload_record,
                sheet_name=selected_sheet,
                subject_column=subject_column or None,
                measurement_columns=measurement_columns,
                primary_x_column=primary_x_column,
                primary_y_column=primary_y_column,
                selected_pair_keys=selected_pair_keys,
                study_design=STUDY_DESIGN_OPTIONS[study_design_label],
                agreement_definition=AGREEMENT_OPTIONS[agreement_label],
                measurement_unit=MEASUREMENT_OPTIONS[measurement_label],
                figure_palette=selected_palette_key,
                figure_action="both",
            )
            st.session_state["analysis_record"] = {"id": uuid4().hex, **analysis_result}
        except AnalysisError as exc:
            st.error(str(exc))
            return

    analysis_record = st.session_state.get("analysis_record")
    if not analysis_record or analysis_record["config"]["upload_id"] != upload_record["id"]:
        return

    st.success(analysis_record["recommendation"]["rationale"])

    source_csv = build_source_data_export_frame(upload_record, analysis_record).to_csv(index=False).encode("utf-8")
    html_bytes = build_html_report(analysis_record).encode("utf-8")
    pdf_bytes = build_pdf_report(analysis_record)
    docx_bytes = build_docx_report(analysis_record)

    download_column_1, download_column_2, download_column_3, download_column_4 = st.columns(4)
    download_column_1.download_button(
        "Download analysed source data (CSV)",
        data=source_csv,
        file_name="reliability-source-data.csv",
        mime="text/csv",
    )
    download_column_2.download_button(
        "Download results report (PDF)",
        data=pdf_bytes,
        file_name="reliability-results.pdf",
        mime="application/pdf",
    )
    download_column_3.download_button(
        "Download results report (DOCX)",
        data=docx_bytes,
        file_name="reliability-results.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )
    download_column_4.download_button(
        "Download results report (HTML)",
        data=html_bytes,
        file_name="reliability-results.html",
        mime="text/html",
    )

    for pair_result in analysis_record["pair_results"]:
        render_pair_section(upload_record, analysis_record, pair_result)
        st.divider()


if __name__ == "__main__":
    main()
