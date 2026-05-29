from __future__ import annotations

import hashlib
import itertools
import math
import statistics
from pathlib import Path
from uuid import uuid4

import matplotlib.pyplot as plt
import pandas as pd
import streamlit as st
from scipy import stats as scipy_stats

from app import (
    ALLOWED_EXTENSIONS,
    AnalysisError,
    DEFAULT_FIGURE_PALETTE,
    UPLOAD_DIR,
    analyse_long_dataset,
    analyse_wide_dataset,
    build_bland_altman_plot,
    build_long_measurement_options,
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
    get_upload_path,
    get_sheet_meta,
    list_sample_datasets,
    load_sample_upload,
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


def default_pair_selections(sheet_meta: dict, analysis_record: dict | None = None) -> list[dict]:
    if analysis_record:
        saved_pair_selections = analysis_record.get("config", {}).get("pair_selections", [])
        if saved_pair_selections:
            return saved_pair_selections

    detected_pairs = sheet_meta.get("detected_pairs", [])
    if detected_pairs:
        return [
            {
                "pair_label": pair["label"],
                "primary_x_column": pair["test_1"],
                "primary_y_column": pair["test_2"],
            }
            for pair in detected_pairs[: min(3, len(detected_pairs))]
        ]

    numeric_columns = sheet_meta.get("numeric_columns", [])
    if len(numeric_columns) >= 2:
        return [
            {
                "pair_label": f"{numeric_columns[0]} vs {numeric_columns[1]}",
                "primary_x_column": numeric_columns[0],
                "primary_y_column": numeric_columns[1],
            }
        ]

    return []


def run_correlation_fragility(x_values: list, y_values: list, n: int) -> dict:
    """Exhaustively test every combination of n replacements to find the worst-case p-value."""
    data = list(zip(x_values, y_values))
    med_x = statistics.median(x_values)
    med_y = statistics.median(y_values)
    median_point = (med_x, med_y)
    orig_r, orig_p = scipy_stats.pearsonr(x_values, y_values)
    replacement_candidates = list(itertools.combinations(data, n))

    p2values: list[float] = []
    r2values: list[float] = []
    for combo in replacement_candidates:
        temp = list(data)
        for point in combo:
            temp.remove(point)
        for _ in range(n):
            temp.append(median_point)
        rx = [pt[0] for pt in temp]
        ry = [pt[1] for pt in temp]
        try:
            r2, p2 = scipy_stats.pearsonr(rx, ry)
        except Exception:
            r2, p2 = float("nan"), 1.0
        p2values.append(p2)
        r2values.append(r2)

    maxpos = p2values.index(max(p2values))
    return {
        "r2values": r2values,
        "p2values": p2values,
        "replacement_candidates": replacement_candidates,
        "maxpos": maxpos,
        "original_r": orig_r,
        "original_p": orig_p,
        "median_x": med_x,
        "median_y": med_y,
        "n": n,
    }


def build_fragility_line_chart(fragility_result: dict):
    r2values = fragility_result["r2values"]
    p2values = fragility_result["p2values"]
    orig_r = fragility_result["original_r"]
    orig_p = fragility_result["original_p"]

    fig, ax1 = plt.subplots(figsize=(8, 4))
    color_r = "tab:red"
    color_p = "tab:blue"

    ax1.set_xlabel("Combination number", fontsize=9)
    ax1.set_ylabel("Pearson r", color=color_r)
    ax1.set_ylim(-1, 1)
    ax1.plot(range(len(r2values)), r2values, color=color_r, linewidth=0.8, label=f"fragile r")
    ax1.axhline(orig_r, linestyle="dotted", color=color_r, label=f"original r = {orig_r:.3f}")
    ax1.tick_params(axis="y", labelcolor=color_r)

    ax2 = ax1.twinx()
    ax2.set_ylabel("p value", color=color_p)
    ax2.plot(range(len(p2values)), p2values, color=color_p, linewidth=0.8, label="fragile p")
    ax2.axhline(orig_p, linestyle="dotted", color=color_p, label=f"original p = {orig_p:.4f}")
    ax2.axhline(0.05, linestyle="dashed", color="gray", linewidth=0.8, label="p = 0.05 threshold")
    ax2.tick_params(axis="y", labelcolor=color_p)

    lines1, labels1 = ax1.get_legend_handles_labels()
    lines2, labels2 = ax2.get_legend_handles_labels()
    ax1.legend(lines1 + lines2, labels1 + labels2, fontsize=7, loc="upper left")

    plt.tight_layout()
    return fig


def build_fragility_scatter(x_values: list, y_values: list, fragility_result: dict):
    maxpos = fragility_result["maxpos"]
    worst_combo = fragility_result["replacement_candidates"][maxpos]
    n = fragility_result["n"]
    med_x = fragility_result["median_x"]
    med_y = fragility_result["median_y"]
    max_p = max(fragility_result["p2values"])
    worst_r = fragility_result["r2values"][maxpos]

    x_replaced = [pt[0] for pt in worst_combo]
    y_replaced = [pt[1] for pt in worst_combo]

    fig, ax = plt.subplots(figsize=(6, 4))
    ax.scatter(x_values, y_values, alpha=0.7, label="Original data")
    ax.scatter(x_replaced, y_replaced, color="red", zorder=5, label=f"Point{'s' if n > 1 else ''} to replace (n={n})")
    ax.scatter(med_x, med_y, color="orange", marker="*", s=150, zorder=6, label="Median (replacement target)")
    ax.set_xlabel("X")
    ax.set_ylabel("Y")
    ax.set_title(f"Worst-case: r = {worst_r:.3f}, p = {max_p:.4f}", fontsize=9)
    ax.legend(fontsize=8)
    plt.tight_layout()
    return fig


def render_pair_section(upload_record: dict, analysis_record: dict, pair_result: dict) -> None:
    st.subheader(pair_result["pair_label"])
    figure_palette = analysis_record["config"].get("figure_palette", DEFAULT_FIGURE_PALETTE)
    metric_columns = st.columns(6)
    metric_columns[0].metric("ICC", pair_result["icc_result"]["estimate_display"])
    metric_columns[1].metric(
        "95% CI",
        f"{pair_result['icc_result']['ci_lower_display']} to {pair_result['icc_result']['ci_upper_display']}",
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

    # --- Correlation Fragility ---
    st.markdown("#### Correlation Fragility")
    with st.expander("What is correlation fragility?", expanded=False):
        st.markdown(
            """
**Correlation fragility** asks a simple but important question: *how many data points need to change
before a statistically significant correlation disappears?*

The analysis works by exhaustively testing every possible combination of **n** data points.
For each combination, those points are temporarily replaced with the dataset median (the middle X value
and the middle Y value) and the Pearson correlation is recalculated.
The combination that produces the highest — least significant — p-value is reported as the worst case.

---

**How to use this tool**

1. After running a reliability analysis, scroll to the *Correlation Fragility* section for any pair.
2. Choose **n** — the number of simultaneous replacements to test. Start with n = 1; increase to n = 2 or 3
   if the correlation survives a single replacement comfortably.
3. Check the combination count shown below the selector. Counts above ~10,000 will be slow;
   reduce n or accept the wait.
4. Click **Run Correlation Fragility** to start the exhaustive search.

---

**Interpreting the results**

| Finding | Implication |
|---|---|
| Worst-case p remains < 0.05 even for n = 2 or 3 | The correlation is **robust** — it survives multiple replacements and is unlikely to be driven by a handful of influential points. |
| Worst-case p ≥ 0.05 for n = 1 | The correlation is **highly fragile** — a single data point is keeping it statistically significant. Treat the original result with caution. |
| Worst-case p ≥ 0.05 for small n (e.g. 2–3 in a dataset of 30+) | The correlation is **moderately fragile**. Consider whether the flagged points are genuine observations or measurement artefacts. |

A fragile correlation does not mean the relationship is false, but it does mean the statistical
significance rests on a small number of observations. In a reliability context this is important:
if the ICC is supported by a significant Pearson correlation, and that correlation is fragile, the
reliability estimate may not generalise well to a larger or more varied sample.

---

*The method used here is described in: Lakens, D. (2014). Performing high-powered studies efficiently
with sequential analyses. European Journal of Social Psychology, 44(7), 701–710, and popularised by
the correlation fragility index approach of Voracek et al.*
"""
        )

    pair_key = pair_result["pair_key"]
    frag_state_key = f"frag-{pair_key}"

    x_vals = pair_frame[pair_result["primary_x_column"]].dropna().tolist()
    y_vals = pair_frame[pair_result["primary_y_column"]].dropna().tolist()
    row_count = len(x_vals)

    if row_count < 3:
        st.info("At least 3 data points are required for correlation fragility analysis.")
    else:
        orig_r, orig_p = scipy_stats.pearsonr(x_vals, y_vals)
        if orig_p >= 0.05:
            st.warning(
                f"The original correlation for this pair is not statistically significant "
                f"(r = {orig_r:.3f}, p = {orig_p:.4f}). "
                "Fragility analysis is most meaningful when the correlation is significant."
            )

        n_replacements = int(
            st.number_input(
                "Number of replacements (n)",
                min_value=1,
                max_value=min(5, row_count - 2),
                value=1,
                step=1,
                key=f"{frag_state_key}-n",
                help=(
                    "How many data points to simultaneously replace with the X/Y median. "
                    "Increasing n tests larger combinations but grows computation time rapidly."
                ),
            )
        )

        combo_count = math.comb(row_count, n_replacements)
        st.caption(f"{row_count} data points → {combo_count:,} combinations to test (n = {n_replacements}).")

        COMBO_WARN_THRESHOLD = 10_000
        if combo_count > COMBO_WARN_THRESHOLD:
            st.warning(
                f"This will test {combo_count:,} combinations, which may be slow. "
                "Consider reducing n."
            )

        if st.button("Run Correlation Fragility", key=f"{frag_state_key}-btn"):
            with st.spinner(f"Testing {combo_count:,} combinations…"):
                frag_result = run_correlation_fragility(x_vals, y_vals, n_replacements)
            st.session_state[frag_state_key] = frag_result

        frag_result = st.session_state.get(frag_state_key)
        if frag_result and frag_result.get("n") == n_replacements:
            max_p = max(frag_result["p2values"])
            worst_r = frag_result["r2values"][frag_result["maxpos"]]
            worst_combo = frag_result["replacement_candidates"][frag_result["maxpos"]]

            frag_metric_cols = st.columns(4)
            frag_metric_cols[0].metric("Original r", f"{frag_result['original_r']:.4f}")
            frag_metric_cols[1].metric("Original p", f"{frag_result['original_p']:.4f}")
            frag_metric_cols[2].metric(f"Worst-case r (n={n_replacements})", f"{worst_r:.4f}")
            frag_metric_cols[3].metric(f"Worst-case p (n={n_replacements})", f"{max_p:.4f}")

            replaced_str = ", ".join(f"({pt[0]:.3g}, {pt[1]:.3g})" for pt in worst_combo)
            st.caption(
                f"Replacing {n_replacements} point(s) — {replaced_str} — with the median "
                f"({frag_result['median_x']:.3g}, {frag_result['median_y']:.3g}) "
                f"achieves the maximum p-value of {max_p:.4f}."
            )

            frag_col1, frag_col2 = st.columns(2)
            with frag_col1:
                st.pyplot(build_fragility_line_chart(frag_result), use_container_width=True)
            with frag_col2:
                st.pyplot(build_fragility_scatter(x_vals, y_vals, frag_result), use_container_width=True)


def main() -> None:
    ensure_storage()
    st.title("Inter-Rater Reliability Tool")
    st.caption(
        "Streamlit deployment branch for CSV/XLSX upload, ICC analysis, plots, and downloadable reports."
    )

    sample_datasets = list_sample_datasets()
    st.markdown("### Start with your own data or a tutorial sample")
    upload_tab, tutorial_tab = st.tabs(["Upload your own file", "Use a tutorial sample"])

    uploaded_file = None
    with upload_tab:
        uploaded_file = st.file_uploader("Choose a CSV or XLSX file", type=["csv", "xlsx"])

    with tutorial_tab:
        st.caption("The tutorial samples open real example workbooks from the repository so you can learn the wide-format and long-format workflows without preparing your own file first.")
        for sample_dataset in sample_datasets:
            st.markdown(f"#### {sample_dataset['title']}")
            st.write(sample_dataset["description"])
            st.caption(sample_dataset["tutorial_notes"])
            st.caption(sample_dataset["audience"])
            st.caption(
                f"Preferred worksheet: {sample_dataset['preferred_sheet']} | Sheets: {sample_dataset['sheet_count']} | File: {sample_dataset['filename']}"
            )
            with st.expander(f"Preview {sample_dataset['title']}", expanded=False):
                st.dataframe(pd.DataFrame(sample_dataset["preview_rows"]), use_container_width=True, hide_index=True)
            if st.button(f"Use {sample_dataset['title']}", key=f"sample-{sample_dataset['key']}"):
                st.session_state["selected_sample_key"] = sample_dataset["key"]
                st.session_state.pop("analysis_record", None)
                st.rerun()
            st.divider()

    upload_record = None
    if uploaded_file is not None:
        try:
            upload_record = persist_uploaded_file(uploaded_file)
            st.session_state.pop("selected_sample_key", None)
        except AnalysisError as exc:
            st.error(str(exc))
            return
    else:
        selected_sample_key = st.session_state.get("selected_sample_key")
        if selected_sample_key:
            try:
                upload_record = load_sample_upload(selected_sample_key)
                st.session_state["upload_record"] = upload_record
            except AnalysisError as exc:
                st.error(str(exc))
                return

    if not upload_record:
        st.info("Upload a dataset or choose one of the tutorial sample files to begin.")
        return

    if upload_record.get("id", "").startswith("sample-"):
        active_sample = next(
            (sample for sample in sample_datasets if f"sample-{sample['key']}" == upload_record["id"]),
            None,
        )
        if active_sample:
            st.success(f"Tutorial sample loaded: {active_sample['title']}")
            st.caption(active_sample["tutorial_notes"])

    sheet_names = [sheet["name"] for sheet in upload_record["sheets"]]
    selected_sheet = st.selectbox("Worksheet", options=sheet_names)
    sheet_meta = get_sheet_meta(upload_record, selected_sheet)
    source_preview_frame = read_dataset(get_upload_path(upload_record), upload_record["file_type"], selected_sheet)

    with st.expander("Preview data", expanded=False):
        st.dataframe(source_preview_frame.head(12), use_container_width=True)

    numeric_columns = sheet_meta.get("numeric_columns", [])
    if len(numeric_columns) < 2:
        st.error("At least two numeric columns are required for reliability analysis.")
        return

    structure_detection = sheet_meta.get("structure_detection", {})
    analysis_record = st.session_state.get("analysis_record")
    if analysis_record and analysis_record["config"].get("upload_id") != upload_record["id"]:
        analysis_record = None
    if analysis_record and analysis_record["config"].get("selected_sheet") != selected_sheet:
        analysis_record = None

    suggested_format = structure_detection.get("format", "wide")
    saved_data_format = analysis_record["config"].get("data_format") if analysis_record else None
    selected_data_format = saved_data_format or ("long" if suggested_format == "long" else "wide")

    headerless = sheet_meta.get("headerless", False)
    if headerless:
        st.info(
            "**No column headers detected.** "
            "This file appears to have no header row — all values in the first row are numeric. "
            "Columns have been automatically named **Column 1**, **Column 2**, etc. "
            "There is no participant ID column available; row numbers will be used instead."
        )

    detection_label = {
        "wide": "Wide format detected",
        "long": "Long format detected",
        "uncertain": "Data layout is uncertain",
    }.get(suggested_format, "Data layout review")
    st.markdown(f"**{detection_label}**")
    for reason in structure_detection.get("reasons", []):
        st.caption(reason)

    selected_data_format = st.radio(
        "Data layout",
        options=["wide", "long"],
        index=0 if selected_data_format == "wide" else 1,
        format_func=lambda value: "Wide format" if value == "wide" else "Long format",
        help="Confirm whether your data is arranged with one column per repeated measurement or with one row per observation-measurement combination.",
        horizontal=True,
    )

    detected_pairs = sheet_meta.get("detected_pairs", [])
    default_pair_rows = default_pair_selections(sheet_meta, analysis_record)
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
        help=(
            "Pick the colour style used for the scatter plot and Bland-Altman plot. "
            "This changes the look of the figures only and does not change the analysis results."
        ),
    )
    selected_palette_key = palette_keys_by_label[selected_palette_label]
    st.session_state["figure_palette"] = selected_palette_key

    palette_preview_frame = build_palette_preview_frame()
    preview_column_1, preview_column_2 = st.columns(2)
    with preview_column_1:
        st.caption("Scatter preview")
        st.pyplot(build_scatter_plot(palette_preview_frame, "Test 1", "Test 2", selected_palette_key), use_container_width=True)
    with preview_column_2:
        st.caption("Bland-Altman preview")
        st.pyplot(build_bland_altman_plot(palette_preview_frame, "Test 1", "Test 2", selected_palette_key), use_container_width=True)
    if selected_data_format == "wide":
        with st.expander("Available measurement columns", expanded=False):
            st.dataframe(
                pd.DataFrame({"Measurement column": numeric_columns}),
                use_container_width=True,
                hide_index=True,
            )

        with st.expander("Suggested column pairs", expanded=False):
            if detected_pairs:
                st.dataframe(
                    pd.DataFrame(
                        [
                            {
                                "Suggested pair": pair["label"],
                                "X column": pair["test_1"],
                                "Y column": pair["test_2"],
                            }
                            for pair in detected_pairs
                        ]
                    ),
                    use_container_width=True,
                    hide_index=True,
                )
                st.caption("Use these suggestions as a guide, then confirm the exact pairs you want below.")
            else:
                st.info("No automatic Test 1/Test 2 suggestions were detected for this worksheet.")

        pair_count = st.number_input(
            "Number of analysis pairs",
            min_value=1,
            max_value=min(12, max(1, len(numeric_columns) * (len(numeric_columns) - 1))),
            value=max(1, len(default_pair_rows)),
            step=1,
            help=(
                "Choose how many X/Y column pairs you want to analyse. "
                "Each pair below becomes a separate reliability analysis with its own plots and results."
            ),
            key=f"pair-count-{upload_record['id']}-{selected_sheet}",
        )
    else:
        all_columns = sheet_meta.get("columns", [])
        suggested_long = structure_detection.get("suggested_columns", {}).get("long", {})
        default_subject_column = suggested_long.get("subject_column") or (analysis_record["config"].get("subject_column") if analysis_record else "") or all_columns[0]
        default_measurement_column = suggested_long.get("measurement_column") or (analysis_record["config"].get("long_measurement_column") if analysis_record else "") or (all_columns[1] if len(all_columns) > 1 else all_columns[0])
        default_rater_column = suggested_long.get("rater_column") or (analysis_record["config"].get("long_rater_column") if analysis_record else "") or (all_columns[2] if len(all_columns) > 2 else all_columns[0])
        default_score_column = suggested_long.get("score_column") or (analysis_record["config"].get("long_score_column") if analysis_record else "") or numeric_columns[-1]

        long_subject_column = st.selectbox(
            "Observation ID column",
            options=all_columns,
            index=all_columns.index(default_subject_column) if default_subject_column in all_columns else 0,
            help="Choose the column that identifies each person, item, or case across repeated rows.",
            key=f"long-subject-{upload_record['id']}-{selected_sheet}",
        )
        long_column_1, long_column_2, long_column_3 = st.columns(3)
        with long_column_1:
            long_measurement_column = st.selectbox(
                "Measurement name column",
                options=all_columns,
                index=all_columns.index(default_measurement_column) if default_measurement_column in all_columns else 0,
                help="Choose the column that names which variable or metric each row belongs to.",
                key=f"long-measurement-{upload_record['id']}-{selected_sheet}",
            )
        with long_column_2:
            long_rater_column = st.selectbox(
                "Repeated-measure level column",
                options=all_columns,
                index=all_columns.index(default_rater_column) if default_rater_column in all_columns else 0,
                help="Choose the column that identifies the repeated measure, such as Test 1/Test 2, session, occasion, or rater.",
                key=f"long-rater-{upload_record['id']}-{selected_sheet}",
            )
        with long_column_3:
            long_score_column = st.selectbox(
                "Score column",
                options=all_columns,
                index=all_columns.index(default_score_column) if default_score_column in all_columns else 0,
                help="Choose the numeric column that contains the actual score or value to analyse.",
                key=f"long-score-{upload_record['id']}-{selected_sheet}",
            )

        rater_values = []
        if long_rater_column in source_preview_frame.columns:
            rater_values = sorted(source_preview_frame[long_rater_column].dropna().astype(str).unique().tolist())
        measurement_values = build_long_measurement_options(source_preview_frame, long_measurement_column, long_rater_column)

        default_selected_measurements = []
        if analysis_record and analysis_record["config"].get("data_format") == "long":
            default_selected_measurements = [
                value for value in analysis_record["config"].get("long_selected_measurements", []) if value in measurement_values
            ]
        if not default_selected_measurements:
            default_selected_measurements = measurement_values[: min(3, len(measurement_values))]

        st.markdown("**Long-format measurements to analyse**")
        selected_measurements = st.multiselect(
            "Measurements",
            options=measurement_values,
            default=default_selected_measurements,
            help="Select one or more measurement names to analyse. Each selected measurement becomes a separate result section.",
        )

        if len(rater_values) < 2:
            st.error("At least two repeated-measure levels are required in the selected long-format level column.")
            return

        default_x_rater = (analysis_record["config"].get("long_x_rater_value") if analysis_record and analysis_record["config"].get("data_format") == "long" else "") or rater_values[0]
        default_y_rater = (analysis_record["config"].get("long_y_rater_value") if analysis_record and analysis_record["config"].get("data_format") == "long" else "") or rater_values[min(1, len(rater_values) - 1)]
        level_column_1, level_column_2 = st.columns(2)
        with level_column_1:
            long_x_rater_value = st.selectbox(
                "First repeated-measure level",
                options=rater_values,
                index=rater_values.index(default_x_rater) if default_x_rater in rater_values else 0,
                help="Choose the first repeated-measure level to compare, such as Test 1.",
            )
        with level_column_2:
            long_y_rater_value = st.selectbox(
                "Second repeated-measure level",
                options=rater_values,
                index=rater_values.index(default_y_rater) if default_y_rater in rater_values else min(1, len(rater_values) - 1),
                help="Choose the second repeated-measure level to compare, such as Test 2.",
            )

        with st.expander("Available long-format columns", expanded=False):
            st.dataframe(
                pd.DataFrame({"Column": all_columns}),
                use_container_width=True,
                hide_index=True,
            )

    with st.form("analysis-form"):
        if selected_data_format == "wide":
            if headerless:
                st.caption("No participant ID column — row numbers will be used as observation labels.")
                subject_column = ""
            else:
                subject_options = [""] + sheet_meta["columns"]
                subject_column = st.selectbox(
                    "Observation ID column",
                    options=subject_options,
                    help=(
                        "Choose the column that identifies each person, item, or case. "
                        "This helps the app keep each row matched correctly across repeated measurements. "
                        "If your file does not have an ID column, use generated row labels."
                    ),
                    format_func=lambda value: "Use generated row labels" if value == "" else value,
                )

            st.markdown("**Pairs to analyse**")
            st.caption("Pick the exact column headers you want to compare. Each row below runs as a separate analysis.")

            pair_selections: list[dict] = []
            for pair_index in range(int(pair_count)):
                default_pair = default_pair_rows[pair_index] if pair_index < len(default_pair_rows) else {}
                pair_columns = st.columns([1.2, 1, 1])
                with pair_columns[0]:
                    pair_label = st.text_input(
                        f"Pair {pair_index + 1} label",
                        value=default_pair.get("pair_label", ""),
                        help="Optional short name for this comparison. If left blank, the app will use the selected X and Y column names.",
                        key=f"pair-label-{upload_record['id']}-{selected_sheet}-{pair_index}",
                    )
                with pair_columns[1]:
                    default_x = default_pair.get("primary_x_column", numeric_columns[0])
                    default_x_index = numeric_columns.index(default_x) if default_x in numeric_columns else 0
                    pair_x_column = st.selectbox(
                        f"Pair {pair_index + 1} X column",
                        options=numeric_columns,
                        index=default_x_index,
                        help="Choose the measurement column to place on the horizontal axis and treat as the first value in this pair.",
                        key=f"pair-x-{upload_record['id']}-{selected_sheet}-{pair_index}",
                    )
                with pair_columns[2]:
                    fallback_y = numeric_columns[min(1, len(numeric_columns) - 1)]
                    default_y = default_pair.get("primary_y_column", fallback_y)
                    default_y_index = numeric_columns.index(default_y) if default_y in numeric_columns else min(1, len(numeric_columns) - 1)
                    pair_y_column = st.selectbox(
                        f"Pair {pair_index + 1} Y column",
                        options=numeric_columns,
                        index=default_y_index,
                        help="Choose the measurement column to place on the vertical axis and treat as the second value in this pair.",
                        key=f"pair-y-{upload_record['id']}-{selected_sheet}-{pair_index}",
                    )

                pair_selections.append(
                    {
                        "pair_label": pair_label,
                        "primary_x_column": pair_x_column,
                        "primary_y_column": pair_y_column,
                    }
                )
        else:
            pair_selections = []
            subject_column = long_subject_column

        setting_column_1, setting_column_2, setting_column_3 = st.columns(3)
        with setting_column_1:
            study_design_label = st.selectbox(
                "Study design",
                options=list(STUDY_DESIGN_OPTIONS.keys()),
                index=1,
                help=(
                    "Choose how the repeated measurements were collected. "
                    "Use one-way random when different people, devices, or occasions may have measured different subjects. "
                    "Use two-way random when the same people, devices, or occasions measured everyone and you want the result to apply more broadly. "
                    "Use two-way mixed when the same specific people or devices measured everyone and you only care about those exact measurers."
                ),
            )
        with setting_column_2:
            agreement_label = st.selectbox(
                "Agreement target",
                options=list(AGREEMENT_OPTIONS.keys()),
                index=0,
                help=(
                    "Choose what kind of matching you want between repeated measurements. "
                    "Use absolute agreement when repeated measurements should be close to the same actual value. "
                    "Use consistency when repeated measurements can differ a little in level, but should still rank people in a similar order."
                ),
            )
        with setting_column_3:
            measurement_label = st.selectbox(
                "Measurement unit",
                options=list(MEASUREMENT_OPTIONS.keys()),
                index=0,
                help=(
                    "Choose whether you want reliability for one measurement or for an average. "
                    "Use single measurement when decisions will be based on one test, one rater, or one occasion at a time. "
                    "Use average measurement when the final score will be the average of multiple tests, raters, or occasions."
                ),
            )

        submitted = st.form_submit_button("Run reliability analysis")

    if submitted:
        try:
            if selected_data_format == "wide":
                measurement_columns = list(
                    dict.fromkeys(
                        column
                        for pair_selection in pair_selections
                        for column in (
                            pair_selection["primary_x_column"],
                            pair_selection["primary_y_column"],
                        )
                    )
                )
                analysis_result = analyse_wide_dataset(
                    upload_record=upload_record,
                    sheet_name=selected_sheet,
                    subject_column=subject_column or None,
                    measurement_columns=measurement_columns,
                    primary_x_column=pair_selections[0]["primary_x_column"] if pair_selections else "",
                    primary_y_column=pair_selections[0]["primary_y_column"] if pair_selections else "",
                    selected_pair_keys=[],
                    study_design=STUDY_DESIGN_OPTIONS[study_design_label],
                    agreement_definition=AGREEMENT_OPTIONS[agreement_label],
                    measurement_unit=MEASUREMENT_OPTIONS[measurement_label],
                    figure_palette=selected_palette_key,
                    figure_action="both",
                    pair_selections=pair_selections,
                )
            else:
                analysis_result = analyse_long_dataset(
                    upload_record=upload_record,
                    sheet_name=selected_sheet,
                    subject_column=long_subject_column,
                    measurement_column=long_measurement_column,
                    rater_column=long_rater_column,
                    score_column=long_score_column,
                    selected_measurements=selected_measurements,
                    x_rater_value=long_x_rater_value,
                    y_rater_value=long_y_rater_value,
                    study_design=STUDY_DESIGN_OPTIONS[study_design_label],
                    agreement_definition=AGREEMENT_OPTIONS[agreement_label],
                    measurement_unit=MEASUREMENT_OPTIONS[measurement_label],
                    figure_palette=selected_palette_key,
                    figure_action="both",
                )
            st.session_state["analysis_record"] = {"id": uuid4().hex, **analysis_result}
            analysis_record = st.session_state["analysis_record"]
        except AnalysisError as exc:
            st.error(str(exc))
            return

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
