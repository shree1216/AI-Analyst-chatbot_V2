"""Read-only spreadsheet analysis shared by the interactive workspace and exports."""
from dataclasses import dataclass
from datetime import date
import logging

import pandas as pd
import streamlit as st

from data_explorer import (
    ChartResult, ChartSpec, auto_charts, build_chart, categories, column_kind, data_overview,
    dates, missing_mask, numbers,
)
from excel_support import EXCEL_MIME, generate_excel_download, read_excel_sheet
from workbook_charts import MAX_CHART_POINTS, apply_excel_dashboard


logger = logging.getLogger(__name__)
AGGREGATES = {"Row count": "count", "Sum": "sum", "Average": "average", "Median": "median", "Minimum": "min", "Maximum": "max"}
CHART_OPTIONS = {
    "Grouped summary": ["Auto", "bar", "column", "pie", "doughnut"],
    "Time trend": ["Auto", "line", "area", "column", "bar"],
    "Distribution": ["histogram"],
    "Relationship": ["scatter"],
}


@dataclass(frozen=True)
class ColumnFilter:
    column: str
    kind: str
    values: tuple[str, ...] = ()
    lower: float | date | None = None
    upper: float | date | None = None
    text: str = ""
    include_missing: bool = False


def filter_rows(frame: pd.DataFrame, filters: list[ColumnFilter]) -> pd.DataFrame:
    """Apply inclusive filters with AND semantics without altering source cells."""
    mask = pd.Series(True, index=frame.index)
    for rule in filters:
        if rule.column not in frame.columns:
            raise ValueError(f"Unknown filter column: {rule.column}.")
        series = frame[rule.column]
        if rule.kind in {"number", "date"}:
            if rule.lower is None or rule.upper is None or rule.lower > rule.upper:
                raise ValueError(f"Choose an ordered lower and upper bound for {rule.column}.")
            cleaned = numbers(series) if rule.kind == "number" else dates(series).dt.date
            selected = cleaned.between(rule.lower, rule.upper).fillna(False)
            missing = cleaned.isna()
        elif rule.kind == "category":
            cleaned = categories(series)
            selected = cleaned.isin(rule.values)
            missing = cleaned.isna()
        elif rule.kind == "contains":
            selected = series.astype(str).str.contains(rule.text.strip(), case=False, regex=False) & ~missing_mask(series)
            missing = missing_mask(series)
        else:
            raise ValueError(f"Unsupported filter type: {rule.kind}.")
        mask &= selected | (missing & rule.include_missing)
    return frame.loc[mask].copy()


def analysis_spec(
    frame: pd.DataFrame, analysis: str, *, dimension: str | None = None, measure: str | None = None,
    aggregate: str = "count", chart_type: str = "Auto", period: str = "month",
) -> ChartSpec:
    if frame.empty:
        raise ValueError("No rows match these filters. Widen or remove a filter.")
    if analysis not in CHART_OPTIONS or chart_type not in CHART_OPTIONS[analysis]:
        raise ValueError("Choose a chart type supported by this analysis.")
    if aggregate not in AGGREGATES.values():
        raise ValueError("Choose a supported summary calculation.")
    if analysis in {"Distribution", "Relationship"} or aggregate != "count":
        if measure not in frame.columns or not numbers(frame[measure]).notna().any():
            raise ValueError("Choose a measure containing numeric values.")
    if analysis != "Distribution" and dimension not in frame.columns:
        raise ValueError("Choose a grouping or X-axis column.")
    if analysis == "Relationship":
        return ChartSpec("scatter", "scatter", measure, dimension, "none")
    if analysis == "Distribution":
        return ChartSpec("histogram", "histogram", measure)
    if analysis == "Time trend":
        if period not in {"month", "year"} or not dates(frame[dimension]).notna().any():
            raise ValueError("Choose a readable date column and a month or year period.")
        return ChartSpec("time", "line" if chart_type == "Auto" else chart_type,
                         measure if aggregate != "count" else None, dimension, aggregate, period)
    groups = categories(frame[dimension]).nunique()
    if groups > MAX_CHART_POINTS:
        raise ValueError(f"Filter to at most {MAX_CHART_POINTS} groups before charting. No groups were dropped.")
    selected_type = chart_type
    if chart_type == "Auto":
        selected_type = "pie" if aggregate == "count" and 2 <= groups <= 6 else "bar"
    return ChartSpec("count" if aggregate == "count" else "group", selected_type,
                     measure if aggregate != "count" else None, dimension, aggregate,
                     category_limit=MAX_CHART_POINTS)


def export_analysis(frame: pd.DataFrame, chart: ChartResult, notes: list[tuple[str, str]]) -> bytes:
    content = generate_excel_download([
        ("Filtered data", frame), ("Analysis notes", pd.DataFrame(notes, columns=["Setting", "Value"])),
    ])
    result = apply_excel_dashboard(
        content, "Filtered data", [(chart.title, chart.table, chart.chart_type)],
        source_rows=len(frame), source_columns=len(frame.columns),
        missing_values=int(data_overview(frame)["Missing values"].sum()), placement="new_sheet",
    )
    return result.content


def _filter_controls(frame: pd.DataFrame, key: str) -> list[ColumnFilter]:
    filters = []
    selected = st.multiselect("Filter columns", list(frame.columns), key=f"{key}_filters")
    for column in selected:
        series = frame[column]
        kind = column_kind(series)
        prefix = f"{key}_filter_{list(frame.columns).index(column)}"
        st.markdown(f"**{column}**")
        include_missing = st.checkbox("Include missing/unreadable values", key=f"{prefix}_missing")
        if kind in {"whole number", "decimal number"}:
            values = numbers(series).dropna()
            left, right = st.columns(2)
            lower = left.number_input("Minimum (inclusive)", value=float(values.min()), key=f"{prefix}_lower")
            upper = right.number_input("Maximum (inclusive)", value=float(values.max()), key=f"{prefix}_upper")
            filters.append(ColumnFilter(column, "number", lower=lower, upper=upper, include_missing=include_missing))
        elif kind == "date":
            values = dates(series).dropna().dt.date
            bounds = st.date_input(
                "Date range (inclusive)", value=(values.min(), values.max()), key=f"{prefix}_dates",
            )
            if not isinstance(bounds, tuple) or len(bounds) != 2:
                raise ValueError(f"Select both a start and end date for {column}.")
            filters.append(ColumnFilter(column, "date", lower=bounds[0], upper=bounds[1], include_missing=include_missing))
        else:
            options = sorted(categories(series).dropna().unique())
            if len(options) <= 200:
                values = st.multiselect("Keep values", options, default=options, key=f"{prefix}_values")
                filters.append(ColumnFilter(column, "category", tuple(values), include_missing=include_missing))
            else:
                text = st.text_input("Contains text (case-insensitive)", key=f"{prefix}_text")
                filters.append(ColumnFilter(column, "contains", text=text, include_missing=include_missing))
    return filters


def render_spreadsheet_analysis(content: bytes, sheet: str, source: str, key: str) -> None:
    with st.container(border=True):
        st.subheader("Spreadsheet analysis")
        st.caption(f"Source: {source}. Uses the full worksheet; filters do not change your workbook or chat answers.")
        try:
            cached = st.session_state.get("spreadsheet_analysis_frame")
            if cached is None or cached[0] != key:
                cached = (key, read_excel_sheet(content, sheet))
                st.session_state["spreadsheet_analysis_frame"] = cached
            frame = cached[1]
            with st.expander("Filters", expanded=True):
                filters = _filter_controls(frame, key)
            filtered = filter_rows(frame, filters)
            metrics = st.columns(3)
            metrics[0].metric("Source rows", f"{len(frame):,}")
            metrics[1].metric("Filtered rows", f"{len(filtered):,}")
            metrics[2].metric("Columns", f"{len(frame.columns):,}")
            if filtered.empty:
                st.warning("No rows match these filters. Widen or remove a filter.")
                return
            kinds = {column: column_kind(filtered[column]) for column in filtered.columns}
            numeric = [column for column, kind in kinds.items() if kind in {"whole number", "decimal number"}]
            date_columns = [column for column, kind in kinds.items() if kind == "date"]
            suggestions = auto_charts(frame)
            suggested_dimension = next((spec.dimension for spec in suggestions if spec.kind in {"count", "group"}), None)
            suggested_measure = next((spec.measure for spec in suggestions if spec.measure is not None), None)
            options = ["Grouped summary"]
            if date_columns:
                options.append("Time trend")
            if numeric:
                options.append("Distribution")
            if len(numeric) > 1:
                options.append("Relationship")
            analysis = st.selectbox("Analysis", options, key=f"{key}_analysis")
            dimension, measure, aggregate, period = None, None, "count", "month"
            if analysis in {"Grouped summary", "Time trend"}:
                choices = date_columns if analysis == "Time trend" else list(filtered.columns)
                dimension = st.selectbox(
                    "Group by", choices, index=choices.index(suggested_dimension) if suggested_dimension in choices else 0,
                    key=f"{key}_{analysis}_dimension",
                )
                calculation = st.selectbox(
                    "Calculation", list(AGGREGATES) if numeric else ["Row count"], key=f"{key}_calculation",
                )
                aggregate = AGGREGATES[calculation]
                if analysis == "Time trend":
                    period = st.selectbox("Period", ["month", "year"], key=f"{key}_period")
            if analysis == "Relationship":
                dimension = st.selectbox("X-axis", numeric, key=f"{key}_x")
            if analysis in {"Distribution", "Relationship"} or aggregate != "count":
                choices = [column for column in numeric if analysis != "Relationship" or column != dimension]
                measure = st.selectbox(
                    "Measure / Y-axis", choices, index=choices.index(suggested_measure) if suggested_measure in choices else 0,
                    key=f"{key}_{analysis}_measure",
                )
            chart_type = st.selectbox("Chart type", CHART_OPTIONS[analysis], key=f"{key}_{analysis}_chart")
            spec = analysis_spec(filtered, analysis, dimension=dimension, measure=measure,
                                 aggregate=aggregate, chart_type=chart_type, period=period)
            chart = build_chart(filtered, spec)
            if len(chart.table) > MAX_CHART_POINTS:
                raise ValueError(f"Filter to at most {MAX_CHART_POINTS} chart points. No points were sampled.")
            if chart_type == "Auto":
                reason = "a small category share" if spec.chart_type == "pie" else (
                    "an ordered time trend" if spec.chart_type == "line" else "comparison across groups"
                )
                st.caption(f"Auto selected {spec.chart_type} for {reason}.")
            chart_tab, table_tab, data_tab = st.tabs(["Chart", "Supporting table", "Filtered data"])
            with chart_tab:
                st.plotly_chart(chart.figure, width="stretch", key=f"{key}_plot")
                st.write(chart.explanation)
                if chart.note:
                    st.warning(chart.note)
            with table_tab:
                st.dataframe(chart.table, hide_index=True, width="stretch")
            with data_tab:
                st.dataframe(filtered, hide_index=True, width="stretch", height=350)
            if st.button("Prepare Excel download", key=f"{key}_export"):
                notes = [
                    ("Source", source), ("Source rows", str(len(frame))), ("Filtered rows", str(len(filtered))),
                    ("Analysis", analysis), ("Calculation", spec.aggregate), ("Chart", spec.chart_type),
                    ("Grouping / X-axis", dimension or ""), ("Measure / Y-axis", measure or ""),
                    ("Period", period if analysis == "Time trend" else ""),
                    ("Filters", "; ".join(str(rule) for rule in filters) or "None"),
                ]
                download = export_analysis(filtered, chart, notes)
                st.download_button("Download analysis with chart", download, file_name="spreadsheet_analysis.xlsx",
                                   mime=EXCEL_MIME, key=f"{key}_download")
                st.caption("Includes filtered source rows, exact supporting numbers, analysis settings and a native editable chart.")
        except ValueError as error:
            logger.warning("Spreadsheet analysis failed: %s", type(error).__name__)
            st.warning(str(error))
