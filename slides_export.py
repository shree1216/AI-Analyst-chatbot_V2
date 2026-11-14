"""PowerPoint export for dashboards and reports, using native (editable) PowerPoint charts."""

from __future__ import annotations

from dataclasses import dataclass, field
from datetime import datetime
from io import BytesIO
import math

import pandas as pd
from pptx import Presentation
from pptx.chart.data import CategoryChartData, XyChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN
from pptx.util import Emu, Inches, Pt

from report_themes import ReportTheme, get_theme

PPTX_MIME = "application/vnd.openxmlformats-officedocument.presentationml.presentation"
MAX_CHART_POINTS = 40
MAX_TABLE_ROWS = 15
_CHART_TYPES = {
    "bar": XL_CHART_TYPE.BAR_CLUSTERED, "column": XL_CHART_TYPE.COLUMN_CLUSTERED,
    "histogram": XL_CHART_TYPE.COLUMN_CLUSTERED, "line": XL_CHART_TYPE.LINE_MARKERS,
    "area": XL_CHART_TYPE.AREA, "pie": XL_CHART_TYPE.PIE, "doughnut": XL_CHART_TYPE.DOUGHNUT,
    "donut": XL_CHART_TYPE.DOUGHNUT, "scatter": XL_CHART_TYPE.XY_SCATTER,
}


@dataclass
class SlideChart:
    title: str
    chart_type: str
    table: pd.DataFrame
    explanation: str = ""


@dataclass
class Deck:
    title: str
    subtitle: str = ""
    kpis: list[tuple[str, str, str]] = field(default_factory=list)
    charts: list[SlideChart] = field(default_factory=list)
    sections: list[tuple[str, list[str]]] = field(default_factory=list)
    tables: list[tuple[str, pd.DataFrame]] = field(default_factory=list)


def _rgb(color: str) -> RGBColor:
    return RGBColor.from_string(color.lstrip("#").upper())


def _text(slide, left, top, width, height, text: str, size: int, color: str, bold: bool = False,
          align=PP_ALIGN.LEFT) -> None:
    frame = slide.shapes.add_textbox(left, top, width, height).text_frame
    frame.word_wrap = True
    paragraph = frame.paragraphs[0]
    paragraph.alignment = align
    run = paragraph.add_run()
    run.text = text
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = _rgb(color)
    run.font.name = "Segoe UI"


def _header(slide, prs, title: str, theme: ReportTheme) -> None:
    band = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, Inches(0.12))
    band.fill.solid()
    band.fill.fore_color.rgb = _rgb(theme.primary)
    band.line.fill.background()
    _text(slide, Inches(0.5), Inches(0.3), prs.slide_width - Inches(1), Inches(0.7), title, 24, theme.ink, True)


def _footer(slide, prs, theme: ReportTheme, number: int) -> None:
    _text(slide, Inches(0.5), prs.slide_height - Inches(0.45), Inches(6), Inches(0.3),
          "Workforce Compass", 9, theme.muted)
    _text(slide, prs.slide_width - Inches(1.5), prs.slide_height - Inches(0.45), Inches(1), Inches(0.3),
          str(number), 9, theme.muted, align=PP_ALIGN.RIGHT)


def _number(value) -> float | None:
    try:
        number = float(value)
    except (TypeError, ValueError):
        return None
    return None if math.isnan(number) or math.isinf(number) else number


def _add_chart(slide, chart: SlideChart, left, top, width, height, theme: ReportTheme) -> bool:
    table = chart.table
    if table is None or table.shape[1] < 2 or table.empty:
        return False
    label_column, value_column = table.columns[0], table.columns[1]
    rows = [(row[0], _number(row[1])) for row in table[[label_column, value_column]].itertuples(index=False)]
    rows = [(label, value) for label, value in rows if value is not None][:MAX_CHART_POINTS]
    if not rows:
        return False
    chart_type = _CHART_TYPES.get(chart.chart_type, XL_CHART_TYPE.COLUMN_CLUSTERED)
    if chart_type == XL_CHART_TYPE.XY_SCATTER:
        points = [(_number(x), y) for x, y in rows if _number(x) is not None]
        if not points:
            return False
        data = XyChartData()
        series = data.add_series(str(value_column))
        for x, y in points:
            series.add_data_point(x, y)
    else:
        if chart_type in {XL_CHART_TYPE.PIE, XL_CHART_TYPE.DOUGHNUT} and any(value < 0 for _, value in rows):
            chart_type = XL_CHART_TYPE.BAR_CLUSTERED
        data = CategoryChartData()
        data.categories = [str(label) for label, _ in rows]
        data.add_series(str(value_column), [value for _, value in rows])
    graphic = slide.shapes.add_chart(chart_type, left, top, width, height, data).chart
    graphic.font.size = Pt(10)
    graphic.font.color.rgb = _rgb(theme.ink)
    plot = graphic.plots[0]
    if chart_type in {XL_CHART_TYPE.PIE, XL_CHART_TYPE.DOUGHNUT}:
        graphic.has_legend = True
        graphic.legend.position = XL_LEGEND_POSITION.RIGHT
        graphic.legend.include_in_layout = False
        for index, point in enumerate(plot.series[0].points):
            point.format.fill.solid()
            point.format.fill.fore_color.rgb = _rgb(theme.palette[index % len(theme.palette)])
        plot.has_data_labels = True
        plot.data_labels.number_format = "0%"
        plot.data_labels.show_percentage = True
        plot.data_labels.show_value = False
    else:
        graphic.has_legend = False
        series = plot.series[0]
        if chart_type in {XL_CHART_TYPE.LINE_MARKERS, XL_CHART_TYPE.XY_SCATTER}:
            series.format.line.color.rgb = _rgb(theme.primary)
            series.marker.format.fill.solid()
            series.marker.format.fill.fore_color.rgb = _rgb(theme.primary)
            if chart_type == XL_CHART_TYPE.XY_SCATTER:
                series.format.line.fill.background()
        else:
            series.format.fill.solid()
            series.format.fill.fore_color.rgb = _rgb(theme.primary)
        if chart_type == XL_CHART_TYPE.BAR_CLUSTERED:
            graphic.category_axis.reverse_order = True
        if chart.chart_type == "histogram":
            plot.gap_width = 10
    return True


def _add_table(slide, frame: pd.DataFrame, left, top, width, theme: ReportTheme) -> None:
    shown = frame.head(MAX_TABLE_ROWS)
    rows, cols = len(shown) + 1, max(1, len(shown.columns))
    shape = slide.shapes.add_table(rows, cols, left, top, width, Inches(0.3) * rows)
    table = shape.table
    for col, name in enumerate(shown.columns):
        cell = table.cell(0, col)
        cell.text = str(name)
        cell.fill.solid()
        cell.fill.fore_color.rgb = _rgb(theme.primary)
        run = cell.text_frame.paragraphs[0].runs[0]
        run.font.size, run.font.bold, run.font.color.rgb = Pt(10), True, _rgb("#FFFFFF")
    for row_index, row in enumerate(shown.itertuples(index=False), start=1):
        for col, value in enumerate(row):
            cell = table.cell(row_index, col)
            number = _number(value) if not isinstance(value, str) else None
            if pd.isna(value):
                cell.text = ""
            elif number is not None:
                cell.text = f"{number:,.2f}".rstrip("0").rstrip(".")
            else:
                cell.text = str(value)
            cell.fill.solid()
            cell.fill.fore_color.rgb = _rgb(theme.band if row_index % 2 == 0 else theme.surface)
            for paragraph in cell.text_frame.paragraphs:
                for run in paragraph.runs:
                    run.font.size, run.font.color.rgb = Pt(9), _rgb(theme.ink)


def build_pptx(deck: Deck, theme_key: str | None = None) -> bytes:
    """Render a 16:9 deck: title, KPI overview, one slide per chart, text sections and tables."""
    theme = get_theme(theme_key)
    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
    blank = prs.slide_layouts[6]
    number = 1

    slide = prs.slides.add_slide(blank)
    background = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, prs.slide_height)
    background.fill.solid()
    background.fill.fore_color.rgb = _rgb(theme.primary)
    background.line.fill.background()
    _text(slide, Inches(0.8), Inches(2.4), Inches(11.5), Inches(1.4), deck.title, 40, "#FFFFFF", True)
    _text(slide, Inches(0.8), Inches(3.9), Inches(11.5), Inches(0.8),
          deck.subtitle or f"Generated {datetime.now():%d %B %Y}", 16, "#E7E2DD")
    _text(slide, Inches(0.8), Inches(6.6), Inches(8), Inches(0.4), "Workforce Compass", 12, "#E7E2DD", True)

    if deck.kpis:
        number += 1
        slide = prs.slides.add_slide(blank)
        _header(slide, prs, "Key metrics", theme)
        count = len(deck.kpis)
        gap = Inches(0.3)
        card_width = Emu(int((prs.slide_width - Inches(1) - gap * (count - 1)) / count))
        for index, (label, value, caption) in enumerate(deck.kpis):
            left = Inches(0.5) + (card_width + gap) * index
            card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, Inches(2.2), card_width, Inches(2.2))
            card.fill.solid()
            card.fill.fore_color.rgb = _rgb(theme.band)
            card.line.color.rgb = _rgb(theme.line)
            card.adjustments[0] = 0.08
            _text(slide, left + Inches(0.25), Inches(2.4), card_width - Inches(0.5), Inches(0.5), label.upper(), 11, theme.muted, True)
            _text(slide, left + Inches(0.25), Inches(2.95), card_width - Inches(0.5), Inches(0.9), value, 28, theme.primary, True)
            if caption:
                _text(slide, left + Inches(0.25), Inches(3.85), card_width - Inches(0.5), Inches(0.4), caption, 11, theme.muted)
        _footer(slide, prs, theme, number)

    for chart in deck.charts:
        slide = prs.slides.add_slide(blank)
        _header(slide, prs, chart.title, theme)
        width = prs.slide_width - Inches(1) if not chart.explanation else Inches(8.3)
        if not _add_chart(slide, chart, Inches(0.5), Inches(1.2), width, Inches(5.6), theme):
            _add_table(slide, chart.table, Inches(0.5), Inches(1.3), width, theme)
        if chart.explanation:
            _text(slide, Inches(9.1), Inches(1.4), Inches(3.8), Inches(5.2), chart.explanation, 13, theme.ink)
        number += 1
        _footer(slide, prs, theme, number)

    for heading, items in deck.sections:
        if not items:
            continue
        slide = prs.slides.add_slide(blank)
        _header(slide, prs, heading, theme)
        frame = slide.shapes.add_textbox(Inches(0.6), Inches(1.3), prs.slide_width - Inches(1.2), Inches(5.6)).text_frame
        frame.word_wrap = True
        for index, item in enumerate(items):
            paragraph = frame.paragraphs[0] if index == 0 else frame.add_paragraph()
            paragraph.space_after = Pt(10)
            run = paragraph.add_run()
            run.text = f"\u2022  {item}"
            run.font.size, run.font.color.rgb, run.font.name = Pt(16), _rgb(theme.ink), "Segoe UI"
        number += 1
        _footer(slide, prs, theme, number)

    for heading, frame in deck.tables:
        if frame is None or frame.empty:
            continue
        slide = prs.slides.add_slide(blank)
        _header(slide, prs, heading, theme)
        _add_table(slide, frame, Inches(0.5), Inches(1.3), prs.slide_width - Inches(1), theme)
        if len(frame) > MAX_TABLE_ROWS:
            _text(slide, Inches(0.5), prs.slide_height - Inches(0.9), Inches(10), Inches(0.3),
                  f"First {MAX_TABLE_ROWS} of {len(frame):,} rows shown.", 10, theme.muted)
        number += 1
        _footer(slide, prs, theme, number)

    output = BytesIO()
    prs.save(output)
    return output.getvalue()
