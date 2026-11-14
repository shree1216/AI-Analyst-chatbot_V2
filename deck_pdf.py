"""PDF export for dashboard decks, matching the PowerPoint export's content and theme."""

from __future__ import annotations

from datetime import datetime
from html import escape
from io import BytesIO

import pandas as pd
from reportlab.graphics.charts.barcharts import HorizontalBarChart, VerticalBarChart
from reportlab.graphics.charts.linecharts import HorizontalLineChart
from reportlab.graphics.charts.piecharts import Pie
from reportlab.graphics.shapes import Drawing
from reportlab.lib import colors
from reportlab.lib.pagesizes import landscape, letter
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import inch
from reportlab.platypus import KeepTogether, Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

from report_themes import ReportTheme, get_theme
from slides_export import MAX_TABLE_ROWS, Deck, SlideChart, _number

MAX_PDF_POINTS = 15


def _drawing(chart: SlideChart, theme: ReportTheme, width: float) -> Drawing | None:
    table = chart.table
    if table is None or table.shape[1] < 2 or chart.chart_type == "scatter":
        return None
    rows = [(str(label)[:24], _number(value)) for label, value in table.iloc[:, :2].itertuples(index=False)]
    rows = [(label, value) for label, value in rows if value is not None][:MAX_PDF_POINTS]
    if not rows:
        return None
    labels, values = [label for label, _ in rows], [value for _, value in rows]
    drawing = Drawing(width, 220)
    if chart.chart_type in {"pie", "doughnut", "donut"} and min(values) >= 0 and sum(values) > 0:
        pie = Pie()
        pie.x, pie.y, pie.width, pie.height = width / 2 - 90, 15, 180, 180
        pie.data = values
        pie.labels = [f"{label}: {value / sum(values):.0%}" for label, value in rows]
        pie.sideLabels = True
        pie.slices.strokeColor = colors.white
        pie.slices.fontSize = 8
        if chart.chart_type != "pie":
            pie.innerRadiusFraction = 0.45
        for index in range(len(values)):
            pie.slices[index].fillColor = colors.HexColor(theme.palette[index % len(theme.palette)])
        drawing.add(pie)
        return drawing
    horizontal = chart.chart_type == "bar"
    if chart.chart_type in {"line", "area"}:
        graphic = HorizontalLineChart()
        graphic.lines[0].strokeColor = colors.HexColor(theme.primary)
        graphic.lines[0].strokeWidth = 2
    else:
        graphic = HorizontalBarChart() if horizontal else VerticalBarChart()
        graphic.bars[0].fillColor = colors.HexColor(theme.primary)
        graphic.bars[0].strokeColor = colors.HexColor(theme.primary)
    graphic.x, graphic.y = (140, 20) if horizontal else (60, 50)
    graphic.width, graphic.height = (width - 170, 185) if horizontal else (width - 80, 155)
    graphic.data = [list(reversed(values)) if horizontal else values]
    graphic.categoryAxis.categoryNames = list(reversed(labels)) if horizontal else labels
    graphic.categoryAxis.labels.fontSize = 7
    if not horizontal:
        graphic.categoryAxis.labels.angle = 30
        graphic.categoryAxis.labels.boxAnchor = "ne"
    graphic.valueAxis.labels.fontSize = 7
    low = min(values)
    graphic.valueAxis.valueMin = min(0, low * 1.1)
    graphic.valueAxis.valueMax = max(max(values) * 1.15, 1)
    graphic.valueAxis.labelTextFormat = lambda value: f"{value:,.0f}"
    drawing.add(graphic)
    return drawing


def _table(frame: pd.DataFrame, theme: ReportTheme, width: float) -> Table:
    shown = frame.head(MAX_TABLE_ROWS)
    cell = ParagraphStyle("DeckCell", fontName="Helvetica", fontSize=8, leading=10)
    head = ParagraphStyle("DeckHead", parent=cell, fontName="Helvetica-Bold", textColor=colors.white)
    rows = [[Paragraph(escape(str(name)), head) for name in shown.columns]]
    for values in shown.itertuples(index=False):
        rendered = []
        for value in values:
            number = None if isinstance(value, str) else _number(value)
            if number is not None:
                text = f"{number:,.2f}".rstrip("0").rstrip(".")
            elif value is None or (not isinstance(value, str) and pd.isna(value)):
                text = ""
            else:
                text = str(value)
            rendered.append(Paragraph(escape(text), cell))
        rows.append(rendered)
    table = Table(rows, colWidths=[width / max(1, len(shown.columns))] * len(shown.columns), repeatRows=1, hAlign="LEFT")
    table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor(theme.primary)),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor(theme.band)]),
        ("LINEBELOW", (0, 1), (-1, -1), 0.25, colors.HexColor(theme.line)),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
    ]))
    return table


def build_deck_pdf(deck: Deck, theme_key: str | None = None) -> bytes:
    theme = get_theme(theme_key)
    output = BytesIO()
    document = SimpleDocTemplate(
        output, pagesize=landscape(letter), leftMargin=0.6 * inch, rightMargin=0.6 * inch,
        topMargin=0.6 * inch, bottomMargin=0.6 * inch, title=deck.title, author="Workforce Compass",
    )
    styles = getSampleStyleSheet()
    title = ParagraphStyle("DeckTitle", parent=styles["Title"], alignment=0, textColor=colors.HexColor(theme.ink))
    heading = ParagraphStyle("DeckHeading", parent=styles["Heading2"], textColor=colors.HexColor(theme.primary))
    body = ParagraphStyle("DeckBody", parent=styles["BodyText"], fontSize=9.5, leading=13,
                          textColor=colors.HexColor(theme.ink))
    muted = ParagraphStyle("DeckMuted", parent=body, fontSize=8, textColor=colors.HexColor(theme.muted))
    story = [
        Paragraph("WORKFORCE COMPASS", ParagraphStyle("Brand", parent=muted, fontName="Helvetica-Bold",
                                                      textColor=colors.HexColor(theme.primary))),
        Paragraph(escape(deck.title), title),
        Paragraph(escape(deck.subtitle or f"Generated {datetime.now():%d %B %Y}"), muted),
    ]
    if deck.kpis:
        label = ParagraphStyle("KpiLabel", parent=muted, fontName="Helvetica-Bold")
        value = ParagraphStyle("KpiValue", parent=body, fontName="Helvetica-Bold", fontSize=16, leading=20,
                               textColor=colors.HexColor(theme.primary))
        cells = [[Paragraph(escape(item[0].upper()), label), Paragraph(escape(item[1]), value),
                  Paragraph(escape(item[2]), muted)] for item in deck.kpis]
        kpis = Table([cells], colWidths=[document.width / len(cells)] * len(cells))
        kpis.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor(theme.band)),
            ("BOX", (0, 0), (-1, -1), 0.5, colors.HexColor(theme.line)),
            ("INNERGRID", (0, 0), (-1, -1), 0.5, colors.white),
            ("TOPPADDING", (0, 0), (-1, -1), 10), ("BOTTOMPADDING", (0, 0), (-1, -1), 10),
        ]))
        story.extend([Spacer(1, 12), kpis])
    for chart in deck.charts:
        block = [Paragraph(escape(chart.title), heading)]
        drawing = _drawing(chart, theme, document.width)
        if drawing is not None:
            block.append(drawing)
        if chart.explanation:
            block.append(Paragraph(escape(chart.explanation), body))
        story.extend([Spacer(1, 8), KeepTogether(block), Spacer(1, 4), _table(chart.table, theme, document.width)])
    for title_text, items in deck.sections:
        if items:
            story.append(Paragraph(escape(title_text), heading))
            story.extend(Paragraph(f"&#8226; {escape(item)}", body) for item in items)
    for title_text, frame in deck.tables:
        if frame is not None and not frame.empty:
            story.extend([Paragraph(escape(title_text), heading), _table(frame, theme, document.width)])

    def footer(canvas, doc) -> None:
        canvas.saveState()
        canvas.setFont("Helvetica", 8)
        canvas.setFillColor(colors.HexColor(theme.muted))
        canvas.drawString(doc.leftMargin, 0.35 * inch, "Workforce Compass")
        canvas.drawRightString(doc.pagesize[0] - doc.rightMargin, 0.35 * inch, f"Page {doc.page}")
        canvas.restoreState()

    document.build(story, onFirstPage=footer, onLaterPages=footer)
    return output.getvalue()
