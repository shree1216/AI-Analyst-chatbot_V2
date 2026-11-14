"""The single corporate palette shared by in-app dashboards, PDF reports and PowerPoint exports."""

from __future__ import annotations

from dataclasses import dataclass


@dataclass(frozen=True)
class ReportTheme:
    key: str
    name: str
    primary: str
    accent: str
    ink: str
    muted: str
    line: str
    surface: str
    band: str
    palette: tuple[str, ...]


# Charcoal and wine with ochre/bronze accents; deliberately free of blue and green.
CORPORATE = ReportTheme(
    "corporate", "Corporate", "#7B2D43", "#C08A3E", "#1F1D1B", "#6B6560", "#E7E2DD", "#FFFFFF", "#F7F3EF",
    ("#7B2D43", "#C08A3E", "#4A4542", "#B5838D", "#8C6D4F", "#D4A373", "#5E3A48", "#A39C95"),
)
THEMES: dict[str, ReportTheme] = {CORPORATE.key: CORPORATE}
DEFAULT_THEME = CORPORATE.key


def get_theme(key: str | None) -> ReportTheme:
    return THEMES.get(key or DEFAULT_THEME, THEMES[DEFAULT_THEME])


def style_figure(figure, theme: ReportTheme):
    """Apply a theme's palette and typography to a Plotly figure in place."""
    figure.update_layout(
        colorway=list(theme.palette), piecolorway=list(theme.palette),
        paper_bgcolor=theme.surface, plot_bgcolor=theme.surface,
        font={"color": theme.ink, "family": "Inter, Segoe UI, sans-serif"},
        title_font={"color": theme.ink, "size": 15},
    )
    figure.update_xaxes(gridcolor=theme.line, linecolor=theme.line)
    figure.update_yaxes(gridcolor=theme.line, linecolor=theme.line)
    for trace in figure.data:
        if trace.type == "pie":
            trace.marker.colors = None
        elif trace.type in {"bar", "scatter"}:
            trace.marker.color = theme.primary
            if trace.type == "scatter" and getattr(trace, "line", None) is not None:
                trace.line.color = theme.primary
    return figure
