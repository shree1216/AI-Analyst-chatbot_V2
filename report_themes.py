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


# Ocean teal, midnight navy and copper on cool neutrals.
CORPORATE = ReportTheme(
    "corporate", "Corporate", "#126E75", "#B57747", "#142D3B", "#526877", "#D8E3E8", "#FFFFFF", "#EDF3F5",
    ("#126E75", "#B57747", "#183E53", "#4E9897", "#687F91", "#D6A679", "#0D4F56", "#A7BDC7"),
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
