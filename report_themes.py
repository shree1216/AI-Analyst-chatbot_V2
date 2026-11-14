"""Visual themes shared by in-app dashboards, PDF reports and PowerPoint exports."""

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


THEMES: dict[str, ReportTheme] = {
    theme.key: theme for theme in (
        ReportTheme(
            "navy", "Navy Corporate", "#1E3A8A", "#3B82F6", "#0F172A", "#64748B", "#E2E8F0", "#FFFFFF", "#F1F5F9",
            ("#1E3A8A", "#3B82F6", "#0EA5E9", "#64748B", "#F59E0B", "#10B981", "#8B5CF6", "#EF4444"),
        ),
        ReportTheme(
            "slate", "Slate Minimal", "#334155", "#0EA5E9", "#0F172A", "#64748B", "#E2E8F0", "#FFFFFF", "#F8FAFC",
            ("#334155", "#0EA5E9", "#94A3B8", "#475569", "#38BDF8", "#CBD5E1", "#1E293B", "#7DD3FC"),
        ),
        ReportTheme(
            "emerald", "Emerald Finance", "#065F46", "#10B981", "#052E26", "#4B6B61", "#D1E7DE", "#FFFFFF", "#ECF7F2",
            ("#065F46", "#10B981", "#F59E0B", "#0EA5E9", "#6B7280", "#84CC16", "#14B8A6", "#DC2626"),
        ),
        ReportTheme(
            "midnight", "Midnight Executive", "#111827", "#6366F1", "#111827", "#6B7280", "#E5E7EB", "#FFFFFF", "#F3F4F6",
            ("#6366F1", "#111827", "#22D3EE", "#F472B6", "#A3A3A3", "#FBBF24", "#34D399", "#F87171"),
        ),
    )
}
DEFAULT_THEME = "navy"


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
