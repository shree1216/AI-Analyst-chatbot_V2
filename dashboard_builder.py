"""Headline KPI cards for Bricks-style dashboards built from uploaded or queried data."""

from __future__ import annotations

from dataclasses import dataclass
import re

import pandas as pd

from data_explorer import (
    AVERAGE_FIRST, _default_noun, _fmt, _profile, _usable_dimension, _usable_measure, categories, dates, numbers,
)

MAX_KPIS = 4
_MONEY = re.compile(r"salary|amount|sales|revenue|price|cost|pay|income|profit|value|total|budget|spend", re.I)


@dataclass(frozen=True)
class Kpi:
    label: str
    value: str
    caption: str = ""


def dashboard_kpis(frame: pd.DataFrame) -> list[Kpi]:
    """Pick up to four headline numbers: record count, key measures, top group and date span."""
    if frame.empty:
        return [Kpi("Records", "0")]
    kinds = _profile(frame)
    columns = [str(column) for column in frame.columns]
    noun = _default_noun(frame)
    kpis = [Kpi(f"Total {noun}", f"{len(frame):,}", f"{len(columns)} columns")]
    measures = sorted(
        (column for column in columns if _usable_measure(column, kinds[column], frame)),
        key=lambda column: (not _MONEY.search(column), -numbers(frame[column]).notna().sum()),
    )
    for measure in measures[:2]:
        values = numbers(frame[measure]).dropna()
        if values.empty:
            continue
        if AVERAGE_FIRST.search(measure):
            kpis.append(Kpi(f"Average {measure}", _fmt(values.mean()), f"Median {_fmt(values.median())}"))
        else:
            kpis.append(Kpi(f"Total {measure}", _fmt(values.sum()), f"Average {_fmt(values.mean())}"))
    dimension = next((column for column in columns if _usable_dimension(column, kinds[column], frame)), None)
    if dimension is not None and len(kpis) < MAX_KPIS:
        counts = categories(frame[dimension]).dropna().value_counts()
        if not counts.empty:
            share = counts.iloc[0] / counts.sum() * 100
            kpis.append(Kpi(f"Largest {dimension}", str(counts.index[0]), f"{share:.0f}% of {noun}"))
    date_column = next((column for column in columns if kinds[column] == "date"), None)
    if date_column is not None and len(kpis) < MAX_KPIS:
        when = dates(frame[date_column]).dropna()
        if not when.empty:
            kpis.append(Kpi(f"{date_column} range", f"{when.min():%b %Y} – {when.max():%b %Y}"))
    return kpis[:MAX_KPIS]
