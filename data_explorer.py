"""Upload conversion, data overview and question-driven charts for uploaded tabular data.

Everything here is deterministic (no AI): questions are matched to real column names, and
values are cleaned only for analysis. The uploaded data itself is never rewritten.
"""
import csv
from dataclasses import dataclass, field
from datetime import date, datetime
from io import BytesIO, StringIO
import math
from pathlib import Path
import re
from typing import Any

import numpy as np
import pandas as pd
import plotly.express as px
import pyarrow as pa
import pyarrow.parquet as pq
from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE

from excel_support import (
    MAX_EXCEL_BYTES, MAX_EXCEL_CELLS, MAX_EXCEL_COLUMNS, MAX_EXCEL_EXPANDED_BYTES, MAX_EXCEL_ROWS,
    MAX_UPLOAD_MB, generate_excel_download,
)


MISSING_TOKENS = {"", "na", "n/a", "nan", "null", "none", "nil", "-", "--", "?", "not available", "missing", "#n/a"}
DATE_FORMATS = (
    "%Y-%m-%d", "%Y/%m/%d", "%d-%m-%Y", "%d/%m/%Y", "%d.%m.%Y", "%m/%d/%Y", "%m-%d-%Y",
    "%d-%b-%y", "%d-%b-%Y", "%d %b %Y", "%d %B %Y", "%b %d, %Y", "%B %d, %Y", "%b %d %Y", "%B %d %Y",
    "%d-%m-%y", "%d/%m/%y", "%Y-%m-%d %H:%M:%S", "%Y-%m-%dT%H:%M:%S", "%d-%m-%Y %H:%M", "%d/%m/%Y %H:%M",
)
ID_NAME = re.compile(r"(?:^|[\s_-])(?:id|ids|no|number|code|phone|mobile|tel|zip|pin|postcode|postal)(?:$|[\s_-])|^(?:id|phone|email|e-mail)", re.I)
TEXT_ONLY_NAME = re.compile(r"\b(?:name|email|e-mail|address|description|comment|notes?|url|link)\b", re.I)
AVERAGE_FIRST = re.compile(r"salary|salaries|pay|wage|rating|score|price|rate|age|percent|grade|tenure|experience", re.I)
COLORS = ["#0F766E", "#2563EB", "#D97706", "#DB2777", "#64748B", "#65A30D", "#7C3AED", "#DC2626"]
MAX_CATEGORIES = 25
MAX_AUTO_CHARTS = 4
MAX_SCATTER_POINTS = 500


# --------------------------------------------------------------------------- upload conversion

def _decode(content: bytes) -> str:
    for encoding in ("utf-8-sig", "cp1252", "latin-1"):
        try:
            return content.decode(encoding)
        except UnicodeDecodeError:
            continue
    raise ValueError("The CSV file could not be read as text.")


def _clean_headers(headers: list[str]) -> list[str]:
    used: set[str] = set()
    result = []
    for index, header in enumerate(headers, 1):
        name = re.sub(r"\s+", " ", str(header or "")).strip() or f"Column {index}"
        base, number = name, 2
        while name.casefold() in used:
            name = f"{base} ({number})"
            number += 1
        used.add(name.casefold())
        result.append(name)
    return result


_PLAIN_NUMBER = re.compile(r"[+-]?(?:\d+\.?\d*|\.\d+)(?:[eE][+-]?\d+)?")


def _numeric_column(values: pd.Series) -> pd.Series | None:
    present = values[values.str.strip() != ""].str.strip()
    if present.empty or not present.str.fullmatch(_PLAIN_NUMBER.pattern).all():
        return None
    # Keep codes such as 00123 as text so leading zeros survive.
    if present.str.fullmatch(r"[+-]?0\d+").any() or present.str.fullmatch(r"[+-]?\d{16,}").any():
        return None
    numbers = pd.to_numeric(values.str.strip().replace("", None), errors="coerce")
    if numbers.dropna().map(lambda value: float(value).is_integer() and abs(value) < 2**53).all():
        return numbers.map(lambda value: None if pd.isna(value) else int(value)).astype(object)
    return numbers.astype(object).where(numbers.notna(), None)


def _table_workbook(frame: pd.DataFrame, file_name: str) -> bytes:
    sheet = re.sub(r"[\[\]:*?/\\]", "_", Path(file_name).stem).strip().strip("'")[:31] or "Data"
    return generate_excel_download([(sheet, frame)])


def csv_to_workbook(content: bytes, file_name: str) -> bytes:
    """Convert an uploaded CSV into a one-sheet .xlsx so every workbook feature can use it."""
    if len(content) > MAX_EXCEL_BYTES:
        raise ValueError(f"Upload a CSV file no larger than {MAX_UPLOAD_MB} MB.")
    if not content.strip():
        raise ValueError("The CSV file is empty. Upload a CSV with a header row and data.")
    text = _decode(content)
    sample = text[:20000]
    try:
        delimiter = csv.Sniffer().sniff(sample, delimiters=",;\t|").delimiter
    except csv.Error:
        delimiter = ","
    try:
        reader = csv.reader(StringIO(text), delimiter=delimiter, strict=True)
        headers = next((row for row in reader if any(value.strip() for value in row)), None)
        if headers is None:
            raise ValueError("The CSV file needs a header row with column names.")
        if len(headers) > MAX_EXCEL_COLUMNS:
            raise ValueError(f"A CSV file can have at most {MAX_EXCEL_COLUMNS} columns.")
        rows = []
        for row in reader:
            if not any(value.strip() for value in row):
                continue
            if len(row) != len(headers):
                raise csv.Error("Row width does not match the header.")
            rows.append(row)
            if len(rows) > MAX_EXCEL_ROWS:
                raise ValueError(f"A CSV file can have at most {MAX_EXCEL_ROWS:,} data rows.")
            if len(rows) * len(headers) > MAX_EXCEL_CELLS:
                raise ValueError(f"A CSV file can have at most {MAX_EXCEL_CELLS:,} data cells. Reduce its rows or columns.")
    except csv.Error as error:
        raise ValueError(
            "Some CSV rows have a different number of values than the header. "
            "Check for unquoted commas inside values, save, and upload again."
        ) from error
    frame = pd.DataFrame(rows, columns=_clean_headers(headers), dtype=str)
    for column in frame.columns:
        converted = _numeric_column(frame[column])
        frame[column] = converted if converted is not None else frame[column].where(frame[column].str.strip() != "", None)
    return _table_workbook(frame, file_name)


def parquet_to_workbook(content: bytes, file_name: str) -> bytes:
    """Read bounded, flat Parquet data into the existing Excel working-copy flow."""
    if not content or len(content) > MAX_EXCEL_BYTES:
        raise ValueError(f"Upload a nonempty Parquet file no larger than {MAX_UPLOAD_MB} MB.")
    try:
        with pq.ParquetFile(BytesIO(content), arrow_extensions_enabled=False) as parquet:
            metadata, schema = parquet.metadata, parquet.schema_arrow
            if not 1 <= len(schema) <= MAX_EXCEL_COLUMNS:
                raise ValueError(f"A Parquet file needs 1 to {MAX_EXCEL_COLUMNS} columns.")
            if metadata.num_rows > MAX_EXCEL_ROWS:
                raise ValueError(f"A Parquet file can have at most {MAX_EXCEL_ROWS:,} data rows.")
            if metadata.num_rows * len(schema) > MAX_EXCEL_CELLS:
                raise ValueError(f"A Parquet file can have at most {MAX_EXCEL_CELLS:,} data cells.")
            if sum(metadata.row_group(index).total_byte_size for index in range(metadata.num_row_groups)) > MAX_EXCEL_EXPANDED_BYTES:
                raise ValueError(
                    f"The expanded Parquet data must be no larger than {MAX_EXCEL_EXPANDED_BYTES // 1_000_000:,} MB."
                )
            types = [field.type.value_type if pa.types.is_dictionary(field.type) else field.type for field in schema]
            for field, kind in zip(schema, types):
                if not (
                    pa.types.is_null(kind) or pa.types.is_boolean(kind) or pa.types.is_integer(kind)
                    or pa.types.is_floating(kind) or pa.types.is_string(kind) or pa.types.is_large_string(kind)
                    or pa.types.is_decimal(kind) or pa.types.is_date(kind)
                    or (pa.types.is_timestamp(kind) and kind.tz is None)
                ):
                    raise ValueError(
                        f"Parquet column '{field.name}' has unsupported type {kind}. "
                        "Flatten nested columns and convert binary, timezone-aware or other unsupported values to text."
                    )
            table = parquet.read()
        headers = _clean_headers(schema.names)
        if any(ILLEGAL_CHARACTERS_RE.search(name) or len(name) > 32_767 for name in headers):
            raise ValueError("The Parquet headers contain control characters or exceed the Excel cell length limit.")
        columns = {}
        for index, (name, kind) in enumerate(zip(headers, types)):
            values = table.column(index).to_pylist()
            if pa.types.is_integer(kind):
                values = [str(value) if value is not None and abs(value) >= 10**15 else value for value in values]
            elif pa.types.is_decimal(kind):
                values = [None if value is None else str(value) for value in values]
            elif pa.types.is_string(kind) or pa.types.is_large_string(kind):
                if any(value is not None and len(value) > 32_767 for value in values):
                    raise ValueError(f"Parquet column '{name}' contains text longer than the 32,767-character Excel cell limit.")
                if any(value is not None and ILLEGAL_CHARACTERS_RE.search(value) for value in values):
                    raise ValueError(f"Parquet column '{name}' contains control characters unsupported by Excel. Remove them and upload again.")
            elif pa.types.is_timestamp(kind):
                if any(
                    value is not None and (pd.Timestamp(value).nanosecond or pd.Timestamp(value).microsecond % 1000)
                    for value in values
                ):
                    raise ValueError(f"Parquet column '{name}' has sub-millisecond timestamps. Cast it to timestamp[ms] or text.")
            columns[name] = pd.Series(values, dtype=object)
        return _table_workbook(pd.DataFrame(columns), file_name)
    except (pa.ArrowException, OSError) as error:
        raise ValueError("Could not read this Parquet file. Upload a valid, unencrypted .parquet file.") from error


# --------------------------------------------------------------------------- cleaning helpers

def _is_missing(value: Any) -> bool:
    if value is None:
        return True
    if isinstance(value, float) and math.isnan(value):
        return True
    if value is pd.NaT or (not isinstance(value, (str, bytes)) and bool(pd.isna(value))):
        return True
    return isinstance(value, str) and value.strip().casefold() in MISSING_TOKENS


def missing_mask(series: pd.Series) -> pd.Series:
    return series.map(_is_missing).astype(bool)


def _to_number(value: Any) -> float | None:
    if _is_missing(value) or isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return float(value) if math.isfinite(value) else None
    if isinstance(value, str):
        text = re.sub(r"(?i)^(?:rs\.?|inr|usd)\s*|[\s,₹$€£]", "", value.strip())
        if _PLAIN_NUMBER.fullmatch(text):
            number = float(text)
            return number if math.isfinite(number) else None
    return None


def numbers(series: pd.Series) -> pd.Series:
    return pd.to_numeric(series.map(_to_number), errors="coerce")


def _to_date(value: Any) -> pd.Timestamp | None:
    if _is_missing(value):
        return None
    if isinstance(value, (datetime, date, pd.Timestamp)):
        return pd.Timestamp(value)
    if not isinstance(value, str):
        return None
    text = re.sub(r"\s+", " ", value.strip())
    for pattern in DATE_FORMATS:
        try:
            return pd.Timestamp(datetime.strptime(text, pattern))
        except ValueError:
            continue
    return None


def dates(series: pd.Series) -> pd.Series:
    return pd.to_datetime(series.map(_to_date), errors="coerce")


def _category_key(value: Any) -> str:
    return re.sub(r"[^\w]+", " ", str(value)).strip().casefold()


def categories(series: pd.Series) -> pd.Series:
    """Group spellings that differ only by capital letters, spaces or punctuation, labelled by the most common form."""
    present = series[~missing_mask(series)]
    labels: dict[str, str] = {}
    for key, group in present.map(lambda value: re.sub(r"\s+", " ", str(value)).strip()).groupby(present.map(_category_key)):
        label = group.value_counts().index[0]
        if label.islower() or (label.isupper() and len(label) > 3):
            label = label.title()
        labels[key] = label
    return series.map(lambda value: None if _is_missing(value) or not _category_key(value) else labels[_category_key(value)])


def possible_duplicates(labels: list[str], limit: int = 3) -> list[tuple[str, str]]:
    """Pairs like ('Prob', 'Probation') where one group looks like an abbreviation of another."""
    pairs = []
    for short in labels:
        stem = short.casefold().rstrip(".")
        for long in labels:
            if long != short and len(long) > len(stem) and long.casefold().startswith(stem) and len(stem) <= 5:
                pairs.append((short, long))
                break
        if len(pairs) >= limit:
            break
    return pairs


def column_kind(series: pd.Series) -> str:
    present = series[~missing_mask(series)]
    if present.empty:
        return "empty"
    if present.map(lambda value: isinstance(value, bool)).all() or present.map(
        lambda value: isinstance(value, str) and value.strip().casefold() in {"true", "false", "yes", "no"}
    ).all():
        return "true/false"
    numeric = numbers(present)
    if numeric.notna().mean() >= 0.9:
        return "whole number" if (numeric.dropna() % 1 == 0).all() else "decimal number"
    if dates(present).notna().mean() >= 0.8:
        return "date"
    return "text"


# --------------------------------------------------------------------------- overview

_OVERVIEW = re.compile(
    r"(?:please\s+)?(?:(?:can|could)\s+you\s+)?(?:open|show|display|view|load|preview|explore|describe|summari[sz]e)"
    r"(?:\s+me)?\s+(?:the\s+|my\s+|this\s+|uploaded\s+|that\s+)*"
    r"(?:file|data|dataset|data\s*set|csv|csv\s+file|uploaded\s+file|spreadsheet|table)(?:\s+(?:please|now))?",
    re.I,
)


def is_overview_request(question: str) -> bool:
    return bool(_OVERVIEW.fullmatch(question.strip().rstrip("?.!").strip()))


def _display(value: Any) -> str:
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    if isinstance(value, (datetime, pd.Timestamp)):
        stamp = pd.Timestamp(value)
        return stamp.date().isoformat() if stamp == stamp.normalize() else str(stamp)
    return str(value).strip()


def data_overview(frame: pd.DataFrame) -> pd.DataFrame:
    rows = []
    for column in frame.columns:
        series = frame[column]
        missing = int(missing_mask(series).sum())
        present = series[~missing_mask(series)]
        example = next((_display(value) for value in present if str(value).strip()), "")
        rows.append({
            "Column": str(column), "Data type": column_kind(series), "Missing values": missing,
            "Missing %": round(missing / len(frame) * 100, 1) if len(frame) else 0.0,
            "Unique values": int(present.map(_category_key).nunique()), "Example": example[:40],
        })
    return pd.DataFrame(rows)


def overview_answer(frame: pd.DataFrame, summary: pd.DataFrame, source: str) -> str:
    missing = summary[summary["Missing values"] > 0]
    worst = missing.sort_values("Missing values", ascending=False).head(3)
    text = (
        f"{source} has {len(frame):,} rows and {len(frame.columns):,} columns: "
        f"{', '.join(summary['Column'])}. "
    )
    if missing.empty:
        text += "No values are missing."
    else:
        text += (
            f"{int(missing['Missing values'].sum()):,} values are missing across {len(missing):,} column(s); "
            "the most are in " + ", ".join(
                f"{column} ({count:,})" for column, count in zip(worst["Column"], worst["Missing values"])
            ) + ". "
            "Blank cells and placeholders such as N/A or 'not available' count as missing."
        )
    return text


# --------------------------------------------------------------------------- chart planning

@dataclass
class ChartSpec:
    kind: str                     # group, count, time, histogram, scatter
    chart_type: str
    measure: str | None = None
    dimension: str | None = None
    aggregate: str = "count"      # count, sum, average, median, min, max
    period: str = "year"          # year or month for time charts
    noun: str = "rows"
    category_limit: int = MAX_CATEGORIES


@dataclass
class ChartResult:
    title: str
    chart_type: str
    table: pd.DataFrame
    explanation: str
    figure: Any = field(default=None, repr=False)
    note: str = ""


_AGGREGATE_WORDS = (
    (r"average|avg|mean", "average"), (r"total|sum(?:\s+of)?", "sum"), (r"median", "median"),
    (r"highest|maximum|max|largest", "max"), (r"lowest|minimum|min|smallest", "min"),
    (r"how\s+many|number\s+of|count(?:\s+of)?|headcount", "count"),
)
_COUNT_NOUNS = r"people|persons?|employees?|staff|workers?|members?|customers?|clients?|users?|students?|orders?|records?|rows?|entries|items?|products?|sales|transactions?|hires?"
_GENERIC_CHARTS = re.compile(
    r"(?:please\s+)?(?:(?:can|could)\s+you\s+)?(?:make|create|generate|build|draw|show|give|plot|produce|suggest)?\s*(?:me\s+)?"
    r"(?:some\s+|an?\s+|a\s+few\s+|few\s+|useful\s+|the\s+|all\s+|any\s+|more\s+|good\s+|basic\s+|quick\s+|auto(?:matic)?\s+)*"
    r"(?:charts?|graphs?|plots?|visuali[sz]ations?|visuals?|dashboard)"
    r"(?:\s+(?:for|from|of|on|about)\s+(?:the\s+|this\s+|my\s+)?(?:data|dataset|file|sheet|csv|table))?(?:\s+please)?",
    re.I,
)
_CHART_WORDS = re.compile(
    r"\b(?:charts?|graphs?|plots?|dashboard|visuali[sz]e|visuali[sz]ation|histogram|pie|trend|scatter|correlation|relationship)\b", re.I,
)


def chart_presentation_question(question: str) -> str:
    text = question.strip().rstrip("?.!").strip()
    text = re.sub(
        r"\s+(?:and|with)\s+(?:(?:show|include|add|give|provide)\s+)?(?:a\s+|the\s+)?"
        r"(?:dashboard\s+)?(?:supporting|support|summary|data)\s+tables?\s*$",
        "", text, flags=re.I,
    )
    if re.search(r"\bdashboard\b", text, re.I):
        text = re.sub(r"\s+with\s+(?:charts?|graphs?)(?:\s+and\s+tables?)?\s*$", "", text, flags=re.I)
        text = re.sub(r"(?<=dashboard)\s+(?:supporting|support|summary|data)\s+tables?\s*$", "", text, flags=re.I)
    return text


def is_generic_chart_request(question: str) -> bool:
    return bool(_GENERIC_CHARTS.fullmatch(chart_presentation_question(question)))


def is_chart_question(question: str) -> bool:
    return bool(
        is_generic_chart_request(question) or _CHART_WORDS.search(question)
        or (
            any(re.search(rf"\b(?:{pattern})\b", question, re.I) for pattern, _ in _AGGREGATE_WORDS)
            and re.search(r"\b(?:by|per|each|every|across|monthly|yearly|annually)\b", question, re.I)
        )
    )


def _words(text: str) -> list[str]:
    return re.findall(r"[a-z0-9]+", text.casefold())


def _stem(word: str) -> str:
    word = word.casefold()
    for suffix in ("ies", "es", "s", "ed", "ing"):
        if len(word) > len(suffix) + 2 and word.endswith(suffix):
            return word[: -len(suffix)] + ("y" if suffix == "ies" else "")
    return word


def _column_tokens(column: str) -> list[str]:
    return [_stem(word) for word in _words(re.sub(r"([a-z])([A-Z])", r"\1 \2", column))]


def _find_column(phrase: str, columns: list[str], allowed: set[str] | None = None) -> str | None:
    """Return the column whose name best matches the start of a phrase (e.g. 'department name' -> department)."""
    tokens = [_stem(word) for word in _words(phrase)]
    if not tokens:
        return None
    best: tuple[int, int, str] | None = None
    for column in columns:
        if allowed is not None and column not in allowed:
            continue
        names = _column_tokens(column)
        if not names:
            continue
        for length in range(min(len(tokens), 4), 0, -1):
            window = tokens[:length]
            exact = window == names
            partial = len(names) > 1 and window == names[: len(window)] and length >= 1 and len(window[0]) > 2
            joined = "".join(window) == "".join(names)
            stem_match = length == 1 and len(names) >= 1 and len(window[0]) >= 4 and (
                names[0].startswith(window[0][:4]) and window[0].startswith(names[0][:4])
            )
            if exact or joined or partial or stem_match:
                score = (3 if exact or joined else 1 if partial else 0, length)
                if best is None or score > best[:2]:
                    best = (*score, column)
                break
    if best is None or best[0] < 3:
        rest = re.sub(rf"^(?:(?:{_COUNT_NOUNS})\s+)+", "", phrase.strip().strip("\"'`"), flags=re.I)
        if rest != phrase.strip().strip("\"'`"):
            found = _find_column(rest, columns, allowed)
            if found is not None:
                return found
    return best[2] if best else None


def _profile(frame: pd.DataFrame) -> dict[str, str]:
    return {str(column): column_kind(frame[column]) for column in frame.columns}


def _usable_measure(column: str, kind: str, frame: pd.DataFrame) -> bool:
    return kind in {"whole number", "decimal number"} and not ID_NAME.search(column) and numbers(frame[column]).nunique() > 1


def _usable_dimension(column: str, kind: str, frame: pd.DataFrame) -> bool:
    if kind not in {"text", "true/false"} or TEXT_ONLY_NAME.search(column) or ID_NAME.search(column):
        return False
    groups = categories(frame[column]).nunique()
    return 2 <= groups <= MAX_CATEGORIES


def _requested_type(text: str) -> str | None:
    for pattern, kind in (
        (r"\b(?:scatter|relationship|correlation)\b", "scatter"),
        (r"\b(?:donut|doughnut)\b", "doughnut"), (r"\b(?:pie|share|proportion|percentage)\b", "pie"),
        (r"\barea\s+(?:chart|graph|plot)\b", "area"), (r"\b(?:line|trend)\b", "line"),
        (r"\b(?:histogram|distribution|spread)\b", "histogram"), (r"\b(?:column|vertical)\b", "column"),
        (r"\bbar\b|horizontal", "bar"),
    ):
        if re.search(pattern, text):
            return kind
    return None


def _default_noun(frame: pd.DataFrame) -> str:
    names = " ".join(str(column) for column in frame.columns).casefold()
    for pattern, noun in ((r"emp|employee|staff|salary|hire|join", "employees"), (r"customer|client", "customers"),
                          (r"order", "orders"), (r"student", "students"), (r"product|sku", "products")):
        if re.search(pattern, names):
            return noun
    return "rows"


def plan_charts(question: str, frame: pd.DataFrame) -> list[ChartSpec] | None:
    question = chart_presentation_question(question)
    specs = _plan_charts(question, frame)
    if specs is None or is_generic_chart_request(question):
        return specs
    # Do not turn filtered, multi-metric or instructional requests into an unfiltered chart.
    spec = specs[0]
    vocabulary = (
        "please can could you me make create generate build draw show give plot produce add insert "
        "a an the my this data file chart charts graph graphs plot plots "
        "bar column vertical horizontal line pie donut doughnut area scatter histogram distribution spread trend "
        "relationship correlation between vs versus against and "
        "average avg mean total sum median highest maximum max largest lowest minimum min smallest "
        "how many number of count headcount by per for each every across in within "
        "year yearly annual annually month monthly over time"
    )
    allowed = {_stem(word) for word in _words(vocabulary)}
    if spec.aggregate == "count":
        allowed.update(_stem(word) for word in _words(spec.noun))
        allowed.update({"people", "person", "employee", "staff", "worker", "row", "record"})
    if spec.kind == "time":
        allowed.update({"join", "hire", "start"})
    for column in (spec.measure, spec.dimension):
        if column is not None:
            allowed.update(_column_tokens(column))
    if any(_stem(word) not in allowed for word in _words(question)):
        return None
    aggregate_count = sum(bool(re.search(rf"\b(?:{pattern})\b", question, re.I))
                          for pattern, _ in _AGGREGATE_WORDS)
    if aggregate_count > 1:
        return None
    return specs


def _plan_charts(question: str, frame: pd.DataFrame) -> list[ChartSpec] | None:
    """Match a question to columns. Returns None when the question is not a chart/grouped question we can answer."""
    if frame is None or frame.empty:
        return None
    if is_generic_chart_request(question):
        return auto_charts(frame)
    text = question.casefold()
    columns = [str(column) for column in frame.columns]
    kinds = _profile(frame)
    measures = {column for column in columns if _usable_measure(column, kinds[column], frame)}
    numeric = {column for column in columns if kinds[column] in {"whole number", "decimal number"}}
    date_columns = [column for column in columns if kinds[column] == "date"]
    requested = _requested_type(text)
    noun_match = re.search(rf"\b({_COUNT_NOUNS})\b", text)
    noun = noun_match[1] if noun_match and noun_match[1] not in {"row", "rows", "record", "records", "sales"} else _default_noun(frame)

    if requested == "scatter":
        pair = re.search(
            r"\b(?:scatter(?:\s+(?:chart|plot|graph))?(?:\s+(?:of|for|showing))?|"
            r"(?:relationship|correlation)\s+between)\s+(?P<x>.+?)\s+"
            r"(?:vs\.?|versus|against|and)\s+(?P<y>.+)", text,
        )
        x = _find_column(pair["x"], columns, numeric) if pair else None
        y = _find_column(pair["y"], columns, numeric) if pair else None
        if x is None or y is None or x == y:
            return None
        return [ChartSpec("scatter", "scatter", y, x, "none", noun=noun)]

    aggregate, measure = None, None
    for pattern, name in _AGGREGATE_WORDS:
        match = re.search(rf"\b(?:{pattern})\b\s*(?:the\s+|of\s+)?(?P<rest>.*)", text)
        if match:
            aggregate = name
            if name != "count":
                measure = _find_column(match["rest"], columns, numeric)
            break
    if measure is None and aggregate != "count":
        chart_of = re.search(r"\b(?:of|for|showing)\s+(?:the\s+)?(?P<rest>.+)", text)
        if chart_of:
            measure = _find_column(chart_of["rest"], columns, numeric)

    period = "month" if re.search(r"\b(?:each|per|by|every|a)\s+month\b|\bmonthly\b", text) else (
        "year" if re.search(r"\b(?:each|per|by|every|a)\s+year\b|\byearly\b|\bannual(?:ly)?\b|\bover\s+(?:the\s+)?years\b|\bover\s+time\b", text)
        else None
    )
    if period and date_columns:
        if requested == "histogram":
            return None
        stems = {_stem(word)[:4] for word in _words(text) if len(word) > 3}
        date_column = next((column for column in date_columns if any(token[:4] in stems for token in _column_tokens(column))),
                           date_columns[0] if len(date_columns) == 1 else None)
        if date_column is None:
            return None
        if measure is not None and aggregate in {None, "count"}:
            aggregate = "average" if AVERAGE_FIRST.search(measure) else "sum"
        return [ChartSpec("time", requested or "line", measure if aggregate != "count" else None,
                          date_column, aggregate or "count", period, noun)]

    dimension = None
    for match in re.finditer(r"\b(?:by|per|for\s+each|for\s+every|across|in\s+each|of\s+each|each|every|within)\s+(?P<rest>[a-z0-9 _-]+)", text):
        dimension = _find_column(match["rest"], columns)
        if dimension is not None and dimension != measure:
            break
        dimension = None

    if requested == "histogram" and measure is None:
        target = re.search(r"\b(?:distribution|spread|histogram)\s+(?:of\s+)?(?:the\s+)?(?P<rest>.+)", text)
        if target:
            found = _find_column(target["rest"], columns)
            if found in measures:
                measure = found
            elif found is not None and dimension is None:
                dimension = found
    if requested == "histogram" and measure is not None and dimension is None:
        if aggregate is not None:
            return None
        return [ChartSpec("histogram", "histogram", measure, aggregate="count", noun=noun)]
    if dimension is None and measure is None and _CHART_WORDS.search(text):
        # "pie chart of status" / "graph of department": one category column, counted.
        target = re.search(r"\b(?:of|for|showing|on)\s+(?:the\s+|each\s+|all\s+)?(?P<rest>.+)", text)
        found = _find_column(target["rest"], columns) if target else None
        if found is not None and _usable_dimension(found, kinds[found], frame):
            dimension = found
        elif found in measures:
            measure = found
    if dimension is None:
        if measure in measures and aggregate is None and requested in {None, "histogram"} and _CHART_WORDS.search(text):
            return [ChartSpec("histogram", "histogram", measure, aggregate="count", noun=noun)]
        return None
    if measure is not None:
        if aggregate in {None, "count"}:
            aggregate = "average" if AVERAGE_FIRST.search(measure) else "sum"
        return [ChartSpec("group", requested if requested in {"bar", "column", "pie", "doughnut", "line", "area"} else "bar",
                          measure, dimension, aggregate, noun=noun)]
    if aggregate == "count" or _CHART_WORDS.search(text) or re.search(r"\b(?:breakdown|split|distribution)\b", text):
        return [ChartSpec("count", requested if requested in {"bar", "column", "pie", "doughnut", "line", "area"} else "bar",
                          None, dimension, "count", noun=noun)]
    return None


def auto_charts(frame: pd.DataFrame) -> list[ChartSpec]:
    kinds = _profile(frame)
    columns = [str(column) for column in frame.columns]
    noun = _default_noun(frame)
    priority = ("department", "dept", "region", "category", "team", "segment", "product", "type", "status", "city",
                "country", "state", "location", "branch", "gender", "grade", "level")

    def rank(column: str) -> tuple[int, int]:
        name = column.casefold()
        position = next((index for index, word in enumerate(priority) if word in name), len(priority))
        return position, categories(frame[column]).nunique()

    dimensions = sorted((column for column in columns if _usable_dimension(column, kinds[column], frame)), key=rank)
    measures = sorted(
        (column for column in columns if _usable_measure(column, kinds[column], frame)),
        key=lambda column: (not re.search(r"salary|amount|sales|revenue|price|cost|pay|income|profit|value|total", column, re.I),
                            -numbers(frame[column]).notna().sum()),
    )
    date_columns = [column for column in columns if kinds[column] == "date"]
    specs: list[ChartSpec] = []
    if dimensions:
        groups = categories(frame[dimensions[0]]).nunique()
        specs.append(ChartSpec("count", "pie" if groups <= 6 else "bar", None, dimensions[0], "count", noun=noun))
    if dimensions and measures:
        measure = measures[0]
        specs.append(ChartSpec("group", "bar", measure, dimensions[0],
                               "average" if AVERAGE_FIRST.search(measure) else "sum", noun=noun))
    if date_columns:
        dated = dates(frame[date_columns[0]]).dropna()
        period = "year" if dated.dt.year.nunique() >= 3 else "month"
        specs.append(ChartSpec("time", "line", None, date_columns[0], "count", period, noun))
    if measures:
        specs.append(ChartSpec("histogram", "histogram", measures[0], aggregate="count", noun=noun))
    if len(specs) < MAX_AUTO_CHARTS and len(measures) > 1:
        pairs = pd.DataFrame({column: numbers(frame[column]) for column in measures[:2]}).dropna()
        if 2 <= len(pairs) <= MAX_SCATTER_POINTS:
            specs.append(ChartSpec("scatter", "scatter", measures[1], measures[0], "none", noun=noun))
    if len(specs) < MAX_AUTO_CHARTS and len(dimensions) > 1:
        groups = categories(frame[dimensions[1]]).nunique()
        specs.append(ChartSpec("count", "pie" if groups <= 6 else "bar", None, dimensions[1], "count", noun=noun))
    if len(specs) < MAX_AUTO_CHARTS and len(measures) > 1 and dimensions:
        specs.append(ChartSpec("group", "bar", measures[1], dimensions[0],
                               "average" if AVERAGE_FIRST.search(measures[1]) else "sum", noun=noun))
    return specs[:MAX_AUTO_CHARTS]


# --------------------------------------------------------------------------- building charts

def _fmt(value: float) -> str:
    if value is None or (isinstance(value, float) and math.isnan(value)):
        return "n/a"
    if float(value).is_integer() or abs(value) >= 1000:
        return f"{value:,.0f}"
    return f"{value:,.2f}".rstrip("0").rstrip(".")


_AGGREGATE_LABEL = {"sum": "Total", "average": "Average", "median": "Median", "min": "Lowest", "max": "Highest", "count": "Count"}


def _left_out(total: int, used: int, what: str) -> str:
    skipped = total - used
    return f" {skipped:,} of {total:,} rows were left out because {what} was missing or unreadable." if skipped else ""


def _chart_figure(table: pd.DataFrame, chart_type: str):
    dimension, value = table.columns
    if chart_type in {"pie", "doughnut"}:
        if (table[value] < 0).any() or table[value].sum() <= 0:
            raise ValueError("Pie and doughnut charts need nonnegative values and a positive total. Use a bar chart for signed values.")
        figure = px.pie(table, names=dimension, values=value, hole=0.45 if chart_type == "doughnut" else 0,
                        color_discrete_sequence=COLORS)
        figure.update_traces(textinfo="label+percent", sort=False)
        return figure
    if chart_type == "scatter":
        return px.scatter(table, x=dimension, y=value, color_discrete_sequence=COLORS)
    if chart_type == "line":
        return px.line(table, x=dimension, y=value, markers=True, color_discrete_sequence=COLORS)
    if chart_type == "area":
        figure = px.area(table, x=dimension, y=value, markers=True, color_discrete_sequence=COLORS)
        figure.update_traces(fill="tozeroy")
        return figure
    if chart_type == "bar":
        figure = px.bar(table, x=value, y=dimension, orientation="h", text=value, color_discrete_sequence=COLORS)
        figure.update_yaxes(autorange="reversed", type="category")
        return figure
    return px.bar(table, x=dimension, y=value, text=value, color_discrete_sequence=COLORS)


def build_chart(frame: pd.DataFrame, spec: ChartSpec) -> ChartResult:
    total = len(frame)
    noun = spec.noun
    if spec.kind == "scatter":
        if spec.dimension == spec.measure:
            raise ValueError("Choose two different numeric columns for a scatter chart.")
        table = pd.DataFrame({
            spec.dimension: numbers(frame[spec.dimension]), spec.measure: numbers(frame[spec.measure]),
        }).dropna().reset_index(drop=True)
        if len(table) < 2:
            raise ValueError("A scatter chart needs at least two rows with both numeric values.")
        if len(table) > MAX_SCATTER_POINTS:
            raise ValueError(f"A scatter chart supports at most {MAX_SCATTER_POINTS} points. Filter the data first; no rows were sampled.")
        title = f"{spec.measure} vs {spec.dimension}"
        explanation = f"Each point represents one row, comparing {spec.dimension} with {spec.measure} ({len(table):,} pairs)."
        if len(table) >= 3 and table.nunique().min() > 1:
            correlation = table.iloc[:, 0].corr(table.iloc[:, 1])
            explanation += f" Pearson correlation is {correlation:.2f}; correlation does not establish cause and effect."
        else:
            explanation += "There is not enough variation or enough pairs to report a correlation."
        explanation += _left_out(total, len(table), f"{spec.dimension} or {spec.measure}")
        return _finish(ChartResult(title, "scatter", table, explanation, _chart_figure(table, "scatter")))
    if spec.kind == "time":
        when = dates(frame[spec.dimension])
        period = when.dt.year if spec.period == "year" else when.dt.to_period("M").astype(str).where(when.notna(), None)
        label = "Year" if spec.period == "year" else "Month"
        data = pd.DataFrame({label: period})
        if spec.measure:
            data["value"] = numbers(frame[spec.measure])
        data = data.dropna()
        if data.empty:
            raise ValueError(f"No readable dates were found in {spec.dimension}.")
        if spec.period == "year":
            data[label] = data[label].astype(int)
        if spec.measure:
            method = {"average": "mean"}.get(spec.aggregate, spec.aggregate)
            value_name = f"{_AGGREGATE_LABEL[spec.aggregate]} {spec.measure}"
            table = data.groupby(label)["value"].agg(method).rename(value_name).reset_index()
        else:
            value_name = f"Number of {noun}"
            table = data.groupby(label).size().rename(value_name).reset_index()
        table = table.sort_values(label).reset_index(drop=True)
        joined = re.search(r"join|hire|start", spec.dimension, re.I)
        what = f"{noun} joined" if joined and not spec.measure else (
            f"{value_name.lower()}" if spec.measure else f"{noun} have a {spec.dimension}"
        )
        title = f"{value_name} by {label.lower()} ({spec.dimension})"
        peak = table.loc[table[value_name].idxmax()]
        low = table.loc[table[value_name].idxmin()]
        explanation = (
            f"This chart shows how many {what} in each {label.lower()}, from {table[label].iloc[0]} to {table[label].iloc[-1]}."
            if not spec.measure else
            f"This chart shows the {value_name.lower()} for each {label.lower()}, from {table[label].iloc[0]} to {table[label].iloc[-1]}."
        )
        explanation += (f" The highest was {peak[label]} with {_fmt(peak[value_name])}, and the lowest was "
                        f"{low[label]} with {_fmt(low[value_name])}.")
        explanation += _left_out(total, len(data), f"the {spec.dimension}" + (f" or {spec.measure}" if spec.measure else ""))
        figure = _chart_figure(table, spec.chart_type)
        if spec.period == "year":
            figure.update_xaxes(dtick=1)
        return _finish(ChartResult(title, spec.chart_type, table, explanation, figure))

    if spec.kind == "histogram":
        values = numbers(frame[spec.measure]).dropna()
        if values.empty:
            raise ValueError(f"No numeric values were found in {spec.measure}.")
        bins = int(min(10, max(1, values.nunique())))
        counts, edges = np.histogram(values, bins=bins)
        table = pd.DataFrame({
            f"{spec.measure} range": [f"{_fmt(low)} to {_fmt(high)}" for low, high in zip(edges[:-1], edges[1:])],
            f"Number of {noun}": counts.astype(int),
        })
        busiest = table.iloc[int(counts.argmax())]
        title = f"Distribution of {spec.measure}"
        explanation = (
            f"This chart shows how {spec.measure} values are spread across {len(table)} ranges. "
            f"The most common range is {busiest.iloc[0]}, with {int(busiest.iloc[1]):,} {noun}. "
            f"The average {spec.measure} is {_fmt(values.mean())} and the median is {_fmt(values.median())}."
        )
        if len(values) < total:
            explanation += f" {total - len(values):,} rows without a usable {spec.measure} were left out."
        figure = px.bar(table, x=table.columns[0], y=table.columns[1], color_discrete_sequence=COLORS)
        figure.update_traces(marker_line_width=0)
        figure.update_layout(bargap=0.05)
        return _finish(ChartResult(title, "histogram", table, explanation, figure))

    groups = categories(frame[spec.dimension])
    if spec.kind == "count":
        data = groups.dropna()
        if data.empty:
            raise ValueError(f"{spec.dimension} has no values to count.")
        value_name = f"Number of {noun}"
        table = data.value_counts().rename_axis(spec.dimension).rename(value_name).reset_index()
    else:
        data = pd.DataFrame({spec.dimension: groups, "value": numbers(frame[spec.measure])}).dropna()
        if data.empty:
            raise ValueError(f"No rows have both a {spec.dimension} and a numeric {spec.measure}.")
        method = {"average": "mean"}.get(spec.aggregate, spec.aggregate)
        value_name = f"{_AGGREGATE_LABEL[spec.aggregate]} {spec.measure}"
        table = data.groupby(spec.dimension)["value"].agg(method).rename(value_name).reset_index()
        table = table.sort_values(value_name, ascending=False, kind="stable").reset_index(drop=True)
    others = 0
    if len(table) > spec.category_limit:
        if spec.chart_type in {"pie", "doughnut"}:
            raise ValueError(f"Too many categories for a complete share chart. Filter to at most {spec.category_limit} groups or use a bar chart.")
        others = len(table) - spec.category_limit
        table = table.head(spec.category_limit)
    top, bottom = table.iloc[0], table.iloc[-1]
    if spec.kind == "count":
        title = f"{value_name} by {spec.dimension}"
        share = top[value_name] / table[value_name].sum() * 100
        explanation = (
            f"This chart shows how many {noun} fall into each {spec.dimension} ({len(table)} groups). "
            f"{top[spec.dimension]} has the most with {int(top[value_name]):,} ({share:.0f}%), and "
            f"{bottom[spec.dimension]} has the fewest with {int(bottom[value_name]):,}."
        )
        explanation += _left_out(total, len(data), f"the {spec.dimension}")
    else:
        title = f"{value_name} by {spec.dimension}"
        word = _AGGREGATE_LABEL[spec.aggregate].lower()
        explanation = (
            f"This chart shows the {word} {spec.measure} for each {spec.dimension} ({len(table)} groups). "
            f"{top[spec.dimension]} is highest at {_fmt(top[value_name])}, and {bottom[spec.dimension]} is lowest at "
            f"{_fmt(bottom[value_name])}."
        )
        explanation += _left_out(total, len(data), f"the {spec.dimension} or {spec.measure}")
    if others:
        explanation += f" Only the top {spec.category_limit} groups are shown; {others} smaller groups were left out."
    if spec.chart_type in {"line", "area"}:
        table = table.sort_values(spec.dimension, kind="stable").reset_index(drop=True)
    figure = _chart_figure(table, spec.chart_type)
    pairs = possible_duplicates([str(label) for label in groups.dropna().unique()])
    note = (
        "Data note: some groups may be the same thing written differently ("
        + "; ".join(f"{short} / {long}" for short, long in pairs)
        + "). They are kept separate because the data spells them differently; clean those values to combine them."
    ) if pairs else ""
    return _finish(ChartResult(title, spec.chart_type, table, explanation, figure, note))


def _finish(result: ChartResult) -> ChartResult:
    figure = result.figure
    horizontal = result.chart_type == "bar"
    if result.chart_type in {"bar", "column"}:
        figure.update_traces(texttemplate="%{" + ("x" if horizontal else "y") + ":,.4~g}", textposition="outside",
                             cliponaxis=False, selector={"type": "bar"})
    figure.update_layout(
        title=result.title, height=max(380, len(result.table) * 26 + 120) if horizontal else 400,
        margin={"l": 24, "r": 48, "t": 64, "b": 40}, paper_bgcolor="#F7F8FA", plot_bgcolor="#FFFFFF", showlegend=False,
    )
    if result.chart_type in {"pie", "doughnut"}:
        figure.update_layout(showlegend=True)
    return result


def build_charts(frame: pd.DataFrame, specs: list[ChartSpec]) -> tuple[list[ChartResult], list[str]]:
    results, problems = [], []
    for spec in specs:
        try:
            results.append(build_chart(frame, spec))
        except ValueError as error:
            problems.append(str(error))
    return results, problems


def charts_answer(results: list[ChartResult], generic: bool) -> str:
    if not results:
        return "I couldn't build a chart from this data."
    intro = (f"I looked at the columns and made {len(results)} charts." if generic and len(results) > 1 else "")
    return "\n\n".join(([intro] if intro else []) + [f"**{result.title}.** {result.explanation}" for result in results])
