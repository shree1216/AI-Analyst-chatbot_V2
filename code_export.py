"""Readable code (pandas, SQL, Excel formulas) for what the backend actually ran, shown in Backend diagnostics.

Generated code is kept in the user's own session only; it is never written to the redacted server log.
"""

from __future__ import annotations

from contextlib import contextmanager
from contextvars import ContextVar
import json
from typing import Any, Callable, Iterator

from openpyxl.utils import get_column_letter

MAX_SNIPPETS = 40
GENERATED: ContextVar[list[dict[str, str]] | None] = ContextVar("generated_code", default=None)
SOURCE_FILE: ContextVar[str] = ContextVar("generated_code_source_file", default="workbook.xlsx")
_PANDAS_AGGREGATE = {
    "sum": "sum()", "average": "mean()", "min": "min()", "max": "max()", "median": "median()",
    "count": "count()", "counta": "count()", "count_distinct": "nunique()",
    "stdev": "std(ddof=1)", "stdev_p": "std(ddof=0)", "var": "var(ddof=1)", "var_p": "var(ddof=0)",
}
_EXCEL_AGGREGATE = {"sum": "SUMIFS", "average": "AVERAGEIFS", "min": "MINIFS", "max": "MAXIFS"}
_OPERATORS = {
    "add": "+", "subtract": "-", "multiply": "*", "divide": "/", "power": "**",
    "eq": "==", "ne": "!=", "gt": ">", "gte": ">=", "lt": "<", "lte": "<=",
}


@contextmanager
def capture(target: list[dict[str, str]], file_name: str | None = None) -> Iterator[list[dict[str, str]]]:
    token = GENERATED.set(target)
    file_token = SOURCE_FILE.set(file_name) if file_name else None
    try:
        yield target
    finally:
        GENERATED.reset(token)
        if file_token is not None:
            SOURCE_FILE.reset(file_token)


def emit(language: str, title: str, code: str) -> None:
    target = GENERATED.get()
    if target is None or not code.strip():
        return
    snippet = {"language": language, "title": title, "code": code.strip() + "\n"}
    if snippet not in target and len(target) < MAX_SNIPPETS:
        target.append(snippet)


def emit_lazy(language: str, title: str, build: Callable[[], str | None]) -> None:
    """Build and record code only while an answer is being captured; never let diagnostics break a result."""
    if GENERATED.get() is None:
        return
    try:
        code = build()
    except Exception as error:  # diagnostics are best-effort and must not affect the answer
        code = f"# Code for this step could not be generated ({type(error).__name__})."
    if code:
        emit(language, title, code)


def load_code(file_name: str, sheet: str | None) -> str:
    lower = file_name.casefold()
    if lower.endswith(".csv"):
        return f"import numpy as np\nimport pandas as pd\n\ndf = pd.read_csv({file_name!r})"
    if lower.endswith(".parquet"):
        return f"import numpy as np\nimport pandas as pd\n\ndf = pd.read_parquet({file_name!r})"
    if lower.endswith(".pdf"):
        return (
            "import numpy as np\nimport pandas as pd\nimport pdfplumber\n\n"
            f"with pdfplumber.open({file_name!r}) as pdf:  # tables were extracted into worksheets\n"
            "    rows = pdf.pages[0].extract_tables()[0]\n"
            "df = pd.DataFrame(rows[1:], columns=rows[0])"
        )
    return f"import numpy as np\nimport pandas as pd\n\ndf = pd.read_excel({file_name!r}, sheet_name={sheet!r})"


def sql_code(sql: str, parameters: dict[str, Any] | None) -> str:
    from backend_diagnostics import diagnostic_sql

    text = diagnostic_sql(sql.strip())
    if parameters:
        values = "\n".join(
            f"--   :{name}  ({type(value).__name__}, value from your question - redacted)"
            for name, value in parameters.items()
        )
        text = f"-- Parameters\n{values}\n{text}"
    return text


# --------------------------------------------------------------------------- workbook plans

_PARAMS: ContextVar[list[str] | None] = ContextVar("generated_code_params", default=None)


def _literal(value: Any, params: list[str] | None) -> str:
    # Values typed in a question are redacted in diagnostics; column names and operations are shown.
    params = params if params is not None else _PARAMS.get()
    if params is None or value is None or isinstance(value, bool):
        return repr(value)
    params.append(type(value).__name__)
    return f"params[{len(params) - 1}]"


def expression_code(node: Any, df: str = "df", params: list[str] | None = None) -> str:
    if not isinstance(node, dict):
        return _literal(node, params)
    if "literal" in node:
        return _literal(node["literal"], params)
    if "column" in node:
        return f"{df}[{node['column']!r}]"
    function, args = node.get("function"), node.get("args", [])
    parts = [expression_code(arg, df, params) for arg in args]
    if function in _OPERATORS and len(parts) == 2:
        return f"({parts[0]} {_OPERATORS[function]} {parts[1]})"
    if function in {"and", "or"}:
        return "(" + f" {'&' if function == 'and' else '|'} ".join(parts) + ")"
    if function == "not":
        return f"~{parts[0]}"
    if function == "in":
        return f"{parts[0]}.isin({parts[1]})"
    if function == "is_blank":
        return f"{parts[0]}.isna()"
    if function == "if":
        return f"np.where({parts[0]}, {parts[1]}, {parts[2]})"
    if function == "coalesce":
        return parts[0] + "".join(f".fillna({part})" for part in parts[1:])
    if function in _PANDAS_AGGREGATE or function == "count_rows":
        source = df if "where" not in node else f"{df}.loc[{expression_code(node['where'], df, params)}]"
        if function == "count_rows":
            return f"len({source})"
        column = args[0].get("column") if args and isinstance(args[0], dict) else None
        return f"{source}[{column!r}].{_PANDAS_AGGREGATE[function]}" if column else f"{function}({', '.join(parts)})"
    if function in {"year", "month", "day"}:
        return f"pd.to_datetime({parts[0]}).dt.{function}"
    if function == "date":
        if len(parts) == 3:
            return f"pd.to_datetime(dict(year={parts[0]}, month={parts[1]}, day={parts[2]}))"
        return f"pd.to_datetime({parts[0]})"
    if function == "datediff":
        unit = args[2] if len(args) > 2 else "days"
        days = f"(pd.to_datetime({parts[1]}) - pd.to_datetime({parts[0]})).dt.days"
        return days if unit == "days" else f"{days} // {30 if unit == 'months' else 365}  # whole {unit}"
    if function == "today":
        return "pd.Timestamp.today().normalize()"
    if function == "concat":
        return " + ".join(f"{part}.astype(str)" for part in parts)
    if function in {"left", "right"}:
        return f"{parts[0]}.str[:{parts[1]}]" if function == "left" else f"{parts[0]}.str[-{parts[1]}:]"
    if function == "contains":
        return f"{parts[0]}.str.contains({parts[1]}, case=False, regex=False)"
    if function in {"starts_with", "ends_with"}:
        method = "startswith" if function == "starts_with" else "endswith"
        return f"{parts[0]}.str.lower().str.{method}({parts[1]}.lower())"
    if function == "trim":
        return f"{parts[0]}.str.strip()"
    if function == "length":
        return f"{parts[0]}.str.len()"
    if function in {"lower", "upper"}:
        return f"{parts[0]}.str.{function}()"
    if function == "abs":
        return f"{parts[0]}.abs()"
    if function == "number":
        return f"pd.to_numeric({parts[0]}, errors='coerce')"
    return f"{function}({', '.join(parts)})"


_AGG_FUNCTION_CODE = {
    "sum": "'sum'", "average": "'mean'", "min": "'min'", "max": "'max'", "median": "'median'",
    "count": "'count'", "counta": "'count'", "count_distinct": "'nunique'", "stdev": "'std'",
    "stdev_p": "lambda s: s.std(ddof=0)", "var": "'var'", "var_p": "lambda s: s.var(ddof=0)",
}


def _step_code(step: dict[str, Any]) -> list[str]:
    operation = step.get("op")
    if operation == "filter":
        return [f"df = df[{expression_code(step.get('where'))}]"]
    if operation == "calculate":
        return [f"df[{item.get('name')!r}] = {expression_code(item.get('expression'))}" for item in step.get("columns", [])]
    if operation == "aggregate":
        groups = step.get("group_by", [])
        metrics = step.get("metrics", [])
        if not groups:
            values = ",\n    ".join(
                f"{metric.get('name')!r}: [" + expression_code({
                    "function": metric.get("function"),
                    "args": [] if metric.get("column", "*") == "*" else [{"column": metric.get("column")}],
                    **({"where": metric["where"]} if "where" in metric else {}),
                }) + "]"
                for metric in metrics
            )
            return [f"df = pd.DataFrame({{\n    {values},\n}})"]
        notes = [
            f"# {metric.get('name')!r} only includes rows where {expression_code(metric['where'])}"
            for metric in metrics if "where" in metric
        ]
        named = ",\n    ".join(
            f"{metric.get('name')!r}: "
            + (f"({groups[0]!r}, 'size')" if metric.get("function") == "count_rows" or metric.get("column", "*") == "*"
               else f"({metric.get('column')!r}, {_AGG_FUNCTION_CODE.get(metric.get('function'), repr(metric.get('function')))})")
            for metric in metrics
        )
        return [*notes, f"df = df.groupby({groups!r}, dropna=False, sort=False).agg(**{{\n    {named},\n}}).reset_index()"]
    if operation == "sort":
        columns = [item.get("column") for item in step.get("by", [])]
        ascending = [item.get("direction") == "asc" for item in step.get("by", [])]
        return [f"df = df.sort_values({columns!r}, ascending={ascending!r}, kind='stable')"]
    if operation == "limit":
        return [f"df = df.head({step.get('count')})"]
    if operation == "select":
        selections = [item if isinstance(item, dict) else {"column": item} for item in step.get("columns", [])]
        columns = [item["column"] for item in selections if "column" in item]
        renames = {item["column"]: item["name"] for item in selections if "column" in item and "name" in item}
        literals = [f"df[{item['name']!r}] = {_literal(item['literal'], None)}" for item in selections if "literal" in item]
        code = [f"df = df[{columns!r}]"] + ([f"df = df.rename(columns={renames!r})"] if renames else [])
        return code + literals
    if operation == "rank":
        groups = step.get("group_by", [])
        source = f"df.groupby({groups!r})[{step.get('column')!r}]" if groups else f"df[{step.get('column')!r}]"
        return [f"df[{step.get('name')!r}] = {source}.rank(method={step.get('method', 'min')!r}, "
                f"ascending={step.get('direction') == 'asc'})"]
    if operation in {"join", "lookup"}:
        columns = [item.get("column") if isinstance(item, dict) else item for item in step.get("columns", [])]
        return [
            f"other = pd.read_excel(WORKBOOK, sheet_name={step.get('sheet')!r})",
            f"df = df.merge(other[{step.get('right_on', [])!r} + {columns!r}], how={step.get('how', 'left')!r}, "
            f"left_on={step.get('left_on')!r}, right_on={step.get('right_on')!r})",
        ]
    if operation == "append":
        return [f"df = pd.concat([df, pd.read_excel(WORKBOOK, sheet_name={step.get('sheet')!r})], ignore_index=True)"]
    if operation == "pivot":
        function = _AGG_FUNCTION_CODE.get(step.get("function"), "'sum'")
        return [f"df = df.pivot_table(index={step.get('index')!r}, columns={step.get('columns')!r}, "
                f"values={step.get('values')!r}, aggfunc={function}"
                + (f", fill_value={step['fill_value']!r}" if "fill_value" in step else "") + ").reset_index()"]
    if operation == "extrema":
        column = step.get("column")
        which = step.get("which", "both")
        picks = {"highest": f"df.nlargest(1, {column!r})", "lowest": f"df.nsmallest(1, {column!r})"}
        chosen = list(picks.values()) if which == "both" else [picks[which]]
        return [f"df = pd.concat([{', '.join(chosen)}])"]
    return [f"# {operation} step: {json.dumps(step, default=str)}"]


def workbook_plan_code(plan: dict[str, Any], file_name: str | None = None, *, redact: bool = True) -> str:
    workbook = file_name or SOURCE_FILE.get()
    # CSV/Parquet uploads are converted to a one-sheet workbook, so read them with pandas' matching reader.
    reader = (
        "df = pd.read_csv(WORKBOOK)" if workbook.casefold().endswith(".csv")
        else "df = pd.read_parquet(WORKBOOK)" if workbook.casefold().endswith(".parquet")
        else f"df = pd.read_excel(WORKBOOK, sheet_name={plan.get('sheet')!r})"
    )
    params: list[str] = []
    token = _PARAMS.set(params if redact else None)
    try:
        body: list[str] = []
        for number, step in enumerate(plan.get("steps", []), start=1):
            if isinstance(step, dict):
                body.append(f"\n# Step {number}: {step.get('op', 'step')}")
                try:
                    body.extend(_step_code(step))
                except (KeyError, TypeError, ValueError, AttributeError):
                    body.append(f"# {step.get('op', 'step')} step: {json.dumps(step, default=str)}")
    finally:
        _PARAMS.reset(token)
    lines = ["import numpy as np", "import pandas as pd", "", f"WORKBOOK = {workbook!r}", reader]
    if params:
        lines.append(
            f"params = [...]  # {len(params)} value(s) from your question, redacted in diagnostics: " + ", ".join(params)
        )
    lines.extend(body)
    lines.append("\nresult = df")
    return "\n".join(lines)


# --------------------------------------------------------------------------- charts and summaries

def chart_spec_code(spec: Any, top: int = 25) -> str:
    noun = getattr(spec, "noun", "rows")
    measure, dimension = spec.measure, spec.dimension
    if spec.kind == "count":
        return (f"table = df[{dimension!r}].str.strip().value_counts().rename('Number of {noun}').reset_index()"
                f"\ntable = table.head({top})")
    if spec.kind == "group":
        method = _PANDAS_AGGREGATE.get(spec.aggregate, "sum()")
        return (
            f"values = pd.to_numeric(df[{measure!r}], errors='coerce')\n"
            f"table = (df.assign(**{{{measure!r}: values}})\n"
            f"         .dropna(subset=[{dimension!r}, {measure!r}])\n"
            f"         .groupby({dimension!r})[{measure!r}].{method}\n"
            f"         .sort_values(ascending=False).head({top}).reset_index())"
        )
    if spec.kind == "time":
        period = "df_dates.dt.year" if spec.period == "year" else "df_dates.dt.to_period('M').astype(str)"
        value = (f".assign(value=pd.to_numeric(df[{measure!r}], errors='coerce')).groupby('period')['value']"
                 f".{_PANDAS_AGGREGATE.get(spec.aggregate, 'sum()')}" if measure else ".groupby('period').size()")
        return (f"df_dates = pd.to_datetime(df[{dimension!r}], errors='coerce')\n"
                f"table = df.assign(period={period}).dropna(subset=['period']){value}.sort_index().reset_index()")
    if spec.kind == "histogram":
        return (f"values = pd.to_numeric(df[{measure!r}], errors='coerce').dropna()\n"
                f"counts, edges = np.histogram(values, bins=min(10, values.nunique()))\n"
                f"table = pd.DataFrame({{'range': [f'{{a:,.0f}} to {{b:,.0f}}' for a, b in zip(edges[:-1], edges[1:])], "
                f"'Number of {noun}': counts}})")
    if spec.kind == "scatter":
        return (f"table = df[[{dimension!r}, {measure!r}]].apply(pd.to_numeric, errors='coerce').dropna()\n"
                f"correlation = table[{dimension!r}].corr(table[{measure!r}])")
    return f"# {spec.kind} chart"


def chart_spec_excel(spec: Any, columns: list[str], last_row: int) -> str | None:
    """Equivalent worksheet formulas for one category row (copy down for each category in column A)."""
    def column_range(name: str | None) -> str | None:
        if name not in columns:
            return None
        letter = get_column_letter(columns.index(name) + 1)
        return f"${letter}$2:${letter}${last_row}"

    dimension, measure = column_range(spec.dimension), column_range(spec.measure)
    if spec.kind == "count" and dimension:
        return (f"' Category list:  =UNIQUE({dimension})\n"
                f"' Count per category (next to the category in A2):  =COUNTIFS({dimension}, A2)")
    if spec.kind == "group" and dimension and measure:
        if spec.aggregate in _EXCEL_AGGREGATE:
            formula = f"={_EXCEL_AGGREGATE[spec.aggregate]}({measure}, {dimension}, A2)"
        elif spec.aggregate == "median":
            formula = f"=MEDIAN(IF({dimension}=A2, {measure}))"
        else:
            formula = f"=COUNTIFS({dimension}, A2)"
        return f"' Category list:  =UNIQUE({dimension})\n' Value per category (next to A2):  {formula}"
    if spec.kind == "histogram" and measure:
        return f"' Bin edges in A2:A11, then:  =FREQUENCY({measure}, A2:A11)"
    if spec.kind == "scatter" and dimension and measure:
        return f"' Pearson correlation:  =CORREL({dimension}, {measure})"
    if spec.kind == "time" and dimension:
        return f"' Rows per year (year in A2):  =SUMPRODUCT(--(YEAR({dimension})=A2))"
    return None


def overview_code() -> str:
    return (
        "summary = pd.DataFrame({\n"
        "    'Column': df.columns,\n"
        "    'Type': df.dtypes.astype(str).values,\n"
        "    'Missing values': df.isna().sum().values,\n"
        "    'Unique values': df.nunique().values,\n"
        "})\n"
        "numeric_summary = df.describe()  # count, mean, std, min, quartiles, max"
    )
