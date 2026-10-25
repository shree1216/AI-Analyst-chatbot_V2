from datetime import date, datetime
from dataclasses import dataclass
import math
from numbers import Real
from typing import Any

import pandas as pd
from backend_diagnostics import BackendTrace


MAX_STEPS = 20
MAX_SHEETS = 20
MAX_ROWS = 50_000
MAX_COLUMNS = 200
MAX_CELLS = 5_000_000


@dataclass(frozen=True)
class JoinDiagnostic:
    left_sources: frozenset[str]
    right_source: str
    left_columns: tuple[str, ...]
    right_columns: tuple[str, ...]
    left_rows: int
    matched_rows: int


AGGREGATES = {"sum", "average", "min", "max", "count", "counta", "count_rows", "count_distinct", "median"}
FUNCTION_ARITY = {
    **{name: (2, 2) for name in (
        "add", "subtract", "multiply", "divide", "power", "eq", "ne", "gt", "gte", "lt", "lte",
        "in", "contains", "starts_with", "ends_with", "left", "right",
    )},
    **{name: (1, 1) for name in (
        "not", "abs", "number", "text", "trim", "lower", "upper", "length", "year", "month", "day", "is_blank",
    )},
    **{name: (1, 1) for name in AGGREGATES - {"count_rows"}},
    "count_rows": (0, 0), "and": (2, 20), "or": (2, 20), "if": (3, 3),
    "concat": (2, 20), "coalesce": (2, 20), "date": (1, 3), "datediff": (3, 3), "today": (0, 0),
}


def _require(condition: bool, message: str) -> None:
    if not condition:
        raise ValueError(message)


def _name(value: Any) -> str:
    _require(isinstance(value, str) and 0 < len(value) <= 256, "Use a nonempty column or worksheet name.")
    return value


def _columns(value: Any, frame: pd.DataFrame, *, empty: bool = False) -> list[str]:
    _require(isinstance(value, list) and (empty or bool(value)) and len(value) <= MAX_COLUMNS,
             "Provide a bounded list of column names.")
    names = [_name(item) for item in value]
    _require(len(set(names)) == len(names), "Column names in a list must be unique.")
    _require(all(name in frame.columns for name in names), "A requested column is not available in this result.")
    return names


def _bounded(frame: pd.DataFrame) -> pd.DataFrame:
    _require(len(frame) <= MAX_ROWS and len(frame.columns) <= MAX_COLUMNS and frame.size <= MAX_CELLS,
             "This calculation is too large. Filter the data or reduce the number of columns first.")
    _require(frame.columns.is_unique, "The calculation produced duplicate column names. Use distinct output names.")
    _require(not frame.select_dtypes(include="number").isin([float("inf"), -float("inf")]).any().any(),
             "The calculation contains an infinite number. Check the inputs or reduce the calculation.")
    return frame.reset_index(drop=True)


def _series(value: Any, frame: pd.DataFrame) -> pd.Series:
    _require(not isinstance(value, list), "Literal lists can only be used with IN conditions.")
    return value.reindex(frame.index) if isinstance(value, pd.Series) else pd.Series(value, index=frame.index)


def _truth(value: Any, frame: pd.DataFrame) -> pd.Series:
    result = _series(value, frame)
    _require(pd.api.types.is_bool_dtype(result) or result.dropna().map(lambda item: isinstance(item, bool)).all(),
             "A logical condition must produce true or false values.")
    return result.astype("boolean").fillna(False)


def _numeric(value: Any, frame: pd.DataFrame, active: pd.Series) -> pd.Series:
    original = _series(value, frame)
    _require(not pd.api.types.is_datetime64_any_dtype(original), "Use YEAR, MONTH, DAY, or a date difference for dates.")
    converted = pd.to_numeric(original, errors="coerce").astype(float)
    invalid = original.notna() & (converted.isna() | converted.isin([float("inf"), -float("inf")]))
    _require(not (invalid & active).any(), "This calculation needs finite numeric values. Check the selected columns.")
    return converted


def _dates(value: Any, frame: pd.DataFrame, active: pd.Series) -> pd.Series:
    original = _series(value, frame)
    numeric = original.map(lambda item: isinstance(item, Real) and not isinstance(item, bool))
    _require(not (numeric & original.notna() & active).any(),
             "Use Excel date cells or ISO date text, not ambiguous numeric date serials.")
    text = original.astype("string")
    iso = text.str.match(r"^\d{4}-\d{2}-\d{2}(?:[ T].*)?$", na=False)
    actual_date = original.map(lambda item: isinstance(item, (date, datetime)))
    _require(not (original.notna() & ~(iso | actual_date) & active).any(),
             "Use unambiguous dates such as 2025-01-31.")
    converted = pd.to_datetime(original, errors="coerce", format="mixed", utc=True).dt.tz_localize(None)
    _require(not (original.notna() & converted.isna() & active).any(), "Some date values are invalid. Check the worksheet.")
    return converted


def _literal(value: Any) -> bool:
    if value is None or isinstance(value, (str, bool, int)):
        return (not isinstance(value, str) or len(value) <= 32767) and (
            not isinstance(value, int) or abs(value) <= 1e308
        )
    if isinstance(value, float):
        return math.isfinite(value)
    return isinstance(value, list) and len(value) <= 100 and all(
        not isinstance(item, list) and _literal(item) for item in value
    )


def _validate_expression(
    node: Any, columns: pd.Index, depth: int = 0, budget: list[int] | None = None, *, allow_list: bool = False,
) -> None:
    budget = [0] if budget is None else budget
    budget[0] += 1
    _require(depth <= 12 and budget[0] <= 200, "The requested formula is too complex. Split it into smaller calculations.")
    if depth > 0 and not isinstance(node, dict):
        _require(_literal(node) and (allow_list or not isinstance(node, list)),
                 "Function arguments must be expressions or finite JSON literals; lists are only valid for IN.")
        return
    _require(isinstance(node, dict), "A calculation expression must be a structured object.")
    if set(node) == {"column"}:
        _require(_name(node["column"]) in columns, "A requested calculation column is not available.")
        return
    if set(node) == {"literal"}:
        _require(_literal(node["literal"]), "Use finite numbers, text, dates, or bounded literal lists.")
        _require(allow_list or not isinstance(node["literal"], list), "Literal lists can only be used with IN conditions.")
        return
    _require(set(node) <= {"function", "args", "where"} and {"function", "args"} <= set(node),
             "The calculation contains unsupported expression fields.")
    function = node["function"]
    _require(isinstance(function, str) and function in FUNCTION_ARITY, "That spreadsheet function is not supported.")
    args = node["args"]
    minimum, maximum = FUNCTION_ARITY[function]
    _require(isinstance(args, list) and minimum <= len(args) <= maximum and not (function == "date" and len(args) == 2),
             "A spreadsheet function has the wrong number of arguments.")
    for index, arg in enumerate(args):
        _validate_expression(arg, columns, depth + 1, budget, allow_list=function == "in" and index == 1)
    if "where" in node:
        _require(function in AGGREGATES, "Only aggregate functions accept a conditional filter.")
        _validate_expression(node["where"], columns, depth + 1, budget)


def _aggregate_values(function: str, series: pd.Series, frame: pd.DataFrame, mask: pd.Series) -> pd.Series:
    if function in {"min", "max"}:
        if pd.api.types.is_datetime64_any_dtype(series):
            return series.where(mask)
        selected = series.loc[mask].dropna()
        if not selected.empty and selected.map(lambda item: isinstance(item, str)).all() and (
            pd.to_numeric(selected, errors="coerce").isna().any()
        ):
            return series.where(mask).astype("string")
    return _numeric(series, frame, mask).where(mask)


def _aggregate(function: str, value: Any, frame: pd.DataFrame, mask: pd.Series) -> Any:
    if function == "count_rows":
        return int(mask.sum())
    series = _series(value, frame)
    selected = series.loc[mask]
    if function == "counta":
        return int(selected.notna().sum())
    if function == "count_distinct":
        return int(selected.nunique(dropna=True))
    if function == "count":
        return int(selected.map(
            lambda item: isinstance(item, (Real, date, datetime)) and not isinstance(item, bool) and not pd.isna(item)
        ).sum())
    values = _aggregate_values(function, series, frame, mask).loc[mask]
    if values.empty:
        return 0 if function == "sum" else None
    if not values.notna().any():
        return None
    if function == "sum":
        return values.sum(min_count=1)
    return getattr(values, "mean" if function == "average" else function)()


def _evaluate(node: Any, frame: pd.DataFrame, active: pd.Series) -> Any:
    if not active.any():
        return None
    if not isinstance(node, dict):
        return node
    if "literal" in node:
        return node["literal"]
    if "column" in node:
        return frame[node["column"]]
    function, args = node["function"], node["args"]
    if function == "coalesce":
        result = _series(_evaluate(args[0], frame, active), frame).astype(object)
        for arg in args[1:]:
            missing = result.isna()
            remaining = active & missing
            if not remaining.any():
                break
            fallback = _series(_evaluate(arg, frame, remaining), frame)
            result = result.where(~missing, fallback)
        return result.infer_objects()
    if function == "if":
        condition = _truth(_evaluate(args[0], frame, active), frame)
        yes = _series(_evaluate(args[1], frame, active & condition), frame)
        no = _series(_evaluate(args[2], frame, active & ~condition), frame)
        return yes.where(condition, no).infer_objects()
    if function in AGGREGATES:
        all_rows = pd.Series(True, index=frame.index)
        mask = _truth(_evaluate(node["where"], frame, all_rows), frame) if "where" in node else all_rows
        value = _evaluate(args[0], frame, mask) if args else None
        return _aggregate(function, value, frame, mask)
    values = [_evaluate(arg, frame, active) for arg in args]
    if function == "today":
        return pd.Timestamp.now(tz="UTC").normalize().tz_localize(None)
    if function in {"and", "or", "not"}:
        conditions = [_truth(value, frame) for value in values]
        result = conditions[0]
        for condition in conditions[1:]:
            result = result & condition if function == "and" else result | condition
        return ~result if function == "not" else result
    if function == "is_blank":
        value = _series(values[0], frame)
        return value.isna() | value.astype("string").str.strip().eq("").fillna(False)
    if function in {"add", "subtract", "multiply", "divide", "power", "abs", "number"}:
        numeric = [_numeric(value, frame, active) for value in values]
        if function in {"abs", "number"}:
            return numeric[0].abs() if function == "abs" else numeric[0]
        left, right = numeric
        if function == "divide":
            _require(not (right.eq(0) & active).any(), "The calculation would divide by zero. Add an IF condition or filter.")
            right = right.mask(right.eq(0))
        if function == "power":
            _require(not (right.abs().gt(1000) & active).any(), "The requested exponent is too large.")
        operation = {"add": "add", "subtract": "sub", "multiply": "mul", "divide": "div", "power": "pow"}[function]
        result = getattr(left, operation)(right)
        invalid = result.isin([float("inf"), -float("inf")]) | (result.isna() & left.notna() & right.notna())
        _require(not (invalid & active).any(), "The calculation produced an undefined or overflowing number.")
        return result
    if function in {"eq", "ne", "gt", "gte", "lt", "lte", "in"}:
        left = _series(values[0], frame)
        if function == "in":
            _require(isinstance(values[1], list), "IN requires a literal list of values.")
            if pd.api.types.is_string_dtype(left):
                return left.astype("string").str.strip().str.casefold().isin([
                    item.strip().casefold() if isinstance(item, str) else item for item in values[1]
                ])
            return left.isin(values[1])
        right = _series(values[1], frame)
        if pd.api.types.is_string_dtype(left) and pd.api.types.is_string_dtype(right):
            left = left.astype("string").str.strip().str.casefold()
            right = right.astype("string").str.strip().str.casefold()
        try:
            result = getattr(left, {"eq": "eq", "ne": "ne", "gt": "gt", "gte": "ge", "lt": "lt", "lte": "le"}[function])(right)
        except TypeError as error:
            raise ValueError("These columns cannot be compared. Use consistent numeric, text, or date values.") from error
        return result.fillna(False) & left.notna() & right.notna()
    if function in {"date", "year", "month", "day", "datediff"}:
        if function == "date" and len(values) == 3:
            components = {name: _numeric(value, frame, active) for name, value in zip(("year", "month", "day"), values)}
            _require(all(not (series.mod(1).ne(0) & series.notna() & active).any() for series in components.values()),
                     "DATE requires whole-number years, months, and days.")
            result = pd.to_datetime(pd.DataFrame(components), errors="coerce")
            _require(not (result.isna() & active).any(), "DATE contains an invalid year, month, or day.")
            return result
        dates = _dates(values[0], frame, active)
        if function == "date":
            return dates
        if function != "datediff":
            return getattr(dates.dt, function)
        end = _dates(values[1], frame, active)
        unit = values[2]
        _require(isinstance(unit, str) and unit in {"days", "months", "years"},
                 "Date differences support days, months, or years.")
        _require(not (end.lt(dates) & active).any(), "The end date must not be earlier than the start date.")
        if unit == "days":
            return (end - dates).dt.days
        if unit == "months":
            return (end.dt.year - dates.dt.year) * 12 + end.dt.month - dates.dt.month - end.dt.day.lt(dates.dt.day)
        anniversary_pending = end.dt.month.lt(dates.dt.month) | (
            end.dt.month.eq(dates.dt.month) & end.dt.day.lt(dates.dt.day)
        )
        return end.dt.year - dates.dt.year - anniversary_pending
    text = [_series(value, frame).astype("string") for value in values]
    if function == "concat":
        lengths = sum(item.str.len().fillna(0) for item in text)
        _require(not lengths.gt(32767).any() and lengths.sum() <= 10_000_000,
                 "The combined text would be too large. Filter the rows or concatenate shorter values.")
        result = text[0].fillna("")
        for item in text[1:]:
            result = result + item.fillna("")
        return result
    if function in {"left", "right"}:
        count = values[1]
        _require(isinstance(count, int) and not isinstance(count, bool) and 0 <= count <= 32767,
                 "LEFT and RIGHT need a constant nonnegative character count.")
        return text[0].str.slice(0, count) if function == "left" or count == 0 else text[0].str.slice(-count)
    if function in {"contains", "starts_with", "ends_with"}:
        _require(isinstance(values[1], str), "Text matching needs a constant text value.")
        source, term = text[0].str.casefold(), values[1].casefold()
        if function == "contains":
            return source.str.contains(term, regex=False, na=False)
        return getattr(source.str, "startswith" if function == "starts_with" else "endswith")(term, na=False)
    if function == "trim":
        return text[0].str.replace(r" +", " ", regex=True).str.strip(" ")
    if function == "length":
        return text[0].str.len()
    if function in {"lower", "upper"}:
        return getattr(text[0].str, function)()
    return text[0]


def evaluate_expression(node: Any, frame: pd.DataFrame) -> Any:
    _validate_expression(node, frame.columns)
    return _evaluate(node, frame, pd.Series(True, index=frame.index))


def _keys(step: dict[str, Any], allowed: set[str], required: set[str]) -> None:
    _require(set(step) <= allowed | {"op"} and required | {"op"} <= set(step),
             "The calculation step contains missing or unsupported fields.")


def _aggregate_frame(frame: pd.DataFrame, groups: list[str], metrics: Any) -> pd.DataFrame:
    _require(isinstance(metrics, list) and 1 <= len(metrics) <= 20, "Provide between 1 and 20 summary metrics.")
    result: dict[str, Any] = {}
    group_index = frame.groupby(groups, dropna=False, sort=False, observed=True).size().index if groups else None
    for metric in metrics:
        _require(isinstance(metric, dict) and set(metric) <= {"name", "function", "column", "where"}
                 and {"name", "function"} <= set(metric), "A summary metric is invalid.")
        name, function = _name(metric["name"]), metric["function"]
        _require(name not in result and name not in groups, "Summary output names must be unique.")
        _require(isinstance(function, str) and function in AGGREGATES, "That summary function is not supported.")
        column = metric.get("column", "*")
        _require((function == "count_rows" and column == "*")
                 or (function != "count_rows" and isinstance(column, str) and column in frame),
                 "The summary needs an available value column.")
        mask = _truth(evaluate_expression(metric["where"], frame), frame) if "where" in metric else pd.Series(True, index=frame.index)
        values = frame[column] if function != "count_rows" else pd.Series(1, index=frame.index)
        if not groups:
            result[name] = [_aggregate(function, values, frame, mask)]
            continue
        if function == "count":
            values = values.map(lambda item: isinstance(item, (Real, date, datetime)) and not isinstance(item, bool)
                                and not pd.isna(item)).astype(int).where(mask, 0)
        elif function in {"count_rows", "counta"}:
            values = values.notna().astype(int).where(mask, 0)
        elif function != "count_distinct":
            values = _aggregate_values(function, values, frame, mask)
        else:
            values = values.where(mask)
        work = frame[groups].copy()
        temporary = next(f"_metric_{index}" for index in range(MAX_COLUMNS + 1) if f"_metric_{index}" not in groups)
        work[temporary] = values
        grouped = work.groupby(groups, dropna=False, sort=False, observed=True)[temporary]
        if function == "sum":
            aggregated = grouped.sum(min_count=1)
            if "where" in metric:
                work[temporary] = mask.astype(int)
                matches = work.groupby(groups, dropna=False, sort=False, observed=True)[temporary].sum()
                aggregated = aggregated.where(matches.ne(0), 0)
        elif function in {"count", "counta", "count_rows"}:
            aggregated = grouped.sum()
        else:
            aggregated = getattr(grouped, {"average": "mean", "count_distinct": "nunique"}.get(function, function))()
        result[name] = aggregated.reindex(group_index).to_numpy()
    output = group_index.to_frame(index=False) if group_index is not None else pd.DataFrame(index=[0])
    for name, values in result.items():
        output[name] = values
    return output


def _join(
    frame: pd.DataFrame, step: dict[str, Any], sheets: dict[str, pd.DataFrame], *, match_null_keys: bool = True,
) -> tuple[pd.DataFrame, int]:
    lookup = step["op"] == "lookup"
    allowed = {"sheet", "left_on", "right_on", "columns", "how", "keep"} if lookup else {
        "sheet", "left_on", "right_on", "columns", "how",
    }
    _keys(step, allowed, {"sheet", "left_on", "right_on", "columns"})
    sheet = _name(step["sheet"])
    _require(sheet in sheets, "A lookup worksheet is unavailable. Check its name and headers.")
    source = _bounded(sheets[sheet])
    left_keys = _columns(step["left_on"], frame)
    right_keys = _columns(step["right_on"], source)
    _require(len(left_keys) == len(right_keys) <= 5, "A join needs the same number of keys on both worksheets (at most five).")
    how = step.get("how", "left")
    _require(isinstance(how, str) and how in ({"left"} if lookup else {"left", "inner", "right", "outer"}),
             "Use a left lookup or a left, inner, right, or outer join.")
    selections = step["columns"]
    _require(isinstance(selections, list) and 1 <= len(selections) <= 100, "Choose the columns to retrieve from the other sheet.")
    _require(len(frame.columns) + len(selections) <= MAX_COLUMNS, "This join would produce too many columns. Select fewer columns first.")
    right = pd.DataFrame(index=source.index)
    temporary_keys = []
    for index, column in enumerate(right_keys):
        temporary = f"__workbook_join_{index}"
        while temporary in frame or temporary in right:
            temporary += "_"
        temporary_keys.append(temporary)
        right[temporary] = source[column]
    for selection in selections:
        _require(isinstance(selection, dict) and set(selection) == {"column", "name"}, "Name each retrieved lookup column.")
        column, name = _name(selection["column"]), _name(selection["name"])
        _require(column in source, "A retrieved lookup column does not exist.")
        _require(name not in frame and name not in right, "A retrieved column would collide with an existing name. Choose an alias.")
        right[name] = source[column]
    left = frame.copy()
    for left_key, right_key in zip(left_keys, temporary_keys):
        left[right_key] = left[left_key].map(lambda value: value.strip().casefold() if isinstance(value, str) else value)
        right[right_key] = right[right_key].map(lambda value: value.strip().casefold() if isinstance(value, str) else value)
    if not match_null_keys:
        temporary = "__workbook_join_valid"
        while temporary in left or temporary in right:
            temporary += "_"
        for table, marker in ((left, 1), (right, 2)):
            valid = table[temporary_keys].notna().all(axis=1)
            for key in temporary_keys:
                valid &= table[key].map(lambda value: not isinstance(value, str) or bool(value.strip()))
            table[temporary] = (~valid).astype(int) * marker
        temporary_keys.append(temporary)
    if lookup:
        keep = step.get("keep", "first")
        _require(isinstance(keep, str) and keep in {"first", "last"}, "Exact lookups support the first or last matching row.")
        right = right.drop_duplicates(temporary_keys, keep=keep)
    keys = [f"k{index}" for index in range(len(temporary_keys))]
    left_counts = left[temporary_keys].set_axis(keys, axis=1).value_counts(dropna=False).rename("left_count").reset_index()
    right_counts = right[temporary_keys].set_axis(keys, axis=1).value_counts(dropna=False).rename("right_count").reset_index()
    try:
        counts = left_counts.merge(right_counts, on=keys, how=how)
        valid_keys = counts[keys].notna().all(axis=1)
        for key in keys:
            valid_keys &= counts[key].map(lambda value: not isinstance(value, str) or bool(value.strip()))
        matched_rows = int(counts.loc[
            valid_keys & counts["left_count"].notna() & counts["right_count"].notna(), "left_count"
        ].sum())
        predicted = int((counts["left_count"].fillna(1) * counts["right_count"].fillna(1)).sum())
        _require(predicted <= MAX_ROWS and predicted * (len(frame.columns) + len(selections)) <= MAX_CELLS,
                 "This join would produce too many rows. Filter the data or use a first-match lookup.")
        output = left.merge(right, on=temporary_keys, how=how, sort=False)
    except (TypeError, ValueError) as error:
        if isinstance(error, ValueError) and str(error).startswith("This join"):
            raise
        raise ValueError("The join keys have incompatible types. Use matching text, numeric, or date keys.") from error
    return output.drop(columns=temporary_keys), matched_rows


def _extrema(frame: pd.DataFrame, step: dict[str, Any]) -> pd.DataFrame:
    _keys(step, {"column", "which", "name"}, {"column"})
    column = _columns([step["column"]], frame)[0]
    which = step.get("which", "both")
    _require(isinstance(which, str) and which in {"highest", "lowest", "both"},
             "Choose highest, lowest, or both extremes.")
    name = _name(step.get("name", "Position"))
    _require(name not in frame.columns, "The extrema label would overwrite an existing column.")
    values = frame[column].dropna()
    try:
        highest, lowest = values.max(), values.min()
    except TypeError as error:
        raise ValueError("The highest/lowest values need comparable data types.") from error
    high = frame[column].eq(highest).fillna(False)
    low = frame[column].eq(lowest).fillna(False)
    selected = high if which == "highest" else low if which == "lowest" else high | low
    result = frame.loc[selected].copy()
    result[name] = "Highest" if which == "highest" else "Lowest" if which == "lowest" else (
        "Highest and lowest" if not values.empty and highest == lowest else
        pd.Series("Lowest", index=result.index).where(~high.loc[result.index], "Highest")
    )
    return result


def _select(frame: pd.DataFrame, selections: Any) -> pd.DataFrame:
    if not (isinstance(selections, list) and selections and all(isinstance(item, dict) for item in selections)):
        return frame[_columns(selections, frame)]
    _require(isinstance(selections, list) and 1 <= len(selections) <= MAX_COLUMNS,
             "Choose a bounded list of output columns.")
    for selection in selections:
        _require(isinstance(selection, dict) and set(selection) in ({"column"}, {"column", "name"}, {"literal", "name"}),
                 "Name each selected source or constant and output column.")
        if "column" in selection:
            _columns([selection["column"]], frame)
        else:
            _require(_literal(selection["literal"]) and not isinstance(selection["literal"], list),
                     "A selected constant must be a finite JSON scalar.")
    names = [_name(selection["name"] if "name" in selection else selection["column"]) for selection in selections]
    _require(len(set(names)) == len(names), "Selected output names must be unique.")
    _require(len(frame) * len(names) <= MAX_CELLS, "The selected output would contain too many cells.")
    output = pd.DataFrame(index=frame.index)
    for selection, name in zip(selections, names):
        output[name] = frame[selection["column"]] if "column" in selection else selection["literal"]
    return output


def _append(frame: pd.DataFrame, step: dict[str, Any], sheets: dict[str, pd.DataFrame]) -> pd.DataFrame:
    _keys(step, {"sheet", "columns"}, {"sheet", "columns"})
    sheet = _name(step["sheet"])
    _require(sheet in sheets, "An append source is unavailable. Check its name and headers.")
    source = _bounded(sheets[sheet])
    rows = len(frame) + len(source)
    _require(rows <= MAX_ROWS and rows * len(frame.columns) <= MAX_CELLS,
             "This append would produce too many rows or cells. Filter the sources first.")
    incoming = _select(source, step["columns"])
    _require(set(incoming.columns) == set(frame.columns),
             "Appended column names must match the current result columns exactly.")
    return pd.concat([frame, incoming[list(frame.columns)]], ignore_index=True)


def execute_workbook_plan(
    plan: Any, sheets: dict[str, pd.DataFrame], *, trace: BackendTrace | None = None,
    join_diagnostics: list[JoinDiagnostic] | None = None, match_null_keys: bool = True,
) -> pd.DataFrame:
    _require(isinstance(plan, dict) and set(plan) == {"understood", "unhandled_conditions", "sheet", "steps"},
             "The workbook interpretation is invalid. Please specify the columns and calculation.")
    _require(plan["understood"] is True and plan["unhandled_conditions"] == [],
             "I couldn't fully interpret this workbook question. Specify the sheets, columns, and calculation.")
    sheet = _name(plan["sheet"])
    _require(sheet in sheets, "The requested worksheet is unavailable. Check its name and headers.")
    steps = plan["steps"]
    _require(isinstance(steps, list) and 1 <= len(steps) <= MAX_STEPS,
             "A workbook question must contain between 1 and 20 calculation steps.")
    if trace:
        trace.record_code("source_selected", execute_workbook_plan, evaluate_expression, _evaluate,
                          _aggregate_frame, _aggregate, _aggregate_values, _join, _extrema, _select, _append,
                          sheet=sheet, rows=len(sheets[sheet]), columns=len(sheets[sheet].columns))
    frame = _bounded(sheets[sheet].copy(deep=True))
    frame_sources = {sheet}
    for step_number, step in enumerate(steps, start=1):
        _require(isinstance(step, dict) and isinstance(step.get("op"), str), "A workbook calculation step is invalid.")
        operation = step["op"]
        _require(operation in {"filter", "calculate", "aggregate", "join", "lookup", "append", "extrema", "select", "sort", "limit", "rank", "pivot"},
                 "That workbook operation is not supported. No calculation was performed.")
        if trace:
            trace.record("step_started", step=step_number, operation=operation, rows=len(frame),
                         columns=len(frame.columns), details=step)
        if operation == "filter":
            _keys(step, {"where"}, {"where"})
            frame = frame.loc[_truth(evaluate_expression(step["where"], frame), frame)]
        elif operation == "calculate":
            _keys(step, {"columns"}, {"columns"})
            _require(isinstance(step["columns"], list) and 1 <= len(step["columns"]) <= 20,
                     "Calculate between 1 and 20 columns per step.")
            for item in step["columns"]:
                _require(isinstance(item, dict) and set(item) == {"name", "expression"}, "Name each calculated column.")
                frame[_name(item["name"])] = _series(evaluate_expression(item["expression"], frame), frame)
                _bounded(frame)
        elif operation == "aggregate":
            _keys(step, {"group_by", "metrics"}, {"metrics"})
            frame = _aggregate_frame(frame, _columns(step.get("group_by", []), frame, empty=True), step["metrics"])
        elif operation == "extrema":
            frame = _extrema(frame, step)
        elif operation in {"join", "lookup"}:
            left_rows = len(frame)
            frame, matched_rows = _join(frame, step, sheets, match_null_keys=match_null_keys)
            diagnostic = JoinDiagnostic(
                frozenset(frame_sources), step["sheet"], tuple(step["left_on"]), tuple(step["right_on"]),
                left_rows, matched_rows,
            )
            if join_diagnostics is not None:
                join_diagnostics.append(diagnostic)
            if trace:
                trace.record("join_matched", source_id=step["sheet"], rows=left_rows,
                             matched_rows=matched_rows, unmatched_rows=left_rows - matched_rows,
                             details={"left_on": step["left_on"], "right_on": step["right_on"]})
            frame_sources.add(step["sheet"])
        elif operation == "append":
            frame = _append(frame, step, sheets)
            frame_sources.add(step["sheet"])
        elif operation == "select":
            _keys(step, {"columns"}, {"columns"})
            frame = _select(frame, step["columns"])
        elif operation == "sort":
            _keys(step, {"by"}, {"by"})
            ordering = step["by"]
            _require(isinstance(ordering, list) and 1 <= len(ordering) <= 20, "Choose a bounded sort order.")
            for item in ordering:
                _require(isinstance(item, dict) and set(item) == {"column", "direction"}
                         and item["direction"] in ("asc", "desc"), "Sort directions must be asc or desc.")
            columns = _columns([item["column"] for item in ordering], frame)
            try:
                frame = frame.sort_values(columns, ascending=[item["direction"] == "asc" for item in ordering],
                                          kind="stable", na_position="last")
            except TypeError as error:
                raise ValueError("These columns contain mixed types that cannot be sorted. Normalize them first.") from error
        elif operation == "limit":
            _keys(step, {"count"}, {"count"})
            count = step["count"]
            _require(isinstance(count, int) and not isinstance(count, bool) and 1 <= count <= MAX_ROWS,
                     "Choose a row limit between 1 and 50,000.")
            frame = frame.head(count)
        elif operation == "rank":
            _keys(step, {"column", "name", "direction", "group_by", "method"}, {"column", "name", "direction"})
            column = _columns([step["column"]], frame)[0]
            name = _name(step["name"])
            groups = _columns(step.get("group_by", []), frame, empty=True)
            method = step.get("method", "min")
            _require(step["direction"] in ("asc", "desc") and method in ("min", "max", "first", "dense", "average"),
                     "Use a supported rank method and asc or desc direction.")
            values = _numeric(frame[column], frame, pd.Series(True, index=frame.index))
            ranked = values.groupby([frame[group] for group in groups], dropna=False, sort=False) if groups else values
            frame[name] = ranked.rank(method=method, ascending=step["direction"] == "asc", na_option="keep")
        elif operation == "pivot":
            _keys(step, {"index", "columns", "values", "function", "fill_value"},
                  {"index", "columns", "values", "function"})
            index = _columns(step["index"], frame)
            column = _columns([step["columns"]], frame)[0]
            _require(column not in index, "Pivot row and column dimensions must be different.")
            function = step["function"]
            _require(isinstance(function, str) and function in AGGREGATES, "That pivot function is not supported.")
            categories = frame[column].nunique(dropna=False)
            rows = len(frame[index].drop_duplicates())
            _require(categories + len(index) <= MAX_COLUMNS and rows * (categories + len(index)) <= MAX_CELLS,
                     "The pivot would be too wide or too large. Filter its categories first.")
            temporary = next(f"_pivot_{i}" for i in range(MAX_COLUMNS + 1) if f"_pivot_{i}" not in index + [column])
            grouped = _aggregate_frame(frame, index + [column], [
                {"name": temporary, "function": function, "column": step["values"]},
            ])
            if grouped.empty:
                frame = pd.DataFrame(columns=index)
            else:
                frame = grouped.set_index(index + [column])[temporary].unstack(column)
                frame.columns = [str(value) if not pd.isna(value) else "(blank)" for value in frame.columns]
                if "fill_value" in step:
                    fill = step["fill_value"]
                    _require(fill is None or (isinstance(fill, (int, float)) and not isinstance(fill, bool) and _literal(fill)),
                             "A pivot fill value must be a finite number or null.")
                    if fill is not None:
                        frame = frame.fillna(fill)
                _require(not any(value in index for value in frame.columns),
                         "A pivot category collides with a row label. Rename the dimension first.")
                frame = frame.reset_index()
        frame = _bounded(frame)
        if trace:
            trace.record("step_completed", step=step_number, operation=operation, rows=len(frame),
                         columns=len(frame.columns), result_columns=[str(column) for column in frame.columns])
    return frame
