from collections import Counter
from contextlib import closing
from copy import copy
from datetime import date, datetime
from io import BytesIO
import math
from numbers import Real
from pathlib import Path
import posixpath
import re
from typing import Any
from zipfile import BadZipFile, ZipFile
from xml.etree import ElementTree as ET
from xml.etree.ElementTree import ParseError

from openpyxl import Workbook, load_workbook
from openpyxl.cell import WriteOnlyCell
from openpyxl.formula import Tokenizer
from openpyxl.formula.tokenizer import TokenizerError
from openpyxl.styles import Font
from openpyxl.utils.cell import coordinate_to_tuple, get_column_letter, range_boundaries
from openpyxl.utils.exceptions import InvalidFileException
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.worksheet.formula import ArrayFormula, DataTableFormula
import pandas as pd


EXCEL_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
MAX_UPLOAD_MB = 20
MAX_EXCEL_BYTES = MAX_UPLOAD_MB * 1024 * 1024
MAX_EXCEL_EXPANDED_BYTES = 1_000_000_000
MAX_EXCEL_ROWS = 1_000_000
MAX_EXCEL_COLUMNS = 100
MAX_EXCEL_CELLS = 5_000_000
FORMULA_CACHE_VERSION = 2
SOURCE_RANGE_NAME = "_WorkforceSourceData"
SALES_SHEET_MISSING = (
    "I couldn't find a sheet named 'Sales' in this workbook. "
    "Add or rename the sales worksheet to 'Sales', save the file, and upload it again."
)


def _excel_access_message(error: OSError) -> str:
    if isinstance(error, FileNotFoundError):
        return "I couldn't find the Excel file. Check the file path and try again."
    if isinstance(error, IsADirectoryError):
        return "That path points to a folder. Choose an .xlsx file instead."
    if isinstance(error, PermissionError) or getattr(error, "winerror", None) in {32, 33}:
        return (
            "I couldn't access the Excel file. It may be locked or you may not have permission. "
            "Close it in Excel or other apps, check your file permissions, and try again."
        )
    return "I couldn't access the Excel file. Check that it is available and try again."


def _excel_content(source: bytes | str | Path) -> bytes:
    if isinstance(source, bytes):
        content = source
    else:
        try:
            path = Path(source)
            if path.is_dir():
                raise IsADirectoryError()
            with path.open("rb") as workbook_file:
                content = workbook_file.read(MAX_EXCEL_BYTES + 1)
        except OSError as error:
            raise ValueError(_excel_access_message(error)) from error
        except ValueError as error:
            raise ValueError("That file path isn't valid. Choose an .xlsx file and try again.") from error
    if not content or len(content) > MAX_EXCEL_BYTES:
        raise ValueError(f"Upload a nonempty .xlsx workbook no larger than {MAX_UPLOAD_MB} MB.")
    return content


def open_excel_workbook(content: bytes | str | Path) -> Workbook:
    content = _excel_content(content)
    try:
        with ZipFile(BytesIO(content)) as archive:
            if sum(item.file_size for item in archive.infolist()) > MAX_EXCEL_EXPANDED_BYTES:
                raise ValueError(f"The expanded workbook must be no larger than {MAX_EXCEL_EXPANDED_BYTES // 1_000_000:,} MB.")
        return load_workbook(BytesIO(content), read_only=True, data_only=True)
    except OSError as error:
        raise ValueError(_excel_access_message(error)) from error
    except (BadZipFile, InvalidFileException, ParseError, KeyError) as error:
        raise ValueError("Could not read this workbook. Upload a valid, unencrypted .xlsx file.") from error


def excel_sheet_names(content: bytes | str | Path) -> list[str]:
    workbook = open_excel_workbook(content)
    try:
        names = [worksheet.title for worksheet in workbook.worksheets]
        if not names:
            raise ValueError("The workbook has no data worksheets.")
        return names
    finally:
        workbook.close()


_MAIN_NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"


def _worksheet_parts(archive: ZipFile) -> dict[str, str]:
    relationships = {
        item.attrib["Id"]: item.attrib["Target"]
        for item in ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
    }
    sheets = ET.fromstring(archive.read("xl/workbook.xml")).find(f"{_MAIN_NS}sheets")
    if sheets is None:
        return {}
    result = {}
    for sheet in sheets:
        identifier = sheet.attrib["{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"]
        target = relationships[identifier]
        path = target.lstrip("/") if target.startswith("/") else posixpath.normpath(posixpath.join("xl", target))
        if path in archive.namelist() and "/worksheets/" in path:
            result[sheet.attrib["name"]] = path
    return result


def excel_source_range(content: bytes, sheet: str) -> tuple[int, int] | None:
    with ZipFile(BytesIO(content)) as archive:
        root = ET.fromstring(archive.read("xl/workbook.xml"))
    sheets = root.find(f"{_MAIN_NS}sheets")
    names = root.find(f"{_MAIN_NS}definedNames")
    if sheets is None or names is None:
        return None
    index = next((index for index, item in enumerate(sheets) if item.get("name") == sheet), None)
    for item in names:
        if item.get("name") != SOURCE_RANGE_NAME or item.get("localSheetId") != str(index):
            continue
        match = re.fullmatch(r"'((?:[^']|'')+)'!\$A\$1:\$([A-Z]+)\$(\d+)", item.text or "")
        if match is None or match[1].replace("''", "'") != sheet:
            raise ValueError("The saved source-data range is invalid. Upload the original workbook.")
        _, _, last_column, last_row = range_boundaries(f"A1:{match[2]}{match[3]}")
        if not (1 <= last_column <= MAX_EXCEL_COLUMNS and 1 <= last_row <= MAX_EXCEL_ROWS + 1):
            raise ValueError("The saved source-data range exceeds the workbook data limits.")
        if (last_row - 1) * last_column > MAX_EXCEL_CELLS:
            raise ValueError("The saved source-data range exceeds the workbook cell limit.")
        return last_row, last_column
    return None


def _workbook_cell_values(content: bytes) -> dict[tuple[str, str], tuple[str, object]]:
    workbook = load_workbook(BytesIO(content), data_only=False)
    try:
        result = {}
        for sheet in workbook:
            for cell in sheet._cells.values():
                value = cell.value
                if value is None:
                    continue
                if isinstance(value, ArrayFormula):
                    value = ("array", value.text, value.ref)
                elif isinstance(value, DataTableFormula):
                    value = ("dataTable", tuple(sorted(dict(value).items())))
                result[(sheet.title, cell.coordinate)] = (cell.data_type, value)
        return result
    finally:
        workbook.close()


def _defined_name_values(content: bytes) -> dict[tuple[str, str], str]:
    with ZipFile(BytesIO(content)) as archive:
        root = ET.fromstring(archive.read("xl/workbook.xml"))
    worksheets = root.find(f"{_MAIN_NS}sheets")
    sheets = [sheet.attrib["name"] for sheet in worksheets] if worksheets is not None else []
    definitions = root.find(f"{_MAIN_NS}definedNames")
    if definitions is None:
        return {}
    return {
        (
            sheets[int(item.attrib["localSheetId"])].casefold() if "localSheetId" in item.attrib else "",
            item.attrib["name"].casefold(),
        ): item.text or ""
        for item in definitions
    }


def _formula_text(value: object) -> str | None:
    if isinstance(value, str):
        return value
    if isinstance(value, tuple) and len(value) == 3 and value[0] == "array" and isinstance(value[1], str):
        return value[1]
    return None


def _formula_affected(
    formula: str, sheet: str, changed: set[tuple[str, str]], names: dict[tuple[str, str], str],
    changed_names: set[tuple[str, str]], seen: frozenset[tuple[str, str]] = frozenset(),
) -> bool:
    if not changed and not changed_names:
        return False
    try:
        tokens = Tokenizer(formula).items
        if any(token.type == "FUNC" and token.value.upper().removeprefix("_XLFN.") in {
            "INDIRECT(", "OFFSET(",
        } for token in tokens):
            return True
        for token in tokens:
            if token.type != "OPERAND" or token.subtype != "RANGE":
                continue
            reference = token.value
            source, area = reference.rsplit("!", 1) if "!" in reference else (sheet, reference)
            if source.startswith("'") and source.endswith("'"):
                source = source[1:-1].replace("''", "'")
            if any(character in source for character in "[]:"):
                return True
            if not re.fullmatch(
                r"\$?[A-Za-z]{1,3}\$?\d+(?::\$?[A-Za-z]{1,3}\$?\d+)?|"
                r"\$?[A-Za-z]{1,3}:\$?[A-Za-z]{1,3}|\$?\d+:\$?\d+", area,
            ):
                candidates = [(source.casefold(), area.casefold()), ("", area.casefold())]
                key = next((key for key in candidates if key in names or key in changed_names), None)
                if key is None or key in changed_names or key in seen or len(seen) >= 20:
                    return True
                if _formula_affected("=" + names[key].lstrip("="), source, changed, names, changed_names, seen | {key}):
                    return True
                continue
            left, top, right, bottom = range_boundaries(area)
            for changed_sheet, address in changed:
                if changed_sheet.casefold() != source.casefold():
                    continue
                row, column = coordinate_to_tuple(address)
                if (left or 1) <= column <= (right or 16384) and (top or 1) <= row <= (bottom or 1048576):
                    return True
    except (ValueError, TypeError, TokenizerError):
        return True
    return False


def preserve_formula_values(original: bytes, updated: bytes, *, reject_changed: bool = False) -> bytes:
    """Restore Excel's saved results only for formulas whose inputs are unchanged."""
    old_values = _workbook_cell_values(original)
    formulas = {key: value for key, (kind, value) in old_values.items() if kind == "f"}
    if not formulas:
        return updated
    new_values = _workbook_cell_values(updated)
    changed = {key for key in old_values.keys() | new_values.keys() if old_values.get(key) != new_values.get(key)}
    current_formulas = {key: value for key, (kind, value) in new_values.items() if kind == "f"}
    old_names = _defined_name_values(original)
    names = _defined_name_values(updated)
    changed_names = {key for key in old_names.keys() | names.keys() if old_names.get(key) != names.get(key)}
    while changed or changed_names:
        affected = set()
        for key, formula in current_formulas.items():
            if key in changed:
                continue
            expression = _formula_text(formula)
            if expression is None or _formula_affected(expression, key[0], changed, names, changed_names):
                affected.add(key)
        if not affected:
            break
        changed.update(affected)
    if reject_changed and any(key in changed for key in current_formulas):
        raise ValueError(
            "This edit changes inputs used by workbook formulas or has dependencies that cannot be verified "
            f"at {', '.join(f'{sheet}!{cell}' for sheet, cell in sorted(current_formulas.keys() & changed)[:5])}. "
            "Recalculate a copy in Excel before uploading it; "
            "the chatbot will not reuse stale formula results. No workbook changes were applied."
        )
    replacements = {}
    with ZipFile(BytesIO(original)) as source, ZipFile(BytesIO(updated)) as destination:
        old_parts = _worksheet_parts(source)
        new_parts = _worksheet_parts(destination)
        for sheet, part in new_parts.items():
            if sheet not in old_parts:
                continue
            cached = {
                cell.attrib["r"]: cell
                for cell in ET.fromstring(source.read(old_parts[sheet])).iter(f"{_MAIN_NS}c")
                if cell.find(f"{_MAIN_NS}f") is not None
            }
            if not cached:
                continue
            root = ET.fromstring(destination.read(part))
            restored = False
            for cell in root.iter(f"{_MAIN_NS}c"):
                address = cell.attrib["r"]
                key = (sheet, address)
                if key in changed or key not in formulas or key not in current_formulas or address not in cached:
                    continue
                saved = cached[address]
                value = saved.find(f"{_MAIN_NS}v")
                if value is None or (value.text is None and saved.get("t") != "str"):
                    continue
                current = cell.find(f"{_MAIN_NS}v")
                if current is not None:
                    cell.remove(current)
                cell.append(ET.fromstring(ET.tostring(value)))
                if "t" in saved.attrib:
                    cell.set("t", saved.attrib["t"])
                else:
                    cell.attrib.pop("t", None)
                restored = True
            if restored:
                replacements[part] = ET.tostring(root, encoding="utf-8", xml_declaration=True)
        if not replacements:
            return updated
        output = BytesIO()
        with ZipFile(output, "w") as archive:
            for item in destination.infolist():
                archive.writestr(item, replacements.get(item.filename, destination.read(item.filename)))
    return output.getvalue()


def _check_formula_values(content: bytes, sheet: str, last_row: int, last_column: int | None = None) -> None:
    with ZipFile(BytesIO(content)) as archive:
        part = _worksheet_parts(archive).get(sheet)
        if part is None:
            return
        missing = []
        with archive.open(part) as document:
            sheet_data = None
            for event, element in ET.iterparse(document, events=("start", "end")):
                if event == "start" and element.tag == f"{_MAIN_NS}sheetData":
                    sheet_data = element
                if event != "end" or element.tag != f"{_MAIN_NS}row":
                    continue
                for cell in element.iter(f"{_MAIN_NS}c"):
                    row, column = coordinate_to_tuple(cell.attrib["r"])
                    if cell.find(f"{_MAIN_NS}f") is None or row > last_row or (last_column is not None and column > last_column):
                        continue
                    value = cell.find(f"{_MAIN_NS}v")
                    if (value is None or (value.text is None and cell.get("t") != "str")) and len(missing) < 5:
                        missing.append(cell.attrib["r"])
                element.clear()
                if sheet_data is not None:
                    sheet_data.clear()
        if missing:
            raise ValueError(
                f"Worksheet '{sheet}' has formulas without saved calculation results at {', '.join(missing[:5])}. "
                "These are not empty source cells. Open the original workbook in Excel, recalculate all formulas "
                "(Ctrl+Alt+F9), save it, and upload the saved file again."
            )


def read_excel_sheet(
    content: bytes | str | Path, sheet: str, *, preview_rows: int | None = None,
) -> pd.DataFrame:
    if preview_rows is not None and (type(preview_rows) is not int or not 1 <= preview_rows <= MAX_EXCEL_ROWS):
        raise ValueError(f"Choose a preview size between 1 and {MAX_EXCEL_ROWS:,} rows.")
    content = _excel_content(content)
    workbook = open_excel_workbook(content)
    try:
        if sheet not in [worksheet.title for worksheet in workbook.worksheets]:
            raise ValueError(
                SALES_SHEET_MISSING if sheet.casefold() == "sales"
                else "I couldn't find the selected worksheet. Choose another worksheet and try again."
            )
        worksheet = workbook[sheet]
        worksheet.reset_dimensions()
        source_range = excel_source_range(content, sheet)
        rows = worksheet.iter_rows(
            values_only=True, max_row=source_range[0] if source_range else None,
            max_col=source_range[1] if source_range else None,
        )
        header = list(next(rows, ()))
        while header and header[-1] is None:
            header.pop()
        if not header or len(header) > MAX_EXCEL_COLUMNS:
            raise ValueError(f"Worksheet '{sheet}': use a header row with 1 to {MAX_EXCEL_COLUMNS} columns in row 1.")
        if any(not isinstance(value, str) or not value.strip() for value in header):
            cells = [f"{get_column_letter(index + 1)}1" for index, value in enumerate(header)
                     if not isinstance(value, str) or not value.strip()]
            raise ValueError(
                f"Worksheet '{sheet}': every column needs a nonempty text header in the first row. "
                f"Check {', '.join(cells)}. Put the actual column headings in row 1, save, and upload again."
            )
        columns = [value.strip() for value in header]
        counts = Counter(value.casefold() for value in columns)
        if len(counts) != len(columns):
            cells = [f"{get_column_letter(index + 1)}1" for index, value in enumerate(columns)
                     if counts[value.casefold()] > 1]
            raise ValueError(
                f"Worksheet '{sheet}': column headers must be unique (ignoring case and surrounding spaces). "
                f"Repeated headers are at {', '.join(cells)}. Give them distinct names, save, and upload again."
            )
        values = []
        populated_rows = 0
        for row_number, row in enumerate(rows, start=2):
            if len(row) > len(columns) and any(value is not None for value in row[len(columns):]):
                column = next(index for index, value in enumerate(row, start=1) if index > len(columns) and value is not None)
                raise ValueError(
                    f"Worksheet '{sheet}': some cells have no column header, for example "
                    f"{get_column_letter(column)}{row_number}. Add a distinct text header in row 1, save, and upload again."
                )
            if len(values) >= MAX_EXCEL_ROWS:
                raise ValueError(f"A worksheet can contain at most {MAX_EXCEL_ROWS:,} data rows.")
            if (len(values) + 1) * len(columns) > MAX_EXCEL_CELLS:
                raise ValueError(f"A worksheet can contain at most {MAX_EXCEL_CELLS:,} data cells. Reduce its rows or columns.")
            values.append(tuple(row[:len(columns)]) + (None,) * max(0, len(columns) - len(row)))
            if any(value is not None for value in row):
                populated_rows += 1
            if preview_rows is not None and populated_rows >= preview_rows:
                break
        _check_formula_values(content, sheet, len(values) + 1, source_range[1] if source_range else None)
        populated = pd.DataFrame(values, columns=columns).dropna(how="all")
        frame = populated.reset_index(drop=True)
        frame.attrs["excel_row_numbers"] = [int(index) + 2 for index in populated.index]
        return frame
    except OSError as error:
        raise ValueError(_excel_access_message(error)) from error
    finally:
        workbook.close()


def spreadsheet_schema(frame: pd.DataFrame) -> tuple[pd.DataFrame, dict[str, str], dict[str, str]]:
    data = frame.copy()
    names = {f"c{index}": str(column) for index, column in enumerate(frame.columns)}
    data.columns = list(names)
    schema = {}
    for column in data:
        series = data[column]
        populated = series.dropna()
        if pd.api.types.is_numeric_dtype(series) and not pd.api.types.is_bool_dtype(series):
            schema[column] = "number"
        elif not populated.empty and all(isinstance(value, (date, datetime)) for value in populated):
            schema[column] = "date"
            data[column] = series.map(lambda value: None if pd.isna(value) else value.date().isoformat()
                                      if isinstance(value, datetime) else value.isoformat())
        else:
            schema[column] = "text"
            data[column] = series.map(lambda value: None if pd.isna(value) else str(value))
    return data, schema, names


def _sales_column(frame: pd.DataFrame, metric: str = "sales") -> str | None:
    headers = {str(column): re.sub(r"[_\s]+", " ", str(column)).strip().casefold() for column in frame}
    candidates = [column for column, header in headers.items() if header == metric]
    if not candidates:
        candidates = [column for column, header in headers.items()
                      if header in {"sales", "total sales", "sales amount", "revenue"}]
    return candidates[0] if len(candidates) == 1 else None


def _sales_values(frame: pd.DataFrame, column: str) -> pd.Series:
    series = frame[column]
    if not series.dropna().empty and (
        not pd.api.types.is_numeric_dtype(series) or pd.api.types.is_bool_dtype(series)
    ):
        raise ValueError(f"The {column} column must contain numeric values to calculate total sales.")
    return series


def region_sales_totals(frame: pd.DataFrame, metric: str = "sales") -> pd.DataFrame:
    column = _sales_column(frame, metric)
    regions = [name for name in frame if str(name).strip().casefold() == "region"]
    if column is None or len(regions) != 1:
        raise ValueError("The Sales sheet needs a Region column and one unambiguous sales column.")
    values = _sales_values(frame, column)
    labels = frame[regions[0]].astype("string").str.strip()
    labels = labels.mask(labels.eq(""), pd.NA)
    keys = labels.str.casefold()
    totals = values.groupby(keys, dropna=False, sort=False).sum(min_count=1)
    display_labels = labels.groupby(keys, dropna=False, sort=False).first()
    return pd.DataFrame({"Region": display_labels.to_numpy(), "Total Sales": totals.to_numpy()})


def regional_sales_plan(question: str, schemas: dict[str, dict[str, str]]) -> dict[str, Any] | None:
    match = re.fullmatch(
        r"(?:which|what)\s+regions?\s+(?:had|has|have)\s+(?:the\s+)?"
        r"(?P<extreme>highest|lowest)\s+(?:total\s+)?sales"
        r"(?P<count>,?\s+and\s+how\s+many\s+(?:sales|transactions)"
        r"\s+(?:did|does|do)\s+(?:it|they)\s+have)?",
        question.strip().rstrip("?.!").strip(), flags=re.IGNORECASE,
    )
    all_regions = re.fullmatch(
        r"(?:(?:give(?: me)?|show(?: me)?|calculate|what (?:is|are))\s+)?(?:the\s+)?"
        r"(?:total|sum(?: of)?)\s+(?:sales amount|sales|revenue)\s+"
        r"(?:for\s+(?:all|each)\s+regions?|by\s+regions?|per\s+regions?)",
        question.strip().rstrip("?.!").strip(), flags=re.IGNORECASE,
    )
    if match is None and all_regions is None:
        return None
    candidates = []
    for sheet, schema in schemas.items():
        regions = [name for name in schema if name.strip().casefold() == "region"]
        headers = {name: re.sub(r"[_\s]+", " ", name).strip().casefold() for name in schema}
        amounts = [name for name, header in headers.items()
                   if header in {"amount", "sales amount", "salesamount", "revenue"}]
        if not amounts:
            amounts = [name for name, header in headers.items() if header == "sales"]
        if len(regions) == 1 and amounts:
            candidates.append((sheet, regions[0], amounts))
    if len(candidates) != 1 or len(candidates[0][2]) != 1:
        return None
    sheet, region, amounts = candidates[0]
    amount = amounts[0]
    metrics = [{"name": "Total Sales", "function": "sum", "column": amount}]
    columns = [region, "Total Sales"]
    if match is not None and match["count"]:
        metrics.append({"name": "Sales Transactions", "function": "count_rows", "column": "*"})
        columns.append("Sales Transactions")
    steps = [{"op": "aggregate", "group_by": [region], "metrics": metrics}]
    if match is not None:
        steps.append({"op": "extrema", "column": "Total Sales", "which": match["extreme"].casefold()})
    steps.append({"op": "select", "columns": columns})
    return {"understood": True, "unhandled_conditions": [], "sheet": sheet, "steps": steps}


def query_sales_total(question: str, frame: pd.DataFrame) -> pd.DataFrame | None:
    match = re.fullmatch(
        r"(?:(?:what (?:are|is)|show(?: me)?|give(?: me)?)\s+)?(?:the\s+)?"
        r"(?:total|sum(?: of)?)\s+(?P<metric>sales amount|total sales|sales|revenue)"
        r"(?:\s+(?:(?P<grouped>(?:by|per|for each)\s+region)|"
        r"(?:for|in)\s+(?:the\s+)?(?P<region>.+?)\s+region))?",
        question.strip().rstrip("?.!").strip(), flags=re.IGNORECASE,
    )
    if match is None:
        return None
    headers = {column: re.sub(r"[_\s]+", " ", str(column)).strip().casefold() for column in frame}
    metric = match["metric"].casefold()
    if match["grouped"]:
        return region_sales_totals(frame, metric)
    column = _sales_column(frame, metric)
    if column is None:
        return None
    selected = frame
    scope = ""
    if match["region"]:
        regions = [column for column, header in headers.items() if header == "region"]
        if len(regions) != 1:
            return None
        region_column = regions[0]
        region = match["region"].strip()
        values = frame[region_column].astype("string").str.strip()
        mask = values.str.casefold().eq(region.casefold()).fillna(False)
        if not mask.any():
            return None
        selected = frame.loc[mask]
        region_label = str(values.loc[mask].iloc[0])
        scope = f" for the {region_label} region"
    series = _sales_values(selected, column)
    result = pd.DataFrame({f"sum {column}": [series.sum(min_count=1)]})
    result.attrs["answer_subject"] = f"The total {metric.removeprefix('total ')}{scope}"
    return result


def write_region_summary(
    content: bytes | str | Path, *, confirmed: bool = False,
) -> tuple[bytes, bytes]:
    """Return an original backup and an updated copy; never modify the source."""
    if confirmed is not True:
        raise ValueError("Please confirm before writing Region_Summary.")
    backup = memoryview(_excel_content(content)).tobytes()
    names = excel_sheet_names(backup)
    if any(name.casefold() == "region_summary" for name in names):
        raise ValueError("Region_Summary already exists. No existing sheet will be overwritten.")
    source = next((name for name in names if name.casefold() == "sales"), None)
    if source is None:
        raise ValueError(SALES_SHEET_MISSING)
    frame = read_excel_sheet(backup, source)
    summary = region_sales_totals(frame)
    try:
        workbook = load_workbook(BytesIO(backup), data_only=False)
    except OSError as error:
        raise ValueError(_excel_access_message(error)) from error
    try:
        if any(name.casefold() == "region_summary" for name in workbook.sheetnames):
            raise ValueError("Region_Summary already exists. No existing sheet will be overwritten.")
        worksheet = workbook.create_sheet("Region_Summary")
        worksheet.append(["Region", "Total Sales"])
        for label, total in summary.itertuples(index=False, name=None):
            worksheet.append([
                None if pd.isna(label) else str(label),
                None if pd.isna(total) else total,
            ])
        worksheet.freeze_panes = "A2"
        worksheet.auto_filter.ref = worksheet.dimensions
        _literal_excel_text(worksheet)
        output = BytesIO()
        workbook.save(output)
        return backup, preserve_formula_values(backup, output.getvalue(), reject_changed=True)
    except OSError as error:
        raise ValueError(_excel_access_message(error)) from error
    finally:
        workbook.close()


def _literal_excel_text(worksheet: Worksheet) -> None:
    for row in worksheet:
        for cell in row:
            if isinstance(cell.value, str):
                cell.data_type = "s"
                cell.hyperlink = None


def is_calendar_year_column(column: object) -> bool:
    return re.sub(r"[\s_-]+", "", str(column)).casefold() in {
        "year", "calendaryear", "fiscalyear", "financialyear", "salesyear", "saleyear",
        "orderyear", "transactionyear", "birthyear", "joiningyear", "joinyear", "hireyear",
    }


def date_derivation_description(result: pd.DataFrame) -> str:
    descriptions = []
    for name, derivation in result.attrs.get("date_derivations", {}).items():
        if name not in result:
            continue
        function, column = derivation["function"], derivation["column"]
        description = (
            f"{name} is derived from the {column} values using {function.upper()}({column}); "
            "it is not a separate column read from the source."
        )
        if function == "year":
            years = pd.to_numeric(result[name], errors="coerce").dropna().unique()
            if len(years):
                labels = ", ".join(f"{int(value):04d}" for value in years[:10])
                description += f" Years in this result: {labels}" + (" (first 10 shown)." if len(years) > 10 else ".")
        descriptions.append(description)
    return "\n\n".join(descriptions)


def spreadsheet_answer(result: pd.DataFrame) -> str:
    def value_text(value: object, column: object) -> str:
        if pd.isna(value):
            return "not available"
        calendar_year = is_calendar_year_column(column) or (
            result.attrs.get("date_derivations", {}).get(column, {}).get("function") == "year"
        )
        if calendar_year and isinstance(value, str) and re.fullmatch(r"\d{1,4}(?:\.0+)?", value.strip()):
            return f"{int(float(value)):04d}"
        if isinstance(value, Real) and not isinstance(value, bool):
            if math.isfinite(value) and value == int(value):
                if calendar_year:
                    return f"{int(value):04d}"
                return f"{int(value):,}"
            return format(value, ",")
        if isinstance(value, (date, datetime)):
            return value.isoformat()
        return str(value)

    def label_text(column: object) -> str:
        label = str(column)
        for prefix, description in (
            ("sum ", "total "), ("average ", "average "), ("min ", "minimum "),
            ("max ", "maximum "), ("count distinct ", "distinct count of "), ("count ", "count of "),
        ):
            if label.startswith(prefix):
                return description + label[len(prefix):]
        return label.replace("_", " ")

    if result.empty:
        return "No matching rows were found in this worksheet."
    if result.shape == (1, 1):
        value = result.iloc[0, 0]
        if str(result.columns[0]) == "row_count":
            return f"There {'is' if value == 1 else 'are'} {value_text(value, result.columns[0])} {'row' if value == 1 else 'rows'}."
        subject = result.attrs.get("answer_subject") or f"The {label_text(result.columns[0])}"
        if pd.isna(value):
            return f"No value is available for {subject.removeprefix('The ')}."
        answer = f"{subject} is {value_text(value, result.columns[0])}."
        description = date_derivation_description(result)
        return answer + (f"\n\n{description}" if description else "")
    if len(result.columns) > 6:
        answer = f"Found {len(result):,} matching {'row' if len(result) == 1 else 'rows'} across {len(result.columns):,} columns. See the data table below."
        description = date_derivation_description(result)
        return answer + (f"\n\n{description}" if description else "")
    lines = [
        "; ".join(f"{label_text(column)}: {value_text(value, column)}" for column, value in zip(result.columns, row)) + "."
        for row in result.itertuples(index=False, name=None)
    ]
    answer = lines[0] if len(lines) == 1 else (
        f"Here are the {len(result):,} matching results:\n\n" + "\n".join(f"- {line}" for line in lines)
    )
    description = date_derivation_description(result)
    return answer + (f"\n\n{description}" if description else "")


def generate_excel_download(tables: list[tuple[str, pd.DataFrame]]) -> bytes:
    output = BytesIO()
    used_names = set()
    workbook = Workbook(write_only=True)
    try:
        for title, table in tables:
            base = re.sub(r"[\[\]:*?/\\]", "_", title).strip("'")[:31] or "Sheet"
            name = base
            suffix = 2
            while name.casefold() in used_names:
                tail = f" ({suffix})"
                name = base[:31 - len(tail)] + tail
                suffix += 1
            used_names.add(name.casefold())
            worksheet = workbook.create_sheet(name)
            worksheet.freeze_panes = "A2"
            worksheet.auto_filter.ref = f"A1:{get_column_letter(max(1, len(table.columns)))}{len(table) + 1}"
            headers = []
            for column in table.columns:
                cell = WriteOnlyCell(worksheet, value=str(column))
                cell.data_type = "s"
                cell.font = Font(bold=True)
                headers.append(cell)
            worksheet.append(headers)
            for row in table.itertuples(index=False, name=None):
                cells = []
                for value in row:
                    if isinstance(value, (list, dict)):
                        value = str(value)
                    if pd.isna(value):
                        value = None
                    elif isinstance(value, datetime) and value.tzinfo is not None:
                        value = value.replace(tzinfo=None)
                    elif isinstance(value, Real) and not math.isfinite(value):
                        value = str(value)
                    cell = WriteOnlyCell(worksheet, value=value)
                    if isinstance(value, str):
                        cell.data_type = "s"
                    cells.append(cell)
                worksheet.append(cells)
        if not tables:
            raise ValueError("Choose at least one result table to download.")
        workbook.save(output)
    finally:
        workbook.close()
    return output.getvalue()


def append_result_sheet(
    content: bytes, result: pd.DataFrame, title: str, *, after_sheet: str,
) -> tuple[bytes, str]:
    if not 1 <= len(result.columns) <= MAX_EXCEL_COLUMNS or len(result) > MAX_EXCEL_ROWS or result.size > MAX_EXCEL_CELLS:
        raise ValueError(
            f"A saved result needs 1 to {MAX_EXCEL_COLUMNS} columns, at most {MAX_EXCEL_ROWS:,} rows "
            f"and {MAX_EXCEL_CELLS:,} cells."
        )
    excel_sheet_names(content)
    with closing(load_workbook(BytesIO(content), data_only=False)) as workbook, closing(
        load_workbook(BytesIO(generate_excel_download([("Results", result)])))
    ) as exported:
        if after_sheet not in workbook.sheetnames:
            raise ValueError("The source worksheet is unavailable. No worksheet was added.")
        base = re.sub(r"[\[\]:*?/\\]", "_", title).strip().strip("'")[:31] or "Results"
        name = base
        suffix = 1
        existing = {sheet.casefold() for sheet in workbook.sheetnames}
        while name.casefold() in existing:
            tail = f"_{suffix}"
            name = base[:31 - len(tail)] + tail
            suffix += 1
        worksheet = workbook.create_sheet(name, workbook.sheetnames.index(after_sheet) + 1)
        for row in exported.worksheets[0]:
            for source in row:
                cell = worksheet.cell(source.row, source.column, source.value)
                cell.data_type = source.data_type
                cell.font = copy(source.font)
                cell.number_format = source.number_format
        worksheet.freeze_panes = "A2"
        worksheet.auto_filter.ref = worksheet.dimensions
        for index, column in enumerate(result.columns, 1):
            worksheet.column_dimensions[get_column_letter(index)].width = min(40, max(14, len(str(column)) + 3))
        output = BytesIO()
        workbook.save(output)
        updated = preserve_formula_values(content, output.getvalue(), reject_changed=True)
        excel_sheet_names(updated)
        return updated, name
