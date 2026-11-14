"""Apply edits from the in-app spreadsheet grid to the workbook working copy."""

from __future__ import annotations

from datetime import date, datetime
from io import BytesIO
from typing import Any

import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

from excel_support import MAX_EXCEL_ROWS, excel_sheet_names, excel_source_range, preserve_formula_values
from workbook_edits import _formula_dependencies

MAX_GRID_ROWS = 5000


def _value(raw: Any, template: pd.Series) -> Any:
    if raw is None or (isinstance(raw, float) and pd.isna(raw)):
        return None
    if isinstance(raw, str):
        text = raw.strip()
        if text == "":
            return None
        is_date = pd.api.types.is_datetime64_any_dtype(template) or any(
            isinstance(value, (date, datetime)) for value in template.dropna().head(20)
        )
        if is_date:
            try:
                return pd.Timestamp(text).to_pydatetime()
            except ValueError:
                return raw
    return raw


def grid_change_count(changes: dict[str, Any] | None) -> int:
    if not changes:
        return 0
    return (
        sum(len(columns) for columns in changes.get("edited_rows", {}).values())
        + len(changes.get("added_rows", [])) + len(changes.get("deleted_rows", []))
    )


def apply_grid_changes(content: bytes, sheet: str, frame: pd.DataFrame, changes: dict[str, Any]) -> bytes:
    """Write st.data_editor deltas to the physical worksheet rows they came from."""
    edited = {int(position): values for position, values in (changes.get("edited_rows") or {}).items()}
    added = list(changes.get("added_rows") or [])
    deleted = sorted({int(position) for position in changes.get("deleted_rows") or []})
    if not (edited or added or deleted):
        raise ValueError("There are no grid changes to apply.")
    if excel_source_range(content, sheet) is not None:
        raise ValueError(
            "This worksheet contains same-sheet charts. Edit the original data or a copy without charts; "
            "no changes were applied."
        )
    row_numbers = frame.attrs.get("excel_row_numbers", list(range(2, len(frame) + 2)))
    columns = [str(column) for column in frame.columns]
    workbook = load_workbook(BytesIO(content), data_only=False)
    try:
        worksheet = workbook[sheet]
        for index, name in enumerate(columns, 1):
            header = worksheet.cell(1, index).value
            if not isinstance(header, str) or header.strip() != name:
                raise ValueError("The worksheet headers changed. Reload the grid and try again.")
        for position, values in edited.items():
            if not 0 <= position < len(row_numbers):
                raise ValueError("An edited row no longer exists. Reload the grid and try again.")
            for name, raw in values.items():
                if name not in columns:
                    raise ValueError(f"Column {name} is not in this worksheet.")
                cell = worksheet.cell(row_numbers[position], columns.index(name) + 1)
                if cell.data_type == "f":
                    raise ValueError(
                        f"Cell {get_column_letter(cell.column)}{cell.row} contains a formula. "
                        "Edit formulas in Excel; no changes were applied."
                    )
                cell.value = _value(raw, frame[name])
        if added:
            if len(frame) - len(deleted) + len(added) > MAX_EXCEL_ROWS:
                raise ValueError(f"A worksheet can have at most {MAX_EXCEL_ROWS:,} data rows.")
            next_row = max([worksheet.max_row, *row_numbers, 1]) + 1
            for offset, values in enumerate(added):
                for name, raw in values.items():
                    if name in columns:
                        worksheet.cell(next_row + offset, columns.index(name) + 1, _value(raw, frame[name]))
        if deleted:
            if _formula_dependencies(workbook, sheet) or worksheet.merged_cells.ranges or worksheet.tables:
                raise ValueError(
                    "Deleting rows could break formulas, merged cells or Excel tables on this sheet. "
                    "Clear the cells instead; no changes were applied."
                )
            for position in reversed(deleted):
                if not 0 <= position < len(row_numbers):
                    raise ValueError("A deleted row no longer exists. Reload the grid and try again.")
                worksheet.delete_rows(row_numbers[position])
        buffer = BytesIO()
        workbook.save(buffer)
    finally:
        workbook.close()
    updated = preserve_formula_values(content, buffer.getvalue())
    if sheet not in excel_sheet_names(updated):
        raise ValueError("The edited workbook could not be verified. No changes were applied.")
    return updated
