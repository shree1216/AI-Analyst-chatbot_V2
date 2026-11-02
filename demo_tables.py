"""Bundled demo tables: one CSV file per table plus ``schema.json`` describing column kinds.

``sync_local_database.py`` writes these files from the local EmployeeDB; demo mode loads them into SQLite.
"""
from dataclasses import dataclass
from datetime import date, datetime, time
from decimal import Decimal
from hashlib import sha256
import csv
import io
import json
from pathlib import Path
from typing import Any, Iterable

import pandas as pd


DEMO_DATA_DIR = Path(__file__).with_name("demo_data")
SCHEMA_FILE = "schema.json"
COLUMN_KINDS = {"text", "integer", "decimal", "boolean", "date", "datetime"}
REQUIRED_EMPLOYEE_COLUMNS = (
    "Employee ID", "Name", "Age", "Gender", "Department", "City", "Status", "Job Title",
    "Salary", "Joining Date", "Email", "Phone Number",
)


@dataclass(frozen=True)
class DemoTable:
    name: str
    columns: tuple[tuple[str, str], ...]
    rows: tuple[tuple[Any, ...], ...]


def demo_data_version(directory: Path = DEMO_DATA_DIR) -> str:
    return sha256((directory / SCHEMA_FILE).read_bytes()).hexdigest()


def format_value(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, bool):
        return "true" if value else "false"
    if isinstance(value, datetime):
        return value.date().isoformat() if value.time() == time() and value.tzinfo is None else value.isoformat(sep=" ")
    if isinstance(value, date):
        return value.isoformat()
    if isinstance(value, Decimal):
        return format(value, "f")
    if isinstance(value, (bytes, bytearray, memoryview)):
        return bytes(value).hex()
    return str(value)


def _table_csv(table: DemoTable) -> str:
    output = io.StringIO()
    writer = csv.writer(output, lineterminator="\n")
    writer.writerow([name for name, _ in table.columns])
    writer.writerows([format_value(value) for value in row] for row in table.rows)
    return output.getvalue()


def write_demo_tables(tables: Iterable[DemoTable], directory: Path = DEMO_DATA_DIR) -> bool:
    """Write tables and schema; return True when any file changed. Files of removed tables are deleted."""
    tables = sorted(tables, key=lambda table: table.name.casefold())
    files: dict[str, str] = {}
    schema: dict[str, Any] = {"tables": []}
    for table in tables:
        invalid = [kind for _, kind in table.columns if kind not in COLUMN_KINDS]
        if invalid:
            raise ValueError(f"Unsupported demo column kind for {table.name}: {invalid[0]}")
        content = _table_csv(table)
        file_name = f"{table.name}.csv"
        files[file_name] = content
        schema["tables"].append({
            "name": table.name, "file": file_name,
            "columns": [{"name": name, "kind": kind} for name, kind in table.columns],
            "rows": len(table.rows), "sha256": sha256(content.encode("utf-8")).hexdigest(),
        })
    files[SCHEMA_FILE] = json.dumps(schema, indent=2) + "\n"
    directory.mkdir(parents=True, exist_ok=True)
    changed = False
    for file_name, content in files.items():
        path = directory / file_name
        if not path.exists() or path.read_text(encoding="utf-8") != content:
            path.write_text(content, encoding="utf-8", newline="\n")
            changed = True
    for path in directory.glob("*.csv"):
        if path.name not in files:
            path.unlink()
            changed = True
    return changed


def load_demo_tables(directory: Path = DEMO_DATA_DIR) -> dict[str, pd.DataFrame]:
    schema = json.loads((directory / SCHEMA_FILE).read_text(encoding="utf-8"))
    tables: dict[str, pd.DataFrame] = {}
    for table in schema["tables"]:
        columns = [(column["name"], column["kind"]) for column in table["columns"]]
        frame = pd.read_csv(
            directory / table["file"], dtype=str, keep_default_na=False, na_values=[""],
        )
        if list(frame.columns) != [name for name, _ in columns]:
            raise ValueError(f"Demo table {table['name']} does not match its schema.")
        for name, kind in columns:
            if kind == "integer":
                frame[name] = pd.to_numeric(frame[name], errors="raise").astype("Int64")
            elif kind == "decimal":
                frame[name] = pd.to_numeric(frame[name], errors="raise").astype("float64")
            elif kind == "boolean":
                frame[name] = frame[name].str.lower().map({"true": True, "false": False}).astype("boolean")
            else:
                frame[name] = frame[name].astype(object).where(frame[name].notna(), None)
        tables[table["name"]] = frame
    return tables
