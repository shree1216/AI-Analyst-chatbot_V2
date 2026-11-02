"""Mirror the local EmployeeDB to Azure SQL and to the bundled demo data used by Streamlit Cloud.

Each run reads every user table in the local database and fingerprints its schema and rows.
Changed tables are recreated in Azure SQL in one transaction (cloud-only tables are never touched),
and the demo CSV files are regenerated and pushed to GitHub from a dedicated clone so the
working copy is never modified. Unchanged runs make no cloud or GitHub requests.

    python sync_local_database.py                  # sync once
    python sync_local_database.py --force          # resync every table
    python sync_local_database.py --export-demo demo_data   # only write demo files locally
    python sync_local_database.py --install-task --interval 1
    python sync_local_database.py --uninstall-task

Settings are read from the environment or the git-ignored .env file:
    SYNC_CLOUD_SQL_SERVER       Azure SQL host, e.g. myserver.database.windows.net (required for cloud sync)
    SYNC_CLOUD_SQL_DATABASE     default EmployeeDB
    SYNC_CLOUD_RESOURCE_GROUP   optional; lets the job add this PC's new public IP to the SQL firewall
    SYNC_CLOUD_READER_USER      optional database user granted SELECT on synced tables
    SYNC_LOCAL_SQL_SERVER       default .\\SQLEXPRESS (Windows authentication)
    SYNC_LOCAL_SQL_DATABASE     default EmployeeDB
    SYNC_GIT_BRANCH             default main
Azure SQL access uses the signed-in Azure CLI account (Microsoft Entra authentication).
"""
from __future__ import annotations

import argparse
from contextlib import contextmanager
from dataclasses import dataclass
from hashlib import sha256
import json
import logging
from logging.handlers import RotatingFileHandler
import os
from pathlib import Path
import re
import shutil
import struct
import subprocess
import sys
import time
from typing import Any, Iterator

from dotenv import load_dotenv
import pyodbc

from demo_tables import REQUIRED_EMPLOYEE_COLUMNS, DemoTable, format_value, write_demo_tables


ROOT = Path(__file__).resolve().parent
STATE_DIR = Path(os.getenv("LOCALAPPDATA") or Path.home() / ".local" / "share") / "WorkforceCompass"
TASK_NAME = "WorkforceCompass Database Sync"
ODBC_DRIVER = "ODBC Driver 18 for SQL Server"
SQL_COPT_SS_ACCESS_TOKEN = 1256
MAX_DEMO_ROWS = 50_000
SYNC_COMMIT_MESSAGE = "Sync demo data from local EmployeeDB"
LENGTH_TYPES = {"varchar", "char", "varbinary", "binary"}
UNICODE_LENGTH_TYPES = {"nvarchar", "nchar"}
PLAIN_TYPES = {
    "int", "bigint", "smallint", "tinyint", "bit", "money", "smallmoney", "real", "float",
    "date", "datetime", "smalldatetime", "uniqueidentifier",
}
REPLACEMENT_TYPES = {
    "text": "varchar(max)", "ntext": "nvarchar(max)", "image": "varbinary(max)",
    "xml": "nvarchar(max)", "timestamp": "binary(8)",
}
DEMO_KINDS = {
    **dict.fromkeys(("int", "bigint", "smallint", "tinyint"), "integer"),
    **dict.fromkeys(("decimal", "numeric", "money", "smallmoney", "real", "float"), "decimal"),
    "bit": "boolean", "date": "date",
    **dict.fromkeys(("datetime", "datetime2", "smalldatetime", "datetimeoffset"), "datetime"),
}

logger = logging.getLogger("workforce_sync")


class SyncError(RuntimeError):
    pass


@dataclass(frozen=True)
class Column:
    name: str
    type_name: str
    max_length: int
    precision: int
    scale: int
    nullable: bool

    def ddl(self) -> str:
        kind = self.type_name.lower()
        if kind in LENGTH_TYPES:
            sql_type = f"{kind}({'max' if self.max_length == -1 else self.max_length})"
        elif kind in UNICODE_LENGTH_TYPES:
            sql_type = f"{kind}({'max' if self.max_length == -1 else self.max_length // 2})"
        elif kind in {"decimal", "numeric"}:
            sql_type = f"{kind}({self.precision},{self.scale})"
        elif kind in {"datetime2", "time", "datetimeoffset"}:
            sql_type = f"{kind}({self.scale})"
        elif kind in REPLACEMENT_TYPES:
            sql_type = REPLACEMENT_TYPES[kind]
        elif kind in PLAIN_TYPES:
            sql_type = kind
        else:
            raise SyncError(f"column {self.name} uses unsupported type {self.type_name}")
        return f"{quote(self.name)} {sql_type} {'NULL' if self.nullable else 'NOT NULL'}"


@dataclass(frozen=True)
class LocalTable:
    schema: str
    name: str
    columns: tuple[Column, ...]
    primary_key: tuple[str, ...]
    rows: tuple[tuple[Any, ...], ...]

    @property
    def qualified(self) -> str:
        return f"{self.schema}.{self.name}"

    @property
    def sql_name(self) -> str:
        return f"{quote(self.schema)}.{quote(self.name)}"

    @property
    def demo_name(self) -> str:
        return self.name if self.schema.casefold() == "dbo" else f"{self.schema}_{self.name}"

    def fingerprint(self) -> str:
        payload = {
            "columns": [column.__dict__ for column in self.columns], "primary_key": self.primary_key,
            "rows": sorted(json.dumps([format_value(value) for value in row]) for row in self.rows),
        }
        return sha256(json.dumps(payload, sort_keys=True).encode("utf-8")).hexdigest()


def quote(identifier: str) -> str:
    return "[" + identifier.replace("]", "]]") + "]"


def odbc_value(value: str) -> str:
    return "{" + value.replace("}", "}}") + "}"


def setting(name: str, default: str = "") -> str:
    return os.getenv(name, default).strip()


def configure_logging(verbose: bool) -> None:
    STATE_DIR.mkdir(parents=True, exist_ok=True)
    handler = RotatingFileHandler(STATE_DIR / "sync.log", maxBytes=1_000_000, backupCount=3, encoding="utf-8")
    handler.setFormatter(logging.Formatter("%(asctime)s %(levelname)s %(message)s"))
    logger.addHandler(handler)
    if sys.stderr is not None:
        logger.addHandler(logging.StreamHandler())
    logger.setLevel(logging.DEBUG if verbose else logging.INFO)


@contextmanager
def single_instance() -> Iterator[bool]:
    STATE_DIR.mkdir(parents=True, exist_ok=True)
    handle = open(STATE_DIR / "sync.lock", "a+")
    try:
        try:
            if os.name == "nt":
                import msvcrt
                msvcrt.locking(handle.fileno(), msvcrt.LK_NBLCK, 1)
            else:
                import fcntl
                fcntl.flock(handle, fcntl.LOCK_EX | fcntl.LOCK_NB)
        except OSError:
            yield False
            return
        yield True
    finally:
        handle.close()


def load_state() -> dict[str, Any]:
    try:
        return json.loads((STATE_DIR / "state.json").read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return {"cloud": {}, "demo": ""}


def save_state(state: dict[str, Any]) -> None:
    path = STATE_DIR / "state.json"
    temporary = path.with_suffix(".tmp")
    temporary.write_text(json.dumps(state, indent=2), encoding="utf-8")
    temporary.replace(path)


def connect_local() -> pyodbc.Connection:
    server = setting("SYNC_LOCAL_SQL_SERVER", r".\SQLEXPRESS")
    database = setting("SYNC_LOCAL_SQL_DATABASE", "EmployeeDB")
    return pyodbc.connect(
        f"DRIVER={{{ODBC_DRIVER}}};SERVER={odbc_value(server)};DATABASE={odbc_value(database)};"
        "Trusted_Connection=yes;Encrypt=yes;TrustServerCertificate=yes;", timeout=15,
    )


def read_local_tables(connection: pyodbc.Connection) -> tuple[list[LocalTable], list[str]]:
    tables: list[LocalTable] = []
    skipped: list[str] = []
    listing = connection.execute(
        "SELECT s.name, t.name, t.object_id FROM sys.tables t JOIN sys.schemas s ON s.schema_id = t.schema_id "
        "WHERE t.is_ms_shipped = 0 AND t.name <> 'sysdiagrams' ORDER BY s.name, t.name"
    ).fetchall()
    for schema, name, object_id in listing:
        columns = tuple(
            Column(row[0], row[1], int(row[2]), int(row[3]), int(row[4]), bool(row[5]))
            for row in connection.execute(
                "SELECT c.name, TYPE_NAME(c.system_type_id), c.max_length, c.precision, c.scale, c.is_nullable "
                "FROM sys.columns c WHERE c.object_id = ? ORDER BY c.column_id", object_id,
            ).fetchall()
        )
        try:
            for column in columns:
                column.ddl()
        except SyncError as error:
            skipped.append(f"{schema}.{name}: {error}")
            continue
        primary_key = tuple(
            row[0] for row in connection.execute(
                "SELECT c.name FROM sys.indexes i JOIN sys.index_columns ic "
                "ON ic.object_id = i.object_id AND ic.index_id = i.index_id "
                "JOIN sys.columns c ON c.object_id = ic.object_id AND c.column_id = ic.column_id "
                "WHERE i.object_id = ? AND i.is_primary_key = 1 ORDER BY ic.key_ordinal", object_id,
            ).fetchall()
        )
        select_list = ", ".join(quote(column.name) for column in columns)
        rows = tuple(tuple(row) for row in connection.execute(
            f"SELECT {select_list} FROM {quote(schema)}.{quote(name)}"
        ).fetchall())
        tables.append(LocalTable(schema, name, columns, primary_key, rows))
    return tables, skipped


def run_command(arguments: list[str], cwd: Path | None = None, check: bool = True) -> subprocess.CompletedProcess[str]:
    result = subprocess.run(
        arguments, cwd=cwd, capture_output=True, text=True,
        env={**os.environ, "GIT_TERMINAL_PROMPT": "0"},
        creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
    )
    if check and result.returncode:
        raise SyncError(f"{Path(arguments[0]).name} {arguments[1]} failed: {(result.stderr or result.stdout).strip()[:500]}")
    return result


def azure_cli() -> str:
    executable = shutil.which("az") or shutil.which("az.cmd")
    if not executable:
        raise SyncError("Azure CLI (az) is not installed; it is required for Azure SQL authentication.")
    return executable


def connect_cloud(server: str, database: str) -> pyodbc.Connection:
    token = run_command([
        azure_cli(), "account", "get-access-token", "--resource", "https://database.windows.net/",
        "--query", "accessToken", "--output", "tsv", "--only-show-errors",
    ]).stdout.strip().encode("utf-16-le")
    return pyodbc.connect(
        f"DRIVER={{{ODBC_DRIVER}}};SERVER=tcp:{server},1433;DATABASE={odbc_value(database)};"
        "Encrypt=yes;TrustServerCertificate=no;Connection Timeout=30;",
        attrs_before={SQL_COPT_SS_ACCESS_TOKEN: struct.pack("<I", len(token)) + token}, autocommit=False,
    )


def allow_client_ip(error: pyodbc.Error, server: str) -> bool:
    match = re.search(r"Client with IP address '([0-9.]+)' is not allowed", str(error))
    resource_group = setting("SYNC_CLOUD_RESOURCE_GROUP")
    if not match or not resource_group:
        return False
    logger.info("Adding this PC's public IP %s to the Azure SQL firewall.", match.group(1))
    run_command([
        azure_cli(), "sql", "server", "firewall-rule", "create", "--resource-group", resource_group,
        "--server", server.split(".")[0], "--name", "WorkforceCompassSync",
        "--start-ip-address", match.group(1), "--end-ip-address", match.group(1),
        "--only-show-errors", "--output", "none",
    ])
    return True


def sync_cloud(tables: list[LocalTable], state: dict[str, Any], force: bool) -> None:
    server = setting("SYNC_CLOUD_SQL_SERVER")
    if not server:
        logger.debug("SYNC_CLOUD_SQL_SERVER is not set; skipping Azure SQL.")
        return
    synced: dict[str, str] = state.setdefault("cloud", {})
    current = {table.qualified: table for table in tables}
    changed = [table for table in tables if force or synced.get(table.qualified) != table.fingerprint()]
    removed = [name for name in synced if name not in current]
    if not changed and not removed:
        return
    database = setting("SYNC_CLOUD_SQL_DATABASE", "EmployeeDB")
    try:
        connection = connect_cloud(server, database)
    except pyodbc.Error as error:
        if not allow_client_ip(error, server):
            raise
        time.sleep(30)
        connection = connect_cloud(server, database)
    reader = setting("SYNC_CLOUD_READER_USER")
    with connection:
        cursor = connection.cursor()
        cursor.fast_executemany = True
        reader_exists = bool(reader) and cursor.execute(
            "SELECT 1 FROM sys.database_principals WHERE name = ?", reader,
        ).fetchone() is not None
        try:
            for name in removed:
                schema, table_name = name.split(".", 1)
                cursor.execute(f"DROP TABLE IF EXISTS {quote(schema)}.{quote(table_name)}")
            for table in changed:
                if table.schema.casefold() != "dbo":
                    cursor.execute(
                        f"IF SCHEMA_ID(?) IS NULL EXEC('CREATE SCHEMA {quote(table.schema)}')", table.schema,
                    )
                cursor.execute(f"DROP TABLE IF EXISTS {table.sql_name}")
                definitions = [column.ddl() for column in table.columns]
                if table.primary_key:
                    definitions.append(f"PRIMARY KEY ({', '.join(quote(name) for name in table.primary_key)})")
                cursor.execute(f"CREATE TABLE {table.sql_name} ({', '.join(definitions)})")
                if table.rows:
                    columns = ", ".join(quote(column.name) for column in table.columns)
                    placeholders = ", ".join("?" for _ in table.columns)
                    cursor.executemany(f"INSERT INTO {table.sql_name} ({columns}) VALUES ({placeholders})",
                                       [list(row) for row in table.rows])
                count = cursor.execute(f"SELECT COUNT_BIG(*) FROM {table.sql_name}").fetchval()
                if count != len(table.rows):
                    raise SyncError(f"{table.qualified}: expected {len(table.rows)} rows in Azure SQL, found {count}")
                if reader_exists:
                    cursor.execute(f"GRANT SELECT ON OBJECT::{table.sql_name} TO {quote(reader)}")
            connection.commit()
        except Exception:
            connection.rollback()
            raise
    for name in removed:
        synced.pop(name, None)
    for table in changed:
        synced[table.qualified] = table.fingerprint()
    save_state(state)
    logger.info(
        "Azure SQL updated: %s%s.",
        ", ".join(f"{table.qualified} ({len(table.rows)} rows)" for table in changed) or "no tables recreated",
        f"; removed {', '.join(removed)}" if removed else "",
    )


def demo_tables(tables: list[LocalTable]) -> list[DemoTable]:
    result = []
    for table in tables:
        if len(table.rows) > MAX_DEMO_ROWS:
            logger.warning("Skipping %s in demo data: more than %s rows.", table.qualified, f"{MAX_DEMO_ROWS:,}")
            continue
        result.append(DemoTable(
            table.demo_name,
            tuple((column.name, DEMO_KINDS.get(column.type_name.lower(), "text")) for column in table.columns),
            table.rows,
        ))
    employees = next((table for table in result if table.name == "Employees"), None)
    if employees is None:
        raise SyncError("the local database has no dbo.Employees table, which the demo app requires")
    missing = set(REQUIRED_EMPLOYEE_COLUMNS) - {name for name, _ in employees.columns}
    if missing:
        raise SyncError(f"dbo.Employees is missing columns required by the app: {', '.join(sorted(missing))}")
    return result


def git(repo: Path, *arguments: str, check: bool = True) -> subprocess.CompletedProcess[str]:
    return run_command(["git", *arguments], cwd=repo, check=check)


def sync_clone() -> Path:
    repo = STATE_DIR / "repo"
    branch = setting("SYNC_GIT_BRANCH", "main")
    if not (repo / ".git").exists():
        url = git(ROOT, "config", "--get", "remote.origin.url").stdout.strip()
        if repo.exists():
            shutil.rmtree(repo)
        run_command(["git", "clone", "--quiet", "--branch", branch, "--single-branch", url, str(repo)])
        for key in ("user.name", "user.email"):
            value = git(ROOT, "config", "--get", key, check=False).stdout.strip()
            if value:
                git(repo, "config", key, value)
    return repo


def push_demo_data(tables: list[DemoTable], state: dict[str, Any], force: bool) -> None:
    fingerprint = sha256(json.dumps(
        [[table.name, table.columns, sorted(json.dumps([format_value(v) for v in row]) for row in table.rows)]
         for table in tables], sort_keys=True,
    ).encode("utf-8")).hexdigest()
    if not force and state.get("demo") == fingerprint:
        return
    branch = setting("SYNC_GIT_BRANCH", "main")
    repo = sync_clone()
    for attempt in range(3):
        git(repo, "fetch", "--quiet", "origin", branch)
        git(repo, "reset", "--quiet", "--hard", f"origin/{branch}")
        if not write_demo_tables(tables, repo / "demo_data"):
            logger.info("Demo data on GitHub is already up to date.")
            break
        git(repo, "add", "--all", "demo_data")
        git(repo, "commit", "--quiet", "-m", SYNC_COMMIT_MESSAGE)
        if git(repo, "push", "--quiet", "origin", f"HEAD:{branch}", check=False).returncode == 0:
            logger.info("Pushed updated demo data to GitHub (%s).", ", ".join(table.name for table in tables))
            break
        logger.info("GitHub push was rejected (attempt %s); retrying with the latest branch.", attempt + 1)
    else:
        raise SyncError("could not push demo data to GitHub after 3 attempts")
    state["demo"] = fingerprint
    save_state(state)


def sync(force: bool = False) -> int:
    load_dotenv(ROOT / ".env", override=False)
    with single_instance() as acquired:
        if not acquired:
            logger.debug("Another sync is already running.")
            return 0
        state = load_state()
        failures = 0
        with connect_local() as connection:
            tables, skipped = read_local_tables(connection)
        for message in skipped:
            logger.warning("Skipping table %s", message)
        try:
            sync_cloud(tables, state, force)
        except (SyncError, pyodbc.Error) as error:
            failures += 1
            logger.error("Azure SQL sync failed: %s", error)
        try:
            push_demo_data(demo_tables(tables), state, force)
        except SyncError as error:
            failures += 1
            logger.error("Demo data sync failed: %s", error)
        return 1 if failures else 0


def install_task(interval: int) -> None:
    pythonw = Path(sys.executable).with_name("pythonw.exe")
    executable = pythonw if pythonw.exists() else Path(sys.executable)
    script = f"""
$action = New-ScheduledTaskAction -Execute '{executable}' -Argument '"{Path(__file__).resolve()}"' -WorkingDirectory '{ROOT}'
$trigger = New-ScheduledTaskTrigger -Once -At (Get-Date) -RepetitionInterval (New-TimeSpan -Minutes {interval})
$settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -StartWhenAvailable `
    -MultipleInstances IgnoreNew -ExecutionTimeLimit (New-TimeSpan -Minutes 10)
Register-ScheduledTask -TaskName '{TASK_NAME}' -Action $action -Trigger $trigger -Settings $settings `
    -Description 'Mirrors the local EmployeeDB to Azure SQL and the Workforce Compass demo data.' -Force | Out-Null
"""
    run_command(["powershell", "-NoProfile", "-NonInteractive", "-Command", script])
    print(f"Installed scheduled task '{TASK_NAME}' (every {interval} minute(s)). Log: {STATE_DIR / 'sync.log'}")


def uninstall_task() -> None:
    run_command(["powershell", "-NoProfile", "-NonInteractive", "-Command",
                 f"Unregister-ScheduledTask -TaskName '{TASK_NAME}' -Confirm:$false"])
    print(f"Removed scheduled task '{TASK_NAME}'.")


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__.split("\n\n")[0])
    parser.add_argument("--force", action="store_true", help="resync every table even if unchanged")
    parser.add_argument("--export-demo", type=Path, metavar="DIR", help="only write demo data files to DIR")
    parser.add_argument("--install-task", action="store_true", help="run the sync automatically (Windows)")
    parser.add_argument("--uninstall-task", action="store_true", help="remove the scheduled task")
    parser.add_argument("--interval", type=int, default=1, help="minutes between automatic syncs (default 1)")
    parser.add_argument("--verbose", action="store_true")
    options = parser.parse_args()
    if options.install_task:
        install_task(max(1, options.interval))
        return 0
    if options.uninstall_task:
        uninstall_task()
        return 0
    configure_logging(options.verbose)
    if options.export_demo:
        load_dotenv(ROOT / ".env", override=False)
        with connect_local() as connection:
            tables, skipped = read_local_tables(connection)
        for message in skipped:
            logger.warning("Skipping table %s", message)
        changed = write_demo_tables(demo_tables(tables), options.export_demo)
        logger.info("Demo data %s in %s.", "updated" if changed else "already current", options.export_demo)
        return 0
    try:
        return sync(options.force)
    except Exception:
        logger.exception("Database sync failed.")
        return 1


if __name__ == "__main__":
    sys.exit(main())
