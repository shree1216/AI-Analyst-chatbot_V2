from datetime import datetime, timezone
from functools import lru_cache
import inspect
import json
import logging
from logging.handlers import RotatingFileHandler
from pathlib import Path
import re
from threading import Lock
from time import perf_counter
from typing import Any, Callable
from uuid import uuid4


LOG_FILE = Path(__file__).with_name("logs") / "backend.log"
LOGGER_LOCK = Lock()
LOG_FIELDS = {
    "route", "sheet", "source_sheets", "unavailable_sheets", "rows", "columns",
    "schemas", "operation", "step", "status", "error_type", "reason_count",
    "workbook_id", "preview_rows",
    "sql", "parameters", "plan", "details", "duration_ms", "result_columns",
    "response_kind", "attempt",
    "source_id",
    "implementation",
    "matched_rows", "unmatched_rows",
}

# Preserve identifiers/placeholders while removing SQL comments and literal values.
SQL_TOKENS = re.compile(
    r"(?P<identifier>\[(?:[^\]]|\]\])*\]|\"(?:[^\"]|\"\")*\"|`(?:[^`]|``)*`)"
    r"|(?P<comment>--[^\r\n]*|/\*[\s\S]*?\*/)"
    r"|(?P<literal>N?'(?:[^']|'')*'|(?<![\w:])(?:0x[0-9a-f]+|\d+(?:\.\d*)?(?:e[+-]?\d+)?)(?!\w))",
    re.IGNORECASE,
)
PLAN_FIELDS = {
    "understood", "sheet", "steps", "op", "columns", "metrics", "filter",
    "all", "any", "group_by", "order_by", "limit", "per_group", "limit_per_group",
    "where", "function", "args", "column", "name", "expression", "by", "direction",
    "left_on", "right_on", "how", "keep", "count", "method", "index", "values",
    "mode", "sql", "excel", "calculation", "query", "source",
}


def diagnostic_sql(sql: str) -> str:
    return SQL_TOKENS.sub(
        lambda match: match.group() if match.lastgroup == "identifier"
        else " " if match.lastgroup == "comment" else "?",
        sql,
    )


def diagnostic_plan(value: Any) -> Any:
    if isinstance(value, dict):
        result = {}
        for key, item in value.items():
            if key not in PLAN_FIELDS:
                result[key] = "[REDACTED]"
            elif key == "args":
                result[key] = [
                    diagnostic_plan(argument) if isinstance(argument, dict) else "[REDACTED]"
                    for argument in item
                ] if isinstance(item, list) else "[REDACTED]"
            else:
                result[key] = diagnostic_plan(item)
        return result
    if isinstance(value, list):
        return [diagnostic_plan(item) for item in value]
    return value


class StrictFileHandler(RotatingFileHandler):
    def handleError(self, record: logging.LogRecord) -> None:
        raise OSError(f"The backend log file could not be written for {record.name}.")


@lru_cache(maxsize=1)
def backend_logger() -> logging.Logger:
    with LOGGER_LOCK:
        LOG_FILE.parent.mkdir(parents=True, exist_ok=True)
        logger = logging.getLogger("workforce.backend")
        logger.setLevel(logging.INFO)
        logger.propagate = False
        if any(isinstance(handler, StrictFileHandler) and handler.baseFilename == str(LOG_FILE.resolve())
               for handler in logger.handlers):
            return logger
        for handler in logger.handlers:
            handler.close()
        logger.handlers.clear()
        handler = StrictFileHandler(LOG_FILE, maxBytes=2_000_000, backupCount=3, encoding="utf-8")
        handler.setFormatter(logging.Formatter("%(message)s"))
        logger.addHandler(handler)
        return logger


def _metadata(value: Any, depth: int = 0) -> Any:
    if depth > 32:
        raise ValueError("Backend log metadata is too deeply nested.")
    if isinstance(value, str):
        return value[:256]
    if value is None or isinstance(value, (bool, int, float)):
        return value
    if isinstance(value, list):
        return [_metadata(item, depth + 1) for item in value[:200]]
    if isinstance(value, dict):
        return {str(key)[:256]: _metadata(item, depth + 1) for key, item in list(value.items())[:100]}
    raise ValueError("Backend logs accept structured metadata only.")


class BackendTrace:
    def __init__(self, events: list[dict[str, Any]], request_id: str | None = None):
        self.events = events
        self.request_id = request_id or uuid4().hex
        self.started = perf_counter()
        self.file_error: str | None = None

    def record_code(self, event: str, *functions: Callable[..., Any], **fields: Any) -> None:
        implementation = []
        for function in functions:
            try:
                lines, first_line = inspect.getsourcelines(function)
                file = inspect.getsourcefile(function)
                if file is None:
                    raise OSError("Source file is unavailable.")
                implementation.append({
                    "function": function.__qualname__, "file": Path(file).name,
                    "line": first_line, "language": "python", "status": "available",
                    "code": "".join(lines),
                })
            except (OSError, TypeError) as error:
                implementation.append({
                    "status": "unavailable", "error_type": type(error).__name__,
                })
                logging.getLogger(__name__).warning("Backend implementation source unavailable: %s", type(error).__name__)
        self.record(event, implementation=implementation, **fields)

    def record(self, event: str, **fields: Any) -> None:
        if set(fields) - LOG_FIELDS:
            raise ValueError("Backend logs must not include unapproved fields, questions, cell values, or credentials.")
        if "sql" in fields:
            fields["sql"] = diagnostic_sql(fields["sql"])
        if "parameters" in fields:
            fields["parameters"] = {
                name: {"type": type(value).__name__, "value": "[REDACTED]"}
                for name, value in fields["parameters"].items()
            }
        for key in ("plan", "details"):
            if key in fields:
                fields[key] = diagnostic_plan(fields[key])
        entry = {
            "timestamp": datetime.now(timezone.utc).isoformat(),
            "request_id": self.request_id,
            "event": event,
            "elapsed_ms": round((perf_counter() - self.started) * 1000, 2),
            **{key: value if key in {"sql", "implementation"} else _metadata(value) for key, value in fields.items()},
        }
        self.events.append(entry)
        if len(self.events) > 200:
            del self.events[:-200]
        if self.file_error:
            return
        try:
            backend_logger().info("%s", json.dumps(entry, ensure_ascii=True, allow_nan=False))
        except OSError as error:
            self.file_error = "The server log file could not be saved. Your session log is still available to download."
            logging.getLogger(__name__).warning("Backend log unavailable: %s", type(error).__name__)


def diagnostic_log(events: list[dict[str, Any]]) -> str:
    return "".join(json.dumps(event, ensure_ascii=True, allow_nan=False) + "\n" for event in events[-1000:])
