from datetime import datetime, timezone
from functools import lru_cache
import json
import logging
from logging.handlers import RotatingFileHandler
from pathlib import Path
from threading import Lock
from time import perf_counter
from typing import Any
from uuid import uuid4


LOG_FILE = Path(__file__).with_name("logs") / "backend.log"
LOGGER_LOCK = Lock()
LOG_FIELDS = {
    "route", "sheet", "source_sheets", "unavailable_sheets", "rows", "columns",
    "schemas", "operation", "step", "status", "error_type", "reason_count",
    "workbook_id", "preview_rows",
}


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
    if depth > 4:
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

    def record(self, event: str, **fields: Any) -> None:
        if set(fields) - LOG_FIELDS:
            raise ValueError("Backend logs must not include questions, cell values, credentials, or query literals.")
        entry = {
            "timestamp": datetime.now(timezone.utc).isoformat(),
            "request_id": self.request_id,
            "event": event,
            "elapsed_ms": round((perf_counter() - self.started) * 1000, 2),
            **{key: _metadata(value) for key, value in fields.items()},
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
