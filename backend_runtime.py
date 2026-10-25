from hashlib import sha256
import importlib
from pathlib import Path
import sys
from threading import Lock


MODULES = ("backend_diagnostics", "excel_support", "workbook_operations", "source_queries")
RELOAD_LOCK = Lock()
LOADED_REVISIONS: dict[str, str] = {}


def refresh_backend_modules() -> None:
    # Streamlit reruns the entry script but can retain imported modules after deployment.
    with RELOAD_LOCK:
        root = Path(__file__).resolve().parent
        revisions = {
            name: sha256((root / f"{name}.py").read_bytes()).hexdigest()
            for name in MODULES
        }
        if revisions == LOADED_REVISIONS:
            return
        importlib.invalidate_caches()
        for name in MODULES:
            if name in sys.modules:
                importlib.reload(sys.modules[name])
            else:
                importlib.import_module(name)
        LOADED_REVISIONS.clear()
        LOADED_REVISIONS.update(revisions)
