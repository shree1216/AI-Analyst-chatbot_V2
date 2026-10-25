from pathlib import Path
import subprocess
import sys
import unittest


ROOT = Path(__file__).resolve().parents[1]


class BackendRuntimeTests(unittest.TestCase):
    def test_hot_deployment_refreshes_stale_modules_and_unchanged_reruns_keep_them(self):
        script = """
from unittest.mock import Mock, patch
import backend_diagnostics
import excel_support
import workbook_operations
import source_queries
from backend_runtime import LOADED_REVISIONS, refresh_backend_modules

class LegacyTrace:
    def record(self, event, **fields):
        raise ValueError("Backend logs must not include questions, cell values, credentials, or query literals.")

backend_diagnostics.BackendTrace = LegacyTrace
backend_diagnostics.LOG_FIELDS = {"route"}
refresh_backend_modules()
assert backend_diagnostics.BackendTrace is not LegacyTrace
assert workbook_operations.BackendTrace is backend_diagnostics.BackendTrace
assert source_queries.BackendTrace is backend_diagnostics.BackendTrace
with patch.object(backend_diagnostics, "backend_logger", return_value=Mock()):
    events = []
    trace = backend_diagnostics.BackendTrace(events)
    trace.record("table_schema", source_id="sql_0", schemas={"Amount": "number"})
    trace.record_code("query_started", workbook_operations.execute_workbook_plan,
                      sql="SELECT COUNT(*) FROM Employees", parameters={})
    assert events[1]["implementation"][0]["function"] == "execute_workbook_plan"
    assert events[1]["sql"] == "SELECT COUNT(*) FROM Employees"

current = backend_diagnostics.BackendTrace
refresh_backend_modules()
assert backend_diagnostics.BackendTrace is current
LOADED_REVISIONS["backend_diagnostics"] = "previous-deployment"
refresh_backend_modules()
assert backend_diagnostics.BackendTrace is not current
assert workbook_operations.BackendTrace is backend_diagnostics.BackendTrace
assert source_queries.BackendTrace is backend_diagnostics.BackendTrace
"""
        result = subprocess.run(
            [sys.executable, "-c", script], cwd=ROOT, capture_output=True, text=True, timeout=60,
        )
        self.assertEqual(result.returncode, 0, result.stdout + result.stderr)


if __name__ == "__main__":
    unittest.main()
