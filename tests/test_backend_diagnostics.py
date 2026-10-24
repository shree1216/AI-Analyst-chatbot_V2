import json
import logging
from pathlib import Path
from tempfile import TemporaryDirectory
import unittest
from unittest.mock import Mock, patch

import backend_diagnostics as diagnostics
import pandas as pd
from workbook_operations import execute_workbook_plan


class BackendDiagnosticTests(unittest.TestCase):
    def setUp(self):
        self.logger = Mock()
        self.enterContext(patch.object(diagnostics, "backend_logger", return_value=self.logger))

    def test_trace_has_timestamps_request_id_metadata_and_no_unapproved_fields(self):
        events = []
        trace = diagnostics.BackendTrace(events)
        trace.record("sheet_schema", sheet="Sales", rows=150, schemas={"Revenue": "number"})
        trace.record("request_completed", rows=2, status="success")
        self.assertEqual(len(events), 2)
        self.assertEqual(events[0]["request_id"], events[1]["request_id"])
        self.assertIn("+00:00", events[0]["timestamp"])
        self.assertGreaterEqual(events[1]["elapsed_ms"], events[0]["elapsed_ms"])
        self.assertEqual(json.loads(diagnostics.diagnostic_log(events).splitlines()[0])["rows"], 150)
        for field in ("question", "cell_values", "password", "api_key", "sql", "filter", "error_message"):
            with self.subTest(field=field), self.assertRaisesRegex(ValueError, "must not include"):
                trace.record("unsafe", **{field: "PRIVATE_VALUE"})
        self.assertNotIn("PRIVATE_VALUE", diagnostics.diagnostic_log(events))

    def test_separate_request_logs_are_isolated_and_bounded(self):
        first, second = [], []
        one, two = diagnostics.BackendTrace(first), diagnostics.BackendTrace(second)
        for _ in range(205):
            one.record("step_completed", rows=1)
        two.record("request_completed", rows=2)
        self.assertEqual(len(first), 200)
        self.assertEqual(len(second), 1)
        self.assertNotIn(two.request_id, diagnostics.diagnostic_log(first))
        self.assertEqual(len(diagnostics.diagnostic_log(first * 6).splitlines()), 1000)

    def test_server_log_failure_is_explicit_but_session_log_survives(self):
        self.logger.info.side_effect = OSError("PRIVATE_PATH")
        events = []
        trace = diagnostics.BackendTrace(events)
        trace.record("request_started", route="excel")
        trace.record("request_completed", rows=1)
        self.assertIn("could not be saved", trace.file_error)
        self.assertEqual(len(events), 2)
        self.assertNotIn("PRIVATE_PATH", diagnostics.diagnostic_log(events))
        self.logger.info.assert_called_once()

    def test_execution_log_records_steps_but_not_filter_values_or_cells(self):
        events = []
        plan = {"understood": True, "unhandled_conditions": [], "sheet": "Sales", "steps": [
            {"op": "filter", "where": {"function": "eq", "args": [
                {"column": "Region"}, {"literal": "PRIVATE_REGION"},
            ]}},
            {"op": "aggregate", "metrics": [{"name": "Total", "function": "sum", "column": "Revenue"}]},
        ]}
        result = execute_workbook_plan(plan, {"Sales": pd.DataFrame({
            "Region": ["PRIVATE_REGION"], "Revenue": [987654321],
        })}, trace=diagnostics.BackendTrace(events))
        self.assertEqual(result.iloc[0, 0], 987654321)
        text = diagnostics.diagnostic_log(events)
        self.assertNotIn("PRIVATE_REGION", text)
        self.assertNotIn("987654321", text)
        self.assertEqual([entry["event"] for entry in events],
                         ["source_selected", "step_started", "step_completed", "step_started", "step_completed"])
        self.assertEqual(events[-1]["rows"], 1)


class BackendLogFileTests(unittest.TestCase):
    def test_real_rotating_file_is_json_lines_and_stays_bounded(self):
        diagnostics.backend_logger.cache_clear()
        try:
            with TemporaryDirectory() as directory, patch.object(diagnostics, "LOG_FILE", Path(directory) / "backend.log"):
                logger = diagnostics.backend_logger()
                handler = logger.handlers[0]
                self.assertIsInstance(handler, logging.handlers.RotatingFileHandler)
                handler.maxBytes = 400
                for index in range(20):
                    diagnostics.BackendTrace([]).record("request_completed", rows=index, status="success")
                handler.flush()
                files = list(Path(directory).glob("backend.log*"))
                self.assertLessEqual(len(files), 4)
                self.assertGreater(len(files), 1)
                for file in files:
                    for line in file.read_text(encoding="utf-8").splitlines():
                        self.assertIn("request_id", json.loads(line))
                handler.close()
                logger.handlers.clear()
        finally:
            diagnostics.backend_logger.cache_clear()


if __name__ == "__main__":
    unittest.main()
