import inspect
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
        for field in ("question", "cell_values", "password", "api_key", "filter", "error_message"):
            with self.subTest(field=field), self.assertRaisesRegex(ValueError, "must not include"):
                trace.record("unsafe", **{field: "PRIVATE_VALUE"})
        self.assertNotIn("PRIVATE_VALUE", diagnostics.diagnostic_log(events))

    def test_sql_keeps_statement_and_identifiers_without_literals_or_parameter_values(self):
        events = []
        sql = (
            "SELECT [Employee ID], \"column2\", `field3`, COUNT(*) FROM Employees "
            "WHERE name = N'PRIVATE_NAME' AND department = :department "
            "AND age > 35 AND salary < 1.25e6 AND emp_id = 0xAB "
            "AND city = 'O''PRIVATE_CITY' /* PRIVATE_COMMENT */ -- PRIVATE_COMMENT\n"
            "ORDER BY [Employee ID], [department], [job_title], [join_date], [salary], [status], [gender], [city]"
        )
        diagnostics.BackendTrace(events).record(
            "query_started", sql=sql, parameters={"department": "PRIVATE_PARAMETER"},
        )
        entry = events[0]
        self.assertGreater(len(entry["sql"]), 256)
        self.assertIn('SELECT [Employee ID], "column2", `field3`, COUNT(*) FROM Employees', entry["sql"])
        self.assertIn("department = :department", entry["sql"])
        self.assertIn("ORDER BY [Employee ID]", entry["sql"])
        self.assertEqual(entry["parameters"]["department"], {"type": "str", "value": "[REDACTED]"})
        for private in ("PRIVATE_NAME", "PRIVATE_PARAMETER", "PRIVATE_CITY", "PRIVATE_COMMENT"):
            self.assertNotIn(private, diagnostics.diagnostic_log(events))
        for literal in ("35", "1.25e6", "0xAB"):
            self.assertNotIn(literal, entry["sql"])
        self.assertEqual(json.loads(self.logger.info.call_args.args[1])["sql"], entry["sql"])

    def test_nested_plans_keep_calculations_but_redact_filter_values(self):
        events = []
        condition = {"column": "Region", "op": "eq", "value": "PRIVATE_REGION"}
        plan = {"filter": {"all": [{"any": [condition]}]}, "metrics": [{"function": "sum", "column": "Revenue"}],
                "having": {"column": "sum_Revenue", "op": "gt", "value": 987654321},
                "unexpected_output": "PRIVATE_MODEL_TEXT"}
        diagnostics.BackendTrace(events).record("planner_completed", plan=plan)
        self.assertEqual(events[0]["plan"]["metrics"], plan["metrics"])
        self.assertNotIn("PRIVATE_REGION", diagnostics.diagnostic_log(events))
        self.assertNotIn("PRIVATE_MODEL_TEXT", diagnostics.diagnostic_log(events))
        self.assertEqual(events[0]["plan"]["having"], {
            "column": "sum_Revenue", "op": "gt", "value": "[REDACTED]",
        })
        self.assertNotIn("987654321", diagnostics.diagnostic_log(events))
        self.assertEqual(condition["value"], "PRIVATE_REGION")

    def test_logging_supports_maximum_valid_workbook_expression_depth(self):
        expression = {"column": "Revenue"}
        for _ in range(12):
            expression = {"function": "abs", "args": [expression]}
        plan = {"understood": True, "unhandled_conditions": [], "sheet": "Sales", "steps": [
            {"op": "calculate", "columns": [{"name": "Absolute Revenue", "expression": expression}]},
        ]}
        events = []
        trace = diagnostics.BackendTrace(events)
        trace.record("planner_received", plan=plan)
        result = execute_workbook_plan(plan, {"Sales": pd.DataFrame({"Revenue": [-10]})}, trace=trace)
        self.assertEqual(result["Absolute Revenue"].tolist(), [10])
        self.assertEqual(events[-1]["event"], "step_completed")

    def test_shorthand_expression_literals_are_redacted_without_changing_the_plan(self):
        expression = {"function": "if", "args": [
            {"function": "eq", "args": [{"column": "Name"}, "PRIVATE_PERSON"]},
            "PRIVATE_LABEL", None,
        ]}
        events = []
        diagnostics.BackendTrace(events).record("planner_received", plan={"expression": expression})
        rendered = diagnostics.diagnostic_log(events)
        self.assertNotIn("PRIVATE_PERSON", rendered)
        self.assertNotIn("PRIVATE_LABEL", rendered)
        self.assertIn('"column": "Name"', rendered)
        self.assertEqual(events[0]["plan"]["expression"]["args"][1:], ["[REDACTED]", "[REDACTED]"])
        self.assertEqual(expression["args"][0]["args"][1], "PRIVATE_PERSON")

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

    def test_python_implementation_is_complete_and_in_both_log_outputs(self):
        events = []
        diagnostics.BackendTrace(events).record_code("query_started", execute_workbook_plan, operation="calculate")
        entry = events[0]["implementation"][0]
        self.assertEqual(entry["function"], "execute_workbook_plan")
        self.assertEqual(entry["file"], "workbook_operations.py")
        self.assertEqual(entry["line"], inspect.getsourcelines(execute_workbook_plan)[1])
        self.assertEqual(entry["code"], inspect.getsource(execute_workbook_plan))
        self.assertGreater(len(entry["code"]), 256)
        self.assertEqual(json.loads(diagnostics.diagnostic_log(events))["implementation"][0], entry)
        self.assertEqual(json.loads(self.logger.info.call_args.args[1])["implementation"][0], entry)

    def test_missing_implementation_source_is_explicit_without_losing_the_event(self):
        events = []
        with patch("backend_diagnostics.inspect.getsourcelines", side_effect=OSError("PRIVATE_PATH")):
            diagnostics.BackendTrace(events).record_code("query_started", execute_workbook_plan)
        self.assertEqual(events[0]["implementation"], [{"status": "unavailable", "error_type": "OSError"}])
        self.assertNotIn("PRIVATE_PATH", diagnostics.diagnostic_log(events))

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
        self.assertEqual(events[1]["details"]["where"]["args"][1], {"literal": "[REDACTED]"})
        self.assertEqual(events[3]["details"]["metrics"][0]["function"], "sum")
        self.assertEqual(events[-1]["result_columns"], ["Total"])
        functions = [item["function"] for item in events[0]["implementation"]]
        self.assertIn("execute_workbook_plan", functions)
        self.assertIn("evaluate_expression", functions)
        self.assertIn("_evaluate", functions)
        self.assertIn("_aggregate", functions)


class BackendLogFileTests(unittest.TestCase):
    def test_real_server_file_contains_the_python_implementation(self):
        diagnostics.backend_logger.cache_clear()
        try:
            with TemporaryDirectory() as directory, patch.object(diagnostics, "LOG_FILE", Path(directory) / "backend.log"):
                logger = diagnostics.backend_logger()
                try:
                    diagnostics.BackendTrace([]).record_code("query_started", execute_workbook_plan)
                    logger.handlers[0].flush()
                    entry = json.loads(diagnostics.LOG_FILE.read_text(encoding="utf-8"))
                    self.assertEqual(entry["implementation"][0]["code"], inspect.getsource(execute_workbook_plan))
                finally:
                    for handler in logger.handlers:
                        handler.close()
                    logger.handlers.clear()
        finally:
            diagnostics.backend_logger.cache_clear()

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
