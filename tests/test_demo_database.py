import importlib
import json
import os
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path
import unittest
from unittest.mock import Mock, patch

import pandas as pd
from sqlalchemy.exc import SQLAlchemyError
from streamlit.testing.v1 import AppTest
from backend_diagnostics import BackendTrace, diagnostic_log


ROOT = Path(__file__).resolve().parents[1]


def employee_app():
    app = AppTest.from_file(str(ROOT / "employee_sql_chatbot.py"))
    app.session_state["data_source"] = "Employee database"
    return app


TEST_ENVIRONMENT = {
    name: value for name, value in os.environ.items()
    if not name.upper().startswith(("AZURE_", "SMTP_", "WORKFORCE_"))
}


class DemoDatabaseTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        with patch("dotenv.load_dotenv"), patch.dict(os.environ, TEST_ENVIRONMENT, clear=True):
            cls.chatbot = importlib.import_module("employee_sql_chatbot")
        cls.employees = pd.read_excel(ROOT / "employee_dataset.xlsx")

    def setUp(self):
        self.enterContext(patch("dotenv.load_dotenv"))
        self.enterContext(patch.dict(os.environ, {**TEST_ENVIRONMENT, "WORKFORCE_DATA_MODE": "demo"}, clear=True))
        self.enterContext(patch("socket.create_connection", side_effect=AssertionError("Unexpected network access")))
        self.enterContext(patch.multiple(
            self.chatbot, DATA_MODE="demo", EMPLOYEE_TABLE="Employees",
            TABLE=self.chatbot.TABLE.replace("dbo.Employees", "Employees"),
            SOURCE_LABEL="Bundled synthetic employee_dataset.xlsx",
        ))
        self.chatbot.get_engine.clear()
        self.chatbot.get_employee_names.clear()
        self.engine = self.chatbot.get_engine()

    def tearDown(self):
        self.engine.dispose()
        self.chatbot.get_engine.clear()
        self.chatbot.get_employee_names.clear()

    def query(self, question, *, known_template=False):
        translate = self.chatbot.translate_known_question if known_template else self.chatbot.translate_question
        request = translate(question)
        self.assertIsNotNone(request)
        self.assertTrue(request.sql, request.explanation)
        return self.chatbot.run_query(request.sql, request.parameters)

    def test_count_and_filtered_count_match_bundled_dataset(self):
        self.assertEqual(self.query("How many employees do we have?").iloc[0, 0], len(self.employees))
        self.assertEqual(
            self.query("How many employees are older than 30?").iloc[0, 0],
            (self.employees["Age"] > 30).sum(),
        )
        self.assertEqual(
            self.query("Count employees whose emails end with .org").iloc[0, 0],
            self.employees["Email"].str.lower().str.rstrip().str.endswith(".org").sum(),
        )

    def test_lookup_and_top_salary_limits(self):
        employee = self.employees.iloc[0]
        result = self.query(f"Who owns {employee['Email']}?")
        self.assertEqual(result["emp_id"].tolist(), [employee["Employee ID"]])
        result = self.query(f"Give details of {employee['Employee ID']}")
        self.assertEqual(result["name"].tolist(), [employee["Name"]])
        result = self.query("Show top 3 salaries", known_template=True)
        self.assertEqual(result["salary"].tolist(), self.employees["Salary"].nlargest(3).tolist())
        with patch.object(self.chatbot, "find_employee_name", return_value=employee["Name"]):
            result = self.query(f"What is the salary of {employee['Name']}?", known_template=True)
        self.assertEqual(result["salary"].tolist(), [employee["Salary"]])

    def test_female_salary_count_matches_bundled_dataset_without_ai(self):
        with patch.object(self.chatbot, "generate_employee_query") as planner:
            result = self.query("How many female employees have a salary above 80000?")
        planner.assert_not_called()
        expected = (
            (self.employees["Gender"].str.lower() == "female") & (self.employees["Salary"] > 80000)
        ).sum()
        self.assertEqual(result.columns.tolist(), ["employee_count"])
        self.assertEqual(result.shape, (1, 1))
        self.assertEqual(result.iloc[0, 0], expected)

    def test_salary_summary_variants_match_real_department_data(self):
        for department, employees in self.employees.groupby("Department"):
            for phrase in ("avrage salry", "give avrage salry", "average salary", "mean pay"):
                with self.subTest(department=department, phrase=phrase), patch.object(
                    self.chatbot, "generate_employee_query", side_effect=AssertionError("Unexpected AI call")
                ):
                    result = self.query(f"{phrase} of {department}")
                    self.assertEqual(result.columns.tolist(), ["average_salary"])
                    self.assertAlmostEqual(result.iloc[0, 0], employees["Salary"].mean())
        for phrase, aggregate in (("total", "sum"), ("minimum", "min"), ("maximum", "max")):
            result = self.query(f"{phrase} salary of Sales")
            salaries = self.employees.loc[self.employees["Department"] == "Sales", "Salary"]
            self.assertAlmostEqual(result.iloc[0, 0], salaries.agg(aggregate))
        result = self.query("average salary by department")
        self.assertEqual(
            result.set_index("department")["average_salary"].to_dict(),
            self.employees.groupby("Department")["Salary"].mean().to_dict(),
        )

    def test_salary_typo_ui_shows_only_the_correct_answer(self):
        app = employee_app().run(timeout=30)
        for question in ("avrage salry of sales", "give avrage salry of sales"):
            app.chat_input[0].set_value(question).run(timeout=30)
            self.assertEqual(len(app.exception), 0)
            expected = self.employees.loc[self.employees["Department"] == "Sales", "Salary"].mean()
            self.assertIn(f"INR {expected:,.2f}", app.session_state["chat_history"][-1]["answer"])
        app.chat_input[0].set_value("avrage salry of enginering").run(timeout=30)
        self.assertEqual(len(app.exception), 0)
        self.assertIn("couldn't identify", app.session_state["chat_history"][-1]["answer"])
        self.assertEqual(len(app.code), 1)
        self.assertEqual(len(app.sidebar.code), 1)
        self.assertIn(":filter_0", app.sidebar.code[0].value)
        self.assertFalse(any("filter_" in message["answer"] for message in app.session_state["chat_history"]))
        self.assertFalse(any("filter_" in caption.value or "department =" in caption.value for caption in app.caption))

    def test_compound_salary_count_excludes_boundary_males_and_missing_values(self):
        fixture = self.employees.head(5).copy()
        fixture["Gender"] = ["Female", "Female", "Female", "Male", "Female"]
        fixture["Salary"] = [80000, 80001, 79999, 90000, None]
        self.engine.dispose()
        self.chatbot.get_engine.clear()
        with patch.object(self.chatbot.pd, "read_excel", return_value=fixture):
            self.engine = self.chatbot.get_engine()
        self.assertEqual(self.query("How many female employees have a salary above 80000?").iloc[0, 0], 1)
        self.assertEqual(self.query("How many female employees have a salary at least 80000?").iloc[0, 0], 2)

    def test_monthly_and_yearly_charts_match_dataset(self):
        dates = pd.to_datetime(self.employees["Joining Date"])
        for question, groups in (
            ("Show a line chart of monthly employee joinings", dates.dt.strftime("%Y-%m-01")),
            ("Show a chart of employee joining year", dates.dt.year),
        ):
            with self.subTest(question=question):
                result = self.query(question)
                actual = result.set_index("chart_group")["chart_value"].to_dict()
                self.assertEqual(actual, groups.value_counts().sort_index().to_dict())

    def test_joining_year_count_matches_dataset_and_date_boundaries(self):
        result = self.query("How many employees joined during 2023?")
        expected = (pd.to_datetime(self.employees["Joining Date"]).dt.year == 2023).sum()
        self.assertEqual(result.iloc[0, 0], expected)
        fixture = self.employees.head(5).copy()
        fixture["Joining Date"] = ["2022-12-31", "2023-01-01", "2023-12-31", "2024-01-01", None]
        self.engine.dispose()
        self.chatbot.get_engine.clear()
        with patch.object(self.chatbot.pd, "read_excel", return_value=fixture):
            self.engine = self.chatbot.get_engine()
        self.assertEqual(self.query("How many employees joined during 2023?").iloc[0, 0], 2)

    def test_department_queries_and_rankings_execute(self):
        result = self.query("Which departments have headcount above average?", known_template=True)
        counts = self.employees.groupby("Department").size()
        self.assertEqual(set(result["department"]), set(counts[counts > counts.mean()].index))
        result = self.query("How many employees earn above their department average?", known_template=True)
        averages = self.employees.groupby("Department")["Salary"].transform("mean")
        self.assertEqual(result.iloc[0, 0], (self.employees["Salary"] > averages).sum())
        result = self.chatbot.translate_known_question("Show top 2 paid employees per department by salary")
        ranked = self.chatbot.run_query(result.sql, result.parameters)
        self.assertTrue((ranked.groupby("department").size() == 2).all())
        self.assertEqual(len(ranked), self.employees["Department"].nunique() * 2)

    def test_compiled_plans_preserve_limits_dates_and_bound_values(self):
        plan = {
            "understood": True, "unhandled_conditions": [], "columns": ["emp_id", "salary"],
            "order_by": [{"column": "salary", "direction": "desc"}], "limit": 4,
        }
        request = self.chatbot.compile_employee_plan(plan)
        result = self.chatbot.run_query(request.sql, request.parameters)
        self.assertEqual(result["salary"].tolist(), self.employees["Salary"].nlargest(4).tolist())
        date = pd.to_datetime(self.employees.iloc[0]["Joining Date"]).strftime("%Y-%m-%d")
        request = self.chatbot.compile_employee_plan({
            "understood": True, "unhandled_conditions": [],
            "metrics": [{"function": "count", "column": "*"}],
            "filter": {"column": "join_date", "op": "eq", "value": date},
        })
        self.assertEqual(
            self.chatbot.run_query(request.sql, request.parameters).iloc[0, 0],
            (pd.to_datetime(self.employees["Joining Date"]).dt.strftime("%Y-%m-%d") == date).sum(),
        )
        request = self.chatbot.compile_employee_plan({
            **plan, "filter": {"column": "emp_id", "op": "eq", "value": "' OR 1=1 --"},
        })
        self.assertTrue(self.chatbot.run_query(request.sql, request.parameters).empty)

    def test_grouped_and_per_group_plans(self):
        plans = [
            {
                "columns": ["department"], "group_by": ["department"],
                "metrics": [{"function": "count", "column": "*"}],
                "order_by": [{"column": "employee_count", "direction": "desc"}], "limit": 2,
            },
            {
                "columns": ["emp_id", "department", "salary"], "per_group": ["department"],
                "limit_per_group": 1, "limit": 3,
                "order_by": [{"column": "salary", "direction": "desc"}],
            },
        ]
        for plan in plans:
            with self.subTest(plan=plan):
                request = self.chatbot.compile_employee_plan({
                    "understood": True, "unhandled_conditions": [], **plan,
                })
                result = self.chatbot.run_query(request.sql, request.parameters)
                self.assertEqual(len(result), plan["limit"])
                self.assertTrue(result["department"].is_unique)

    def test_demo_database_rejects_writes_and_supports_parallel_reads(self):
        with self.engine.connect() as connection:
            with self.assertRaises(SQLAlchemyError):
                connection.exec_driver_sql("DELETE FROM Employees")
        with ThreadPoolExecutor(max_workers=4) as executor:
            counts = list(executor.map(
                lambda _: self.chatbot.run_query("SELECT COUNT(*) FROM Employees").iloc[0, 0],
                range(12),
            ))
        self.assertEqual(counts, [len(self.employees)] * 12)

    def test_invalid_demo_dataset_fails_explicitly(self):
        self.chatbot.get_engine.clear()
        with patch.object(self.chatbot.pd, "read_excel", return_value=pd.DataFrame({"Name": ["Demo"]})):
            with self.assertRaisesRegex(ValueError, "Demo dataset is missing columns"):
                self.chatbot.get_engine()

    def test_report_has_real_data_and_generates_pdf(self):
        question = "Create a PDF report comparing department pay"
        result = self.query(question)
        self.assertEqual(len(result), len(self.employees))
        report = self.chatbot.build_business_report(question, result)
        pdf = self.chatbot.generate_business_report_pdf(question, report, None)
        self.assertTrue(pdf.startswith(b"%PDF-"))
        self.assertGreater(len(pdf), 1000)

    def test_cloud_demo_ui_answers_without_sql_server_or_azure(self):
        app = employee_app().run(timeout=30)
        self.assertEqual(len(app.exception), 0)
        self.assertFalse(any("portfolio demo" in caption.value.lower() for caption in app.caption))
        self.assertFalse(any("credentials" in caption.value.lower() for caption in app.caption))
        self.assertTrue(app.get("audio_input")[0].disabled)
        self.assertTrue(app.toggle[0].disabled)
        self.assertEqual(len(app.caption), 0)
        app.chat_input[0].set_value("How many employees do we have?").run(timeout=30)
        self.assertEqual(len(app.exception), 0)
        self.assertIn(f"{len(self.employees)} employees", app.session_state["chat_history"][0]["answer"])
        first = app.session_state["chat_history"][0]
        first_events = list(first["backend_log"])
        self.assertEqual(first_events[0]["route"], "sqlite_demo")
        self.assertTrue(any(event["event"] == "query_completed" and event["rows"] == 1 for event in first_events))
        self.assertTrue(any(event["event"] == "export_completed" for event in first_events))
        app.chat_input[0].set_value("How many female employees have a salary above 80000?").run(timeout=30)
        expected = ((self.employees["Gender"].str.lower() == "female") & (self.employees["Salary"] > 80000)).sum()
        self.assertEqual(len(app.exception), 0)
        self.assertIn(str(expected), app.session_state["chat_history"][1]["answer"])
        self.assertIn("gender = female", app.session_state["chat_history"][1]["request"].explanation)
        self.assertIn("salary > 80000", app.session_state["chat_history"][1]["request"].explanation)
        app.chat_input[0].set_value("How many employees joined during 2023?").run(timeout=30)
        self.assertEqual(len(app.exception), 0)
        expected = (pd.to_datetime(self.employees["Joining Date"]).dt.year == 2023).sum()
        self.assertEqual(app.session_state["chat_history"][2]["answer"], f"Employee count: {expected}.")
        self.assertFalse(any("join date" in caption.value or "gender =" in caption.value for caption in app.caption))
        self.assertEqual(len(app.code), 1)
        self.assertEqual(app.session_state["chat_history"][0]["backend_log"], first_events)
        events = [json.loads(line) for line in app.code[0].value.splitlines()]
        self.assertEqual(len({event["request_id"] for event in events}), 3)
        self.assertNotIn("80000", app.code[0].value)
        self.assertEqual(len(app.sidebar.code), 1)
        self.assertIn("SELECT", app.sidebar.code[0].value)
        self.assertFalse(any("SELECT" in message["answer"] for message in app.session_state["chat_history"]))
        self.assertTrue(any(button.label == "Download backend log" for button in app.get("download_button")))
        self.assertTrue(all(caption.value.startswith("Sources used:") for caption in app.caption))
        self.assertTrue(any("SQLite demo" in caption.value for caption in app.caption))

    def test_backend_status_and_connection_check_show_real_demo_count(self):
        app = employee_app().run(timeout=30)
        self.assertEqual(len(app.caption), 0)
        app.button(key="test_employee_connection").click().run(timeout=30)
        self.assertEqual(len(app.exception), 0)
        self.assertTrue(any(f"{len(self.employees):,} employee rows" in status.value for status in app.success))
        self.assertEqual(app.session_state["employee_connection_log"][0]["operation"], "connection_check")
        self.assertEqual(app.session_state["employee_connection_log"][-1]["event"], "request_completed")
        app.run(timeout=30)
        self.assertTrue(any("Verified SQLite demo" in status.value for status in app.success))
        with patch("pandas.read_sql", side_effect=SQLAlchemyError("private connection details")):
            app.button(key="test_employee_connection").click().run(timeout=30)
        self.assertEqual(len(app.exception), 0)
        self.assertTrue(any("Could not connect" in status.value for status in app.error))
        self.assertFalse(any("private connection details" in status.value for status in app.error))
        self.assertFalse(any("Verified" in status.value for status in app.success))
        self.assertTrue(any(event["event"] == "query_failed" for event in app.session_state["employee_connection_log"]))
        self.assertNotIn("private connection details", app.code[0].value)

    def test_sql_server_trace_includes_sql_and_retry_without_parameter_values_or_cells(self):
        events = []
        token = self.chatbot.SQL_TRACE.set(BackendTrace(events))
        self.addCleanup(self.chatbot.SQL_TRACE.reset, token)
        with patch.object(self.chatbot, "DATA_MODE", "sqlserver"), patch.object(
            self.chatbot, "get_engine", return_value=self.engine,
        ) as get_engine, patch("pandas.read_sql", side_effect=[
            SQLAlchemyError("PRIVATE_CONNECTION_STRING"),
            pd.DataFrame({"private_column": ["PRIVATE_RESULT"]}),
        ]) as read:
            result = self.chatbot.run_query("SELECT :private_literal", {"private_literal": "PRIVATE_PARAMETER"})
        self.assertEqual(result.iloc[0, 0], "PRIVATE_RESULT")
        self.assertEqual(read.call_count, 2)
        get_engine.clear.assert_called_once()
        self.assertEqual([event["event"] for event in events], ["query_started", "query_retry", "query_completed"])
        self.assertEqual(events[0]["route"], "sql_server")
        self.assertEqual(events[0]["implementation"][0]["function"], "run_query")
        self.assertIn("pd.read_sql", events[0]["implementation"][0]["code"])
        self.assertEqual(events[0]["sql"], "SELECT :private_literal")
        self.assertEqual(events[0]["parameters"]["private_literal"], {"type": "str", "value": "[REDACTED]"})
        self.assertEqual(events[1]["attempt"], 2)
        self.assertGreaterEqual(events[-1]["duration_ms"], 0)
        self.assertEqual(events[-1]["result_columns"], ["private_column"])
        log = diagnostic_log(events)
        for secret in ("PRIVATE_CONNECTION_STRING", "PRIVATE_RESULT", "PRIVATE_PARAMETER"):
            self.assertNotIn(secret, log)

    def test_employee_trace_includes_sql_and_result_details(self):
        message = {"question": "How many employees?", "request": None, "pending_query": True}
        self.assertIn("150 employees", self.chatbot.display_employee_message(message, "detailed_log_schema"))
        self.assertEqual(message["backend_log"][0]["route"], "sqlite_demo")
        self.assertEqual(message["backend_log"][-1]["event"], "request_completed")
        queries = [event for event in message["backend_log"] if event["event"] == "query_started"]
        self.assertTrue(queries)
        self.assertTrue(all("SELECT" in event["sql"] for event in queries))
        implementation = {
            item["function"]: item["code"]
            for event in message["backend_log"] for item in event.get("implementation", [])
        }
        for function in ("run_query", "compile_data_plan", "deterministic_response", "generate_excel_download"):
            self.assertIn(function, implementation)
            self.assertIn(f"def {function}(", implementation[function])

    def test_sql_failure_is_friendly_traced_and_context_is_reset(self):
        message = {"question": "How many employees?", "request": None, "pending_query": True}
        with patch("pandas.read_sql", side_effect=SQLAlchemyError("PRIVATE_CONNECTION_STRING")), patch.object(
            self.chatbot, "get_employee_names", return_value=(),
        ):
            self.assertIsNone(self.chatbot.display_employee_message(message, "failed"))
        self.assertIsNone(self.chatbot.SQL_TRACE.get())
        self.assertTrue(any(event["event"] == "query_failed" for event in message["backend_log"]))
        self.assertEqual(message["backend_log"][-1]["event"], "request_failed")
        self.assertNotIn("PRIVATE_CONNECTION_STRING", diagnostic_log(message["backend_log"]))

    def test_sql_ai_plan_cache_rejection_and_logging_failure_are_explicit(self):
        events = []
        trace = BackendTrace(events)
        token = self.chatbot.SQL_TRACE.set(trace)
        self.addCleanup(self.chatbot.SQL_TRACE.reset, token)
        self.chatbot.st.session_state["employee_query_plan_cache"] = {}
        plan = {
            "understood": True, "unhandled_conditions": [], "columns": [],
            "metrics": [{"function": "count", "column": "*"}],
            "filter": {"column": "name", "op": "eq", "value": "PRIVATE_NAME"},
            "group_by": [], "order_by": [], "limit": 100,
        }
        choice = Mock(finish_reason="stop")
        choice.message.content = json.dumps(plan)
        client = Mock()
        client.chat.completions.create.return_value.choices = [choice]
        with patch.dict(os.environ, {"AZURE_OPENAI_DEPLOYMENT_MINI": "test"}), patch.object(
            self.chatbot, "get_azure_openai_client", return_value=client,
        ):
            for _ in range(2):
                self.assertTrue(self.chatbot.generate_employee_query("How many employees named PRIVATE_NAME?").sql)
            self.assertEqual(client.chat.completions.create.call_count, 1)
            choice.message.content = '{"understood": false, "unhandled_conditions": ["PRIVATE_REASON"]}'
            self.assertIsNone(self.chatbot.generate_employee_query("Unsupported PRIVATE_QUESTION"))
        self.assertEqual([event["event"] for event in events], [
            "planner_requested", "planner_completed", "cache_hit", "planner_requested", "planner_rejected",
        ])
        for secret in ("PRIVATE_NAME", "PRIVATE_REASON", "PRIVATE_QUESTION"):
            self.assertNotIn(secret, diagnostic_log(events))
        message = {"question": "How many employees?", "request": None, "pending_query": True}
        with patch("backend_diagnostics.backend_logger", side_effect=OSError("PRIVATE_PATH")):
            self.assertIsNotNone(self.chatbot.display_employee_message(message, "log_failure"))
        self.assertIn("could not be saved", message["backend_log_error"])
        self.assertEqual(message["backend_log"][-1]["event"], "request_completed")
        self.assertNotIn("PRIVATE_PATH", diagnostic_log(message["backend_log"]))

    def test_cloud_voice_controls_enable_when_speech_settings_are_present(self):
        with patch.dict(os.environ, {"AZURE_SPEECH_KEY": "test-key", "AZURE_SPEECH_REGION": "eastus"}):
            app = employee_app().run(timeout=30)
        self.assertEqual(len(app.exception), 0)
        self.assertFalse(app.get("audio_input")[0].disabled)
        self.assertFalse(app.toggle[0].disabled)
        self.assertFalse(any("Voice is not enabled" in caption.value for caption in app.caption))


if __name__ == "__main__":
    unittest.main()
