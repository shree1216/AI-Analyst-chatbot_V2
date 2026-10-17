import importlib
import os
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path
import unittest
from unittest.mock import patch

import pandas as pd
from sqlalchemy.exc import SQLAlchemyError
from streamlit.testing.v1 import AppTest


ROOT = Path(__file__).resolve().parents[1]
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
        app = AppTest.from_file(str(ROOT / "employee_sql_chatbot.py")).run(timeout=30)
        self.assertEqual(len(app.exception), 0)
        self.assertFalse(any("portfolio demo" in caption.value.lower() for caption in app.caption))
        self.assertFalse(any("credentials" in caption.value.lower() for caption in app.caption))
        self.assertTrue(app.get("audio_input")[0].disabled)
        self.assertTrue(app.toggle[0].disabled)
        self.assertTrue(any("Voice is not enabled" in caption.value for caption in app.caption))
        app.chat_input[0].set_value("How many employees do we have?").run(timeout=30)
        self.assertEqual(len(app.exception), 0)
        self.assertIn(f"{len(self.employees)} employees", app.session_state["chat_history"][0]["answer"])
        app.chat_input[0].set_value("How many female employees have a salary above 80000?").run(timeout=30)
        expected = ((self.employees["Gender"].str.lower() == "female") & (self.employees["Salary"] > 80000)).sum()
        self.assertEqual(len(app.exception), 0)
        self.assertIn(str(expected), app.session_state["chat_history"][1]["answer"])
        self.assertIn("gender = female", app.session_state["chat_history"][1]["request"].explanation)
        self.assertIn("salary > 80000", app.session_state["chat_history"][1]["request"].explanation)

    def test_cloud_voice_controls_enable_when_speech_settings_are_present(self):
        with patch.dict(os.environ, {"AZURE_SPEECH_KEY": "test-key", "AZURE_SPEECH_REGION": "eastus"}):
            app = AppTest.from_file(str(ROOT / "employee_sql_chatbot.py")).run(timeout=30)
        self.assertEqual(len(app.exception), 0)
        self.assertFalse(app.get("audio_input")[0].disabled)
        self.assertFalse(app.toggle[0].disabled)
        self.assertFalse(any("Voice is not enabled" in caption.value for caption in app.caption))


if __name__ == "__main__":
    unittest.main()
