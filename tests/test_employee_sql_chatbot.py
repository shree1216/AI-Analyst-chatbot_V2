import importlib
import json
from io import BytesIO
import os
from pathlib import Path
import smtplib
import unittest
from unittest.mock import patch

import pandas as pd
from streamlit.testing.v1 import AppTest


ROOT = Path(__file__).resolve().parents[1]
PAY_QUESTION = "Create a PDF report comparing the highest and lowest department pay"
TEST_ENVIRONMENT = {
    name: value for name, value in os.environ.items()
    if not name.upper().startswith(("AZURE_", "SMTP_"))
}
SMTP_SETTINGS = {
    "SMTP_HOST": "smtp.example.com",
    "SMTP_PORT": "587",
    "SMTP_USERNAME": "sender@example.com",
    "SMTP_APP_PASSWORD": "test-only-password",
    "SMTP_FROM_EMAIL": "sender@example.com",
}


def employee_fixture():
    return pd.DataFrame(
        {
            "emp_id": [1, 2, 3, 4, 5, 6],
            "department": ["Finance", "Finance", "IT", "IT", "Sales", "Sales"],
            "city": ["Delhi", "Mumbai", "Delhi", "Mumbai", "Delhi", "Mumbai"],
            "salary": [100000, 120000, 60000, 80000, 80000, 100000],
            "status": ["Active", "On Leave", "Active", "Inactive", "Active", "Active"],
        }
    )


def employee_query_plan(**changes):
    plan = {
        "understood": True, "unhandled_conditions": [], "columns": [],
        "metrics": [{"function": "count", "column": "*"}], "filter": None,
    }
    plan.update(changes)
    return plan


class WorkforceTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        with patch("dotenv.load_dotenv"), patch.dict(os.environ, TEST_ENVIRONMENT, clear=True):
            cls.chatbot = importlib.import_module("employee_sql_chatbot")

    def setUp(self):
        self.chatbot.st.session_state.pop("employee_query_plan_cache", None)
        self.enterContext(patch.dict(os.environ, TEST_ENVIRONMENT, clear=True))
        self.enterContext(patch("dotenv.load_dotenv"))
        self.enterContext(patch("socket.create_connection", side_effect=AssertionError("Unexpected network access")))
        self.result = employee_fixture()

    def test_chart_request_routes_before_generic_count(self):
        request = self.chatbot.translate_question("Show a bar chart of employee count by department")
        self.assertEqual(request.response_kind, "chart_employee_count")
        self.assertIn("GROUP BY [department]", request.sql)

    def test_job_title_question_reaches_schema_query_path(self):
        expected = self.chatbot.SqlRequest(
            "SELECT COUNT(*) AS employee_count FROM employee_data WHERE job_title = :job_title",
            {"job_title": "Site engineer"}, "Count Site engineers", response_kind="column_query",
        )
        with patch.object(self.chatbot, "generate_employee_query", return_value=expected) as generate:
            request = self.chatbot.translate_question("how many Site engineer we have")
        generate.assert_called_once_with("how many Site engineer we have")
        self.assertIs(request, expected)

    def test_unresolved_column_question_never_becomes_total_count(self):
        with patch.object(self.chatbot, "generate_employee_query", return_value=None):
            request = self.chatbot.translate_question("count employees aged over 35 with missing phone numbers")
        self.assertEqual(request.response_kind, "query_clarification")
        self.assertEqual(request.sql, "")

    def test_every_employee_column_supports_bound_filters(self):
        values = {
            "emp_id": "EMP001", "name": "Some Person", "age": 30, "gender": "Female",
            "department": "Sales", "city": "Delhi", "status": "Active", "job_title": "Site engineer",
            "salary": 75000, "join_date": "2024-01-01", "email": "person@example.org", "phone_number": "555-0100",
        }
        for column, value in values.items():
            with self.subTest(column=column):
                request = self.chatbot.compile_employee_plan(employee_query_plan(filter={"column": column, "op": "eq", "value": value}))
                self.assertIn(f"[{column}]", request.sql.split(" WHERE ")[1])
                self.assertEqual(list(request.parameters.values()), [value.lower() if isinstance(value, str) else value])
                self.assertEqual(request.response_kind, "column_query")

    def test_compound_filters_preserve_and_or_dates_numbers_and_missing_values(self):
        request = self.chatbot.compile_employee_plan(employee_query_plan(filter={"all": [
            {"column": "age", "op": "gt", "value": 35},
            {"column": "join_date", "op": "between", "value": ["2020-01-01", "2024-12-31"]},
            {"any": [
                {"column": "status", "op": "in", "value": ["Active", "On Leave"]},
                {"column": "phone_number", "op": "missing", "value": None},
            ]},
        ]}))
        self.assertIn("[age] > :filter_0", request.sql)
        self.assertIn("BETWEEN :filter_1 AND :filter_2", request.sql)
        self.assertIn(" OR ", request.sql)
        self.assertIn("[phone_number] IS NULL", request.sql)
        self.assertIn(" = ''", request.sql)

    def test_string_filters_escape_sql_wildcards_and_bind_injection_text(self):
        value = "x%_['; DROP TABLE Employees;--"
        request = self.chatbot.compile_employee_plan(employee_query_plan(filter={"column": "name", "op": "contains", "value": value}))
        self.assertNotIn("DROP TABLE", request.sql)
        self.assertIn("ESCAPE '!'", request.sql)
        self.assertIn("!%", request.parameters["filter_0"])
        self.assertIn("!_", request.parameters["filter_0"])
        self.assertIn("![", request.parameters["filter_0"])

    def test_grouped_aggregates_and_top_per_group_are_compiled(self):
        request = self.chatbot.compile_employee_plan(employee_query_plan(
            columns=["department"], group_by=["department"],
            metrics=[{"function": "count", "column": "*"}, {"function": "avg", "column": "salary"}],
            order_by=[{"column": "average_salary", "direction": "desc"}],
        ))
        self.assertIn("GROUP BY [department]", request.sql)
        self.assertIn("AVG(CAST([salary] AS FLOAT))", request.sql)
        ranked = self.chatbot.compile_employee_plan(employee_query_plan(
            columns=["emp_id", "name", "department", "salary"], metrics=[], per_group=["department"],
            limit_per_group=3, order_by=[{"column": "salary", "direction": "desc"}],
        ))
        self.assertIn("ROW_NUMBER() OVER (PARTITION BY [department]", ranked.sql)
        self.assertEqual(ranked.parameters["rank_limit"], 3)

    def test_invalid_or_partially_understood_plans_are_rejected(self):
        for changes in (
            {"understood": False}, {"unhandled_conditions": ["Unknown salary condition"]},
            {"filter": {"column": "rating", "op": "eq", "value": 5}},
            {"filter": {"column": "age", "op": "eq", "value": "thirty"}},
            {"filter": {"column": "join_date", "op": "eq", "value": "2024-99-99"}},
            {"filter": {"column": "salary", "op": "execute", "value": 1}},
            {"columns": ["name"], "group_by": []}, {"limit": 100000},
            {"order_by": [{"column": "employee_count;DROP", "direction": "desc"}]},
            {"sql": "DELETE FROM dbo.Employees"},
        ):
            with self.subTest(changes=changes), self.assertRaises(ValueError):
                self.chatbot.compile_employee_plan(employee_query_plan(**changes))

    def test_query_planner_parses_json_and_never_uses_model_sql(self):
        plan = employee_query_plan(filter={"column": "job_title", "op": "eq", "value": "Site engineer"})
        with patch.dict(os.environ, {"AZURE_OPENAI_DEPLOYMENT_MINI": "test"}), patch.object(self.chatbot, "get_azure_openai_client") as client:
            choice = client.return_value.chat.completions.create.return_value.choices[0]
            choice.message.content = json.dumps(plan)
            choice.finish_reason = "stop"
            request = self.chatbot.generate_employee_query("how many Site engineer we have")
            self.assertEqual(request.parameters, {"filter_0": "site engineer"})
            self.chatbot.st.session_state.pop("employee_query_plan_cache", None)
            choice.message.content = json.dumps(employee_query_plan())
            self.assertIsNone(self.chatbot.generate_employee_query("how many Site engineer we have"))
            choice.message.content = "SELECT COUNT(*) FROM dbo.Employees"
            self.assertIsNone(self.chatbot.generate_employee_query("how many Site engineer we have"))

    def test_incomplete_counts_and_filtered_profiles_use_schema_path(self):
        for question in (
            "count employees with status", "count employees with", "count of emails",
            "give details of employees aged over 35", "How old is Belinda Mccullough?",
        ):
            with self.subTest(question=question), patch.object(self.chatbot, "generate_employee_query", return_value=None) as generate:
                request = self.chatbot.translate_question(question)
                generate.assert_called_once_with(question)
                self.assertEqual(request.response_kind, "query_clarification")

    def test_simple_status_count_is_filtered_and_does_not_call_ai(self):
        request = self.chatbot.translate_question("give count of employee having status Active")
        self.assertEqual(request.parameters, {"status": "Active"})
        with patch.object(self.chatbot, "summarize_with_azure_openai", side_effect=AssertionError("Unexpected AI call")):
            answer = self.chatbot.natural_language_response(pd.DataFrame({"employee_count": [50]}), request, "count Active employees")
        self.assertIn("50", answer)
        self.assertIn("Active", answer)

    def test_unresolved_queries_do_not_execute_sql(self):
        request = self.chatbot.SqlRequest("", {}, "Please clarify the requested condition.", response_kind="query_clarification")
        with patch.object(self.chatbot, "run_query", side_effect=AssertionError("Unexpected query")):
            self.assertIn("clarify", self.chatbot.display_response(request))

    def test_email_suffix_count_is_filtered_before_answering(self):
        for question in (
            "give count of emails ending with .org",
            "i want count of employees ending with .org",
            "how many employees email is ending with .ORG?",
        ):
            with self.subTest(question=question):
                request = self.chatbot.translate_question(question)
                self.assertEqual(request.response_kind, "email_suffix_count")
                self.assertEqual(request.parameters["email_suffix"], ".org")
                self.assertIn("RIGHT(LOWER(RTRIM([email])), LEN(:email_suffix)) = :email_suffix", request.sql)

    def test_email_suffix_count_answer_uses_database_count_without_ai(self):
        request = self.chatbot.translate_question("count of employees ending with .org")
        with patch.object(self.chatbot, "summarize_with_azure_openai", side_effect=AssertionError("Unexpected AI call")):
            for count in (0, 1, 23):
                answer = self.chatbot.natural_language_response(
                    pd.DataFrame({"employee_count": [count]}), request, "count of employees ending with .org"
                )
                self.assertIn(str(count), answer)
                self.assertIn(".org", answer)

    def test_employee_details_route_without_ai_or_name_catalog(self):
        with patch.object(self.chatbot, "generate_employee_query") as generate, patch.object(
            self.chatbot, "get_employee_names"
        ) as names:
            request = self.chatbot.translate_question("give details of pooja singh")
        self.assertIsNotNone(request)
        self.assertEqual(request.response_kind, "employee_details")
        self.assertEqual(request.parameters, {"name": "pooja singh"})
        self.assertNotIn("pooja singh", request.sql)
        generate.assert_not_called()
        names.assert_not_called()

    def test_employee_profile_wording_and_name_normalization(self):
        for question in (
            "give details of pooja singh", "Please show me the details for Pooja   Singh.",
            "Could you provide employee information about Pooja Singh please?",
            "profile for 'Pooja Singh'",
        ):
            with self.subTest(question=question):
                request = self.chatbot.translate_question(question)
                self.assertEqual(request.response_kind, "employee_details")
                self.assertEqual(request.parameters, {"name": "pooja singh"})

    def test_profiles_accept_employee_ids_without_losing_leading_zeroes(self):
        for question in (
            "give details of EMP015", "give details of employee_id EMP015",
            "Show profile for Employee ID EMP015", "give details of emp_id 'emp015'",
        ):
            with self.subTest(question=question), patch.object(self.chatbot, "generate_employee_query") as generate:
                request = self.chatbot.translate_question(question)
                self.assertEqual(request.response_kind, "employee_details")
                self.assertEqual(request.parameters, {"emp_id": "emp015"})
                self.assertIn("WHERE LOWER(LTRIM(RTRIM([emp_id]))) = :emp_id", request.sql)
                generate.assert_not_called()

    def test_id_not_found_message_does_not_call_id_a_name(self):
        request = self.chatbot.translate_question("give details of EMP999")
        answer = self.chatbot.deterministic_response(pd.DataFrame(), request)
        self.assertIn("employee ID EMP999", answer)
        self.assertNotIn("named", answer)

    def test_database_column_spellings_reach_the_query_planner(self):
        for column in ("Employee ID", "employee_id", "emp_id", "Name", "Age", "Gender", "Department", "City", "Status", "Job Title", "job_title", "Salary", "Joining Date", "joining_date", "Email", "Phone Number", "phone_number"):
            question = f"Show {column} for EMP015"
            with self.subTest(column=column), patch.object(self.chatbot, "generate_employee_query", return_value=None) as generate:
                request = self.chatbot.translate_question(question)
                generate.assert_called_once_with(question)
                self.assertEqual(request.response_kind, "query_clarification")

    def test_column_conditions_are_not_mistaken_for_employee_names(self):
        for question in (
            "give details of job_title Site engineer", "give details of salary above 70000",
            "give details of gender Female", "give details of EMP015 with status Active",
        ):
            with self.subTest(question=question), patch.object(self.chatbot, "generate_employee_query", return_value=None) as generate:
                request = self.chatbot.translate_question(question)
                generate.assert_called_once_with(question)
                self.assertEqual(request.response_kind, "query_clarification")

    def test_employee_profile_response_uses_database_fields_without_ai(self):
        request = self.chatbot.translate_question("give details of pooja singh")
        result = pd.DataFrame([{
            "emp_id": 42, "name": "Pooja Singh", "department": "Finance",
            "salary": 95000, "city": "Delhi", "email": None,
        }])
        with patch.object(self.chatbot, "summarize_with_azure_openai", side_effect=AssertionError("Unexpected AI call")):
            answer = self.chatbot.natural_language_response(result, request, "give details of pooja singh")
        for text in ("Name: Pooja Singh", "Employee ID: 42", "Salary: INR 95,000.00", "Email: not available"):
            self.assertIn(text, answer)

    def test_employee_profile_not_found_is_explicit(self):
        request = self.chatbot.translate_question("give details of pooja singh")
        answer = self.chatbot.deterministic_response(pd.DataFrame(), request)
        self.assertIn("No employee named pooja singh was found", answer)

    def test_duplicate_employee_profiles_are_not_silently_discarded(self):
        request = self.chatbot.translate_question("give details of pooja singh")
        result = pd.DataFrame([
            {"emp_id": 42, "name": "Pooja Singh", "department": "Finance"},
            {"emp_id": 43, "name": "Pooja Singh", "department": "IT"},
        ])
        answer = self.chatbot.deterministic_response(result, request)
        self.assertIn("2 matching employee records", answer)
        self.assertIn("Employee ID: 42", answer)
        self.assertIn("Employee ID: 43", answer)

    def test_ambiguous_profile_injection_does_not_execute_sql(self):
        with patch.object(self.chatbot, "generate_employee_query", return_value=None):
            request = self.chatbot.translate_question("give details of Robert'; DROP TABLE Employees;--")
        self.assertEqual(request.response_kind, "query_clarification")
        self.assertEqual(request.sql, "")
        with patch.object(self.chatbot, "run_query", side_effect=AssertionError("Unexpected query")):
            self.chatbot.display_response(request)

    def test_profile_apostrophes_are_bound_not_interpolated(self):
        request = self.chatbot.translate_question("give details of Sean O'Connor")
        self.assertEqual(request.parameters, {"name": "sean o'connor"})
        self.assertNotIn("O'Connor", request.sql)

    def test_salary_analysis_without_department_uses_distribution(self):
        request = self.chatbot.translate_question("I want the analysis on salaries.")
        self.assertEqual(request.response_kind, "business_report")
        report = self.chatbot.build_business_report("I want the analysis on salaries.", self.result)
        self.assertEqual(report.chart_type, "histogram")
        self.assertIn("salary", report.title.lower())
        chart = self.chatbot.create_report_chart(report)
        self.assertEqual(sum(chart.data[0].y), len(self.result))
        self.assertIn("salary", chart.layout.xaxis.title.text.lower())
        self.assertEqual(chart.layout.bargap, 0)

    def test_department_pay_comparison_uses_ranked_bars(self):
        request = self.chatbot.translate_question("Analyze highest and lowest department salaries")
        self.assertEqual(request.response_kind, "business_report")
        report = self.chatbot.build_business_report("Analyze highest and lowest department salaries", self.result)
        self.assertEqual(report.chart_type, "bar")
        chart = self.chatbot.create_report_chart(report)
        self.assertEqual(list(chart.data[0].y), ["Finance", "Sales", "IT"])
        self.assertEqual(list(chart.data[0].x), [110000, 90000, 70000])

    def test_salary_chart_uses_requested_city_grouping(self):
        report = self.chatbot.build_business_report("Analyze average salaries by city", self.result)
        chart = self.chatbot.create_report_chart(report)
        self.assertEqual(list(chart.data[0].y), ["Mumbai", "Delhi"])
        self.assertEqual(list(chart.data[0].x), [100000, 80000])

    def test_salary_histogram_pdf_counts_match_screen(self):
        report = self.chatbot.build_business_report("I want the analysis on salaries.", self.result)
        chart = self.chatbot.create_report_chart(report)
        drawing = self.chatbot.report_pdf_chart(report)
        self.assertEqual(list(drawing.contents[-1].data[0]), list(chart.data[0].y))
        self.assertTrue(self.chatbot.generate_business_report_pdf("Salary analysis", report, None).startswith(b"%PDF-"))

    def test_salary_distribution_handles_constant_missing_and_empty_values(self):
        constant = self.chatbot.salary_distribution(pd.Series([50000, 50000, None]))
        self.assertEqual(constant["Headcount"].tolist(), [2])
        self.assertEqual(constant["Salary band (INR)"].tolist(), ["50,000.00"])
        report = self.chatbot.build_business_report("Salary analysis", self.result.assign(salary=None))
        self.assertIsNone(self.chatbot.create_report_chart(report))
        self.assertIn("No recorded salaries", report.executive_summary)

    def test_salary_report_missing_requested_group_does_not_substitute_headcount(self):
        report = self.chatbot.build_business_report("Analyze average salaries by city", self.result.drop(columns="city"))
        self.assertIsNone(self.chatbot.create_report_chart(report))
        self.assertIn("not available", report.executive_summary)

    def test_explicit_column_chart_overrides_salary_default(self):
        report = self.chatbot.build_business_report("Create an average department salaries report with a column chart", self.result)
        chart = self.chatbot.create_report_chart(report)
        self.assertEqual(chart.data[0].orientation, "v")
        self.assertEqual(list(chart.data[0].y), [110000, 90000, 70000])

    def test_standalone_salary_histogram_uses_salary_records(self):
        request = self.chatbot.translate_question("Show a histogram of salaries")
        self.assertEqual(request.response_kind, "chart_salary_distribution")
        self.assertTrue(request.sql.startswith("SELECT [salary] FROM"))

    def test_image_can_be_resubmitted_after_failure_and_with_new_question(self):
        output = BytesIO()
        self.chatbot.Image.new("RGB", (32, 32), "red").save(output, format="PNG")
        image_bytes = output.getvalue()
        with patch.object(self.chatbot, "analyze_uploaded_image", side_effect=[None, "Red image", "32 pixels"]) as analyze:
            first = self.chatbot.process_image_submission(image_bytes, "test.png", "Describe it")
            second = self.chatbot.process_image_submission(image_bytes, "test.png", "Describe it")
            third = self.chatbot.process_image_submission(image_bytes, "test.png", "What size?")
        self.assertIsNone(first["answer"])
        self.assertIn("empty image analysis", first["error"])
        self.assertEqual(second["answer"], "Red image")
        self.assertEqual(third["answer"], "32 pixels")
        self.assertEqual(analyze.call_count, 3)
        self.assertEqual(analyze.call_args.args[2], "What size?")
        self.assertTrue(second["image_bytes"].startswith(b"\x89PNG"))

    def test_invalid_and_oversized_images_do_not_call_azure(self):
        with patch.object(self.chatbot, "analyze_uploaded_image") as analyze:
            self.assertIsNone(self.chatbot.process_image_submission(b"invalid", "bad.png", ""))
            self.assertIsNone(self.chatbot.process_image_submission(b"x" * (self.chatbot.MAX_IMAGE_BYTES + 1), "large.png", ""))
        analyze.assert_not_called()

    def test_vision_request_contains_high_detail_image_and_default_prompt(self):
        with patch.dict(os.environ, {"AZURE_OPENAI_DEPLOYMENT_MINI": "test-vision"}), patch.object(
            self.chatbot, "get_azure_openai_client"
        ) as client:
            client.return_value.chat.completions.create.return_value.choices[0].message.content = "Image description"
            client.return_value.chat.completions.create.return_value.choices[0].message.refusal = None
            answer = self.chatbot.analyze_uploaded_image(b"test-image", "image/png", "   ")
        request = client.return_value.chat.completions.create.call_args.kwargs
        content = request["messages"][-1]["content"]
        self.assertEqual(answer, "Image description")
        self.assertEqual(content[0]["text"], "Analyze this uploaded image.")
        self.assertEqual(content[1]["image_url"]["detail"], "high")
        self.assertEqual(content[1]["image_url"]["url"], "data:image/png;base64,dGVzdC1pbWFnZQ==")
        self.assertEqual(request["timeout"], 60.0)
        self.assertEqual(request["max_completion_tokens"], 2000)

    def test_image_service_failure_can_be_retried_without_reupload(self):
        message = {"image_bytes": b"test", "question": "Describe it", "answer": None}
        with patch.object(self.chatbot, "analyze_uploaded_image", side_effect=[
            self.chatbot.ImageAnalysisError("Azure image analysis timed out."), "Image description",
        ]) as analyze:
            self.chatbot.complete_image_analysis(message)
            self.assertIn("timed out", message["error"])
            self.chatbot.complete_image_analysis(message)
        self.assertEqual(message["answer"], "Image description")
        self.assertNotIn("error", message)
        self.assertEqual(analyze.call_count, 2)

    def test_vision_empty_response_budget_is_distinguishable(self):
        with patch.dict(os.environ, {"AZURE_OPENAI_DEPLOYMENT_MINI": "test"}), patch.object(
            self.chatbot, "get_azure_openai_client"
        ) as client:
            choice = client.return_value.chat.completions.create.return_value.choices[0]
            choice.message.content = None
            choice.message.refusal = None
            choice.finish_reason = "length"
            with self.assertRaisesRegex(self.chatbot.ImageAnalysisError, "response budget"):
                self.chatbot.analyze_uploaded_image(b"test", "image/png", "")

    def test_vision_unexpected_error_does_not_expose_details(self):
        with patch.dict(os.environ, {"AZURE_OPENAI_DEPLOYMENT_MINI": "test"}), patch.object(
            self.chatbot, "get_azure_openai_client"
        ) as client:
            client.return_value.chat.completions.create.side_effect = RuntimeError("private-token-and-request")
            with self.assertRaises(self.chatbot.ImageAnalysisError) as caught:
                self.chatbot.analyze_uploaded_image(b"test", "image/png", "")
        self.assertNotIn("private-token", str(caught.exception))

    def test_image_retry_ui_updates_existing_message_once(self):
        image = BytesIO()
        self.chatbot.Image.new("RGB", (32, 32), "red").save(image, format="PNG")
        self.chatbot.get_azure_openai_client.clear()
        settings = {
            "AZURE_OPENAI_ENDPOINT": "https://example.test", "AZURE_OPENAI_API_KEY": "test",
            "AZURE_OPENAI_API_VERSION": "test", "AZURE_OPENAI_DEPLOYMENT_MINI": "test-vision",
        }
        try:
            with patch.dict(os.environ, settings), patch("openai.AzureOpenAI") as client:
                choice = client.return_value.chat.completions.create.return_value.choices[0]
                choice.message.content = "A red square."
                choice.message.refusal = None
                choice.finish_reason = "stop"
                app = AppTest.from_file(str(ROOT / "employee_sql_chatbot.py")).run(timeout=30)
                app.session_state["chat_history"] = [{
                    "kind": "image", "image_name": "test.png", "image_bytes": image.getvalue(),
                    "question": "Describe it", "answer": None, "error": "Azure image analysis timed out.",
                }]
                app.run(timeout=30)
                self.assertEqual(len(app.exception), 0)
                self.assertIn("timed out", app.warning[0].value)
                app.button(key="retry_image_0").click().run(timeout=30)
                self.assertEqual(len(app.exception), 0)
                self.assertEqual(len(app.session_state["chat_history"]), 1)
                self.assertEqual(app.session_state["chat_history"][0]["answer"], "A red square.")
                self.assertNotIn("error", app.session_state["chat_history"][0])
                app.run(timeout=30)
                client.return_value.chat.completions.create.assert_called_once()
        finally:
            self.chatbot.get_azure_openai_client.clear()

    def test_pdf_request_preserves_sales_filter(self):
        request = self.chatbot.translate_question("Create a Sales department PDF")
        self.assertEqual(request.response_kind, "business_report")
        self.assertEqual(request.parameters, {"department": "Sales"})
        self.assertIn("WHERE [department] = :department", request.sql)

    def test_status_pie_includes_all_named_statuses(self):
        request = self.chatbot.translate_question(
            "Create a pie chart showing the proportion of Active, On Leave and Inactive employees"
        )
        self.assertEqual(request.response_kind, "chart_employee_count")
        self.assertNotIn("status", request.parameters)
        self.assertIn("GROUP BY [status]", request.sql)

    def test_chart_types_follow_explicit_request(self):
        result = pd.DataFrame({"group": ["IT", "Sales"], "value": [3, 4]})
        for kind, trace_type, orientation in (
            ("pie", "pie", None), ("donut", "pie", None),
            ("bar", "bar", "h"), ("column", "bar", "v"),
            ("line", "scatter", None),
        ):
            with self.subTest(kind=kind):
                chart = self.chatbot.create_query_chart(result, "Employees", "Employees", kind)
                self.assertEqual(chart.data[0].type, trace_type)
                if orientation:
                    self.assertEqual(chart.data[0].orientation, orientation)

    def test_monthly_chart_uses_join_dates_and_chronological_order(self):
        request = self.chatbot.translate_question("Show a line chart of monthly employee joinings")
        self.assertIn("DATEFROMPARTS", request.sql)
        self.assertIn("ORDER BY chart_group", request.sql)

    def test_status_report_preserves_requested_pie_chart(self):
        report = self.chatbot.build_business_report(
            "Create a PDF report with a pie chart of Active, On Leave and Inactive employees",
            self.result,
        )
        chart = self.chatbot.create_report_chart(report)
        self.assertEqual(chart.data[0].type, "pie")
        self.assertEqual(set(chart.data[0].labels), {"Active", "On Leave", "Inactive"})
        self.assertEqual(sum(chart.data[0].values), len(self.result))

    def test_automatic_chart_selection(self):
        for question, dimension, expected in (
            ("Show employee proportions", "department", "pie"),
            ("Chart employee status", "status", "pie"),
            ("Chart monthly joinings", "join_date", "line"),
            ("Chart employees by job title", "job_title", "bar"),
            ("Chart average salary by department", "department", "bar"),
        ):
            with self.subTest(question=question):
                self.assertEqual(self.chatbot.chart_type_for(question, dimension), expected)

    def test_chart_keeps_explicit_single_status_filter(self):
        request = self.chatbot.translate_question("Show a column chart of Active employees by department")
        self.assertEqual(request.parameters, {"status": "Active"})
        self.assertIn("GROUP BY [department]", request.sql)

    def test_temporal_report_sorted_by_date(self):
        result = self.result.assign(join_date=["2024-03-01", "2024-01-01", "2024-03-02", "2024-02-01", "2024-01-02", "2024-03-03"])
        report = self.chatbot.build_business_report("Monthly joinings PDF report", result)
        chart = self.chatbot.create_report_chart(report)
        self.assertEqual(chart.data[0].type, "scatter")
        self.assertEqual(list(chart.data[0].x), ["2024-01", "2024-02", "2024-03"])
        self.assertEqual(list(chart.data[0].y), [2, 1, 3])
        self.assertIn("largest group is 2024-03", report.executive_summary)

    def test_pdf_chart_types_match_request(self):
        for kind, expected_class in (
            ("pie", "Pie"), ("donut", "Pie"), ("bar", "HorizontalBarChart"),
            ("column", "VerticalBarChart"), ("line", "HorizontalLineChart"),
        ):
            with self.subTest(kind=kind):
                question = f"Create a department report with a {kind} chart"
                report = self.chatbot.build_business_report(question, self.result)
                drawing = self.chatbot.report_pdf_chart(report)
                self.assertEqual(type(drawing.contents[-1]).__name__, expected_class)
                pdf = self.chatbot.generate_business_report_pdf(question, report, None)
                self.assertTrue(pdf.startswith(b"%PDF-"))
                self.assertTrue(pdf.rstrip().endswith(b"%%EOF"))

    def test_long_chart_summary_is_concise(self):
        result = pd.DataFrame({"chart_group": [f"Group {index}" for index in range(30)], "chart_value": [2] * 30})
        request = self.chatbot.translate_question("Show a line chart of monthly employee joinings")
        answer = self.chatbot.deterministic_response(result, request)
        self.assertIn("60 employees across 30 groups", answer)
        self.assertLess(len(answer), 200)

    def test_pdf_table_retains_all_groups(self):
        table = pd.DataFrame({"Group": range(30), "Headcount": [2] * 30})
        self.assertEqual(len(self.chatbot.report_pdf_table(table)._cellvalues), 31)

    def test_pronoun_it_does_not_filter_department(self):
        request = self.chatbot.translate_question("Create a department pay report and email it to me")
        self.assertEqual(request.parameters, {})

    def test_explicit_it_department_remains_filterable(self):
        request = self.chatbot.translate_question("Create a report for the IT department")
        self.assertEqual(request.parameters, {"department": "IT"})

    def test_status_breakdown_keeps_all_named_statuses(self):
        request = self.chatbot.translate_question(
            "Create an employment status report for Active, On Leave and Inactive employees"
        )
        self.assertNotIn("status", request.parameters)

    def test_pay_report_calculates_department_extremes(self):
        report = self.chatbot.build_business_report(PAY_QUESTION, self.result)
        salary_table = next(table for heading, table, _ in report.tables if heading == "Department Salary Profile")
        self.assertEqual(salary_table["Department"].tolist(), ["Finance", "Sales", "IT"])
        self.assertEqual(salary_table["Average salary"].tolist(), [110000, 90000, 70000])
        self.assertIn("Finance has the highest average pay at INR 110,000.00", report.executive_summary)
        self.assertIn("IT has the lowest at INR 70,000.00", report.executive_summary)
        self.assertTrue(any("INR 40,000.00" in finding for finding in report.findings))

    def test_headcounts_reconcile_to_input_rows(self):
        report = self.chatbot.build_business_report("Department and location report", self.result)
        for _, table, _ in report.tables:
            self.assertEqual(int(table["Headcount"].sum()), len(self.result))
            self.assertAlmostEqual(table["Percentage of total"].sum(), 100, delta=0.2)

    def test_salary_chart_uses_average_pay(self):
        report = self.chatbot.build_business_report(PAY_QUESTION, self.result)
        chart = self.chatbot.create_report_chart(report)
        self.assertEqual(list(chart.data[0].x), [110000, 90000, 70000])
        self.assertEqual(chart.layout.xaxis.title.text, "Average salary (INR)")
        self.assertEqual(chart.data[0].texttemplate, "%{x:,.2f}")

    def test_pdf_generation(self):
        report = self.chatbot.build_business_report(PAY_QUESTION, self.result)
        pdf = self.chatbot.generate_business_report_pdf(PAY_QUESTION, report, None)
        self.assertTrue(pdf.startswith(b"%PDF-"))
        self.assertTrue(pdf.rstrip().endswith(b"%%EOF"))
        self.assertGreater(len(pdf), 1000)

    def test_copy_component_passes_text_as_data_with_independent_keys(self):
        answer = '<script>alert("test")</script>\nEmployee count: 6'
        with patch.object(self.chatbot, "copy_response_component") as component:
            self.chatbot.render_copy_response(answer, "first")
            self.chatbot.render_copy_response("Another answer", "second")
        self.assertEqual(component.call_args_list[0].kwargs, {
            "data": {"text": answer}, "key": "copy_response_first",
        })
        self.assertEqual(component.call_args_list[1].kwargs["key"], "copy_response_second")

    def test_report_copy_includes_tables_and_findings(self):
        report = self.chatbot.build_business_report(PAY_QUESTION, self.result)
        with patch.object(self.chatbot, "render_copy_response") as copy_response:
            summary = self.chatbot.display_business_report(PAY_QUESTION, self.result, "report_test")
        copy_response.assert_called_once_with(self.chatbot.business_report_markdown(report), "report_test")
        self.assertEqual(summary, report.executive_summary)
        self.assertIn("Department Salary Profile", copy_response.call_args.args[0])
        self.assertIn("Key Findings", copy_response.call_args.args[0])

    def test_recall_uses_previous_question_without_model(self):
        turns = (("user", PAY_QUESTION), ("assistant", "Finance has the highest average pay."))
        with patch.object(self.chatbot, "get_azure_openai_client") as client:
            answer = self.chatbot.answer_general_question("do you remember what i asked you before", turns)
        self.assertIn(PAY_QUESTION, answer)
        client.assert_not_called()

    def test_recall_without_history_does_not_invent_messages(self):
        answer = self.chatbot.answer_general_question("What was my previous question?", ())
        self.assertIn("no earlier messages", answer)

    def test_general_answer_receives_prior_turns_once(self):
        turns = (("user", "Explain headcount"), ("assistant", "It counts employees."))
        with patch.dict(os.environ, {"AZURE_OPENAI_DEPLOYMENT_MINI": "test"}), patch.object(
            self.chatbot, "get_azure_openai_client"
        ) as client:
            client.return_value.chat.completions.create.return_value.choices[0].message.content = "A useful metric."
            self.chatbot.answer_general_question("Why is that useful?", turns)
        messages = client.return_value.chat.completions.create.call_args.kwargs["messages"]
        self.assertEqual(messages[1:], [
            {"role": "user", "content": "Explain headcount"},
            {"role": "assistant", "content": "It counts employees."},
            {"role": "user", "content": "Why is that useful?"},
        ])

    def test_context_is_bounded_and_excludes_internal_fields(self):
        messages = [{"question": f"Question {index}", "answer": "a" * 5000, "sql": "private SQL"} for index in range(10)]
        context = self.chatbot.conversation_context(messages)
        self.assertEqual(len(context), 12)
        self.assertEqual(context[0], ("user", "Question 4"))
        self.assertTrue(context[1][1].endswith("[truncated]"))
        self.assertNotIn("private SQL", str(context))

    def test_recall_question_is_not_routed_to_employee_sql(self):
        with patch.object(self.chatbot, "generate_employee_query") as query:
            request = self.chatbot.translate_question("What did I ask about employee salary before?")
        self.assertIsNone(request)
        query.assert_not_called()

    def test_chat_recall_survives_rerun_and_new_chat_clears_context(self):
        with patch("sqlalchemy.engine.Engine.connect"), patch("pandas.read_sql", return_value=self.result):
            app = AppTest.from_file(str(ROOT / "employee_sql_chatbot.py")).run(timeout=30)
            app.chat_input[0].set_value(PAY_QUESTION).run(timeout=30)
            app.chat_input[0].set_value("do you remember what i asked you before").run(timeout=30)
            self.assertEqual(len(app.exception), 0)
            answer = app.session_state["chat_history"][-1]["answer"]
            self.assertIn(PAY_QUESTION, answer)
            app.run(timeout=30)
            self.assertEqual(app.session_state["chat_history"][-1]["answer"], answer)
            app.button(key="new_chat").click().run(timeout=30)
            app.chat_input[0].set_value("do you remember what i asked you before").run(timeout=30)
            self.assertEqual(len(app.exception), 0)
            self.assertIn("no earlier messages", app.session_state["chat_history"][-1]["answer"])

    def test_repeated_reports_have_independent_widgets(self):
        with patch("sqlalchemy.engine.Engine.connect"), patch("pandas.read_sql", return_value=self.result):
            app = AppTest.from_file(str(ROOT / "employee_sql_chatbot.py")).run(timeout=30)
            for _ in range(2):
                app.chat_input[0].set_value(PAY_QUESTION).run(timeout=30)
                self.assertEqual(len(app.exception), 0)
            self.assertEqual(len(app.get("plotly_chart")), 2)
            self.assertEqual(len(app.get("download_button")), 2)
            self.assertEqual(len(app.text_input), 2)
            self.assertNotEqual(app.text_input[0].key, app.text_input[1].key)
            app.text_input[0].set_value("manager@example.com").run(timeout=30)
            self.assertEqual(len(app.exception), 0)
            self.assertEqual(app.text_input[1].value, "")

    def test_old_messages_reuse_results_but_new_questions_query_fresh_data(self):
        with patch("sqlalchemy.engine.Engine.connect"), patch("pandas.read_sql", return_value=pd.DataFrame({"employee_count": [1]})) as read_sql:
            app = AppTest.from_file(str(ROOT / "employee_sql_chatbot.py")).run(timeout=30)
            app.chat_input[0].set_value("How many employees do we have?").run(timeout=30)
            first_answer = app.session_state["chat_history"][0]["answer"]
            self.assertEqual(read_sql.call_count, 1)
            read_sql.return_value = pd.DataFrame({"employee_count": [2]})
            app.chat_input[0].set_value("How many employees do we have?").run(timeout=30)
            self.assertEqual(len(app.exception), 0)
            self.assertEqual(read_sql.call_count, 2)
            self.assertEqual(app.session_state["chat_history"][0]["answer"], first_answer)
            self.assertIn("2 employees", app.session_state["chat_history"][1]["answer"])
            app.run(timeout=30)
            self.assertEqual(read_sql.call_count, 2)

    def test_query_plan_cache_reuses_interpretation_not_database_results(self):
        plan = employee_query_plan(filter={"column": "job_title", "op": "eq", "value": "Site engineer"})
        with patch.dict(os.environ, {"AZURE_OPENAI_DEPLOYMENT_MINI": "test"}), patch.object(self.chatbot, "get_azure_openai_client") as client:
            choice = client.return_value.chat.completions.create.return_value.choices[0]
            choice.message.content = json.dumps(plan)
            choice.finish_reason = "stop"
            first = self.chatbot.generate_employee_query("how many Site engineer we have")
            second = self.chatbot.generate_employee_query("how many Site engineer we have")
            client.return_value.chat.completions.create.assert_called_once()
            self.assertIsNot(first, second)
            self.assertEqual(first.sql, second.sql)
            with patch.object(self.chatbot, "run_query", side_effect=[pd.DataFrame({"employee_count": [1]}), pd.DataFrame({"employee_count": [2]})]) as query:
                self.assertIn("1", self.chatbot.display_response(first, response_cache={}))
                self.assertIn("2", self.chatbot.display_response(second, response_cache={}))
                self.assertEqual(query.call_count, 2)

    def test_report_reruns_reuse_pdf_and_report_but_new_messages_do_not(self):
        cache = {}
        with patch.object(self.chatbot, "build_business_report", wraps=self.chatbot.build_business_report) as build, patch.object(self.chatbot, "generate_business_report_pdf", return_value=b"%PDF-test") as pdf:
            self.chatbot.display_business_report(PAY_QUESTION, self.result, "cached", cache)
            self.chatbot.display_business_report(PAY_QUESTION, self.result, "cached", cache)
            self.assertEqual(build.call_count, 1)
            self.assertEqual(pdf.call_count, 1)
            self.chatbot.display_business_report(PAY_QUESTION, self.result, "new", {})
            self.assertEqual(pdf.call_count, 2)

    def test_failed_queries_are_not_cached(self):
        cache = {}
        request = self.chatbot.SqlRequest("SELECT COUNT(*) AS employee_count FROM dbo.Employees", {}, "Count employees", response_kind="employee_count")
        with patch.object(self.chatbot, "run_query", side_effect=[self.chatbot.SQLAlchemyError("test"), pd.DataFrame({"employee_count": [3]})]) as query:
            self.assertIsNone(self.chatbot.display_response(request, response_cache=cache))
            self.assertNotIn("result", cache)
            self.assertIn("3 employees", self.chatbot.display_response(request, response_cache=cache))
            self.assertEqual(query.call_count, 2)

    def test_email_uses_tls_and_attaches_exact_pdf(self):
        report = self.chatbot.build_business_report(PAY_QUESTION, self.result)
        pdf = self.chatbot.generate_business_report_pdf(PAY_QUESTION, report, None)
        with patch.dict(os.environ, SMTP_SETTINGS), patch.object(self.chatbot.smtplib, "SMTP") as smtp:
            sent, _ = self.chatbot.send_report_email("manager@example.com", report, pdf)
            self.assertTrue(sent)
            smtp.assert_called_once_with("smtp.example.com", 587, timeout=20)
            server = smtp.return_value.__enter__.return_value
            self.assertEqual(
                [call[0] for call in server.method_calls],
                ["ehlo", "starttls", "ehlo", "login", "send_message"],
            )
            message = server.send_message.call_args.args[0]
            self.assertEqual(message["To"], "manager@example.com")
            attachments = list(message.iter_attachments())
            self.assertEqual(len(attachments), 1)
            self.assertEqual(attachments[0].get_content_type(), "application/pdf")
            self.assertEqual(attachments[0].get_payload(decode=True), pdf)

    def test_invalid_recipient_does_not_connect(self):
        report = self.chatbot.build_business_report(PAY_QUESTION, self.result)
        with patch.object(self.chatbot.smtplib, "SMTP") as smtp:
            sent, _ = self.chatbot.send_report_email("not-an-email", report, b"%PDF-test")
            self.assertFalse(sent)
            smtp.assert_not_called()

    def test_missing_smtp_configuration_does_not_connect(self):
        report = self.chatbot.build_business_report(PAY_QUESTION, self.result)
        with patch.object(self.chatbot.smtplib, "SMTP") as smtp:
            sent, status = self.chatbot.send_report_email("manager@example.com", report, b"%PDF-test")
            self.assertFalse(sent)
            self.assertIn("not configured", status)
            smtp.assert_not_called()

    def test_authentication_failure_is_reported_without_secrets(self):
        report = self.chatbot.build_business_report(PAY_QUESTION, self.result)
        with patch.dict(os.environ, SMTP_SETTINGS), patch.object(self.chatbot.smtplib, "SMTP") as smtp:
            server = smtp.return_value.__enter__.return_value
            server.login.side_effect = smtplib.SMTPAuthenticationError(535, b"Rejected")
            sent, status = self.chatbot.send_report_email("manager@example.com", report, b"%PDF-test")
            self.assertFalse(sent)
            self.assertNotIn(SMTP_SETTINGS["SMTP_APP_PASSWORD"], status)
            server.send_message.assert_not_called()


if __name__ == "__main__":
    unittest.main()