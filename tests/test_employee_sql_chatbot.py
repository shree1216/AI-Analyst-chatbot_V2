import importlib
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


class WorkforceTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        with patch("dotenv.load_dotenv"), patch.dict(os.environ, TEST_ENVIRONMENT, clear=True):
            cls.chatbot = importlib.import_module("employee_sql_chatbot")

    def setUp(self):
        self.enterContext(patch.dict(os.environ, TEST_ENVIRONMENT, clear=True))
        self.enterContext(patch("dotenv.load_dotenv"))
        self.enterContext(patch("socket.create_connection", side_effect=AssertionError("Unexpected network access")))
        self.result = employee_fixture()

    def test_chart_request_routes_before_generic_count(self):
        request = self.chatbot.translate_question("Show a bar chart of employee count by department")
        self.assertEqual(request.response_kind, "chart_employee_count")
        self.assertIn("GROUP BY [department]", request.sql)

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
            ("Chart average salary by department", "department", "column"),
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
        self.assertEqual(list(chart.data[0].y), [110000, 90000, 70000])
        self.assertEqual(chart.layout.yaxis.title.text, "Average salary (INR)")
        self.assertEqual(chart.data[0].texttemplate, "%{y:,.2f}")

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