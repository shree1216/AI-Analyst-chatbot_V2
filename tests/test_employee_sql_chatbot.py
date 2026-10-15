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