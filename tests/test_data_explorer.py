from io import BytesIO
from zipfile import ZipFile
from hashlib import sha256
import json
import os
from pathlib import Path
import unittest
from unittest.mock import patch

from openpyxl import load_workbook
import pandas as pd
from streamlit.testing.v1 import AppTest

from data_explorer import (
    build_charts, csv_to_workbook, data_overview, is_overview_request, plan_charts,
)
from excel_support import excel_sheet_names, generate_excel_download, read_excel_sheet


ROOT = Path(__file__).resolve().parents[1]
TEST_ENVIRONMENT = {
    name: value for name, value in os.environ.items()
    if not name.upper().startswith(("AZURE_", "SMTP_", "WORKFORCE_"))
}


def staff_frame():
    return pd.DataFrame({
        "Department": ["IT", "Sales", "IT", "Sales"] * 4,
        "Salary": list(range(100, 1700, 100)),
        "Joining Date": [f"{2020 + index % 4}-01-01" for index in range(16)],
        "Status": ["Active", "On Leave"] * 8,
    })


class DataExplorerTests(unittest.TestCase):
    def test_csv_preserves_headers_codes_large_ids_and_literal_formulas(self):
        content = csv_to_workbook(
            b' Name ,Name,,Code,Employee ID,Amount,Formula\n'
            b'Ada,A,yes,00123,12345678901234567,12.5,=1+1\n'
            b'Lin,B,no,00456,98765432109876543,14,"text, with comma"\n',
            "staff:upload.csv",
        )
        self.assertEqual(excel_sheet_names(content), ["staff_upload"])
        frame = read_excel_sheet(content, "staff_upload")
        self.assertEqual(list(frame.columns), [
            "Name", "Name (2)", "Column 3", "Code", "Employee ID", "Amount", "Formula",
        ])
        self.assertEqual(frame["Code"].tolist(), ["00123", "00456"])
        self.assertEqual(frame["Employee ID"].iloc[0], "12345678901234567")
        self.assertEqual(frame["Amount"].tolist(), [12.5, 14])
        self.assertEqual(frame["Formula"].tolist(), ["=1+1", "text, with comma"])
        workbook = load_workbook(BytesIO(content), data_only=False)
        self.addCleanup(workbook.close)
        self.assertEqual(workbook.active["G2"].data_type, "s")

    def test_csv_accepts_delimiters_bom_quotes_and_blank_rows(self):
        for delimiter in (",", ";", "\t", "|"):
            with self.subTest(delimiter=delimiter):
                text = f'\ufeffName{delimiter}Salary\n"Ada{delimiter}Lovelace"{delimiter}100\n\nLin{delimiter}200\n'
                content = csv_to_workbook(text.encode("utf-8"), "data.csv")
                frame = read_excel_sheet(content, "data")
                self.assertEqual(len(frame), 2)
                self.assertEqual(frame["Salary"].tolist(), [100, 200])
                self.assertEqual(frame["Name"].iloc[0], f"Ada{delimiter}Lovelace")

    def test_invalid_csv_reports_empty_and_mismatched_rows(self):
        for content in (b"", b" \n\n", b"A,B\n1,2,3\n", b"A,B\n1\n", b'A,B\n"broken,2\n'):
            with self.subTest(content=content), self.assertRaises(ValueError):
                csv_to_workbook(content, "bad.csv")

    def test_overview_reports_types_and_missing_values_without_changing_data(self):
        frame = pd.DataFrame({
            "Name": ["Ada", "Lin", "N/A"],
            "Salary": ["$1,000", "2000", None],
            "Joining Date": ["2022-01-01", "2023-01-01", "not available"],
            "Enabled": [True, False, pd.NA],
        })
        original = frame.copy(deep=True)
        summary = data_overview(frame).set_index("Column")
        self.assertEqual(summary["Missing values"].to_dict(), {
            "Name": 1, "Salary": 1, "Joining Date": 1, "Enabled": 1,
        })
        self.assertEqual(summary["Data type"].to_dict(), {
            "Name": "text", "Salary": "whole number", "Joining Date": "date", "Enabled": "true/false",
        })
        pd.testing.assert_frame_equal(frame, original)
        for question in ("open the file", "show me the data", "Show my CSV file please!"):
            self.assertTrue(is_overview_request(question), question)
        self.assertFalse(is_overview_request("show me the data for active employees"))

    def test_requested_charts_compute_exact_averages_and_yearly_counts(self):
        frame = staff_frame()
        for question, expected in (
            (
                "show average salary by department",
                frame.groupby("Department")["Salary"].mean().reset_index(name="Average Salary"),
            ),
            (
                "how many people joined each year",
                pd.DataFrame({"Year": [2020, 2021, 2022, 2023], "Number of people": [4, 4, 4, 4]}),
            ),
        ):
            with self.subTest(question=question):
                specs = plan_charts(question, frame)
                self.assertIsNotNone(specs)
                charts, problems = build_charts(frame, specs)
                self.assertEqual(problems, [])
                self.assertEqual(len(charts), 1)
                actual = charts[0].table.sort_values(charts[0].table.columns[0]).reset_index(drop=True)
                pd.testing.assert_frame_equal(actual, expected, check_dtype=False)
                self.assertGreaterEqual(len(charts[0].explanation.split(". ")), 2)
                self.assertLessEqual(len(charts[0].explanation.split(". ")), 3)

    def test_employee_salary_chart_uses_the_measure_not_an_id_or_worksheet(self):
        frame = staff_frame()
        frame.insert(0, "Employee ID", range(1, 17))
        for question in (
            "create a chart for employee salary", "create a chart for salary",
            "draw a histogram of employee salary",
        ):
            with self.subTest(question=question):
                specs = plan_charts(question, frame)
                self.assertIsNotNone(specs)
                self.assertEqual(specs[0].measure, "Salary")
                self.assertEqual(specs[0].kind, "histogram")
                charts, problems = build_charts(frame, specs)
                self.assertEqual(problems, [])
                self.assertEqual(charts[0].table.iloc[:, 1].sum(), 16)

    def test_employee_salary_chart_prefers_an_exact_full_column_name(self):
        frame = staff_frame()
        frame["Employee Salary"] = frame["Salary"] * 10
        specs = plan_charts("create a chart for employee salary", frame)
        self.assertIsNotNone(specs)
        self.assertEqual(specs[0].measure, "Employee Salary")

    def test_generic_charts_use_real_columns_and_histogram_table_matches_plot(self):
        frame = staff_frame()
        specs = plan_charts("make charts", frame)
        charts, problems = build_charts(frame, specs)
        self.assertEqual(problems, [])
        self.assertIn(len(charts), (3, 4))
        histogram = next(chart for chart in charts if chart.chart_type == "histogram")
        self.assertEqual(histogram.table.iloc[:, 1].sum(), len(frame))
        self.assertEqual(list(histogram.figure.data[0].y), histogram.table.iloc[:, 1].tolist())
        counts = next(chart for chart in charts if chart.chart_type == "pie")
        self.assertEqual(counts.table.iloc[:, 1].sum(), len(frame))
        self.assertEqual(list(counts.figure.data[0].values), counts.table.iloc[:, 1].tolist())

    def test_dashboard_and_support_table_presentation_prompts_are_recognized(self):
        for question in (
            "create a dashboard", "make a dashboard with charts and supporting tables",
            "show a dashboard with support tables", "give dashboard support tables",
            "draw charts and give dashboard support tables",
        ):
            with self.subTest(question=question):
                charts, problems = build_charts(staff_frame(), plan_charts(question, staff_frame()))
                self.assertEqual(problems, [])
                self.assertEqual(len(charts), 4)

    def test_complex_questions_are_not_silently_simplified(self):
        for question in (
            "show average salary by department for active employees",
            "show top 3 departments by average salary",
            "show average salary and joining date by department",
            "show average salary by department excluding Sales",
            "how many people joined each year after 2021",
            "how do I make a chart of salary by department",
            "do not show a chart of salary by department",
            "show total and average salary by department",
            "create a chart for employee salary for active employees",
            "create a chart for employee salary and joining date",
            "create a chart of average salary",
            "histogram of average salary",
        ):
            with self.subTest(question=question):
                self.assertIsNone(plan_charts(question, staff_frame()))

    def test_missing_values_are_excluded_and_explained(self):
        frame = pd.DataFrame({
            "Department": ["IT", " it ", "Sales", None],
            "Salary": [100, 300, "N/A", 500],
        })
        charts, problems = build_charts(frame, plan_charts("show average salary by department", frame))
        self.assertEqual(problems, [])
        self.assertEqual(charts[0].table.to_dict("list"), {"Department": ["IT"], "Average Salary": [200.0]})
        self.assertIn("2 of 4 rows were left out", charts[0].explanation)


class UploadedDataUITests(unittest.TestCase):
    def setUp(self):
        self.enterContext(patch("dotenv.load_dotenv"))
        self.enterContext(patch.dict(os.environ, {**TEST_ENVIRONMENT, "WORKFORCE_DATA_MODE": "demo"}, clear=True))
        self.enterContext(patch("socket.create_connection", side_effect=AssertionError("Unexpected network access")))

    def start_app(self, content, name, mode="Automatic (SQL + Excel)"):
        self.upload = BytesIO(content)
        self.upload.name = name
        self.enterContext(patch(
            "streamlit.file_uploader",
            side_effect=lambda *args, **kwargs: self.upload if kwargs.get("key") == "excel_upload_0" else None,
        ))
        app = AppTest.from_file(str(ROOT / "employee_sql_chatbot.py"))
        app.session_state["data_source"] = mode
        app.run(timeout=30)
        self.assertEqual(len(app.exception), 0)
        return app

    def test_csv_upload_overview_and_four_charts_use_all_rows_and_survive_reruns(self):
        frame = staff_frame()
        app = self.start_app(frame.to_csv(index=False).encode(), "staff.csv")
        app.chat_input[0].set_value("open the file").run(timeout=30)
        self.assertEqual(len(app.exception), 0)
        overview = app.session_state["chat_history"][-1]
        self.assertEqual(overview["kind"], "uploaded_data")
        self.assertEqual(len(overview["result"]), 16)
        self.assertEqual(len(app.chat_message[-1].dataframe[0].value), 16)
        self.assertEqual(app.chat_message[-1].metric[0].value, "16")
        self.assertEqual(app.chat_message[-1].metric[1].value, "4")
        app.chat_input[0].set_value("make charts").run(timeout=30)
        self.assertEqual(len(app.exception), 0)
        message = app.session_state["chat_history"][-1]
        self.assertNotIn("error", message)
        self.assertEqual(len(message["charts"]), 4)
        self.assertEqual(len(app.chat_message[-1].get("plotly_chart")), 4)
        self.assertEqual(len(app.chat_message[-1].dataframe), 4)
        for table, chart in zip(app.chat_message[-1].dataframe, message["charts"]):
            pd.testing.assert_frame_equal(table.value, chart.table)
        download = message["excel_download"]
        self.assertEqual(excel_sheet_names(download), ["staff"])
        workbook = load_workbook(BytesIO(download))
        self.addCleanup(workbook.close)
        self.assertEqual(len(workbook["staff"]._charts), 4)
        events = list(message["backend_log"])
        app.run(timeout=30)
        self.assertEqual(len(app.exception), 0)
        self.assertEqual(app.session_state["chat_history"][-1]["backend_log"], events)
        self.assertEqual(app.session_state["chat_history"][-1]["excel_download"], download)

    def test_multisheet_workbook_selection_controls_chart_input(self):
        other = pd.DataFrame({"Department": ["Other"], "Salary": [999]})
        frame = staff_frame()
        content = generate_excel_download([("First", other), ("Staff", frame)])
        app = self.start_app(content, "staff.xlsx", "Excel workbook")
        app.selectbox[0].set_value("Staff").run(timeout=30)
        app.chat_input[0].set_value("show average salary by department").run(timeout=30)
        self.assertEqual(len(app.exception), 0)
        message = app.session_state["chat_history"][-1]
        self.assertEqual(message["source"], "staff.xlsx / Staff")
        self.assertEqual(message["charts"][0].table["Average Salary"].tolist(), [900.0, 800.0])
        app.chat_input[0].set_value("how many people joined each year").run(timeout=30)
        self.assertEqual(len(app.exception), 0)
        chart = app.session_state["chat_history"][-1]["charts"][0]
        self.assertEqual(chart.table["Number of people"].tolist(), [4, 4, 4, 4])
        spec = json.loads(app.chat_message[-1].get("plotly_chart")[0].proto.spec)
        self.assertEqual(spec["data"][0]["type"], "scatter")

    def test_malformed_upload_shows_error_and_can_be_replaced_and_retried(self):
        app = self.start_app(b"A,B\n1,2,3\n", "broken.csv")
        self.assertTrue(any("different number" in warning.value for warning in app.warning))
        app.chat_input[0].set_value("show me the data").run(timeout=30)
        self.assertEqual(len(app.exception), 0)
        self.assertIn("different number", app.session_state["chat_history"][-1]["error"])
        self.upload = BytesIO(b"A,B\n1,2\n")
        self.upload.name = "fixed.csv"
        app.run(timeout=30)
        app.button(key="retry_uploaded_data_0").click().run(timeout=30)
        self.assertEqual(len(app.exception), 0)
        message = app.session_state["chat_history"][-1]
        self.assertNotIn("error", message)
        self.assertEqual(message["source"], "fixed.csv / fixed")
        self.assertEqual(len(message["result"]), 1)

    def test_missing_upload_and_unusable_chart_columns_show_actionable_errors(self):
        app = AppTest.from_file(str(ROOT / "employee_sql_chatbot.py")).run(timeout=30)
        app.chat_input[0].set_value("show me the data").run(timeout=30)
        self.assertEqual(len(app.exception), 0)
        self.assertIn("Upload an Excel", app.session_state["chat_history"][-1]["error"])
        app = self.start_app(b"Name\nAda\nLin\n", "names.csv")
        app.chat_input[0].set_value("make charts").run(timeout=30)
        self.assertEqual(len(app.exception), 0)
        self.assertIn("No useful charts", app.session_state["chat_history"][-1]["error"])

    def test_explicit_sql_and_native_workbook_charts_keep_existing_routes(self):
        app = self.start_app(staff_frame().to_csv(index=False).encode(), "staff.csv")
        app.chat_input[0].set_value("Show average salary by department in SQL").run(timeout=30)
        self.assertEqual(len(app.exception), 0)
        self.assertNotEqual(app.session_state["chat_history"][-1]["kind"], "uploaded_data")
        app.chat_input[0].set_value("Add a column chart of average Salary by Department in the workbook").run(timeout=30)
        self.assertEqual(len(app.exception), 0)
        message = app.session_state["chat_history"][-1]
        self.assertEqual(message["kind"], "excel_chart")
        self.assertNotIn("error", message)
        workbook = load_workbook(BytesIO(message["edited_workbook"]))
        self.addCleanup(workbook.close)
        self.assertEqual(len(workbook[message["edited_sheet"]]._charts), 1)

    def test_compound_open_and_draw_charts_creates_native_dashboard(self):
        frame = staff_frame()
        for question in (
            "open employees_messy and draw charts in the excel sheet",
            "open worksheet employees_messy and draw charts in the excel sheet",
        ):
            with self.subTest(question=question):
                app = self.start_app(frame.to_csv(index=False).encode(), "employees_messy.csv")
                original = app.session_state["excel_workbook"]["content"]
                app.chat_input[0].set_value(question).run(timeout=30)
                self.assertEqual(len(app.exception), 0)
                message = app.session_state["chat_history"][-1]
                self.assertEqual(message["kind"], "uploaded_data")
                self.assertNotIn("error", message)
                self.assertTrue(message["dashboard"])
                self.assertEqual(message["source_sheet"], "employees_messy")
                self.assertEqual(len(message["charts"]), 4)
                workbook = load_workbook(BytesIO(message["dashboard_workbook"]))
                self.addCleanup(workbook.close)
                self.assertEqual(len(workbook[message["dashboard_sheet"]]._charts), 4)
                self.assertEqual(message["original_workbook"], original)
                events = list(message["backend_log"])
                digest = app.session_state["excel_workbook"]["digest"]
                app.run(timeout=30)
                self.assertEqual(len(app.exception), 0)
                self.assertEqual(app.session_state["excel_workbook"]["digest"], digest)
                self.assertEqual(app.session_state["chat_history"][-1]["backend_log"], events)

    def test_employee_salary_prompt_and_cached_error_build_the_requested_chart(self):
        for question in (
            "open employees_messy and create a chart for employee salary",
            "create a chart for employee salary",
            "open employees_messy and create a chart for salary",
        ):
            with self.subTest(question=question):
                app = self.start_app(staff_frame().to_csv(index=False).encode(), "employees_messy.csv")
                app.chat_input[0].set_value(question).run(timeout=30)
                self.assertEqual(len(app.exception), 0)
                message = app.session_state["chat_history"][-1]
                self.assertEqual(message["kind"], "uploaded_data")
                self.assertNotIn("error", message)
                self.assertEqual(message["source_sheet"], "employees_messy")
                self.assertEqual(len(message["charts"]), 1)
                self.assertEqual(message["charts"][0].chart_type, "histogram")
                self.assertEqual(message["charts"][0].table.iloc[:, 1].sum(), 16)
                exported = load_workbook(BytesIO(message["excel_download"]))
                self.addCleanup(exported.close)
                self.assertEqual(len(exported["employees_messy"]._charts), 1)
                self.assertEqual(read_excel_sheet(message["excel_download"], "employees_messy").shape, (16, 4))
        app.session_state["chat_history"] = [{
            "kind": "uploaded_data", "question": "open employees_messy and create a chart for employee salary",
            "error": "I couldn't find worksheet employee salary. Available worksheets: employees_messy.",
        }]
        app.run(timeout=30)
        app.button(key="retry_uploaded_data_0").click().run(timeout=30)
        self.assertEqual(len(app.exception), 0)
        self.assertNotIn("error", app.session_state["chat_history"][-1])
        self.assertEqual(app.session_state["chat_history"][-1]["source_sheet"], "employees_messy")

    def test_employee_salary_chart_respects_an_explicit_source_and_unknown_sheet(self):
        content = generate_excel_download([
            ("First", pd.DataFrame({"Salary": [1, 2]})), ("employees_messy", staff_frame()),
        ])
        app = self.start_app(content, "staff.xlsx")
        app.chat_input[0].set_value(
            "open employees_messy and create a chart for employee salary",
        ).run(timeout=30)
        self.assertEqual(len(app.exception), 0)
        message = app.session_state["chat_history"][-1]
        self.assertNotIn("error", message)
        self.assertEqual(message["source_sheet"], "employees_messy")
        self.assertEqual(message["charts"][0].table.iloc[:, 1].sum(), 16)
        current = app.session_state["excel_workbook"]["content"]
        app.chat_input[0].set_value("create a chart for worksheet Salary").run(timeout=30)
        self.assertEqual(len(app.exception), 0)
        message = app.session_state["chat_history"][-1]
        self.assertIn("couldn't find worksheet Salary", message["error"])
        self.assertEqual(app.session_state["excel_workbook"]["content"], current)

    def test_dashboard_with_support_tables_has_metrics_tabs_and_matching_chart_data(self):
        frame = staff_frame().astype(object)
        frame.loc[0, "Salary"] = "$100"
        frame.loc[1, "Department"] = " sales "
        frame.loc[2, "Salary"] = "N/A"
        frame.loc[3, "Joining Date"] = "not available"
        app = self.start_app(frame.to_csv(index=False).encode(), "employees_messy.csv")
        app.chat_input[0].set_value(
            "create a dashboard for employees_messy with charts and supporting tables",
        ).run(timeout=30)
        self.assertEqual(len(app.exception), 0)
        message = app.session_state["chat_history"][-1]
        self.assertNotIn("error", message)
        self.assertTrue(message["dashboard"])
        self.assertEqual(message["dashboard_rows"], 16)
        self.assertEqual(message["dashboard_missing"], 2)
        response = app.chat_message[-1]
        self.assertEqual([tab.label for tab in response.tabs], ["Dashboard", "Supporting tables", "Data summary"])
        self.assertEqual([metric.value for metric in response.metric], ["16", "4", "2", "4"])
        for chart, table in zip(message["charts"], response.tabs[0].dataframe):
            pd.testing.assert_frame_equal(chart.table, table.value)
        self.assertEqual(len(response.tabs[1].dataframe), 4)
        self.assertIn("dashboard_workbook", message)
        self.assertEqual(excel_sheet_names(message["dashboard_workbook"]), ["employees_messy"])
        self.assertFalse(app.session_state["excel_workbook"].get("edited", False))

    def test_compound_request_selects_named_sheet_not_current_preview(self):
        content = generate_excel_download([
            ("First", pd.DataFrame({"Name": ["Other"]})), ("Employees and Sales", staff_frame()),
        ])
        app = self.start_app(content, "staff.xlsx")
        app.chat_input[0].set_value(
            'open worksheet "Employees and Sales" then create a dashboard with supporting tables',
        ).run(timeout=30)
        self.assertEqual(len(app.exception), 0)
        message = app.session_state["chat_history"][-1]
        self.assertNotIn("error", message)
        self.assertEqual(message["source_sheet"], "Employees and Sales")
        self.assertEqual(message["dashboard_rows"], 16)
        self.assertEqual(app.selectbox[0].value, "Employees and Sales")

    def test_unknown_sheet_and_unsupported_compound_chart_do_not_modify_workbook(self):
        for question, expected in (
            ("open missing and draw charts in the excel sheet", "Available worksheets: employees_messy"),
            (
                "open employees_messy and show average Salary by Department for active employees",
                "The worksheet was found, but this chart request",
            ),
        ):
            with self.subTest(question=question):
                app = self.start_app(staff_frame().to_csv(index=False).encode(), "employees_messy.csv")
                content = app.session_state["excel_workbook"]["content"]
                app.chat_input[0].set_value(question).run(timeout=30)
                self.assertEqual(len(app.exception), 0)
                message = app.session_state["chat_history"][-1]
                self.assertEqual(message["kind"], "uploaded_data")
                self.assertIn(expected, message["error"])
                self.assertNotIn("answer", message)
                self.assertEqual(app.session_state["excel_workbook"]["content"], content)

    def test_native_dashboard_with_support_tables_and_specific_compound_chart(self):
        for question in (
            "create a dashboard in Excel with supporting tables",
            "open employees_messy and add a column chart of average Salary by Department in the workbook",
            "open the file and draw charts and give dashboard support tables in the excel sheet",
        ):
            with self.subTest(question=question):
                app = self.start_app(staff_frame().to_csv(index=False).encode(), "employees_messy.csv")
                app.chat_input[0].set_value(question).run(timeout=30)
                self.assertEqual(len(app.exception), 0)
                message = app.session_state["chat_history"][-1]
                self.assertNotIn("error", message)
                self.assertIn("dashboard_workbook", message)
                count = 1 if "average Salary" in question else 4
                self.assertEqual(len(message["charts"]), count)
                workbook = load_workbook(BytesIO(message["dashboard_workbook"]))
                self.addCleanup(workbook.close)
                self.assertEqual(len(workbook[message["dashboard_sheet"]]._charts), count)
                self.assertTrue(any(
                    button.label == "Download Excel with charts" for button in app.get("download_button")
                ))

    def test_old_missing_sheet_response_can_retry_the_original_compound_prompt(self):
        app = self.start_app(staff_frame().to_csv(index=False).encode(), "employees_messy.csv")
        app.session_state["chat_history"] = [{
            "kind": "workbook_command",
            "question": "open employees_messy and draw charts in the excel sheet",
            "answer": "I couldn't find that worksheet. Available worksheets: employees_messy.",
        }]
        app.run(timeout=30)
        app.button(key="retry_workbook_chart_0").click().run(timeout=30)
        self.assertEqual(len(app.exception), 0)
        message = app.session_state["chat_history"][-1]
        self.assertNotIn("error", message)
        self.assertEqual(message["kind"], "uploaded_data")
        self.assertEqual(len(message["charts"]), 4)
        self.assertIn("dashboard_workbook", message)

    def test_dashboard_export_failure_leaves_original_workbook_and_shows_no_success(self):
        app = self.start_app(staff_frame().to_csv(index=False).encode(), "employees_messy.csv")
        original = app.session_state["excel_workbook"]["content"]
        with patch("workbook_charts.apply_excel_dashboard", side_effect=ValueError("Cannot create dashboard package.")):
            app.chat_input[0].set_value("create a dashboard in Excel").run(timeout=30)
        self.assertEqual(len(app.exception), 0)
        message = app.session_state["chat_history"][-1]
        self.assertIn("Cannot create dashboard package", message["error"])
        self.assertNotIn("answer", message)
        self.assertNotIn("dashboard_workbook", message)
        self.assertEqual(app.session_state["excel_workbook"]["content"], original)
        self.assertEqual(len(app.chat_message[-1].success), 0)

    def test_same_sheet_is_default_and_explicit_new_tab_is_respected(self):
        app = self.start_app(staff_frame().to_csv(index=False).encode(), "employees_messy.csv")
        self.assertEqual(app.selectbox(key="excel_chart_placement").value, "same_sheet")
        app.chat_input[0].set_value("draw charts in the same worksheet").run(timeout=30)
        self.assertEqual(len(app.exception), 0)
        message = app.session_state["chat_history"][-1]
        self.assertEqual(message["chart_placement"], "same_sheet")
        self.assertEqual(message["dashboard_sheet"], "employees_messy")
        self.assertEqual(excel_sheet_names(message["excel_download"]), ["employees_messy"])
        app.chat_input[0].set_value("draw charts on a new worksheet").run(timeout=30)
        self.assertEqual(len(app.exception), 0)
        new_tab = app.session_state["chat_history"][-1]
        self.assertEqual(new_tab["chart_placement"], "new_sheet")
        self.assertEqual(excel_sheet_names(new_tab["excel_download"]), ["employees_messy", "employees_messy Dashboard"])
        labels = [button.label for button in app.chat_message[-1].get("download_button")]
        self.assertIn("Download Excel with charts", labels)
        self.assertIn("Download supporting tables only", labels)
        self.assertNotIn("Download Excel", labels)

    def test_source_questions_still_work_after_native_same_sheet_dashboard_and_reupload(self):
        app = self.start_app(staff_frame().to_csv(index=False).encode(), "employees_messy.csv", "Excel workbook")
        app.chat_input[0].set_value("make charts in the same worksheet").run(timeout=30)
        exported = app.session_state["chat_history"][-1]["excel_download"]
        app.chat_input[0].set_value("show me the data").run(timeout=30)
        self.assertEqual(len(app.exception), 0)
        self.assertEqual(app.session_state["chat_history"][-1]["result"].shape, (16, 4))
        self.upload = BytesIO(exported)
        self.upload.name = "employees_messy_charts.xlsx"
        app.run(timeout=30)
        app.chat_input[0].set_value("show average salary by department").run(timeout=30)
        self.assertEqual(len(app.exception), 0)
        message = app.session_state["chat_history"][-1]
        self.assertNotIn("error", message)
        self.assertEqual(message["charts"][0].table["Average Salary"].tolist(), [900.0, 800.0])

    def test_clean_named_table_routes_locally_and_preserves_same_sheet_charts(self):
        frame = staff_frame().astype(object)
        frame.loc[0, "Salary"] = "$100"
        frame.loc[1, "Department"] = " sales "
        frame.loc[2, "Salary"] = "N/A"
        app = self.start_app(frame.to_csv(index=False).encode(), "employees_messy.csv")
        app.chat_input[0].set_value("make charts in the same worksheet").run(timeout=30)
        charted = app.session_state["excel_workbook"]["content"]
        app.chat_input[0].set_value("clean the employees_messy table").run(timeout=30)
        self.assertEqual(len(app.exception), 0)
        message = app.session_state["chat_history"][-1]
        self.assertEqual(message["kind"], "excel_clean")
        self.assertNotIn("error", message)
        self.assertEqual(len(message["result"]), 16)
        self.assertEqual(message["result"]["Department"].iloc[1], "Sales")
        self.assertEqual(message["result"]["Salary"].iloc[0], 100)
        self.assertTrue(pd.isna(message["result"]["Salary"].iloc[2]))
        book = load_workbook(BytesIO(message["edited_workbook"]))
        self.addCleanup(book.close)
        self.assertEqual(len(book["employees_messy"]._charts), 4)
        with ZipFile(BytesIO(charted)) as before, ZipFile(BytesIO(message["edited_workbook"])) as after:
            for name in before.namelist():
                if name != "xl/worksheets/sheet1.xml":
                    self.assertEqual(before.read(name), after.read(name), name)
        app.chat_input[0].set_value("show average salary by department").run(timeout=30)
        self.assertEqual(len(app.exception), 0)
        self.assertNotIn("error", app.session_state["chat_history"][-1])

    def test_failed_automatic_cleanup_can_retry_as_a_local_workbook_action(self):
        app = self.start_app(staff_frame().to_csv(index=False).encode(), "employees_messy.csv")
        app.session_state["chat_history"] = [{
            "kind": "automatic", "question": "clean the employees_messy table",
            "error": "The cross-source calculation was not fully understood.",
            "source_workbook_digest": sha256(app.session_state["excel_workbook"]["content"]).hexdigest(),
        }]
        app.run(timeout=30)
        app.button(key="retry_automatic_0").click().run(timeout=30)
        self.assertEqual(len(app.exception), 0)
        message = app.session_state["chat_history"][-1]
        self.assertEqual(message["kind"], "excel_clean")
        self.assertNotIn("error", message)
        self.assertIn("Cleanup", app.chat_message[-1].subheader[0].value)

    def test_clean_unknown_table_and_sql_requests_do_not_change_uploaded_data(self):
        for question in ("clean the missing table", "clean the SQL Employees table", "clean both sources"):
            with self.subTest(question=question):
                app = self.start_app(staff_frame().to_csv(index=False).encode(), "employees_messy.csv")
                original = app.session_state["excel_workbook"]["content"]
                app.chat_input[0].set_value(question).run(timeout=30)
                self.assertEqual(len(app.exception), 0)
                message = app.session_state["chat_history"][-1]
                self.assertEqual(message["kind"], "excel_clean")
                self.assertIn("error", message)
                self.assertEqual(app.session_state["excel_workbook"]["content"], original)

    def test_failed_cleanup_validation_does_not_publish_a_changed_workbook(self):
        frame = staff_frame().astype(object)
        frame.loc[0, "Department"] = " sales "
        app = self.start_app(frame.to_csv(index=False).encode(), "employees_messy.csv")
        original = app.session_state["excel_workbook"]["content"]
        with patch("workbook_edits._clean_cell_xml", side_effect=lambda original, *_: original):
            app.chat_input[0].set_value("clean the table").run(timeout=30)
        self.assertEqual(len(app.exception), 0)
        message = app.session_state["chat_history"][-1]
        self.assertIn("verify the saved source values", message["error"])
        self.assertNotIn("answer", message)
        self.assertEqual(app.session_state["excel_workbook"]["content"], original)


if __name__ == "__main__":
    unittest.main()
