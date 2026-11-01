from datetime import datetime
from io import BytesIO
import importlib
import json
import os
from pathlib import Path
from types import SimpleNamespace
import unittest
from unittest.mock import Mock, patch
from zipfile import ZipFile
from xml.etree import ElementTree as ET

from openpyxl import Workbook, load_workbook
from openpyxl.chart import LineChart, Reference
from openpyxl.drawing.image import Image as ExcelImage
from PIL import Image
import pandas as pd
from streamlit.testing.v1 import AppTest

from excel_support import excel_sheet_names, read_excel_sheet
from workbook_charts import apply_excel_chart, apply_excel_dashboard, excel_chart_prompt, is_excel_chart_request, local_excel_chart_plan
from workbook_edits import is_excel_edit_request


ROWS = [
    (datetime(2026, 1, 5), "North", "Asha", 50, 1000), (datetime(2026, 1, 12), "South", "Ben", 4, 2000),
    (datetime(2026, 1, 20), "East", "Chen", 30, 1800), (datetime(2026, 1, 28), "North", "Asha", 10, 1500),
    (datetime(2026, 2, 3), "South", "Ben", 40, 800), (datetime(2026, 2, 10), "East", "Chen", 3, 1500),
    (datetime(2026, 2, 18), "North", "Dev", 25, 1500), (datetime(2026, 2, 25), "South", "Ben", 12, 1800),
    (datetime(2026, 3, 4), "East", "Dev", 60, 1200), (datetime(2026, 3, 15), "North", "Asha", 5, 2500),
]
SCHEMAS = {"Sales": {name: "text" for name in ("Date", "Region", "Salesperson", "Quantity", "Amount")},
           "staff": {"Name": "text", "Staff ID": "text"}}


def sales_content(*, extra_part: bool = False) -> bytes:
    workbook = Workbook()
    sales = workbook.active
    sales.title = "Sales"
    sales.append(["Date", "Region", "Salesperson", "Quantity", "Amount"])
    for row in ROWS:
        sales.append(row)
    staff = workbook.create_sheet("staff")
    staff.append(["Name", "Staff ID"])
    staff.append(["Asha Rao", "NOR-1021"])
    output = BytesIO()
    workbook.save(output)
    if not extra_part:
        return output.getvalue()
    # Simulate a part openpyxl would drop on re-save, such as a slicer.
    patched = BytesIO()
    with ZipFile(BytesIO(output.getvalue())) as source, ZipFile(patched, "w") as target:
        for item in source.infolist():
            target.writestr(item, source.read(item.filename))
        target.writestr("xl/slicers/slicer1.xml", b"<slicers/>")
    return patched.getvalue()


def plan(**overrides):
    return {
        "understood": True, "unhandled_conditions": [], "sheet": "Sales", "chart_type": "column",
        "category": "Region", "values": ["Amount"], "aggregate": "sum", "title": None, "sort": "none", "limit": None,
        **overrides,
    }


class ChartIntentTests(unittest.TestCase):
    def test_workbook_chart_prompts_are_detected_before_column_edits(self):
        for question in (
            "Add a column chart of Amount by Region in Sales sheet",
            "Create a pie chart of total Amount by Region in my Excel file",
            "Build a line chart of Amount by Date in the workbook",
            "insert a bar graph of average Amount per Salesperson",
            "Make an Excel chart of Amount by Region",
            "Make a chart of Amount by Region on a new sheet",
        ):
            with self.subTest(question=question):
                self.assertTrue(is_excel_chart_request(question))
        self.assertTrue(is_excel_edit_request("Add a column chart of Amount by Region in Excel"))

    def test_chat_charts_how_to_and_column_edits_keep_existing_routes(self):
        for question in (
            "show a chart of headcount by department", "plot Amount by Region", "How do I add a chart in Excel?",
            "Add column Reviewed to Sales", "Add a column without a chart to the workbook",
            "Delete row 2 from Sales",
        ):
            with self.subTest(question=question):
                self.assertFalse(is_excel_chart_request(question))

    def test_local_plans_resolve_type_aggregate_columns_and_sheet(self):
        self.assertEqual(
            local_excel_chart_plan("Create a pie chart of total amount by region in my Excel file", SCHEMAS, "staff"),
            plan(chart_type="pie"),
        )
        self.assertEqual(
            local_excel_chart_plan("insert a bar graph of average Amount per Salesperson", SCHEMAS, "Sales"),
            plan(chart_type="bar", category="Salesperson", aggregate="average"),
        )
        self.assertEqual(
            local_excel_chart_plan("Add a chart of Quantity and Amount by Region to the workbook", SCHEMAS, "Sales"),
            plan(chart_type="auto", values=["Quantity", "Amount"], aggregate="auto"),
        )
        self.assertEqual(
            local_excel_chart_plan("Add an excel chart of count of sales by Region", SCHEMAS, "Sales"),
            plan(chart_type="auto", values=[], aggregate="count"),
        )
        self.assertIsNone(local_excel_chart_plan("Add a chart of Profit by Region to the workbook", SCHEMAS, "Sales"))
        self.assertIsNone(local_excel_chart_plan("Add a chart for the north region in Excel", SCHEMAS, "Sales"))

    def test_prompt_lists_exact_plan_keys_and_untrusted_data(self):
        prompt = excel_chart_prompt([{"sheet": "Sales"}], "Sales")
        for key in ("understood", "unhandled_conditions", "chart_type", "category", "values", "aggregate", "limit"):
            self.assertIn(key, prompt)
        self.assertIn("untrusted", prompt)


class ApplyChartTests(unittest.TestCase):
    def test_same_sheet_preserves_existing_chart_image_formula_cache_and_repeated_exports(self):
        from data_explorer import build_charts, plan_charts

        workbook = load_workbook(BytesIO(sales_content()))
        existing = LineChart()
        existing.add_data(Reference(workbook["Sales"], min_col=5, min_row=1, max_row=11), titles_from_data=True)
        workbook["Sales"].add_chart(existing, "J2")
        png = BytesIO()
        Image.new("RGB", (20, 20), "white").save(png, format="PNG")
        workbook["Sales"].add_image(ExcelImage(png), "G15")
        buffer = BytesIO()
        workbook.save(buffer)
        workbook.close()
        cached = BytesIO()
        with ZipFile(BytesIO(buffer.getvalue())) as source, ZipFile(cached, "w") as target:
            for item in source.infolist():
                value = source.read(item.filename)
                if item.filename == "xl/worksheets/sheet1.xml":
                    value = value.replace(b'<c r="E2" t="n"><v>1000</v></c>', b'<c r="E2"><f>500+500</f><v>1000</v></c>')
                target.writestr(item, value)
        original = cached.getvalue()
        frame = read_excel_sheet(original, "Sales")
        charts, _ = build_charts(frame, plan_charts("make charts", frame))
        arguments = {
            "source_rows": len(frame), "source_columns": len(frame.columns), "missing_values": 0, "placement": "same_sheet",
        }
        tables = [(chart.title, chart.table, chart.chart_type) for chart in charts]
        first = apply_excel_dashboard(original, "Sales", tables, **arguments)
        repeated = apply_excel_dashboard(first.content, "Sales", tables, **arguments)
        self.assertEqual(excel_sheet_names(repeated.content), ["Sales", "staff"])
        pd.testing.assert_frame_equal(read_excel_sheet(repeated.content, "Sales"), frame)
        final = load_workbook(BytesIO(repeated.content))
        self.addCleanup(final.close)
        self.assertEqual(final["Sales"]["E2"].value, "=500+500")
        self.assertEqual(len(final["Sales"]._charts), 9)
        self.assertEqual(len(final["Sales"]._images), 1)
        with ZipFile(BytesIO(original)) as before, ZipFile(BytesIO(repeated.content)) as after:
            for name in before.namelist():
                if name.startswith(("xl/charts/", "xl/media/")) or name == "xl/worksheets/sheet2.xml":
                    self.assertEqual(before.read(name), after.read(name), name)
            old_anchors = list(ET.fromstring(before.read("xl/drawings/drawing1.xml")))
            new_anchors = list(ET.fromstring(after.read("xl/drawings/drawing1.xml")))
            for old, new in zip(old_anchors, new_anchors):
                self.assertEqual(ET.tostring(old), ET.tostring(new))
            ids = [
                item.attrib["id"] for item in ET.fromstring(after.read("xl/drawings/drawing1.xml")).iter()
                if item.tag.endswith("}cNvPr")
            ]
            self.assertEqual(len(ids), len(set(ids)))

    def test_same_sheet_source_range_preserves_blank_row_coordinates_and_blocks_destructive_edits(self):
        from workbook_edits import apply_workbook_edit

        workbook = Workbook()
        sheet = workbook.active
        sheet.title = "Team's data"
        sheet.append(["Department", "Salary"])
        sheet.append(["IT", 100])
        sheet.append([None, None])
        sheet.append(["Sales", 200])
        output = BytesIO()
        workbook.save(output)
        workbook.close()
        table = pd.DataFrame({"Department": ["IT", "Sales"], "Average Salary": [100, 200]})
        result = apply_excel_dashboard(
            output.getvalue(), sheet.title, [("Average salary", table, "bar")],
            source_rows=2, source_columns=2, missing_values=0, placement="same_sheet",
        )
        frame = read_excel_sheet(result.content, sheet.title)
        self.assertEqual(frame.attrs["excel_row_numbers"], [2, 4])
        self.assertEqual(frame.to_dict("list"), {"Department": ["IT", "Sales"], "Salary": [100, 200]})
        with self.assertRaisesRegex(ValueError, "same-sheet charts"):
            apply_workbook_edit(result.content, {
                "understood": True, "unhandled_conditions": [], "sheet": sheet.title,
                "steps": [{"op": "add_column", "name": "Reviewed", "expression": {"literal": True}}],
            })

    def test_same_sheet_dashboard_keeps_source_cells_and_can_be_queried_again(self):
        from data_explorer import build_charts, plan_charts

        original = sales_content(extra_part=True)
        frame = read_excel_sheet(original, "Sales")
        charts, problems = build_charts(frame, plan_charts("make charts", frame))
        self.assertEqual(problems, [])
        result = apply_excel_dashboard(
            original, "Sales", [(chart.title, chart.table, chart.chart_type) for chart in charts],
            source_rows=len(frame), source_columns=len(frame.columns), missing_values=0, placement="same_sheet",
        )
        self.assertEqual(result.sheet, "Sales")
        self.assertEqual(excel_sheet_names(result.content), ["Sales", "staff"])
        pd.testing.assert_frame_equal(read_excel_sheet(result.content, "Sales"), frame)
        before = load_workbook(BytesIO(original))
        after = load_workbook(BytesIO(result.content))
        self.addCleanup(before.close)
        self.addCleanup(after.close)
        for row in before["Sales"]:
            for cell in row:
                self.assertEqual(after["Sales"][cell.coordinate].value, cell.value, cell.coordinate)
        self.assertEqual(len(after["Sales"]._charts), 4)
        self.assertGreater(after["Sales"]._charts[0].anchor._from.col, len(frame.columns))
        with ZipFile(BytesIO(original)) as source, ZipFile(BytesIO(result.content)) as saved:
            self.assertEqual(source.read("xl/slicers/slicer1.xml"), saved.read("xl/slicers/slicer1.xml"))
            self.assertEqual(source.read("xl/worksheets/sheet2.xml"), saved.read("xl/worksheets/sheet2.xml"))

    def test_dashboard_charts_use_exact_tables_and_preserve_all_existing_parts(self):
        from data_explorer import build_charts, plan_charts

        original = apply_excel_chart(sales_content(extra_part=True), plan()).content
        frame = read_excel_sheet(original, "Sales")
        charts, problems = build_charts(frame, plan_charts("make charts", frame))
        self.assertEqual(problems, [])
        tables = [(chart.title, chart.table, chart.chart_type) for chart in charts]
        result = apply_excel_dashboard(
            original, "Sales", tables, source_rows=len(frame), source_columns=len(frame.columns), missing_values=0,
        )
        self.assertEqual(result.sheet, "Sales Dashboard")
        with ZipFile(BytesIO(original)) as before, ZipFile(BytesIO(result.content)) as after:
            for name in before.namelist():
                if name not in {"xl/workbook.xml", "xl/_rels/workbook.xml.rels", "[Content_Types].xml"}:
                    self.assertEqual(before.read(name), after.read(name), name)
        workbook = load_workbook(BytesIO(result.content))
        self.addCleanup(workbook.close)
        dashboard = workbook[result.sheet]
        self.assertEqual(len(dashboard._charts), 4)
        self.assertEqual(dashboard["B3"].value, len(frame))
        self.assertEqual(dashboard["D3"].value, len(frame.columns))
        self.assertEqual(dashboard["D4"].value, 4)
        start = 7
        for chart, table in zip(dashboard._charts, charts):
            self.assertEqual(dashboard.cell(start - 1, 1).value, table.title)
            saved = pd.DataFrame(
                [
                    [dashboard.cell(row, column).value for column in (1, 2)]
                    for row in range(start + 1, start + 1 + len(table.table))
                ],
                columns=table.table.columns,
            )
            pd.testing.assert_frame_equal(saved, table.table, check_dtype=False)
            series = chart.series[0]
            values = series.val.numRef.f
            self.assertEqual(values, f"'{result.sheet}'!$B${start + 1}:$B${start + len(table.table)}")
            start += max(len(table.table) + 5, 24)
        repeated = apply_excel_dashboard(
            result.content, "Sales", tables, source_rows=len(frame), source_columns=len(frame.columns), missing_values=0,
        )
        self.assertEqual(repeated.sheet, "Sales Dashboard (2)")

    def test_dashboard_limits_and_unknown_sources_fail_explicitly(self):
        table = pd.DataFrame({"Group": ["One"], "Value": [1]})
        for source, charts in (
            ("Missing", [("Chart", table, "bar")]),
            ("Sales", []),
            ("Sales", [("Chart", table, "bar")] * 5),
            ("Sales", [("Chart", table, "unsupported")]),
            ("Sales", [("Chart", pd.DataFrame({"Group": range(501), "Value": range(501)}), "line")]),
        ):
            with self.subTest(source=source, charts=len(charts)), self.assertRaises(ValueError):
                apply_excel_dashboard(
                    sales_content(), source, charts, source_rows=10, source_columns=5, missing_values=0,
                )

    def test_summed_chart_is_added_on_new_sheet_without_touching_existing_parts(self):
        original = sales_content(extra_part=True)
        chart = apply_excel_chart(original, plan())
        self.assertEqual(chart.sheet, "Total Amount by Region")
        self.assertEqual(chart.aggregate, "sum")
        self.assertEqual(chart.frame.to_dict("list"), {"Region": ["North", "South", "East"], "Total Amount": [6500, 4600, 4500]})
        self.assertEqual(excel_sheet_names(chart.content), ["Sales", "staff", "Total Amount by Region"])
        self.assertTrue(read_excel_sheet(chart.content, chart.sheet).equals(chart.frame))
        with ZipFile(BytesIO(original)) as before, ZipFile(BytesIO(chart.content)) as after:
            for name in before.namelist():
                if name not in {"xl/workbook.xml", "xl/_rels/workbook.xml.rels", "[Content_Types].xml"}:
                    self.assertEqual(before.read(name), after.read(name), name)
        for sheet in ("Sales", "staff"):
            self.assertTrue(read_excel_sheet(original, sheet).equals(read_excel_sheet(chart.content, sheet)), sheet)
        book = load_workbook(BytesIO(chart.content))
        try:
            charts = book[chart.sheet]._charts
            self.assertEqual(len(charts), 1)
            self.assertEqual(charts[0].type, "col")
            self.assertEqual(charts[0].series[0].val.numRef.f, "'Total Amount by Region'!$B$2:$B$4")
        finally:
            book.close()

    def test_repeated_charts_accumulate_with_unique_sheet_names(self):
        content = sales_content()
        names = []
        for chart_type in ("pie", "pie", "bar"):
            chart = apply_excel_chart(content, plan(chart_type=chart_type))
            content = chart.content
            names.append(chart.sheet)
        self.assertEqual(names, ["Total Amount by Region", "Total Amount by Region (2)", "Total Amount by Region (3)"])
        book = load_workbook(BytesIO(content))
        try:
            self.assertEqual([len(book[name]._charts) for name in names], [1, 1, 1])
        finally:
            book.close()

    def test_aggregates_sorting_limits_dates_and_titles(self):
        content = sales_content()
        average = apply_excel_chart(content, plan(category="Salesperson", aggregate="average", sort="desc", limit=2))
        self.assertEqual(average.frame.columns.tolist(), ["Salesperson", "Average Amount"])
        self.assertEqual(average.frame["Salesperson"].tolist(), ["Asha", "Chen"])
        self.assertAlmostEqual(average.frame["Average Amount"].iloc[0], 5000 / 3)
        self.assertEqual(average.frame["Average Amount"].iloc[1], 1650)
        counted = apply_excel_chart(content, plan(values=[], aggregate="count", title="Orders per region"))
        self.assertEqual((counted.sheet, counted.frame["Count"].tolist()), ("Orders per region", [4, 3, 3]))
        timeline = apply_excel_chart(content, plan(chart_type="auto", category="Date", aggregate="auto"))
        self.assertEqual((timeline.chart_type, timeline.aggregate), ("line", "none"))
        self.assertEqual(timeline.frame["Date"].iloc[0], "2026-01-05")
        both = apply_excel_chart(content, plan(values=["Quantity", "Amount"], aggregate="auto"))
        self.assertEqual(both.frame.columns.tolist(), ["Region", "Total Quantity", "Total Amount"])

    def test_non_numeric_cells_are_reported_and_invalid_plans_change_nothing(self):
        workbook = Workbook()
        sheet = workbook.active
        sheet.title = "Sales"
        sheet.append(["Region", "Amount", "Note"])
        for row in (("North", 10, "a"), ("North", "Check code", "b"), ("South", 5, "c")):
            sheet.append(row)
        output = BytesIO()
        workbook.save(output)
        chart = apply_excel_chart(output.getvalue(), plan())
        self.assertEqual(chart.frame["Total Amount"].tolist(), [10, 5])
        self.assertIn("Ignored 1 non-numeric Amount cell(s).", chart.notes)
        for bad, message in (
            (plan(values=["Profit"]), "no column named 'Profit'"),
            (plan(values=["Note"]), "no numeric values"),
            (plan(chart_type="pie", values=["Amount", "Amount"]), "listed twice"),
            (plan(chart_type="radar"), "chart types"),
            (plan(understood=False, unhandled_conditions=["Filters are not supported."]), "Filters are not supported."),
            ({**plan(), "save_as": "x.xlsx"}, "exactly"),
            (plan(sheet="Missing"), "unavailable"),
        ):
            with self.subTest(message=message), self.assertRaisesRegex(ValueError, message):
                apply_excel_chart(output.getvalue(), bad)

    def test_pie_requires_one_series_and_sheet_names_are_sanitized(self):
        content = sales_content()
        with self.assertRaisesRegex(ValueError, "exactly one value column"):
            apply_excel_chart(content, plan(chart_type="pie", values=["Quantity", "Amount"]))
        chart = apply_excel_chart(content, plan(title="Q1: Sales & 'Returns' / Region [draft] with a very long name"))
        self.assertEqual(chart.sheet, "Q1_ Sales & 'Returns' _ Region")
        self.assertEqual(chart.title, "Q1: Sales & 'Returns' / Region [draft] with a very long name")
        self.assertIn(chart.sheet, excel_sheet_names(chart.content))


class ChartChatTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        with patch("dotenv.load_dotenv"):
            cls.chatbot = importlib.import_module("employee_sql_chatbot")

    def setUp(self):
        environment = {
            key: value for key, value in os.environ.items()
            if not key.upper().startswith(("AZURE_", "WORKFORCE_", "SMTP_"))
        }
        self.enterContext(patch.dict(os.environ, environment, clear=True))
        self.enterContext(patch("dotenv.load_dotenv"))
        self.enterContext(patch("backend_diagnostics.backend_logger", return_value=Mock()))
        self.enterContext(patch.object(self.chatbot, "get_engine", side_effect=AssertionError("SQL access attempted")))
        self.client = Mock()
        self.enterContext(patch.object(self.chatbot, "get_azure_openai_client", return_value=self.client))
        self.enterContext(patch("openai.AzureOpenAI", return_value=self.client))
        self.enterContext(patch("sqlalchemy.create_engine", side_effect=AssertionError("SQL access attempted")))
        self.chatbot.st.cache_resource.clear()
        self.addCleanup(self.chatbot.st.cache_resource.clear)
        self.upload = BytesIO(sales_content())
        self.upload.name = "sales.xlsx"
        self.enterContext(patch("streamlit.file_uploader", return_value=self.upload))
        self.app = AppTest.from_file(str(Path(__file__).resolve().parents[1] / "employee_sql_chatbot.py"))

    def test_chart_prompt_adds_native_chart_to_working_copy_without_ai(self):
        self.app.run(timeout=30)
        self.app.chat_input[0].set_value("Add a column chart of total Amount by Region in the workbook").run(timeout=30)
        self.assertEqual(len(self.app.exception), 0)
        message = self.app.session_state["chat_history"][-1]
        self.assertNotIn("error", message)
        self.assertEqual(message["kind"], "excel_chart")
        self.assertEqual(message["edited_sheet"], "Sales")
        self.assertIn("native Excel column chart", message["answer"])
        cache = self.app.session_state["excel_workbook"]
        self.assertEqual(cache["original_content"], self.upload.getvalue())
        book = load_workbook(BytesIO(cache["content"]))
        try:
            self.assertEqual(len(book["Sales"]._charts), 1)
        finally:
            book.close()
        self.app.chat_input[0].set_value("Add a pie chart of Amount by Salesperson to my Excel file").run(timeout=30)
        self.assertNotIn("error", self.app.session_state["chat_history"][-1])
        self.assertEqual(excel_sheet_names(self.app.session_state["excel_workbook"]["content"]),
                         ["Sales", "staff"])
        self.client.chat.completions.create.assert_not_called()
        labels = [button.label for button in self.app.get("download_button")]
        self.assertIn("Download updated workbook", labels)
        self.assertGreaterEqual(len(self.app.get("plotly_chart")), 2)

    def test_ai_chart_plan_is_repaired_once_before_commit(self):
        valid = plan(chart_type="line", category="Date", aggregate="none", title="Daily sales")
        self.client.chat.completions.create.side_effect = [
            SimpleNamespace(choices=[SimpleNamespace(finish_reason="stop", message=SimpleNamespace(content=json.dumps(value)))])
            for value in ({**valid, "values": ["Revenue"]}, valid)
        ]
        with patch.dict(os.environ, {
            "AZURE_OPENAI_DEPLOYMENT_MINI": "test", "AZURE_OPENAI_ENDPOINT": "https://unit-test.invalid",
            "AZURE_OPENAI_API_KEY": "test-only", "AZURE_OPENAI_API_VERSION": "2024-02-01",
        }):
            self.app.run(timeout=30)
            self.app.chat_input[0].set_value("Add a trend chart of daily sales to the workbook").run(timeout=30)
        self.assertEqual(len(self.app.exception), 0)
        self.assertEqual(self.client.chat.completions.create.call_count, 2)
        message = self.app.session_state["chat_history"][-1]
        self.assertNotIn("error", message)
        self.assertEqual((message["edited_sheet"], message["chart_type"]), ("Sales", "line"))

    def test_failed_chart_keeps_workbook_and_without_ai_explains_format(self):
        self.app.run(timeout=30)
        original = self.app.session_state["excel_workbook"]["content"]
        self.app.chat_input[0].set_value("Add a chart of Profit by Region to the workbook").run(timeout=30)
        self.assertEqual(len(self.app.exception), 0)
        message = self.app.session_state["chat_history"][-1]
        self.assertIn("Azure OpenAI", message["error"])
        self.assertEqual(self.app.session_state["excel_workbook"]["content"], original)

    def test_database_mode_does_not_build_excel_charts(self):
        self.app.run(timeout=30)
        original = self.app.session_state["excel_workbook"]["content"]
        self.app.radio(key="data_source").set_value("Employee database").run(timeout=30)
        self.app.chat_input[0].set_value("Add a column chart of Amount by Region").run(timeout=30)
        self.assertEqual(len(self.app.exception), 0)
        self.assertIn("Excel charts", self.app.session_state["chat_history"][-1]["error"])
        self.assertEqual(self.app.session_state["excel_workbook"]["content"], original)
        self.client.chat.completions.create.assert_not_called()


if __name__ == "__main__":
    unittest.main()
