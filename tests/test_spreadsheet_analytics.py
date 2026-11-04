from datetime import date
from io import BytesIO
import os
from pathlib import Path
import unittest
from unittest.mock import patch

from openpyxl import load_workbook
import pandas as pd
from streamlit.testing.v1 import AppTest

from data_explorer import build_chart
from excel_support import generate_excel_download, read_excel_sheet
from spreadsheet_analytics import ColumnFilter, analysis_spec, export_analysis, filter_rows


ROOT = Path(__file__).resolve().parents[1]
TEST_ENVIRONMENT = {
    name: value for name, value in os.environ.items()
    if not name.upper().startswith(("AZURE_", "SMTP_", "WORKFORCE_"))
}


def sales_frame():
    return pd.DataFrame({
        "Region": ["North", "South", "North", "South", "North", "South"],
        "Revenue": [100, 200, 300, 400, 500, 600],
        "Cost": [50, 100, 150, 200, 250, 300],
        "Date": ["2025-03-01", "2025-01-01", "2025-02-01"] * 2,
    })


class SpreadsheetAnalyticsTests(unittest.TestCase):
    def test_filters_are_inclusive_combined_and_do_not_rewrite_source(self):
        frame = sales_frame()
        original = frame.copy(deep=True)
        filtered = filter_rows(frame, [
            ColumnFilter("Region", "category", ("North",)),
            ColumnFilter("Revenue", "number", lower=100, upper=300),
            ColumnFilter("Date", "date", lower=date(2025, 2, 1), upper=date(2025, 3, 1)),
        ])
        self.assertEqual(filtered.index.tolist(), [0, 2])
        pd.testing.assert_frame_equal(frame, original)
        filtered.iloc[0, 1] = -999
        pd.testing.assert_frame_equal(frame, original)

    def test_missing_and_unreadable_values_can_be_explicitly_included(self):
        frame = pd.DataFrame({"Amount": ["$100", "N/A", "bad", "1e999", 200]})
        rules = [ColumnFilter("Amount", "number", lower=100, upper=100)]
        self.assertEqual(filter_rows(frame, rules).index.tolist(), [0])
        rules = [ColumnFilter("Amount", "number", lower=100, upper=100, include_missing=True)]
        self.assertEqual(filter_rows(frame, rules).index.tolist(), [0, 1, 2, 3])
        frame = pd.DataFrame({"Region": ["north", " North ", None, "South"]})
        self.assertEqual(filter_rows(frame, [ColumnFilter("Region", "category", ("North",))]).index.tolist(), [0, 1])
        self.assertEqual(filter_rows(frame, [ColumnFilter("Region", "category")]).index.tolist(), [])

    def test_contains_filter_is_literal_not_regex(self):
        frame = pd.DataFrame({"Code": ["A.B", "axb", None, "n/a"]})
        self.assertEqual(filter_rows(frame, [ColumnFilter("Code", "contains", text="a.b")]).index.tolist(), [0])

    def test_invalid_filters_fail_explicitly(self):
        for rule in (
            ColumnFilter("Unknown", "category"),
            ColumnFilter("Revenue", "number", lower=10, upper=1),
            ColumnFilter("Revenue", "number"),
            ColumnFilter("Revenue", "execute"),
        ):
            with self.subTest(rule=rule), self.assertRaises(ValueError):
                filter_rows(sales_frame(), [rule])

    def test_all_summary_calculations_use_exact_filtered_values(self):
        frame = filter_rows(sales_frame(), [ColumnFilter("Revenue", "number", lower=200, upper=500)])
        expected = {"count": [2, 2], "sum": [800, 600], "average": [400, 300],
                    "median": [400, 300], "min": [300, 200], "max": [500, 400]}
        for aggregate, values in expected.items():
            with self.subTest(aggregate=aggregate):
                spec = analysis_spec(frame, "Grouped summary", dimension="Region", measure="Revenue", aggregate=aggregate)
                chart = build_chart(frame, spec)
                table = chart.table.sort_values("Region")
                self.assertEqual(table.iloc[:, 1].tolist(), values)
                self.assertEqual(spec.chart_type, "pie" if aggregate == "count" else "bar")

    def test_time_trend_sorts_chronologically_and_area_values_match(self):
        frame = sales_frame()
        spec = analysis_spec(frame, "Time trend", dimension="Date", measure="Revenue",
                             aggregate="sum", chart_type="area")
        chart = build_chart(frame, spec)
        self.assertEqual(chart.table["Month"].tolist(), ["2025-01", "2025-02", "2025-03"])
        self.assertEqual(chart.table["Total Revenue"].tolist(), [700, 900, 500])
        self.assertEqual(list(chart.figure.data[0].y), [700, 900, 500])
        self.assertEqual(chart.figure.data[0].fill, "tozeroy")

    def test_auto_bar_handles_signed_measures_and_pie_rejects_them(self):
        frame = pd.DataFrame({"Region": ["A", "B"], "Profit": [-10, 20]})
        spec = analysis_spec(frame, "Grouped summary", dimension="Region", measure="Profit", aggregate="sum")
        self.assertEqual(spec.chart_type, "bar")
        self.assertEqual(build_chart(frame, spec).table.iloc[:, 1].tolist(), [20, -10])
        for chart_type in ("pie", "doughnut"):
            for values in ([-10, 20], [0, 0]):
                with self.subTest(chart_type=chart_type, values=values), self.assertRaisesRegex(ValueError, "positive total"):
                    frame["Profit"] = values
                    build_chart(frame, analysis_spec(
                        frame, "Grouped summary", dimension="Region", measure="Profit",
                        aggregate="sum", chart_type=chart_type,
                    ))

    def test_full_group_summary_does_not_stop_at_25_categories(self):
        frame = pd.DataFrame({"Group": [f"Group {i}" for i in range(40)], "Value": range(40)})
        chart = build_chart(frame, analysis_spec(frame, "Grouped summary", dimension="Group", chart_type="bar"))
        self.assertEqual(len(chart.table), 40)
        self.assertEqual(chart.table.iloc[:, 1].sum(), 40)
        too_many = pd.DataFrame({"Group": [str(i) for i in range(501)]})
        with self.assertRaisesRegex(ValueError, "No groups were dropped"):
            analysis_spec(too_many, "Grouped summary", dimension="Group")

    def test_invalid_analysis_inputs_and_empty_filters_are_reported(self):
        for kwargs in (
            {"analysis": "Execute"}, {"analysis": "Distribution", "measure": "Region", "chart_type": "histogram"},
            {"analysis": "Time trend", "dimension": "Region"},
            {"analysis": "Grouped summary", "dimension": "Region", "chart_type": "scatter"},
            {"analysis": "Grouped summary", "dimension": "Region", "aggregate": "eval"},
        ):
            with self.subTest(kwargs=kwargs), self.assertRaises(ValueError):
                analysis_spec(sales_frame(), **kwargs)
        with self.assertRaisesRegex(ValueError, "No rows match"):
            analysis_spec(sales_frame().iloc[:0], "Grouped summary", dimension="Region")

    def test_export_contains_filtered_data_settings_and_exact_native_chart(self):
        frame = filter_rows(sales_frame(), [ColumnFilter("Region", "category", ("North",))])
        for analysis, kwargs, expected_type in (
            ("Grouped summary", {"dimension": "Region", "aggregate": "sum", "measure": "Revenue", "chart_type": "doughnut"}, "DoughnutChart"),
            ("Time trend", {"dimension": "Date", "aggregate": "sum", "measure": "Revenue", "chart_type": "area"}, "AreaChart"),
            ("Relationship", {"dimension": "Cost", "measure": "Revenue", "chart_type": "scatter"}, "ScatterChart"),
            ("Distribution", {"measure": "Revenue", "chart_type": "histogram"}, "BarChart"),
        ):
            with self.subTest(analysis=analysis):
                chart = build_chart(frame, analysis_spec(frame, analysis, **kwargs))
                content = export_analysis(frame, chart, [("Source", "test"), ("Filters", "Region = North")])
                pd.testing.assert_frame_equal(read_excel_sheet(content, "Filtered data"), frame.reset_index(drop=True),
                                              check_dtype=False, check_flags=False)
                workbook = load_workbook(BytesIO(content))
                self.addCleanup(workbook.close)
                dashboard = workbook["Filtered data Dashboard"]
                self.assertEqual(type(dashboard._charts[0]).__name__, expected_type)
                self.assertEqual(workbook["Analysis notes"]["B3"].value, "Region = North")
                for index, row in enumerate(chart.table.itertuples(index=False, name=None), 8):
                    self.assertEqual(tuple(dashboard.cell(index, column).value for column in (1, 2)), row)


class SpreadsheetAnalyticsUITests(unittest.TestCase):
    def setUp(self):
        self.enterContext(patch("dotenv.load_dotenv"))
        self.enterContext(patch.dict(os.environ, {**TEST_ENVIRONMENT, "WORKFORCE_DATA_MODE": "demo"}, clear=True))
        self.enterContext(patch("socket.create_connection", side_effect=AssertionError("Unexpected network access")))

    def start_app(self):
        self.upload = BytesIO(sales_frame().to_csv(index=False).encode())
        self.upload.name = "sales.csv"
        self.enterContext(patch(
            "streamlit.file_uploader",
            side_effect=lambda *args, **kwargs: self.upload if kwargs.get("key") == "excel_upload_0" else None,
        ))
        app = AppTest.from_file(str(ROOT / "employee_sql_chatbot.py")).run(timeout=30)
        self.assertEqual(len(app.exception), 0)
        self.assertEqual(len(app.get("plotly_chart")), 0)
        app.toggle(key="show_spreadsheet_analysis").set_value(True).run(timeout=30)
        self.assertEqual(len(app.exception), 0)
        self.key = f"analysis_{app.session_state['excel_workbook']['digest']}_sales"
        return app

    def test_interactive_filter_summary_chart_and_export_use_full_sheet(self):
        app = self.start_app()
        original = app.session_state["excel_workbook"]["content"]
        app.multiselect(key=f"{self.key}_filters").set_value(["Region"]).run(timeout=30)
        app.multiselect(key=f"{self.key}_filter_0_values").set_value(["North"]).run(timeout=30)
        app.selectbox(key=f"{self.key}_calculation").set_value("Sum").run(timeout=30)
        app.selectbox(key=f"{self.key}_Grouped summary_chart").set_value("doughnut").run(timeout=30)
        self.assertEqual(len(app.exception), 0)
        self.assertEqual(app.metric[1].value, "3")
        self.assertEqual(app.tabs[1].dataframe[0].value.to_dict("list"), {"Region": ["North"], "Total Revenue": [900]})
        self.assertEqual(app.session_state["excel_workbook"]["content"], original)
        app.button(key=f"{self.key}_export").click().run(timeout=30)
        self.assertEqual(len(app.exception), 0)
        self.assertTrue(any(button.label == "Download analysis with chart" for button in app.get("download_button")))
        app.chat_input[0].set_value("show total revenue by region").run(timeout=30)
        self.assertEqual(len(app.exception), 0)
        self.assertEqual(app.session_state["chat_history"][-1]["charts"][0].table.iloc[:, 1].sum(), 2100)

    def test_empty_filter_shows_actionable_warning_and_no_stale_chart(self):
        app = self.start_app()
        app.multiselect(key=f"{self.key}_filters").set_value(["Region"]).run(timeout=30)
        app.multiselect(key=f"{self.key}_filter_0_values").set_value([]).run(timeout=30)
        self.assertEqual(len(app.exception), 0)
        self.assertTrue(any("No rows match" in item.value for item in app.warning))
        self.assertEqual(len(app.get("plotly_chart")), 0)

    def test_numeric_and_date_controls_apply_inclusive_bounds_together(self):
        app = self.start_app()
        app.multiselect(key=f"{self.key}_filters").set_value(["Revenue", "Date"]).run(timeout=30)
        app.number_input(key=f"{self.key}_filter_1_lower").set_value(200.0)
        app.number_input(key=f"{self.key}_filter_1_upper").set_value(500.0)
        app.date_input(key=f"{self.key}_filter_3_dates").set_value((date(2025, 2, 1), date(2025, 3, 1))).run(timeout=30)
        self.assertEqual(len(app.exception), 0)
        self.assertEqual(app.metric[1].value, "2")
        self.assertEqual(app.tabs[2].dataframe[0].value["Revenue"].tolist(), [300, 400])

    def test_replacing_upload_refreshes_cached_analysis_and_controls(self):
        app = self.start_app()
        frame = pd.DataFrame({"Department": ["IT", "HR", "HR"], "Salary": [10, 20, 30]})
        self.upload = BytesIO(generate_excel_download([("Staff", frame)]))
        self.upload.name = "staff.xlsx"
        app.run(timeout=30)
        self.assertEqual(len(app.exception), 0)
        self.assertEqual(app.metric[0].value, "3")
        self.assertEqual(app.tabs[1].dataframe[0].value.columns.tolist(), ["Department", "Number of rows"])
        self.assertNotEqual(app.session_state["spreadsheet_analysis_frame"][0], self.key)

    def test_switching_analysis_types_renders_native_matching_charts(self):
        app = self.start_app()
        app.selectbox(key=f"{self.key}_analysis").set_value("Time trend").run(timeout=30)
        app.selectbox(key=f"{self.key}_Time trend_chart").set_value("area").run(timeout=30)
        self.assertEqual(len(app.exception), 0)
        self.assertEqual(app.tabs[1].dataframe[0].value.iloc[:, 0].tolist(), ["2025-01", "2025-02", "2025-03"])
        app.selectbox(key=f"{self.key}_analysis").set_value("Relationship").run(timeout=30)
        self.assertEqual(len(app.exception), 0)
        self.assertEqual(len(app.tabs[1].dataframe[0].value), 6)
        self.assertEqual(app.tabs[1].dataframe[0].value.columns.tolist(), ["Revenue", "Cost"])

    def test_automatic_mode_analysis_uses_more_than_sidebar_preview(self):
        frame = pd.concat([sales_frame()] * 3, ignore_index=True)
        self.upload = BytesIO(frame.to_csv(index=False).encode())
        self.upload.name = "sales.csv"
        self.enterContext(patch(
            "streamlit.file_uploader",
            side_effect=lambda *args, **kwargs: self.upload if kwargs.get("key") == "excel_upload_0" else None,
        ))
        app = AppTest.from_file(str(ROOT / "employee_sql_chatbot.py")).run(timeout=30)
        app.toggle(key="show_spreadsheet_analysis").set_value(True).run(timeout=30)
        self.assertEqual(len(app.exception), 0)
        self.assertEqual(len(app.session_state["excel_workbook"]["frame"]), 10)
        self.assertEqual(app.metric[0].value, "18")
        self.assertEqual(app.tabs[1].dataframe[0].value.iloc[:, 1].sum(), 18)

    def test_default_grouping_prefers_a_business_category_over_identifiers(self):
        frame = sales_frame()
        frame.insert(0, "Record ID", range(len(frame)))
        self.upload = BytesIO(frame.to_csv(index=False).encode())
        self.upload.name = "sales.csv"
        self.enterContext(patch(
            "streamlit.file_uploader",
            side_effect=lambda *args, **kwargs: self.upload if kwargs.get("key") == "excel_upload_0" else None,
        ))
        app = AppTest.from_file(str(ROOT / "employee_sql_chatbot.py")).run(timeout=30)
        app.toggle(key="show_spreadsheet_analysis").set_value(True).run(timeout=30)
        self.assertEqual(len(app.exception), 0)
        grouping = next(widget for widget in app.selectbox if widget.label == "Group by")
        self.assertEqual(grouping.value, "Region")
        self.assertEqual(app.tabs[1].dataframe[0].value.iloc[:, 1].sum(), 6)

    def test_richer_chat_charts_have_matching_editable_excel_exports(self):
        app = self.start_app()
        for question, expected_type in (
            ("draw a doughnut chart of region", "DoughnutChart"),
            ("show an area chart of total revenue monthly", "AreaChart"),
            ("show a scatter plot of revenue vs cost", "ScatterChart"),
        ):
            with self.subTest(question=question):
                app.chat_input[0].set_value(question).run(timeout=30)
                self.assertEqual(len(app.exception), 0)
                message = app.session_state["chat_history"][-1]
                self.assertNotIn("error", message)
                self.assertEqual(message["kind"], "uploaded_data")
                workbook = load_workbook(BytesIO(message["excel_download"]))
                self.addCleanup(workbook.close)
                self.assertEqual(type(workbook["sales"]._charts[0]).__name__, expected_type)

    def test_native_scatter_and_area_previews_do_not_turn_into_line_charts(self):
        import json

        app = self.start_app()
        for question, native_type in (
            ("Add a scatter chart of Revenue by Cost in the workbook", "scatter"),
            ("Add an area chart of total Revenue by Region in the workbook", "area"),
        ):
            with self.subTest(question=question):
                app.chat_input[0].set_value(question).run(timeout=30)
                self.assertEqual(len(app.exception), 0)
                message = app.session_state["chat_history"][-1]
                self.assertNotIn("error", message)
                self.assertEqual(message["chart_type"], native_type)
                spec = json.loads(app.chat_message[-1].get("plotly_chart")[0].proto.spec)
                trace = spec["data"][0]
                if native_type == "scatter":
                    self.assertEqual(trace["mode"], "markers")
                else:
                    self.assertEqual(trace["fill"], "tozeroy")


if __name__ == "__main__":
    unittest.main()
