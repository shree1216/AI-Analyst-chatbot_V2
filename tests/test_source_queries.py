import importlib
from io import BytesIO
import json
import os
from pathlib import Path
from types import SimpleNamespace
import unittest
from unittest.mock import Mock, patch

import pandas as pd
from sqlalchemy import Numeric, String, create_engine, text
from sqlalchemy.dialects.mssql import MONEY, SMALLMONEY, dialect
from sqlalchemy.exc import SQLAlchemyError
from sqlalchemy.pool import StaticPool
from streamlit.testing.v1 import AppTest

from backend_diagnostics import BackendTrace, diagnostic_log
from excel_support import read_excel_sheet
from source_queries import (
    allows_unmatched_sources, compile_source_query, discover_excel_sources, discover_sql_sources,
    execute_source_plan, requested_source_kinds, source_planner_prompt,
)


ROOT = Path(__file__).resolve().parents[1]
TEST_ENVIRONMENT = {
    key: value for key, value in os.environ.items()
    if not key.upper().startswith(("AZURE_", "SMTP_", "WORKFORCE_"))
}


def query(**changes):
    plan = {
        "understood": True, "unhandled_conditions": [], "columns": ["Department", "Amount"],
        "metrics": [], "filter": None, "group_by": [], "order_by": [],
    }
    plan.update(changes)
    return plan


def data_plan(sql=None, excel=None, calculation=None):
    return {"mode": "data", "message": "", "sql": sql or [], "excel": excel or [], "calculation": calculation}


def workbook_bytes():
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        pd.DataFrame({"Department": ["A", "B", "C", "D"], "Budget": [350, 70, 90, 60]}).to_excel(
            writer, sheet_name="Targets", index=False,
        )
        pd.DataFrame({"Code": ["PRIVATE_SAMPLE"], "Name": ["Lookup"]}).to_excel(
            writer, sheet_name="Lookup", index=False,
        )
    return output.getvalue()


class SourceQueryTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        with patch("dotenv.load_dotenv"), patch.dict(os.environ, TEST_ENVIRONMENT, clear=True):
            cls.chatbot = importlib.import_module("employee_sql_chatbot")

    def setUp(self):
        self.enterContext(patch("backend_diagnostics.backend_logger", return_value=Mock()))
        self.enterContext(patch("dotenv.load_dotenv"))
        self.enterContext(patch.dict(os.environ, TEST_ENVIRONMENT, clear=True))
        self.enterContext(patch.object(self.chatbot, "DATA_MODE", "demo"))
        self.enterContext(patch("socket.create_connection", side_effect=AssertionError("Unexpected network")))
        self.engine = create_engine("sqlite://")
        self.addCleanup(self.engine.dispose)
        with self.engine.begin() as connection:
            pd.DataFrame({"Department": ["A", "A", "B"], "Amount": [100, 200, 75]}).to_sql(
                "Actuals", connection, index=False,
            )
            pd.DataFrame({"Product": ["PRIVATE_SQL_DATA"], "Stock": [8]}).to_sql("Inventory", connection, index=False)
        self.events = []
        self.trace = BackendTrace(self.events)
        self.sql_sources, self.unavailable = discover_sql_sources(self.engine, self.trace)
        self.actual = next(key for key, value in self.sql_sources.items() if value.label.endswith(".Actuals"))
        self.excel_sources, _ = discover_excel_sources(workbook_bytes(), self.trace)
        self.sources = {**self.sql_sources, **self.excel_sources}
        self.read = Mock(side_effect=self.read_sql)

    def compiler(self, plan, schema, table):
        return self.chatbot.compile_data_plan(plan, schema, table, True, max_rows=50_001)

    def read_sql(self, sql, parameters):
        with self.engine.connect() as connection:
            return pd.read_sql(text(sql), connection, params=parameters)

    def execute(self, plan):
        return execute_source_plan(plan, self.sources, self.compiler, self.read, self.trace)

    def combined_plan(self):
        return data_plan(
            sql=[{"source": self.actual, "query": query(
                columns=["Department"], group_by=["Department"],
                metrics=[{"function": "sum", "column": "Amount"}],
            )}],
            excel=["excel_0"],
            calculation={
                "understood": True, "unhandled_conditions": [], "sheet": self.actual, "steps": [
                    {"op": "join", "sheet": "excel_0", "left_on": ["Department"], "right_on": ["Department"],
                     "columns": [{"column": "Budget", "name": "Budget"}], "how": "left"},
                    {"op": "calculate", "columns": [{"name": "Variance", "expression": {
                        "function": "subtract", "args": [{"column": "sum_Amount"}, {"column": "Budget"}],
                    }}]},
                ],
            },
        )

    def test_discovery_reads_real_sql_schema_and_excel_samples_without_sql_data(self):
        self.assertEqual(len(self.sql_sources), 2)
        self.assertEqual(self.sql_sources[self.actual].schema, {"Department": "text", "Amount": "number"})
        self.assertIn("Amount", self.sql_sources[self.actual].sql_types)
        self.assertEqual(self.excel_sources["excel_0"].frame.shape, (3, 2))
        prompt = source_planner_prompt(self.sources, [])
        self.assertIn("Targets", prompt)
        self.assertIn("Budget", prompt)
        self.assertIn("PRIVATE_SAMPLE", prompt)
        self.assertNotIn("PRIVATE_SQL_DATA", prompt)
        log = diagnostic_log(self.events)
        self.assertNotIn("PRIVATE_SAMPLE", log)
        self.assertNotIn("PRIVATE_SQL_DATA", log)

    def test_sql_only_does_not_load_or_calculate_excel(self):
        with patch("source_queries.read_excel_sheet", side_effect=AssertionError("Unneeded Excel read")):
            result, labels = self.execute(data_plan(sql=[{"source": self.actual, "query": query(
                columns=[], metrics=[{"function": "sum", "column": "Amount"}],
            )}]))
        self.assertEqual(result.iloc[0, 0], 375)
        self.assertEqual(labels, ["SQL: main.Actuals"])
        self.read.assert_called_once()
        self.assertEqual(self.events[-1]["route"], "sql")

    def test_explicit_source_mentions_identify_sql_excel_and_combined_questions(self):
        cases = {
            "How many employees are there?": set(),
            "How many employees from SQL?": {"sql"},
            "Get the total in the database": {"sql"},
            "Count rows from Excel": {"excel"},
            "Sum budgets in the workbook": {"excel"},
            "Compare SQL actuals with targets in Excel": {"sql", "excel"},
            "Use SQL Server and the Excel lookup": {"sql", "excel"},
            "Use SQL only, not Excel": {"sql"},
            "Count Excel rows without using SQL": {"excel"},
            "Give me top 5 sales people by using both sources": {"sql", "excel"},
            "Combine the sources to rank salespeople": {"sql", "excel"},
            "Use all available data sources": {"sql", "excel"},
            "Use both SQL sources": {"sql"},
            "Use both workbook sources": {"excel"},
            "Use both sources, excluding inactive employees": {"sql", "excel"},
            "Use both sources but exclude Excel": {"sql"},
            "Do not use both sources": set(),
        }
        for question, expected in cases.items():
            with self.subTest(question=question):
                self.assertEqual(requested_source_kinds(question), expected)

    def test_explicit_sources_override_incompatible_sidebar_modes(self):
        for question, selected, expected in (
            ("Give total sales by region by using both sources", "excel", "automatic"),
            ("Join SQL products with Excel sales", "sql", "automatic"),
            ("Get products from SQL only", "excel", "automatic"),
            ("Get staff from Excel only", "sql", "automatic"),
            ("Use both workbook sources", "excel", "excel"),
            ("Use both SQL sources", "sql", "sql"),
            ("Total sales by region", "excel", "excel"),
            ("Count staff", "automatic", "automatic"),
            ("Export employee details to Excel", "sql", "sql"),
            ("Create an Excel report for SQL employees", "sql", "sql"),
            ("Export sales from Excel to Excel", "sql", "automatic"),
        ):
            with self.subTest(question=question, selected=selected):
                self.assertEqual(self.chatbot.source_mode_for_question(question, selected), expected)

    def test_explicit_sql_excel_or_both_cannot_be_silently_replaced_with_other_sources(self):
        for plan, required in (
            (data_plan(excel=["excel_0"]), {"sql"}),
            (data_plan(sql=[{"source": self.actual, "query": query()}]), {"excel"}),
            (data_plan(excel=["excel_0"]), {"sql", "excel"}),
        ):
            with self.subTest(required=required), self.assertRaisesRegex(ValueError, "requested SQL/Excel"):
                execute_source_plan(plan, self.sources, self.compiler, self.read, self.trace, required_kinds=required)
        self.read.assert_not_called()

    def test_both_sources_phrase_requires_and_repairs_a_real_combined_plan(self):
        client = Mock()
        client.chat.completions.create.side_effect = [
            SimpleNamespace(choices=[
                SimpleNamespace(finish_reason="stop", message=SimpleNamespace(content=json.dumps(plan))),
            ]) for plan in (data_plan(excel=["excel_0"]), self.combined_plan())
        ]
        with patch.dict(os.environ, {"AZURE_OPENAI_DEPLOYMENT_MINI": "test"}), patch.object(
            self.chatbot, "get_azure_openai_client", return_value=client,
        ), patch.object(self.chatbot, "get_engine", return_value=self.engine):
            message = {"question": "Compare actual amounts with budgets using both sources",
                       "workbook_content": workbook_bytes()}
            self.chatbot.display_automatic_message(message, "both_sources")
        self.assertNotIn("error", message)
        self.assertEqual(client.chat.completions.create.call_count, 2)
        self.assertEqual(message["source_labels"], ["SQL: main.Actuals", "EXCEL: Targets"])
        self.assertEqual(message["result"]["Variance"].tolist(), [-50, 5])
        self.assertTrue(any(event["event"] == "planner_repair_requested" for event in message["backend_log"]))

    def test_unmatched_sources_do_not_offer_to_ignore_a_both_sources_requirement(self):
        with self.engine.begin() as connection:
            connection.execute(text("UPDATE Actuals SET Department = 'No overlap'"))
        clarification = {
            "mode": "clarification", "message": "Confirm Excel-only output instead?",
            "sql": [], "excel": [], "calculation": None,
        }
        client = Mock()
        client.chat.completions.create.side_effect = [
            SimpleNamespace(choices=[
                SimpleNamespace(finish_reason="stop", message=SimpleNamespace(content=json.dumps(plan))),
            ]) for plan in (self.combined_plan(), clarification)
        ]
        with patch.dict(os.environ, {"AZURE_OPENAI_DEPLOYMENT_MINI": "test"}), patch.object(
            self.chatbot, "get_azure_openai_client", return_value=client,
        ), patch.object(self.chatbot, "get_engine", return_value=self.engine):
            message = {"question": "Compare actual amounts with budgets using both sources",
                       "workbook_content": workbook_bytes()}
            self.chatbot.display_automatic_message(message, "unmatched_both")
        self.assertIn("no records matched", message["error"])
        self.assertNotIn("Confirm", message["error"])
        self.assertNotIn("answer", message)
        self.assertNotIn("excel_download", message)
        self.assertEqual(client.chat.completions.create.call_count, 2)

    def test_explicit_excel_question_does_not_inspect_sql(self):
        client = Mock()
        client.chat.completions.create.return_value = SimpleNamespace(choices=[
            SimpleNamespace(finish_reason="stop", message=SimpleNamespace(content=json.dumps(data_plan(excel=["excel_0"])))),
        ])
        with patch.dict(os.environ, {"AZURE_OPENAI_DEPLOYMENT_MINI": "test"}), patch.object(
            self.chatbot, "get_azure_openai_client", return_value=client,
        ), patch.object(self.chatbot, "get_engine", side_effect=AssertionError("Unneeded SQL discovery")):
            message = {"question": "Preview the targets from Excel", "workbook_content": workbook_bytes()}
            self.chatbot.display_automatic_message(message, "explicit_excel")
        self.assertNotIn("error", message)
        self.assertEqual(message["source_labels"], ["EXCEL: Targets"])
        self.assertFalse(message["source_warnings"])

    def test_explicit_sql_question_does_not_inspect_uploaded_excel(self):
        client = Mock()
        client.chat.completions.create.return_value = SimpleNamespace(choices=[
            SimpleNamespace(finish_reason="stop", message=SimpleNamespace(content=json.dumps(
                data_plan(sql=[{"source": self.actual, "query": query()}]),
            ))),
        ])
        with patch.dict(os.environ, {"AZURE_OPENAI_DEPLOYMENT_MINI": "test"}), patch.object(
            self.chatbot, "get_azure_openai_client", return_value=client,
        ), patch.object(self.chatbot, "get_engine", return_value=self.engine), patch.object(
            self.chatbot, "discover_excel_sources", side_effect=AssertionError("Unneeded Excel discovery"),
        ):
            message = {"question": "Get actuals from SQL", "workbook_content": workbook_bytes()}
            self.chatbot.display_automatic_message(message, "explicit_sql")
        self.assertNotIn("error", message)
        self.assertEqual(message["source_labels"], ["SQL: main.Actuals"])

    def test_sql_server_money_columns_are_numeric_for_actual_budget_comparisons(self):
        inspector = Mock(default_schema_name="main")
        inspector.get_table_names.return_value = ["Ledger"]
        inspector.get_columns.return_value = [
            {"name": "Actual", "type": MONEY()}, {"name": "Budget", "type": SMALLMONEY()},
        ]
        with patch("source_queries.inspect", return_value=inspector):
            sources, _ = discover_sql_sources(self.engine, self.trace)
        self.assertEqual(sources["sql_0"].schema, {"Actual": "number", "Budget": "number"})

    def test_sql_server_discovers_all_user_schemas_and_compiles_bounded_parameterized_reads(self):
        inspector = Mock()
        inspector.get_schema_names.return_value = ["dbo", "db_finance", "sys", "INFORMATION_SCHEMA"]
        inspector.get_table_names.return_value = ["Actuals"]
        inspector.get_columns.return_value = [
            {"name": "Department", "type": String()}, {"name": "Amount", "type": Numeric()},
        ]
        with patch("source_queries.inspect", return_value=inspector):
            sources, _ = discover_sql_sources(Mock(dialect=dialect()), self.trace)
        self.assertEqual({source.label for source in sources.values()}, {"dbo.Actuals", "db_finance.Actuals"})
        statement, _ = compile_source_query(query(
            filter={"column": "Department", "op": "eq", "value": "A"},
        ), sources["sql_1"], lambda plan, schema, table: self.chatbot.compile_data_plan(
            plan, schema, table, False, max_rows=50_001,
        ))
        self.assertIn("SELECT TOP 50001", statement.sql)
        self.assertIn("[db_finance].[Actuals]", statement.sql)
        self.assertIn(":filter_0", statement.sql)
        self.assertEqual(statement.parameters, {"filter_0": "a"})

    def test_excel_only_loads_all_rows_not_just_samples_and_never_executes_sql(self):
        compiler = Mock(side_effect=AssertionError("Unneeded SQL compilation"))
        plan = data_plan(excel=["excel_0"], calculation={
            "understood": True, "unhandled_conditions": [], "sheet": "excel_0", "steps": [
                {"op": "aggregate", "metrics": [{"name": "Total Budget", "function": "sum", "column": "Budget"}]},
            ],
        })
        result, labels = execute_source_plan(plan, self.sources, compiler, self.read, self.trace)
        self.assertEqual(result.iloc[0, 0], 570)
        self.assertEqual(labels, ["EXCEL: Targets"])
        self.read.assert_not_called()
        compiler.assert_not_called()
        self.assertEqual(self.events[-1]["rows"], 1)

    def test_combined_sales_ranking_matches_real_keys_and_rejects_fictitious_enrichment(self):
        employees = pd.DataFrame({
            "Employee ID": ["E1", "E2"], "Name": [" ava ", "BEN"], "Department": ["Sales", "Sales"],
        })
        with self.engine.begin() as connection:
            employees.to_sql("Employees", connection, index=False)
        sql_sources, _ = discover_sql_sources(self.engine, self.trace)
        employee_id = next(key for key, value in sql_sources.items() if value.label.endswith(".Employees"))
        output = BytesIO()
        pd.DataFrame({
            "Salesperson": ["AVA", "AVA", "ben", "Unmatched"], "amount": [100, 200, 50, 30],
        }).to_excel(output, sheet_name="Sales", index=False)
        excel_sources, _ = discover_excel_sources(output.getvalue(), self.trace)
        sources = {**sql_sources, **excel_sources}
        proposed = data_plan(
            sql=[{"source": employee_id, "query": query(columns=["Employee ID", "Name", "Department"])}],
            excel=["excel_0"], calculation={
                "understood": True, "unhandled_conditions": [], "sheet": "excel_0", "steps": [
                    {"op": "aggregate", "group_by": ["Salesperson"], "metrics": [
                        {"name": "Total Sales", "function": "sum", "column": "amount"},
                    ]},
                    {"op": "join", "sheet": employee_id, "left_on": ["Salesperson"], "right_on": ["Name"],
                     "columns": [{"column": "Employee ID", "name": "Employee ID"},
                                 {"column": "Department", "name": "Department"}], "how": "left"},
                    {"op": "sort", "by": [{"column": "Total Sales", "direction": "desc"}]},
                    {"op": "limit", "count": 5},
                ],
            },
        )
        result, labels = execute_source_plan(
            proposed, sources, self.compiler, self.read, self.trace, required_kinds={"sql", "excel"},
        )
        self.assertEqual(result["Total Sales"].tolist(), [300, 50, 30])
        self.assertEqual(result["Employee ID"].iloc[:2].tolist(), ["E1", "E2"])
        self.assertTrue(pd.isna(result["Employee ID"].iloc[2]))
        self.assertEqual(len(labels), 2)
        self.assertEqual(result.attrs["source_matches"][0]["matched_rows"], 2)
        self.assertEqual(result.attrs["source_matches"][0]["unmatched_rows"], 1)
        self.assertTrue(any(event.get("matched_rows") == 2 for event in self.events))
        with self.engine.begin() as connection:
            connection.execute(text("UPDATE Employees SET Name = 'No overlap'"))
        for how in ("left", "inner", "outer"):
            proposed["calculation"]["steps"][1]["how"] = how
            with self.subTest(how=how), self.assertRaisesRegex(ValueError, "no records matched"):
                execute_source_plan(proposed, sources, self.compiler, self.read, self.trace)
        proposed["calculation"]["steps"][1]["how"] = "left"
        unmatched, _ = execute_source_plan(
            proposed, sources, self.compiler, self.read, self.trace, allow_unmatched=True,
        )
        self.assertEqual(len(unmatched), 3)
        self.assertTrue(unmatched["Employee ID"].isna().all())

    def test_combined_distinct_products_use_union_and_only_matching_catalog_sources(self):
        with self.engine.begin() as connection:
            pd.DataFrame({"code": ["A", "B", "B", None], "product": ["One", "Two", "Two", None]}).to_sql(
                "sql_products", connection, index=False,
            )
        sql_sources, _ = discover_sql_sources(self.engine, self.trace)
        product_id = next(key for key, value in sql_sources.items() if value.label.endswith(".sql_products"))
        output = BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            pd.DataFrame({"Code": [" a ", "C", "", "  "], "product": ["One", "Three", "", "  "]}).to_excel(
                writer, index=False, sheet_name="products",
            )
            pd.DataFrame({"Code": ["D"], "product": ["Transaction-only"], "Amount": [100]}).to_excel(
                writer, index=False, sheet_name="Sales",
            )
        excel_sources, _ = discover_excel_sources(output.getvalue(), self.trace)
        sources = {**sql_sources, **excel_sources}
        combined = data_plan(
            sql=[{"source": product_id, "query": query(columns=["code"])}],
            excel=["excel_0"],
            calculation={"understood": True, "unhandled_conditions": [], "sheet": product_id, "steps": [
                {"op": "append", "sheet": "excel_0", "columns": [{"column": "Code", "name": "code"}]},
                {"op": "calculate", "columns": [{
                    "name": "key", "expression": {"function": "lower", "args": [
                        {"function": "trim", "args": [{"column": "code"}]},
                    ]},
                }]},
                {"op": "filter", "where": {"function": "not", "args": [
                    {"function": "is_blank", "args": [{"column": "key"}]},
                ]}},
                {"op": "aggregate", "metrics": [{
                    "name": "Distinct product codes across SQL and Excel", "function": "count_distinct", "column": "key",
                }]},
            ]},
        )
        result, labels = execute_source_plan(
            combined, sources, self.compiler, self.read, self.trace, required_kinds={"sql", "excel"},
        )
        self.assertEqual(result.iloc[0, 0], 3)
        self.assertEqual(labels, ["SQL: main.sql_products", "EXCEL: products"])
        self.assertEqual(self.read.call_count, 1)
        implementation = {
            item["function"] for event in self.events for item in event.get("implementation", [])
        }
        self.assertIn("_append", implementation)
        self.assertTrue(any(event.get("operation") == "append" for event in self.events))
        prompt = source_planner_prompt(sources, [])
        self.assertIn("Do not ask the user to confirm tables", prompt)
        self.assertIn("ONE combined distinct count", prompt)
        self.assertIn("prefer a shared stable ID/code", prompt)
        self.assertIn("master/catalog sheet", prompt)
        self.read.reset_mock()
        combined["excel"].append("excel_1")
        with self.assertRaisesRegex(ValueError, "unused or missing sources"):
            execute_source_plan(combined, sources, self.compiler, self.read, self.trace)
        self.read.assert_not_called()

    def test_combined_question_aggregates_actuals_and_joins_targets_at_same_grain(self):
        result, labels = self.execute(self.combined_plan())
        self.assertEqual(result.to_dict("list"), {
            "Department": ["A", "B"], "sum_Amount": [300, 75], "Budget": [350, 70], "Variance": [-50, 5],
        })
        self.assertEqual(labels, ["SQL: main.Actuals", "EXCEL: Targets"])
        self.read.assert_called_once()
        self.assertIn("GROUP BY", self.read.call_args.args[0])
        self.assertEqual(
            next(event for event in self.events if event["event"] == "sources_selected")["route"], "combined",
        )

    def test_sales_products_and_staff_combine_without_duplicate_sales_or_guessed_ids(self):
        with self.engine.begin() as connection:
            pd.DataFrame({"code": ["A", "A", "B"], "product": ["One", "One", "Two"]}).to_sql(
                "sql_products", connection, index=False,
            )
            pd.DataFrame({"Employee ID": ["E1", "E2"], "Name": ["Asha", "Ben"],
                          "Department": ["Sales", "Operations"]}).to_sql("StaffDirectory", connection, index=False)
        sql_sources, _ = discover_sql_sources(self.engine, self.trace)
        catalog = next(key for key, source in sql_sources.items() if source.label.endswith(".sql_products"))
        directory = next(key for key, source in sql_sources.items() if source.label.endswith(".StaffDirectory"))
        output = BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            pd.DataFrame({"code": ["A", "B", "X"], "product": ["One", "Two", "Other"], "price": [50, 60, 70]}).to_excel(
                writer, sheet_name="products", index=False,
            )
            pd.DataFrame({
                "Code": ["A", "A", "B", "X"], "Region": ["North", "North", "South", "West"],
                "Salesperson": [" ASHA ", "Asha", "BEN", "Ben"], "amount": [10, 20, 5, 8],
                "Quantity": [100, 100, 100, 100], "price": [50, 50, 60, 70], "Total Sales": [9999] * 4,
            }).to_excel(writer, sheet_name="Sales", index=False)
            pd.DataFrame({"Name": ["Asha", "Ben"], "Staff ID": ["S1", "S2"]}).to_excel(
                writer, sheet_name="staff", index=False,
            )
        excel_sources, _ = discover_excel_sources(output.getvalue(), self.trace)
        sources = {**sql_sources, **excel_sources}
        proposed = data_plan(
            sql=[{"source": catalog, "query": query(
                columns=["code"], group_by=["code"], metrics=[{"function": "min", "column": "product", "name": "Product"}],
            )}], excel=["excel_1"],
            calculation={"understood": True, "unhandled_conditions": [], "sheet": "excel_1", "steps": [
                {"op": "join", "sheet": catalog, "left_on": ["Code"], "right_on": ["code"], "how": "left",
                 "columns": [{"column": "code", "name": "SQL Code"}, {"column": "Product", "name": "SQL Product"}]},
                {"op": "aggregate", "group_by": ["Region"], "metrics": [
                    {"name": "Total Sales", "function": "sum", "column": "amount"},
                ]},
            ]},
        )
        result, labels = execute_source_plan(
            proposed, sources, self.compiler, self.read, self.trace, required_kinds={"sql", "excel"},
        )
        self.assertEqual(result.to_dict("list"), {
            "Region": ["North", "South", "West"], "Total Sales": [30, 5, 8],
        })
        self.assertEqual(result.attrs["source_matches"][0]["unmatched_rows"], 1)
        self.assertEqual(len(labels), 2)
        proposed["sql"] = [{"source": directory, "query": query(columns=["Employee ID", "Name", "Department"])}]
        proposed["excel"] = ["excel_1", "excel_2"]
        proposed["calculation"]["steps"] = [
            {"op": "join", "sheet": "excel_2", "left_on": ["Salesperson"], "right_on": ["Name"], "how": "left",
             "columns": [{"column": "Staff ID", "name": "Staff ID"}, {"column": "Name", "name": "Staff Name"}]},
            {"op": "join", "sheet": directory, "left_on": ["Staff Name"], "right_on": ["Name"], "how": "left",
             "columns": [{"column": "Employee ID", "name": "Employee ID"}, {"column": "Department", "name": "Department"}]},
            {"op": "aggregate", "group_by": ["Staff Name", "Staff ID", "Employee ID", "Department"], "metrics": [
                {"name": "Total Sales", "function": "sum", "column": "amount"},
            ]},
        ]
        result, labels = execute_source_plan(
            proposed, sources, self.compiler, self.read, self.trace, required_kinds={"sql", "excel"},
        )
        self.assertEqual(result.to_dict("list"), {
            "Staff Name": ["Asha", "Ben"], "Staff ID": ["S1", "S2"], "Employee ID": ["E1", "E2"],
            "Department": ["Sales", "Operations"], "Total Sales": [30, 13],
        })
        self.assertEqual(len(labels), 3)
        prompt = source_planner_prompt(sources, [])
        self.assertIn("recorded row-level amount", prompt)
        self.assertIn("different ID namespaces", prompt)
        self.assertIn("one row per code", prompt)

    def test_sql_table_joins_grouping_text_metrics_and_having(self):
        with self.engine.begin() as connection:
            connection.execute(text('INSERT INTO "Actuals" VALUES (:department, :amount)'), {"department": "D", "amount": 25})
            pd.DataFrame({"Department": ["A", "B", "C"], "Label": ["Alpha", "Beta", "Gamma"]}).to_sql(
                "Departments", connection, index=False,
            )
        self.sources, _ = discover_sql_sources(self.engine, self.trace)
        catalog = next(key for key, value in self.sources.items() if value.label.endswith(".Departments"))
        for how, expected in (("inner", {"A", "B"}), ("left", {"A", "B", "D"}),
                              ("right", {"A", "B", "C"}), ("outer", {"A", "B", "C", "D"})):
            proposed = data_plan(sql=[
                {"source": self.actual, "query": query()},
                {"source": catalog, "query": query(columns=["Department", "Label"])},
            ], calculation={"understood": True, "unhandled_conditions": [], "sheet": self.actual, "steps": [
                {"op": "join", "sheet": catalog, "left_on": ["Department"], "right_on": ["Department"], "how": how,
                 "columns": [{"column": "Department", "name": "Catalog Department"}, {"column": "Label", "name": "Label"}]},
                {"op": "calculate", "columns": [{"name": "Department", "expression": {
                    "function": "coalesce", "args": [{"column": "Department"}, {"column": "Catalog Department"}],
                }}]},
                {"op": "aggregate", "group_by": ["Department"], "metrics": [
                    {"name": "Total", "function": "sum", "column": "Amount"},
                    {"name": "Label", "function": "min", "column": "Label"},
                ]},
            ]})
            with self.subTest(how=how):
                result, labels = self.execute(proposed)
                self.assertEqual(set(result["Department"]), expected)
                self.assertEqual(result.loc[result["Department"].eq("A"), "Total"].iloc[0], 300)
                self.assertEqual(result.loc[result["Department"].eq("A"), "Label"].iloc[0], "Alpha")
                self.assertTrue(all(label.startswith("SQL:") for label in labels))
                proposed["calculation"]["steps"].append({
                    "op": "filter", "where": {"function": "gt", "args": [{"column": "Total"}, 80]},
                })
                having, _ = self.execute(proposed)
                self.assertEqual(having["Department"].tolist(), ["A"])

    def test_sql_metric_names_and_aggregate_filters_compile_as_parameterized_having(self):
        proposed = query(
            columns=["Department"], group_by=["Department"],
            metrics=[{"function": "count", "column": "*", "name": "Employee Count"},
                     {"function": "avg", "column": "Amount", "name": "Average Amount"}],
            filter={"all": [{"column": "Amount", "op": "gt", "value": 50},
                            {"column": "Employee Count", "op": "gt", "value": 1}]},
            order_by=[{"column": "Average Amount", "direction": "desc"}],
        )
        result, _ = self.execute(data_plan(sql=[{"source": self.actual, "query": proposed}]))
        self.assertEqual(result.to_dict("records"), [{"Department": "A", "Employee Count": 2, "Average Amount": 150}])
        self.assertIn(" HAVING COUNT(*) > :filter_1", self.read.call_args.args[0])
        self.assertIn(" WHERE ", self.read.call_args.args[0])
        self.assertEqual(self.read.call_args.args[1], {"filter_0": 50, "filter_1": 1})
        proposed["filter"] = None
        proposed["having"] = {"column": "Employee Count", "op": "eq", "value": 1}
        proposed["order_by"] = [{"column": "Department", "direction": "asc"}]
        proposed["limit"] = 1
        result, _ = self.execute(data_plan(sql=[{"source": self.actual, "query": proposed}]))
        self.assertEqual(result["Department"].tolist(), ["B"])
        request, labels = compile_source_query(
            proposed, self.sources[self.actual],
            lambda plan, schema, table: self.chatbot.compile_data_plan(plan, schema, table, False, max_rows=50_001),
        )
        self.assertIn("SELECT TOP 1", request.sql)
        self.assertIn("HAVING COUNT(*) = :filter_0", request.sql)
        self.assertEqual(labels["row_count"], "Employee Count")
        proposed["having"] = None
        proposed["limit"] = 2
        proposed["filter"] = {"any": [{"column": "Department", "op": "eq", "value": "B"},
                                     {"column": "Employee Count", "op": "gt", "value": 1}]}
        result, _ = self.execute(data_plan(sql=[{"source": self.actual, "query": proposed}]))
        self.assertEqual(result["Department"].tolist(), ["A", "B"])
        for invalid_filter in (
            {"any": [{"column": "Amount", "op": "gt", "value": 0},
                     {"column": "Employee Count", "op": "gt", "value": 1}]},
            {"all": [None, {"column": "Department", "op": "eq", "value": "A"}]},
        ):
            proposed["filter"] = invalid_filter
            with self.subTest(filter=invalid_filter), self.assertRaises(ValueError):
                self.execute(data_plan(sql=[{"source": self.actual, "query": proposed}]))
        proposed["filter"] = None
        for metric in (
            {"function": "count", "column": "*", "name": None},
            {"function": "count", "column": "*", "name": "Department"},
            {"function": "count", "column": "*", "name": "Rows", "execute": "bad"},
        ):
            proposed["metrics"] = [metric]
            with self.subTest(metric=metric), self.assertRaises(ValueError):
                self.execute(data_plan(sql=[{"source": self.actual, "query": proposed}]))

    def test_zero_sales_combines_catalog_and_transactions_including_entirely_unmatched_sales(self):
        with self.engine.begin() as connection:
            pd.DataFrame({"code": ["A", "B", "C", "D", "E", "F"], "product": [
                "Returned", "Sold", "Zero", "Unsold", "Unknown amount", "Partly unknown",
            ]}).to_sql("sql_products", connection, index=False)
        self.sources, _ = discover_sql_sources(self.engine, self.trace)
        catalog = next(key for key, value in self.sources.items() if value.label.endswith(".sql_products"))
        proposed = data_plan(
            sql=[{"source": catalog, "query": query(columns=["code", "product"])}], excel=["excel_0"],
            calculation={"understood": True, "unhandled_conditions": [], "sheet": catalog, "steps": [
                {"op": "join", "sheet": "excel_0", "left_on": ["code"], "right_on": ["Code"], "how": "left",
                 "columns": [{"column": "Code", "name": "Sales Code"}, {"column": "amount", "name": "Sales Amount"}]},
                {"op": "aggregate", "group_by": ["code"], "metrics": [
                    {"name": "Product", "function": "min", "column": "product"},
                    {"name": "Total Sales", "function": "sum", "column": "Sales Amount"},
                    {"name": "Transactions", "function": "counta", "column": "Sales Code"},
                    {"name": "Known Amounts", "function": "counta", "column": "Sales Amount"},
                ]},
                {"op": "filter", "where": {"function": "or", "args": [
                    {"function": "and", "args": [
                        {"function": "eq", "args": [{"column": "Total Sales"}, 0]},
                        {"function": "eq", "args": [{"column": "Known Amounts"}, {"column": "Transactions"}]},
                    ]},
                    {"function": "eq", "args": [{"column": "Transactions"}, 0]},
                ]}},
                {"op": "calculate", "columns": [{"name": "Total Sales", "expression": {
                    "function": "coalesce", "args": [{"column": "Total Sales"}, 0],
                }}]},
            ]},
        )
        for codes, amounts, expected in (
            ([" a ", "A", "B", "C", "E", "F", "F"], [10, -10, 5, 0, None, 0, None], {"A", "C", "D"}),
            (["Unrelated"], [100], {"A", "B", "C", "D", "E", "F"}),
        ):
            output = BytesIO()
            pd.DataFrame({"Code": codes, "amount": amounts}).to_excel(output, sheet_name="Sales", index=False)
            client = Mock()
            client.chat.completions.create.return_value = SimpleNamespace(choices=[
                SimpleNamespace(finish_reason="stop", message=SimpleNamespace(content=json.dumps(proposed))),
            ])
            with self.subTest(codes=codes), patch.dict(os.environ, {"AZURE_OPENAI_DEPLOYMENT_MINI": "test"}), patch.object(
                self.chatbot, "get_azure_openai_client", return_value=client,
            ), patch.object(self.chatbot, "get_engine", return_value=self.engine):
                message = {"question": "Which products have zero sales, combine data from both sources?" if len(codes) > 1 else "both",
                           "workbook_content": output.getvalue()}
                self.chatbot.display_automatic_message(message, "zero_sales", (
                    ("user", "Which products have zero sales, combine data from both sources?"),
                    ("user", "You can match on code"),
                ))
                self.assertNotIn("error", message)
                self.assertEqual(set(message["result"]["code"]), expected)
                self.assertTrue(message["result"]["Total Sales"].eq(0).all())
                self.assertEqual(len(message["source_labels"]), 2)
                self.assertIn("excel_download", message)
                self.assertEqual(message["source_warnings"], [])
                client.chat.completions.create.assert_called_once()
        for question in ("Show unsold products", "Products without any sales", "Products never sold",
                         "Products with no recorded transactions", "Employees missing in Excel"):
            self.assertTrue(allows_unmatched_sources(question), question)
        for question in ("Top salespeople", "Employees with zero salary", "Join product prices"):
            self.assertFalse(allows_unmatched_sources(question), question)
        self.assertFalse(allows_unmatched_sources("both", prior_questions=(
            "Which products have zero sales?", "Show top salespeople",
        )))

    def test_planner_uses_followup_context_without_sending_prior_result_rows(self):
        client = Mock()
        client.chat.completions.create.return_value = SimpleNamespace(choices=[
            SimpleNamespace(finish_reason="stop", message=SimpleNamespace(content=json.dumps(data_plan()))),
        ])
        history = (("user", "Which products have zero sales?"), ("assistant", "PRIVATE_SQL_RESULT"),
                   ("user", "You can match on code"))
        for repair in (False, True):
            self.chatbot.generate_source_plan(
                client, "test", "Use the current schemas", "both", self.trace, prior_turns=history,
                previous_plan=data_plan() if repair else None, validation_error="Invalid calculation",
            )
            messages = client.chat.completions.create.call_args.kwargs["messages"]
            self.assertIn({"role": "user", "content": "You can match on code"}, messages)
            self.assertIn({"role": "user", "content": "Which products have zero sales?"}, messages)
            self.assertNotIn("PRIVATE_SQL_RESULT", json.dumps(messages))
            self.assertIn({"role": "user", "content": "both"}, messages)

    def test_equivalent_planner_calculation_wrapper_is_unwrapped_then_fully_validated(self):
        for key in ("PLAN", "plan"):
            plan = self.combined_plan()
            plan["calculation"] = {key: plan["calculation"]}
            with self.subTest(key=key):
                result, labels = self.execute(plan)
                self.assertEqual(result["Variance"].tolist(), [-50, 5])
                self.assertEqual(labels, ["SQL: main.Actuals", "EXCEL: Targets"])
        self.assertTrue(any(event["event"] == "plan_normalized" for event in self.events))
        invalid = self.combined_plan()
        invalid["calculation"] = {"PLAN": invalid["calculation"], "unhandled_conditions": ["extra condition"]}
        self.read.reset_mock()
        with self.assertRaisesRegex(ValueError, "calculation plan"):
            self.execute(invalid)
        self.read.assert_not_called()

    def test_sql_extraction_is_not_implicitly_truncated_to_chat_preview_size(self):
        with self.engine.begin() as connection:
            pd.DataFrame({"Department": ["A"] * 151, "Amount": list(range(151))}).to_sql(
                "Actuals", connection, if_exists="replace", index=False,
            )
        result, _ = self.execute(data_plan(sql=[{"source": self.actual, "query": query()}]))
        self.assertEqual(len(result), 151)
        limited, _ = self.execute(data_plan(sql=[{"source": self.actual, "query": query(limit=10)}]))
        self.assertEqual(len(limited), 10)

    def test_excess_sql_rows_fail_explicitly_instead_of_returning_partial_totals(self):
        self.read.side_effect = None
        self.read.return_value = pd.DataFrame({"c0": ["A"] * 50_001, "c1": [1] * 50_001})
        with self.assertRaisesRegex(ValueError, "50,000"):
            self.execute(data_plan(sql=[{"source": self.actual, "query": query()}]))

    def test_filters_are_parameterized_and_result_aliases_remain_readable(self):
        result, _ = self.execute(data_plan(sql=[{"source": self.actual, "query": query(
            columns=[], metrics=[{"function": "sum", "column": "Amount"}],
            filter={"column": "Department", "op": "eq", "value": "A"},
            order_by=[{"column": "sum_Amount", "direction": "desc"}],
        )}]))
        self.assertEqual(result.columns.tolist(), ["sum_Amount"])
        self.assertEqual(result.iloc[0, 0], 300)
        sql, parameters = self.read.call_args.args
        self.assertIn(":filter_0", sql)
        self.assertEqual(parameters, {"filter_0": "a"})
        self.assertNotIn("'a'", sql)

    def test_sql_metrics_accept_only_equivalent_workbook_aliases_and_case_variants(self):
        for function, column, expected_name, expected_value in (
            ("COUNT", "*", "row_count", 3),
            ("count_rows", "*", "row_count", 3),
            ("average", "Amount", "average_Amount", 125),
            ("SUM", "Amount", "sum_Amount", 375),
        ):
            with self.subTest(function=function):
                result, _ = self.execute(data_plan(sql=[{"source": self.actual, "query": query(
                    columns=[], metrics=[{"function": function, "column": column}],
                )}]))
                self.assertEqual(result.columns.tolist(), [expected_name])
                self.assertEqual(result.iloc[0, 0], expected_value)
        with self.assertRaisesRegex(ValueError, "count_rows"):
            self.execute(data_plan(sql=[{"source": self.actual, "query": query(
                columns=[], metrics=[{"function": "count_rows", "column": "Amount"}],
            )}]))

    def test_discovered_boolean_columns_use_typed_parameters(self):
        with self.engine.begin() as connection:
            pd.DataFrame({"Name": ["A", "B"], "Enabled": [True, False]}).to_sql("Flags", connection, index=False)
        sources, _ = discover_sql_sources(self.engine, self.trace)
        source = next(value for value in sources.values() if value.label.endswith(".Flags"))
        self.assertEqual(source.schema["Enabled"], "boolean")
        statement, _ = compile_source_query(query(
            columns=["Name"], filter={"column": "Enabled", "op": "eq", "value": True},
        ), source, self.compiler)
        self.assertIs(statement.parameters["filter_0"], True)
        result = self.read_sql(statement.sql, statement.parameters)
        self.assertEqual(result.iloc[:, 0].tolist(), ["A"])

    def test_quoted_table_and_column_names_cannot_change_query_structure(self):
        with self.engine.begin() as connection:
            pd.DataFrame({'Amount]"': [42]}).to_sql('Odd"; DROP TABLE Actuals;--', connection, index=False)
        sources, _ = discover_sql_sources(self.engine, self.trace)
        source_id = next(key for key, value in sources.items() if "Odd" in value.label)
        statement, labels = compile_source_query(query(columns=['Amount]"']), sources[source_id], self.compiler)
        result = self.read_sql(statement.sql, statement.parameters).rename(columns=labels)
        self.assertEqual(result.iloc[0, 0], 42)
        self.assertEqual(self.read_sql('SELECT COUNT(*) FROM "Actuals"', {}).iloc[0, 0], 3)

    def test_invalid_or_unused_sources_are_rejected_before_any_sql_reads(self):
        unknown_column = data_plan(sql=[{"source": self.actual, "query": query(columns=["Unknown"])}])
        raw_sql = data_plan(sql=[{"source": self.actual, "query": {**query(), "sql": "DROP TABLE Actuals"}}])
        extra_source = self.combined_plan()
        extra_source["excel"].append("excel_1")
        missing_source = self.combined_plan()
        missing_source["excel"] = []
        bad_join = self.combined_plan()
        bad_join["calculation"]["steps"][0]["sheet"] = []
        for invalid in (
            unknown_column, raw_sql, extra_source, missing_source, bad_join,
            data_plan(sql=[{"source": "unknown", "query": query()}]),
            data_plan(sql=[{"source": self.actual, "query": query(limit=50_001)}]),
            data_plan(sql=[{"source": self.actual, "query": query()}, {"source": self.actual, "query": query()}]),
        ):
            with self.subTest(plan=invalid), self.assertRaises(ValueError):
                self.execute(invalid)
        self.read.assert_not_called()

    def test_preview_skips_blank_rows_and_full_query_still_validates_remaining_rows(self):
        content = BytesIO()
        with pd.ExcelWriter(content, engine="openpyxl") as writer:
            pd.DataFrame({"A": [None, None, 1, 2, 3, 4]}).to_excel(writer, index=False, sheet_name="Data")
        self.assertEqual(read_excel_sheet(content.getvalue(), "Data", preview_rows=3)["A"].tolist(), [1, 2, 3])
        self.assertEqual(len(read_excel_sheet(content.getvalue(), "Data")), 4)
        for invalid in (0, -1, True, 50_001):
            with self.subTest(preview=invalid), self.assertRaises(ValueError):
                read_excel_sheet(content.getvalue(), "Data", preview_rows=invalid)

    def test_automatic_response_caches_results_and_keeps_samples_out_of_logs(self):
        client = Mock()
        client.chat.completions.create.return_value = SimpleNamespace(choices=[
            SimpleNamespace(finish_reason="stop", message=SimpleNamespace(content=json.dumps(self.combined_plan()))),
        ])
        with patch.dict(os.environ, {"AZURE_OPENAI_DEPLOYMENT_MINI": "test"}), patch.object(
            self.chatbot, "discover_sql_sources", return_value=(self.sql_sources, []),
        ), patch.object(self.chatbot, "get_azure_openai_client", return_value=client):
            with patch("employee_sql_chatbot.get_engine", return_value=self.engine):
                message = {"kind": "automatic", "question": "Compare SQL actual amounts against Excel budgets",
                           "workbook_content": workbook_bytes()}
                self.chatbot.display_automatic_message(message, "automatic")
                self.assertNotIn("error", message)
                self.assertEqual(message["result"]["Variance"].tolist(), [-50, 5])
                self.assertEqual(message["source_labels"], ["SQL: main.Actuals", "EXCEL: Targets"])
                self.assertNotIn("workbook_content", message)
                self.assertNotIn("PRIVATE_SAMPLE", diagnostic_log(message["backend_log"]))
                implementation = {
                    item["function"]: item["code"]
                    for event in message["backend_log"] for item in event.get("implementation", [])
                }
                for function in (
                    "execute_source_plan", "compile_source_query", "run_query", "execute_workbook_plan",
                    "_evaluate", "_aggregate", "_join", "spreadsheet_answer", "generate_excel_download",
                ):
                    self.assertIn(function, implementation)
                    self.assertIn(f"def {function}(", implementation[function])
                original_events = list(message["backend_log"])
                self.chatbot.display_automatic_message(message, "automatic")
                self.assertEqual(message["backend_log"], original_events)
                self.assertEqual(client.chat.completions.create.call_count, 1)

    def test_automatic_ui_queries_both_sources_and_displays_sources_and_download(self):
        employees = pd.read_excel(ROOT / "employee_dataset.xlsx")
        budgets = employees.groupby("Department")["Salary"].sum().reset_index(name="Budget")
        budgets["Budget"] += 1000
        output = BytesIO()
        budgets.to_excel(output, sheet_name="Targets", index=False)
        ui_plan = self.combined_plan()
        ui_plan["sql"][0]["source"] = "sql_0"
        ui_plan["sql"][0]["query"]["metrics"] = [{"function": "sum", "column": "Salary"}]
        ui_plan["calculation"]["sheet"] = "sql_0"
        ui_plan["calculation"]["steps"][1]["columns"][0]["expression"]["args"][0]["column"] = "sum_Salary"
        client = Mock()
        client.chat.completions.create.return_value = SimpleNamespace(choices=[
            SimpleNamespace(finish_reason="stop", message=SimpleNamespace(content=json.dumps(ui_plan))),
        ])
        upload = BytesIO(output.getvalue())
        upload.name = "targets.xlsx"
        environment = {
            "AZURE_OPENAI_ENDPOINT": "https://unit-test.invalid", "AZURE_OPENAI_API_KEY": "test-only",
            "AZURE_OPENAI_API_VERSION": "2024-02-01", "AZURE_OPENAI_DEPLOYMENT_MINI": "test",
            "WORKFORCE_DATA_MODE": "demo",
        }
        self.chatbot.st.cache_resource.clear()
        self.addCleanup(self.chatbot.st.cache_resource.clear)
        with patch.dict(os.environ, environment), patch("openai.AzureOpenAI", return_value=client), patch(
            "streamlit.file_uploader",
            side_effect=lambda *args, **kwargs: upload if kwargs.get("key") == "excel_upload_0" else None,
        ):
            app = AppTest.from_file(str(ROOT / "employee_sql_chatbot.py")).run(timeout=30)
            self.assertEqual(len(app.exception), 0)
            self.assertEqual(app.radio(key="data_source").value, "Automatic (SQL + Excel)")
            app.chat_input[0].set_value("Compare SQL actual amounts against Excel budgets").run(timeout=30)
            self.assertEqual(len(app.exception), 0)
            message = app.session_state["chat_history"][-1]
            self.assertNotIn("error", message)
            self.assertEqual(message["result"]["Variance"].tolist(), [-1000] * len(budgets))
            self.assertTrue(any("SQL: main.Employees; EXCEL: Targets" in caption.value for caption in app.caption))
            self.assertEqual(len(app.caption), 1)
            self.assertTrue(all(caption.value.startswith("Sources used:") for caption in app.caption))
            self.assertTrue(any(button.label == "Download Excel" for button in app.get("download_button")))
            self.assertTrue(any(button.label == "Download backend log" for button in app.get("download_button")))
            original_events = list(message["backend_log"])
            app.run(timeout=30)
            self.assertEqual(app.session_state["chat_history"][-1]["backend_log"], original_events)
            self.assertEqual(client.chat.completions.create.call_count, 1)
            self.assertIsNone(self.chatbot.SQL_TRACE.get())
            app.chat_input[0].set_value("open Targets").run(timeout=30)
            self.assertEqual(app.radio(key="data_source").value, "Automatic (SQL + Excel)")
            app.chat_input[0].set_value("close Excel").run(timeout=30)
            self.assertEqual(len(app.exception), 0)
            self.assertEqual(app.radio(key="data_source").value, "Automatic (SQL + Excel)")
            self.assertNotIn("excel_workbook", app.session_state)
            self.assertEqual(client.chat.completions.create.call_count, 1)

    def test_explicit_both_sources_works_from_excel_database_and_previous_failed_messages(self):
        employees = pd.read_excel(ROOT / "employee_dataset.xlsx")
        budgets = employees.groupby("Department")["Salary"].sum().reset_index(name="Budget")
        budgets["Budget"] += 1000
        output = BytesIO()
        budgets.to_excel(output, sheet_name="Targets", index=False)
        proposed = self.combined_plan()
        proposed["sql"][0]["source"] = "sql_0"
        proposed["sql"][0]["query"]["metrics"] = [{"function": "sum", "column": "Salary"}]
        proposed["calculation"]["sheet"] = "sql_0"
        proposed["calculation"]["steps"][1]["columns"][0]["expression"]["args"][0]["column"] = "sum_Salary"
        environment = {
            "AZURE_OPENAI_ENDPOINT": "https://unit-test.invalid", "AZURE_OPENAI_API_KEY": "test-only",
            "AZURE_OPENAI_API_VERSION": "2024-02-01", "AZURE_OPENAI_DEPLOYMENT_MINI": "test",
            "WORKFORCE_DATA_MODE": "demo",
        }
        for selected_mode, previous_failure in (
            ("Excel workbook", False), ("Employee database", False), ("Excel workbook", True),
        ):
            client = Mock()
            client.chat.completions.create.return_value = SimpleNamespace(choices=[
                SimpleNamespace(finish_reason="stop", message=SimpleNamespace(content=json.dumps(proposed))),
            ])
            upload = BytesIO(output.getvalue())
            upload.name = "targets.xlsx"
            self.chatbot.st.cache_resource.clear()
            self.addCleanup(self.chatbot.st.cache_resource.clear)
            with self.subTest(mode=selected_mode, old_error=previous_failure), patch.dict(
                os.environ, environment,
            ), patch("openai.AzureOpenAI", return_value=client), patch(
                "streamlit.file_uploader",
                side_effect=lambda *args, **kwargs: upload if kwargs.get("key") == "excel_upload_0" else None,
            ):
                app = AppTest.from_file(str(ROOT / "employee_sql_chatbot.py")).run(timeout=30)
                app.radio(key="data_source").set_value(selected_mode).run(timeout=30)
                question = "Compare SQL salaries against Excel budgets using both sources"
                if previous_failure:
                    app.session_state["chat_history"] = [{
                        "kind": "excel", "question": question, "excel_source": "targets.xlsx / Targets",
                        "error": "Choose A, B or C",
                    }]
                    app.run(timeout=30)
                else:
                    app.chat_input[0].set_value(question).run(timeout=30)
                self.assertEqual(len(app.exception), 0)
                message = app.session_state["chat_history"][-1]
                self.assertEqual(message["kind"], "automatic")
                self.assertNotIn("error", message)
                self.assertEqual(message["result"]["Variance"].tolist(), [-1000] * len(budgets))
                self.assertTrue(any("SQL: main.Employees; EXCEL: Targets" in caption.value for caption in app.caption))
                self.assertEqual(client.chat.completions.create.call_count, 1)
                app.run(timeout=30)
                self.assertEqual(client.chat.completions.create.call_count, 1)

    def test_invalid_model_plan_is_repaired_automatically_once_and_result_is_cached(self):
        invalid = data_plan(excel=["excel_0"], calculation={
            "understood": True, "unhandled_conditions": [], "sheet": "excel_0",
            "steps": [{"op": "aggregate", "metrics": [{"name": "Total Budget", "function": "unsupported", "column": "Budget"}]}],
        })
        valid = data_plan(excel=["excel_0"], calculation={
            "understood": True, "unhandled_conditions": [], "sheet": "excel_0",
            "steps": [{"op": "aggregate", "metrics": [{"name": "Total Budget", "function": "sum", "column": "Budget"}]}],
        })
        for corrected, expected_calls in ((valid, 2), (invalid, 2)):
            client = Mock()
            client.chat.completions.create.side_effect = [
                SimpleNamespace(choices=[
                    SimpleNamespace(finish_reason="stop", message=SimpleNamespace(content=json.dumps(proposed))),
                ]) for proposed in (invalid, corrected)
            ]
            with self.subTest(repaired=corrected is valid), patch.dict(
                os.environ, {"AZURE_OPENAI_DEPLOYMENT_MINI": "test"},
            ), patch.object(self.chatbot, "get_azure_openai_client", return_value=client), patch.object(
                self.chatbot, "get_engine", side_effect=AssertionError("Excel question must not inspect SQL"),
            ):
                message = {"question": "What is the total budget in Excel?", "workbook_content": workbook_bytes()}
                self.chatbot.display_automatic_message(message, "repair")
                self.assertEqual(client.chat.completions.create.call_count, expected_calls)
                self.assertTrue(any(event["event"] == "planner_repair_requested" for event in message["backend_log"]))
                repair_messages = client.chat.completions.create.call_args.kwargs["messages"]
                self.assertIn("unsupported", repair_messages[-2]["content"])
                self.assertIn("That summary function is not supported", repair_messages[-1]["content"])
                self.assertIn("Preserve every requested condition", repair_messages[-1]["content"])
                if corrected is valid:
                    self.assertNotIn("error", message)
                    self.assertEqual(message["result"].iloc[0, 0], 570)
                    self.assertIn("excel_download", message)
                    self.chatbot.display_automatic_message(message, "repair")
                    self.assertEqual(client.chat.completions.create.call_count, expected_calls)
                else:
                    self.assertIn("error", message)
                    self.assertNotIn("answer", message)
                    self.assertEqual(message["backend_log"][-1]["event"], "request_failed")

    def test_routine_clarification_is_reconsidered_without_asking_the_user(self):
        clarification = {"mode": "clarification", "message": "Which sources should I use? Choose A, B or C.",
                         "sql": [], "excel": [], "calculation": None}
        invalid = self.combined_plan()
        invalid["calculation"]["steps"].append({"op": "execute_python"})
        for corrected in (self.combined_plan(), clarification, invalid):
            client = Mock()
            client.chat.completions.create.side_effect = [
                SimpleNamespace(choices=[
                    SimpleNamespace(finish_reason="stop", message=SimpleNamespace(content=json.dumps(proposed))),
                ]) for proposed in (clarification, corrected)
            ]
            with self.subTest(mode=corrected["mode"]), patch.dict(
                os.environ, {"AZURE_OPENAI_DEPLOYMENT_MINI": "test"},
            ), patch.object(self.chatbot, "get_azure_openai_client", return_value=client), patch.object(
                self.chatbot, "get_engine", return_value=self.engine,
            ):
                message = {"question": "Compare actual amounts with budgets using both sources",
                           "workbook_content": workbook_bytes()}
                self.chatbot.display_automatic_message(message, "clarification_repair")
                self.assertEqual(client.chat.completions.create.call_count, 2)
                feedback = client.chat.completions.create.call_args.kwargs["messages"][-1]["content"]
                self.assertIn("automatic defaults", feedback)
                if corrected == self.combined_plan():
                    self.assertNotIn("error", message)
                    self.assertEqual(message["result"]["Variance"].tolist(), [-50, 5])
                else:
                    self.assertIn("error", message)
                    self.assertNotIn("answer", message)

    def test_excel_only_can_answer_when_sql_is_unavailable_with_an_explicit_warning(self):
        client = Mock()
        client.chat.completions.create.return_value = SimpleNamespace(choices=[
            SimpleNamespace(finish_reason="stop", message=SimpleNamespace(content=json.dumps(
                data_plan(excel=["excel_0"]),
            ))),
        ])
        with patch.dict(os.environ, {"AZURE_OPENAI_DEPLOYMENT_MINI": "test"}), patch.object(
            self.chatbot, "get_azure_openai_client", return_value=client,
        ), patch.object(self.chatbot, "get_engine", side_effect=SQLAlchemyError("PRIVATE_CONNECTION")), patch.object(
            self.chatbot, "run_query", side_effect=AssertionError("Unexpected SQL read"),
        ):
            message = {"question": "Preview the targets", "workbook_content": workbook_bytes()}
            self.chatbot.display_automatic_message(message, "excel_only")
        self.assertNotIn("error", message)
        self.assertEqual(len(message["result"]), 4)
        self.assertEqual(message["source_labels"], ["EXCEL: Targets"])
        self.assertTrue(message["source_warnings"])

    def test_database_mode_queries_both_tables_without_using_uploaded_excel(self):
        engine = create_engine(
            "sqlite://", poolclass=StaticPool, connect_args={"check_same_thread": False},
        )
        self.addCleanup(engine.dispose)
        with engine.begin() as connection:
            pd.DataFrame({
                "code": ["A", "B", "C", "D"], "product": ["One", "Two", "Three", "Four"],
                "price": [10, 20, 30, 40],
            }).to_sql("sql_products", connection, index=False)
        plans = [
            data_plan(sql=[{"source": "sql_1", "query": query(
                columns=[], metrics=[{"function": "avg", "column": "price"}],
            )}]),
            data_plan(sql=[{"source": "sql_1", "query": query(
                columns=[], metrics=[{"function": "count_distinct", "column": "product"}],
            )}]),
            data_plan(sql=[{"source": "sql_0", "query": query(
                columns=[], metrics=[{"function": "count", "column": "*"}],
            )}]),
            data_plan(sql=[{"source": "sql_0", "query": query(columns=["Department", "Salary"])}]),
        ]
        client = Mock()
        client.chat.completions.create.side_effect = [
            SimpleNamespace(choices=[
                SimpleNamespace(finish_reason="stop", message=SimpleNamespace(content=json.dumps(plan))),
            ]) for plan in plans
        ]
        environment = {
            "AZURE_OPENAI_ENDPOINT": "https://unit-test.invalid", "AZURE_OPENAI_API_KEY": "test-only",
            "AZURE_OPENAI_API_VERSION": "2024-02-01", "AZURE_OPENAI_DEPLOYMENT_MINI": "test",
            "WORKFORCE_DATA_MODE": "demo",
        }
        self.chatbot.st.cache_resource.clear()
        self.addCleanup(self.chatbot.st.cache_resource.clear)
        with patch.dict(os.environ, environment), patch("openai.AzureOpenAI", return_value=client), patch(
            "sqlalchemy.create_engine", return_value=engine,
        ), patch("source_queries.discover_excel_sources", side_effect=AssertionError("Unexpected Excel discovery")), patch(
            "streamlit.file_uploader", return_value=None,
        ):
            app = AppTest.from_file(str(ROOT / "employee_sql_chatbot.py")).run(timeout=30)
            app.radio(key="data_source").set_value("Employee database").run(timeout=30)
            app.session_state["excel_workbook"] = {"content": workbook_bytes()}
            for question, expected, source in (
                ("What is the average price in sql_products?", 25, "sql_products"),
                ("How many distinct products exist in sql_products?", 4, "sql_products"),
                ("How many employees?", 150, "Employees"),
            ):
                app.chat_input[0].set_value(question).run(timeout=30)
                self.assertEqual(len(app.exception), 0)
                message = app.session_state["chat_history"][-1]
                self.assertEqual(message["kind"], "database")
                self.assertNotIn("error", message)
                self.assertEqual(message["result"].iloc[0, 0], expected)
                self.assertEqual(message["source_labels"], [f"SQL: main.{source}"])
                self.assertNotIn("EXCEL", " ".join(message["source_labels"]))
                self.assertTrue(any(event["event"] == "query_completed" for event in message["backend_log"]))
                self.assertIn("excel_download", message)
            app.chat_input[0].set_value("Create a PDF report comparing department pay").run(timeout=30)
            self.assertEqual(len(app.exception), 0)
            report_message = app.session_state["chat_history"][-1]
            self.assertTrue(report_message["employee_report"])
            self.assertEqual(report_message["request"].response_kind, "business_report")
            self.assertIn("business_report", report_message)
            app.run(timeout=30)
            self.assertEqual(client.chat.completions.create.call_count, 4)

    def test_database_mode_rejects_excel_plans_without_reading_either_source(self):
        client = Mock()
        client.chat.completions.create.return_value = SimpleNamespace(choices=[
            SimpleNamespace(finish_reason="stop", message=SimpleNamespace(content=json.dumps(
                data_plan(excel=["excel_0"]),
            ))),
        ])
        with patch.dict(os.environ, {"AZURE_OPENAI_DEPLOYMENT_MINI": "test"}), patch.object(
            self.chatbot, "get_azure_openai_client", return_value=client,
        ), patch.object(self.chatbot, "get_engine", return_value=self.engine), patch.object(
            self.chatbot, "discover_excel_sources", side_effect=AssertionError("Unexpected Excel discovery"),
        ), patch.object(self.chatbot, "run_query", side_effect=AssertionError("Unexpected SQL read")):
            message = {
                "kind": "database", "sql_only": True, "question": "Show the targets",
                "workbook_content": workbook_bytes(),
            }
            self.chatbot.display_automatic_message(message, "database_wrong_source")
        self.assertIn("error", message)
        self.assertNotIn("answer", message)
        self.assertEqual(message["backend_log"][-1]["event"], "request_failed")

    def test_ambiguous_question_and_sql_outage_do_not_run_a_substitute_query(self):
        client = Mock()
        client.chat.completions.create.return_value = SimpleNamespace(choices=[
            SimpleNamespace(finish_reason="stop", message=SimpleNamespace(content=json.dumps({
                "mode": "clarification", "message": "Which department key and budget period should be matched?",
                "sql": [], "excel": [], "calculation": None,
            }))),
        ])
        with patch.dict(os.environ, {"AZURE_OPENAI_DEPLOYMENT_MINI": "test"}), patch.object(
            self.chatbot, "get_azure_openai_client", return_value=client,
        ), patch.object(self.chatbot, "get_engine", side_effect=SQLAlchemyError("PRIVATE_CONNECTION")), patch.object(
            self.chatbot, "run_query", side_effect=AssertionError("Unexpected SQL read"),
        ):
            message = {"question": "Compare actuals with targets", "workbook_content": workbook_bytes()}
            self.chatbot.display_automatic_message(message, "clarify")
        self.assertIn("budget period", message["error"])
        self.assertTrue(any("SQL is unavailable" in warning for warning in message["source_warnings"]))
        self.assertNotIn("PRIVATE_CONNECTION", diagnostic_log(message["backend_log"]))
        self.assertEqual(message["backend_log"][-1]["event"], "request_failed")
        self.assertIsNone(self.chatbot.SQL_TRACE.get())


if __name__ == "__main__":
    unittest.main()
