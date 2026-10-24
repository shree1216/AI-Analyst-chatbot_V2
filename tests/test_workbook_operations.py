from datetime import date
import unittest
from unittest.mock import patch

import pandas as pd

from workbook_operations import evaluate_expression, execute_workbook_plan


def col(name):
    return {"column": name}


def lit(value):
    return {"literal": value}


def fn(name, *args, **fields):
    return {"function": name, "args": list(args), **fields}


def metric(name, function, column="*", **fields):
    return {"name": name, "function": function, "column": column, **fields}


def plan(*steps, sheet="Sales"):
    return {"understood": True, "unhandled_conditions": [], "sheet": sheet, "steps": list(steps)}


class WorkbookOperationTests(unittest.TestCase):
    def setUp(self):
        self.sales = pd.DataFrame({
            "Region": ["North", "North", "South", "South"],
            "Sales": [10.0, 30.0, 20.0, None],
            "Cost": [0, 10, 5, 0],
            "Product": ["A", "B", "A", "C"],
            "Date": pd.to_datetime(["2024-01-31", "2024-02-29", "2025-01-01", None]),
        })
        self.products = pd.DataFrame({
            "Code": ["a", "A", "B", None], "Name": ["First", "Last", "Second", "Blank"],
        })
        self.sheets = {"Sales": self.sales, "Products": self.products}

    def run_plan(self, *steps, **fields):
        return execute_workbook_plan(plan(*steps, **fields), self.sheets)

    def test_all_aggregate_functions_and_count_semantics(self):
        result = self.run_plan({"op": "aggregate", "metrics": [
            metric(name, name, "Sales") for name in ("sum", "average", "min", "max", "median", "count", "counta", "count_distinct")
        ] + [metric("rows", "count_rows"), metric("dates", "count", "Date"), metric("text", "count", "Region")]})
        self.assertEqual(result.iloc[0].to_dict(), {
            "sum": 60, "average": 20, "min": 10, "max": 30, "median": 20,
            "count": 3, "counta": 3, "count_distinct": 3, "rows": 4, "dates": 3, "text": 0,
        })

    def test_conditional_sum_count_average_multiple_nested_conditions(self):
        condition = fn("and", fn("eq", col("Region"), lit(" NORTH ")),
                       fn("or", fn("gt", col("Sales"), lit(15)), fn("not", fn("eq", col("Product"), lit("A")))))
        result = self.run_plan({"op": "aggregate", "metrics": [
            metric("SUMIFS", "sum", "Sales", where=condition),
            metric("COUNTIFS", "count_rows", where=condition),
            metric("AVERAGEIFS", "average", "Sales", where=condition),
        ]})
        self.assertEqual(result.iloc[0].tolist(), [30, 1, 30])

    def test_grouping_conditional_metrics_and_missing_no_match(self):
        condition = fn("eq", col("Product"), lit("missing"))
        result = self.run_plan({"op": "aggregate", "group_by": ["Region"], "metrics": [
            metric("Total", "sum", "Sales"), metric("Count", "count_rows"),
            metric("Absent sum", "sum", "Sales", where=condition),
            metric("Absent count", "count_rows", where=condition),
            metric("Absent average", "average", "Sales", where=condition),
        ]})
        self.assertEqual(result["Total"].tolist(), [40, 20])
        self.assertEqual(result["Count"].tolist(), [2, 2])
        self.assertEqual(result["Absent sum"].tolist(), [0, 0])
        self.assertEqual(result["Absent count"].tolist(), [0, 0])
        self.assertTrue(result["Absent average"].isna().all())
        self.sales["Sales"] = None
        result = self.run_plan({"op": "aggregate", "metrics": [metric("Total", "sum", "Sales")]})
        self.assertIsNone(result.iloc[0, 0])

    def test_filter_then_aggregate_uses_all_rows_not_display_limit(self):
        self.sheets["Sales"] = pd.concat([self.sales] * 100, ignore_index=True)
        result = self.run_plan({"op": "filter", "where": fn("in", col("Region"), lit(["north"]))},
                               {"op": "aggregate", "metrics": [metric("Total", "sum", "Sales")]})
        self.assertEqual(result.iloc[0, 0], 4000)

    def test_empty_filtered_results_and_empty_aggregates(self):
        steps = [{"op": "filter", "where": fn("gt", col("Sales"), lit(1000))}]
        self.assertTrue(self.run_plan(*steps).empty)
        result = self.run_plan(*steps, {"op": "aggregate", "metrics": [
            metric("Sum", "sum", "Sales"), metric("Count", "count_rows"), metric("Avg", "average", "Sales"),
        ]})
        self.assertEqual(result.iloc[0, :2].tolist(), [0, 0])
        self.assertIsNone(result.iloc[0, 2])
        result = self.run_plan(*steps, {"op": "aggregate", "group_by": ["Region"], "metrics": [metric("Sum", "sum", "Sales")]})
        self.assertEqual(result.columns.tolist(), ["Region", "Sum"])
        self.assertTrue(result.empty)

    def test_nested_if_guards_zero_division_without_changing_percentage_denominator(self):
        expression = fn("if", fn("eq", col("Cost"), lit(0)), lit(0),
                        fn("if", fn("gt", col("Sales"), lit(0)),
                           fn("multiply", fn("divide", col("Sales"), fn("sum", col("Sales"))), lit(100)), lit(-1)))
        result = evaluate_expression(expression, self.sales)
        for actual, expected in zip(result, [0, 50, 100 / 3, 0]):
            self.assertAlmostEqual(actual, expected)
        guarded = evaluate_expression(fn("if", fn("eq", col("Cost"), lit(0)), lit(0),
                                          fn("divide", col("Sales"), col("Cost"))), self.sales)
        self.assertEqual(guarded.tolist(), [0, 3, 4, 0])
        with self.assertRaisesRegex(ValueError, "divide by zero"):
            evaluate_expression(fn("divide", col("Sales"), col("Cost")), self.sales)

    def test_unused_if_branch_still_validates_function_and_column(self):
        for invalid in (fn("execute_python", lit("bad")), col("Missing")):
            with self.subTest(invalid=invalid), self.assertRaises(ValueError):
                evaluate_expression(fn("if", lit(True), lit(1), invalid), self.sales)

    def test_arithmetic_and_conditional_scalar_aggregates(self):
        expression = fn("subtract", fn("add", fn("power", lit(2), lit(3)), fn("abs", lit(-2))),
                        fn("multiply", lit(3), lit(2)))
        self.assertEqual(evaluate_expression(expression, self.sales).tolist(), [4] * 4)
        result = evaluate_expression(fn("sum", col("Sales"), where=fn("eq", col("Region"), lit("South"))), self.sales)
        self.assertEqual(result, 20)

    def lookup(self, **fields):
        return {"op": "lookup", "sheet": "Products", "left_on": ["Product"], "right_on": ["Code"],
                "columns": [{"column": "Name", "name": "Product Name"}], **fields}

    def test_exact_lookups_first_last_missing_and_original_keys_preserved(self):
        for keep, expected in (("first", "First"), ("last", "Last")):
            result = self.run_plan(self.lookup(keep=keep))
            self.assertEqual(result["Product"].tolist(), self.sales["Product"].tolist())
            self.assertEqual(result["Product Name"].iloc[:3].tolist(), [expected, "Second", expected])
            self.assertTrue(pd.isna(result["Product Name"].iloc[3]))
        self.sales.loc[3, "Product"] = None
        self.assertEqual(self.run_plan(self.lookup())["Product Name"].iloc[3], "Blank")

    def test_join_modes_duplicates_multiple_keys_and_no_mutation(self):
        original = {name: frame.copy(deep=True) for name, frame in self.sheets.items()}
        for how, count in (("left", 6), ("inner", 5), ("right", 6), ("outer", 7)):
            step = {**self.lookup(), "op": "join", "how": how}
            result = self.run_plan(step)
            self.assertEqual(len(result), count)
        for name in self.sheets:
            pd.testing.assert_frame_equal(self.sheets[name], original[name])
        self.products["Market"] = ["North", "South", "North", "South"]
        step = {**self.lookup(), "left_on": ["Product", "Region"], "right_on": ["Code", "Market"]}
        self.assertEqual(self.run_plan(step)["Product Name"].iloc[:3].tolist(), ["First", "Second", "Last"])

    def test_join_cardinality_checked_before_materializing_rows(self):
        self.sheets["Sales"] = pd.DataFrame({"Product": ["A"] * 225})
        self.sheets["Products"] = pd.DataFrame({"Code": ["A"] * 225, "Name": ["X"] * 225})
        with self.assertRaisesRegex(ValueError, "too many rows"):
            self.run_plan({**self.lookup(), "op": "join"})
        self.assertEqual(len(self.run_plan(self.lookup())), 225)

    def test_join_rejects_collisions_approximate_lookup_and_incompatible_keys(self):
        for fields in (
            {"columns": [{"column": "Name", "name": "Region"}]},
            {"keep": "nearest"}, {"how": "outer"}, {"match_mode": "approximate"},
            {"sheet": "Missing"}, {"right_on": ["Unknown"]},
        ):
            with self.subTest(fields=fields), self.assertRaises(ValueError):
                self.run_plan(self.lookup(**fields))
        self.products["Code"] = [1, 2, 3, 4]
        with self.assertRaisesRegex(ValueError, "incompatible types"):
            self.run_plan(self.lookup())

    def test_top_bottom_per_group_and_rank_ties(self):
        self.sales.loc[1, "Sales"] = 10
        for method, expected in (("min", [1, 1, 1]), ("dense", [1, 1, 1]), ("first", [1, 2, 1])):
            result = self.run_plan({"op": "rank", "column": "Sales", "name": "Rank", "direction": "desc",
                                    "group_by": ["Region"], "method": method})
            self.assertEqual(result["Rank"].iloc[:3].tolist(), expected)
        result = self.run_plan(
            {"op": "rank", "column": "Sales", "name": "Rank", "direction": "desc", "group_by": ["Region"], "method": "first"},
            {"op": "filter", "where": fn("lte", col("Rank"), lit(1))}, {"op": "select", "columns": ["Region", "Sales"]})
        self.assertEqual(result["Sales"].tolist(), [10, 20])
        result = self.run_plan({"op": "sort", "by": [{"column": "Sales", "direction": "asc"}]}, {"op": "limit", "count": 2})
        self.assertEqual(result["Sales"].tolist(), [10, 10])

    def test_pivots_numeric_totals_counts_missing_and_cardinality(self):
        result = self.run_plan({"op": "pivot", "index": ["Region"], "columns": "Product",
                                "values": "Sales", "function": "sum", "fill_value": 0})
        self.assertEqual(result.to_dict("list"), {"Region": ["North", "South"], "A": [10, 20], "B": [30, 0], "C": [0, 0]})
        result = self.run_plan({"op": "pivot", "index": ["Region"], "columns": "Product",
                                "values": "*", "function": "count_rows", "fill_value": 0})
        self.assertEqual(result["C"].tolist(), [0, 1])
        self.sheets["Sales"] = pd.DataFrame({"Region": ["N"] * 200, "Product": list(range(200)), "Sales": [1] * 200})
        with self.assertRaisesRegex(ValueError, "too wide"):
            self.run_plan({"op": "pivot", "index": ["Region"], "columns": "Product", "values": "Sales", "function": "sum"})

    def test_text_functions_and_zero_length_slices(self):
        self.sales["Text"] = ["  A   B  ", "XYZ", None, "=1+1"]
        for function, expected in (("left", ["", "", None, ""]), ("right", ["", "", None, ""])):
            result = evaluate_expression(fn(function, col("Text"), lit(0)), self.sales)
            self.assertEqual(result.fillna("<blank>").tolist(), [value if value is not None else "<blank>" for value in expected])
        self.assertEqual(evaluate_expression(fn("trim", col("Text")), self.sales).iloc[0], "A B")
        self.assertEqual(evaluate_expression(fn("right", col("Text"), lit(2)), self.sales).iloc[1], "YZ")
        self.assertEqual(evaluate_expression(fn("concat", fn("trim", col("Text")), lit("!")), self.sales).tolist(),
                         ["A B!", "XYZ!", "!", "=1+1!"])
        self.assertEqual(evaluate_expression(fn("upper", fn("lower", lit("AbC"))), self.sales).tolist(), ["ABC"] * 4)
        self.assertEqual(evaluate_expression(fn("length", lit("abc")), self.sales).tolist(), [3] * 4)
        for function in ("contains", "starts_with", "ends_with"):
            self.assertTrue(evaluate_expression(fn(function, lit("ABC"), lit("abc")), self.sales).all())
        self.assertEqual(evaluate_expression(fn("is_blank", col("Text")), self.sales).tolist(), [False, False, True, False])

    def test_date_components_constructor_comparison_and_completed_differences(self):
        self.assertEqual(evaluate_expression(fn("year", col("Date")), self.sales).iloc[:3].tolist(), [2024, 2024, 2025])
        self.assertEqual(evaluate_expression(fn("month", col("Date")), self.sales).iloc[:3].tolist(), [1, 2, 1])
        self.assertEqual(evaluate_expression(fn("day", col("Date")), self.sales).iloc[:3].tolist(), [31, 29, 1])
        result = evaluate_expression(fn("date", lit(2024), lit(2), lit(29)), self.sales)
        self.assertEqual(result.iloc[0], pd.Timestamp("2024-02-29"))
        self.assertTrue(evaluate_expression(fn("gte", fn("date", col("Date")), fn("date", lit("2024-01-31"))), self.sales).iloc[0])
        for unit, expected in (("days", 394), ("months", 12), ("years", 1)):
            result = evaluate_expression(fn("datediff", lit("2024-01-31"), lit("2025-02-28"), lit(unit)), self.sales)
            self.assertEqual(result.iloc[0], expected)
        self.assertEqual(evaluate_expression(fn("today"), self.sales).date(), pd.Timestamp.now(tz="UTC").date())

    def test_date_min_max_and_count_grouped_numeric_text_and_boolean_counts(self):
        result = self.run_plan({"op": "aggregate", "group_by": ["Region"], "metrics": [
            metric("Earliest", "min", "Date"), metric("Latest", "max", "Date"),
            metric("Numeric", "count", "Sales"), metric("Unique", "count_distinct", "Product"),
            metric("Nonempty", "counta", "Sales"),
        ]})
        self.assertEqual(result["Earliest"].tolist(), [pd.Timestamp("2024-01-31"), pd.Timestamp("2025-01-01")])
        self.assertEqual(result["Latest"].tolist(), [pd.Timestamp("2024-02-29"), pd.Timestamp("2025-01-01")])
        self.assertEqual(result["Numeric"].tolist(), [2, 1])
        self.assertEqual(result["Unique"].tolist(), [2, 2])
        self.assertEqual(result["Nonempty"].tolist(), [2, 1])
        self.sales["Mixed"] = pd.Series([True, "100", date(2025, 1, 1), 5], dtype=object)
        self.assertEqual(evaluate_expression(fn("count", col("Mixed")), self.sales), 2)

    def test_text_expansion_is_checked_before_concatenating(self):
        with self.assertRaisesRegex(ValueError, "combined text"):
            evaluate_expression(fn("concat", lit("A" * 32767), lit("B")), self.sales)
        large = pd.DataFrame({"Text": ["a"] * 1000})
        with self.assertRaisesRegex(ValueError, "combined text"):
            evaluate_expression(fn("concat", lit("A" * 10001), lit("B")), large)

    def test_pivot_name_collisions_and_missing_dimensions(self):
        self.sales.loc[0, "Region"] = None
        self.sales.loc[0, "Product"] = None
        result = self.run_plan({"op": "pivot", "index": ["Region"], "columns": "Product",
                                "values": "Sales", "function": "sum"})
        self.assertIn("(blank)", result.columns)
        self.assertEqual(result.loc[result["Region"].isna(), "(blank)"].iloc[0], 10)
        self.sales.loc[0, "Product"] = "Region"
        with self.assertRaisesRegex(ValueError, "collides"):
            self.run_plan({"op": "pivot", "index": ["Region"], "columns": "Product", "values": "Sales", "function": "sum"})

    def test_malformed_expression_shapes_arities_and_literal_values(self):
        for expression in (
            {"function": "sum", "args": []}, {"function": "date", "args": [lit(2025), lit(1)]},
            {"function": "sum", "args": [col("Sales")], "sql": "invalid"},
            fn("add", lit(1), lit(2), where=lit(True)), lit(float("nan")), lit(float("inf")),
            lit(10 ** 400), {"column": ["Sales"]}, lit({"python": "invalid"}),
            fn("in", col("Region"), lit("North")), lit(["a"]),
        ):
            with self.subTest(expression=expression), self.assertRaises(ValueError):
                evaluate_expression(expression, self.sales)

    def test_calculated_columns_preserve_input_and_can_be_used_by_later_steps(self):
        original = self.sales.copy(deep=True)
        result = self.run_plan(
            {"op": "calculate", "columns": [{"name": "Profit", "expression": fn("subtract", col("Sales"), col("Cost"))}]},
            {"op": "aggregate", "metrics": [metric("Total Profit", "sum", "Profit")]})
        self.assertEqual(result.iloc[0, 0], 45)
        pd.testing.assert_frame_equal(self.sales, original)

    def test_monthly_and_yearly_summaries(self):
        result = self.run_plan(
            {"op": "calculate", "columns": [
                {"name": "Year", "expression": fn("year", col("Date"))},
                {"name": "Month", "expression": fn("month", col("Date"))},
            ]},
            {"op": "filter", "where": fn("not", fn("is_blank", col("Year")))},
            {"op": "aggregate", "group_by": ["Year", "Month"], "metrics": [metric("Total", "sum", "Sales")]})
        self.assertEqual(result["Total"].tolist(), [10, 30, 20])
        self.assertEqual(result["Year"].tolist(), [2024, 2024, 2025])

    def test_invalid_dates_numeric_values_and_undefined_arithmetic_are_friendly(self):
        for expression in (
            fn("date", lit("01/02/2025")), fn("date", lit(45000)), fn("date", lit("2025-02-31")),
            fn("date", lit(2025), lit(2), lit(31)), fn("date", lit(2025.5), lit(2), lit(1)),
            fn("datediff", lit("2025-01-01"), lit("2024-01-01"), lit("days")),
            fn("datediff", lit("2025-01-01"), lit("2025-02-01"), lit("hours")),
            fn("multiply", lit(1e308), lit(1e308)), fn("power", lit(-1), lit(0.5)),
            fn("number", lit("not a number")), fn("left", lit("abc"), lit(-1)),
        ):
            with self.subTest(expression=expression), self.assertRaises(ValueError):
                evaluate_expression(expression, self.sales)
        self.sales.loc[0, "Sales"] = float("inf")
        with self.assertRaisesRegex(ValueError, "infinite number"):
            self.run_plan({"op": "select", "columns": ["Sales"]})

    def test_malformed_plans_and_unsupported_operations_never_run_partial_results(self):
        for invalid in (
            {}, [], plan({"op": "python", "code": "print(1)"}), plan({"op": "select", "columns": ["Missing"]}),
            {**plan({"op": "limit", "count": 1}), "unhandled_conditions": ["unsupported"]},
            plan({"op": "limit", "count": True}), plan({"op": "filter", "where": lit("yes")}),
            plan({"op": "select", "columns": ["Sales"]}, sheet="Missing"),
            plan({"op": "select", "columns": ["Sales"], "sql": "DROP TABLE data"}),
            plan({"op": "aggregate", "metrics": [metric("x", "eval", "Sales")]}),
            plan({"op": "sort", "by": [{"column": "Sales", "direction": "sideways"}]}),
            plan({"op": "rank", "column": "Sales", "name": "R", "direction": "desc", "method": "invalid"}),
            plan(*[{"op": "limit", "count": 1}] * 21),
        ):
            with self.subTest(invalid=invalid), self.assertRaises(ValueError):
                execute_workbook_plan(invalid, self.sheets)

    def test_expression_depth_and_intermediate_cell_bounds(self):
        expression = col("Sales")
        for _ in range(14):
            expression = fn("abs", expression)
        with self.assertRaisesRegex(ValueError, "too complex"):
            evaluate_expression(expression, self.sales)
        with patch("workbook_operations.MAX_CELLS", 19):
            with self.assertRaisesRegex(ValueError, "too large"):
                self.run_plan({"op": "select", "columns": ["Sales"]})
        with patch("workbook_operations.MAX_COLUMNS", 5):
            with self.assertRaisesRegex(ValueError, "too large"):
                self.run_plan({"op": "calculate", "columns": [{"name": "Extra", "expression": lit(1)}]})


if __name__ == "__main__":
    unittest.main()
