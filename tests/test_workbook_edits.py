import importlib
from io import BytesIO
from hashlib import sha256
import json
import os
from pathlib import Path
from types import SimpleNamespace
import unittest
from unittest.mock import Mock, patch

from chat_upload import patch_chat_upload
from xml.etree import ElementTree as ET
from zipfile import ZipFile

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font
import pandas as pd
from streamlit.testing.v1 import AppTest

from excel_support import append_result_sheet, preserve_formula_values, read_excel_sheet, spreadsheet_answer
from workbook_edits import apply_safe_workbook_clean, apply_workbook_edit, excel_edit_prompt, is_excel_edit_request, local_excel_edit_plan, result_sheet_request, workbook_clean_request
from data_grid import apply_grid_changes


def content():
    workbook = Workbook()
    sales = workbook.active
    sales.title = "Sales"
    for row in [
        ["Salesperson", "amount", "Code"],
        ["Amy", 2000, "A"], ["Bob", 3000, "B"], ["Cara", 2000, "C"], ["Dan", 1000, "D"],
    ]:
        sales.append(row)
    sales["B2"].number_format = "#,##0.00"
    sales["B2"].font = Font(bold=True)
    sales.column_dimensions["B"].width = 22
    sales.freeze_panes = "A2"
    sales.auto_filter.ref = "A1:C5"
    staff = workbook.create_sheet("staff")
    for row in [["Name", "Join date"], ["Amy", "2025-01-01"], ["Bob", "2023-01-01"], ["Cara", "2024-01-01"]]:
        staff.append(row)
    output = BytesIO()
    workbook.save(output)
    workbook.close()
    return output.getvalue()


class GridEditTests(unittest.TestCase):
    def workbook(self, rows):
        workbook = Workbook()
        sheet = workbook.active
        sheet.title = "Sales"
        for row in rows:
            sheet.append(row)
        output = BytesIO()
        workbook.save(output)
        return output.getvalue()

    def test_grid_edits_use_physical_rows_across_blank_rows(self):
        original = self.workbook([["Name", "Amount"], ["Amy", 10], [None, None], ["Bob", 20], ["Cara", 30]])
        frame = read_excel_sheet(original, "Sales")
        self.assertEqual(frame.attrs["excel_row_numbers"], [2, 4, 5])
        updated = apply_grid_changes(original, "Sales", frame, {
            "edited_rows": {1: {"Amount": 25}}, "added_rows": [{"Name": "Dan", "Amount": 40}],
            "deleted_rows": [0],
        })
        result = read_excel_sheet(updated, "Sales")
        self.assertEqual(result.to_dict("list"), {"Name": ["Bob", "Cara", "Dan"], "Amount": [25, 30, 40]})

    def test_grid_refuses_formula_cells_and_empty_changes(self):
        original = formula_content()
        frame = read_excel_sheet(original, "Sales")
        with self.assertRaisesRegex(ValueError, "F2 contains a formula"):
            apply_grid_changes(original, "Sales", frame, {"edited_rows": {0: {"Average sales": 5}}})
        with self.assertRaisesRegex(ValueError, "no grid changes"):
            apply_grid_changes(original, "Sales", frame, {"edited_rows": {}})
        with self.assertRaisesRegex(ValueError, "Deleting rows could break formulas"):
            apply_grid_changes(original, "Sales", frame, {"deleted_rows": [1]})


class SafeCleanTests(unittest.TestCase):
    def test_cleanup_prompt_recognizes_named_and_selected_tables_without_instructions(self):
        for question, target in (
            ("clean the employees_messy table", "employees_messy"),
            ("clean worksheet employees_messy", "employees_messy"),
            ("please tidy up the table", None),
            ("clean my data", None),
        ):
            self.assertEqual(workbook_clean_request(question).target, target)
            self.assertTrue(is_excel_edit_request(question))
        for question in ("How do I clean the table?", "Do not clean the table", "Show clean data"):
            self.assertIsNone(workbook_clean_request(question))

    def test_safe_cleanup_keeps_duplicates_ids_formulas_styles_and_ambiguous_values(self):
        book = Workbook()
        sheet = book.active
        sheet.title = "employees_messy"
        sheet.append(["Employee ID", "Name", "Email", "Department", "Salary", "Joining Date", "Comment"])
        sheet.append(["00123", " Ada ", "ADA@EXAMPLE.COM ", " it ", "$1,200.50", "25/12/2025", "Keep  double spaces"])
        duplicate = ["12345678901234567", "Lin", "Lin@example.com", "IT", "12,5", "01/02/2025", "=SUM(1,2)"]
        sheet.append(duplicate)
        sheet.append(["00999", "Bo", "bo@example.com", "Engg", "=600+600", "2025-01-01", "N/A"])
        sheet.append(duplicate)
        sheet["G3"].data_type = "s"
        sheet["G5"].data_type = "s"
        sheet["E2"].font = Font(bold=True, color="FF0000")
        sheet.freeze_panes = "A2"
        sheet.auto_filter.ref = "A1:G5"
        raw = BytesIO()
        book.save(raw)
        book.close()
        cached = BytesIO()
        with ZipFile(BytesIO(raw.getvalue())) as source, ZipFile(cached, "w") as target:
            for item in source.infolist():
                data = source.read(item.filename)
                if item.filename == "xl/worksheets/sheet1.xml":
                    root = ET.fromstring(data)
                    namespace = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
                    root.find(f".//{namespace}c[@r='E4']/{namespace}v").text = "1200"
                    data = ET.tostring(root, encoding="utf-8")
                target.writestr(item, data)
        original = cached.getvalue()
        result = apply_safe_workbook_clean(original, "employees_messy")
        self.assertEqual(len(result.frame), 4)
        self.assertEqual(result.frame["Employee ID"].tolist(), ["00123", "12345678901234567", "00999", "12345678901234567"])
        self.assertEqual(result.frame["Name"].iloc[0], "Ada")
        self.assertEqual(result.frame["Email"].iloc[0], "ADA@EXAMPLE.COM")
        self.assertEqual(result.frame["Department"].tolist(), ["IT", "IT", "Engg", "IT"])
        self.assertEqual(result.frame["Salary"].tolist(), [1200.5, "12,5", 1200, "12,5"])
        self.assertEqual(result.frame["Joining Date"].iloc[0], "2025-12-25")
        self.assertEqual(result.frame["Joining Date"].iloc[1], "01/02/2025")
        self.assertEqual(result.frame["Comment"].iloc[0], "Keep  double spaces")
        self.assertEqual(result.frame["Comment"].iloc[1], "=SUM(1,2)")
        self.assertTrue(pd.isna(result.frame["Comment"].iloc[2]))
        self.assertEqual(result.protected_cells, 1)
        self.assertEqual(result.ambiguous_values, 4)
        self.assertEqual(int(result.summary["Changed cells"].sum()), result.changed_cells)
        cleaned = load_workbook(BytesIO(result.content))
        self.addCleanup(cleaned.close)
        self.assertEqual(cleaned["employees_messy"]["E4"].value, "=600+600")
        self.assertTrue(cleaned["employees_messy"]["E2"].font.bold)
        self.assertEqual(cleaned["employees_messy"].freeze_panes, "A2")
        self.assertEqual(cleaned["employees_messy"].auto_filter.ref, "A1:G5")
        with ZipFile(BytesIO(original)) as before, ZipFile(BytesIO(result.content)) as after:
            for name in before.namelist():
                if name != "xl/worksheets/sheet1.xml":
                    self.assertEqual(before.read(name), after.read(name), name)
        again = apply_safe_workbook_clean(result.content, "employees_messy")
        self.assertEqual(again.changed_cells, 0)
        self.assertEqual(again.content, result.content)

    def test_safe_cleanup_does_not_guess_department_abbreviations_or_numeric_precision(self):
        book = Workbook()
        sheet = book.active
        sheet.title = "Data"
        sheet.append(["Department", "Salary", "Date"])
        sheet.append(["engg", "1234567890123456.78", "01/02/20"])
        sheet.append(["Engineering", "1e999", "2025-01-01"])
        output = BytesIO()
        book.save(output)
        book.close()
        cleaned = apply_safe_workbook_clean(output.getvalue(), "Data")
        self.assertEqual(cleaned.frame["Department"].tolist(), ["Engg", "Engineering"])
        self.assertEqual(cleaned.frame["Salary"].tolist(), ["1234567890123456.78", "1e999"])
        self.assertEqual(cleaned.frame["Date"].iloc[0], "01/02/20")
        self.assertEqual(cleaned.ambiguous_values, 3)

    def test_placeholder_only_rows_remain_at_original_worksheet_coordinates(self):
        book = Workbook()
        sheet = book.active
        sheet.title = "Data"
        sheet.append(["Department", "Salary"])
        sheet.append([" sales ", "100"])
        sheet.append([None, None])
        sheet.append(["N/A", "null"])
        sheet.append(["engineering", "$200"])
        sheet.append(["NULL", "N/A"])
        sheet["A4"].font = Font(bold=True)
        output = BytesIO()
        book.save(output)
        book.close()
        cleaned = apply_safe_workbook_clean(output.getvalue(), "Data")
        self.assertEqual(len(cleaned.frame), 4)
        self.assertEqual(cleaned.frame.attrs["excel_row_numbers"], [2, 4, 5, 6])
        self.assertTrue(cleaned.frame.iloc[[1, 3]].isna().all().all())
        reopened = load_workbook(BytesIO(cleaned.content))
        self.addCleanup(reopened.close)
        self.assertEqual(reopened["Data"].max_row, 6)
        self.assertEqual(reopened["Data"]["A5"].value, "Engineering")
        self.assertTrue(reopened["Data"]["A4"].font.bold)
        with ZipFile(BytesIO(output.getvalue())) as before, ZipFile(BytesIO(cleaned.content)) as after:
            namespace = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
            old_rows = [row.attrib for row in ET.fromstring(before.read("xl/worksheets/sheet1.xml")).iter(f"{namespace}row")]
            new_rows = [row.attrib for row in ET.fromstring(after.read("xl/worksheets/sheet1.xml")).iter(f"{namespace}row")]
            self.assertEqual(old_rows, new_rows)
        again = apply_safe_workbook_clean(cleaned.content, "Data")
        self.assertEqual(again.changed_cells, 0)
        self.assertEqual(again.content, cleaned.content)

    def test_ambiguous_date_formats_and_numeric_spacing_are_not_inferred(self):
        book = Workbook()
        sheet = book.active
        sheet.title = "Data"
        sheet.append(["Date", "Amount"])
        sheet.append(["01.02.2025", "1 23"])
        sheet.append(["25-Dec-20", "1e-999"])
        sheet.append(["01/01/2025", "1 234.50"])
        output = BytesIO()
        book.save(output)
        book.close()
        cleaned = apply_safe_workbook_clean(output.getvalue(), "Data")
        self.assertEqual(cleaned.frame["Date"].tolist(), ["01.02.2025", "25-Dec-20", "2025-01-01"])
        self.assertEqual(cleaned.frame["Amount"].tolist(), ["1 23", "1e-999", 1234.5])
        self.assertEqual(cleaned.ambiguous_values, 4)

    def test_rich_inline_and_metadata_cells_are_preserved(self):
        book = Workbook()
        sheet = book.active
        sheet.title = "Data"
        sheet.append(["Department", "Amount"])
        sheet.append([" sales ", "$100"])
        sheet.append([" it ", "$200"])
        sheet.append(["engineering", "$300"])
        output = BytesIO()
        book.save(output)
        book.close()
        protected = BytesIO()
        namespace = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
        with ZipFile(BytesIO(output.getvalue())) as source, ZipFile(protected, "w") as target:
            for item in source.infolist():
                data = source.read(item.filename)
                if item.filename == "xl/worksheets/sheet1.xml":
                    root = ET.fromstring(data)
                    inline = root.find(f".//{namespace}c[@r='A2']/{namespace}is")
                    text = inline.find(f"{namespace}t")
                    inline.remove(text)
                    ET.SubElement(inline, f"{namespace}r").append(text)
                    root.find(f".//{namespace}c[@r='A3']").set("cm", "1")
                    data = ET.tostring(root, encoding="utf-8")
                target.writestr(item, data)
        cleaned = apply_safe_workbook_clean(protected.getvalue(), "Data")
        self.assertEqual(cleaned.protected_cells, 2)
        self.assertEqual(cleaned.frame["Department"].tolist(), [" sales ", " it ", "Engineering"])
        with ZipFile(BytesIO(protected.getvalue())) as before, ZipFile(BytesIO(cleaned.content)) as after:
            original = ET.fromstring(before.read("xl/worksheets/sheet1.xml"))
            updated = ET.fromstring(after.read("xl/worksheets/sheet1.xml"))
            for reference in ("A2", "A3"):
                self.assertEqual(
                    ET.tostring(original.find(f".//{namespace}c[@r='{reference}']")),
                    ET.tostring(updated.find(f".//{namespace}c[@r='{reference}']")),
                )


def plan(*steps):
    return {"understood": True, "unhandled_conditions": [], "sheet": "Sales", "steps": list(steps)}


def positional_content():
    book = load_workbook(BytesIO(content()))
    for number in range(6, 13):
        book["Sales"].append([f"Person {number}", number * 100, f"Code {number}"])
    book["Sales"]["A7"] = None
    book["Sales"]["B7"] = None
    book["Sales"]["C7"] = None
    output = BytesIO()
    book.save(output)
    book.close()
    return output.getvalue()


def formula_content():
    book = load_workbook(BytesIO(content()))
    sales = book["Sales"]
    sales.append(["Eve", 2500, "E"])
    for address, value in {
        "D1": "product", "E1": "Rating", "F1": "Average sales",
        "D2": '="Widget"', "E2": '=IF(B2>1500,"High","Low")', "F2": "=AVERAGE(B2:B100)",
    }.items():
        sales[address] = value
    notes = book.create_sheet("Notes")
    notes.append(["Average"])
    notes.append(["=Sales!F2"])
    raw = BytesIO()
    book.save(raw)
    book.close()
    output = BytesIO()
    ns = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
    with ZipFile(BytesIO(raw.getvalue())) as source, ZipFile(output, "w") as target:
        for item in source.infolist():
            data = source.read(item.filename)
            if item.filename in {"xl/worksheets/sheet1.xml", "xl/worksheets/sheet3.xml"}:
                root = ET.fromstring(data)
                for cell in root.iter(f"{ns}c"):
                    if cell.find(f"{ns}f") is None:
                        continue
                    cached = cell.find(f"{ns}v")
                    cached.text = {"D2": "Widget", "E2": "High"}.get(cell.get("r"), "2100")
                    cell.set("t", "str" if cell.get("r") in {"D2", "E2"} else "n")
                data = ET.tostring(root)
            target.writestr(item, data)
    return output.getvalue()


def modern_formula_content(*, named_ranges=False):
    ns = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
    original = formula_content()
    output = BytesIO()
    with ZipFile(BytesIO(original)) as source, ZipFile(output, "w") as target:
        for item in source.infolist():
            data = source.read(item.filename)
            if item.filename == "xl/worksheets/sheet1.xml":
                root = ET.fromstring(data)
                for cell in root.iter(f"{ns}c"):
                    formula = cell.find(f"{ns}f")
                    if formula is not None:
                        formula.set("t", "array")
                        formula.set("ref", cell.attrib["r"])
                        if named_ranges and cell.attrib["r"] == "F2":
                            formula.text = "AVERAGE(SalesAmounts)"
                data = ET.tostring(root)
            if named_ranges and item.filename == "xl/workbook.xml":
                root = ET.fromstring(data)
                names = root.find(f"{ns}definedNames")
                if names is None:
                    names = ET.SubElement(root, f"{ns}definedNames")
                definition = ET.SubElement(names, f"{ns}definedName", {"name": "SalesAmounts"})
                definition.text = "'Sales'!$B$2:$B$100"
                data = ET.tostring(root)
            target.writestr(item, data)
    return output.getvalue()


def regional_content():
    original = modern_formula_content(named_ranges=True)
    book = load_workbook(BytesIO(original))
    book["Sales"]["G1"] = "Region"
    for row, region in enumerate(["North", "South", "North", "East", "South"], 2):
        book["Sales"].cell(row, 7, region)
    products = book.create_sheet("products")
    products.append(["code", "product"])
    products.append(["A", "Widget"])
    output = BytesIO()
    book.save(output)
    book.close()
    return preserve_formula_values(original, output.getvalue(), reject_changed=True)


def col(name):
    return {"column": name}


def fn(name, *args):
    return {"function": name, "args": list(args)}


class WorkbookEditTests(unittest.TestCase):
    def test_result_sheet_request_extracts_calculation_and_optional_name(self):
        calculation = "give sum of sales for all regions"
        self.assertEqual(result_sheet_request(calculation + " and save in other sheet in same workbook"),
                         (calculation, None))
        self.assertEqual(result_sheet_request(
            calculation + ' and save results into a new worksheet named "Regional totals" in same workbook',
        ), (calculation, "Regional totals"))
        self.assertIsNone(result_sheet_request(calculation))
        self.assertIsNone(result_sheet_request(calculation + " and do not save in another sheet"))
        self.assertIsNone(result_sheet_request("How to save results in a new worksheet"))

    def test_result_sheets_never_replace_existing_sheets_and_preserve_formula_values(self):
        original = regional_content()
        frame = read_excel_sheet(original, "Sales")[["Region", "amount"]]
        first, name = append_result_sheet(original, frame, "total_sales_by region", after_sheet="Sales")
        self.assertEqual(name, "total_sales_by region")
        second, name2 = append_result_sheet(first, frame, "TOTAL_SALES_BY REGION", after_sheet="Sales")
        self.assertEqual(name2, "TOTAL_SALES_BY REGION_1")
        book = load_workbook(BytesIO(second))
        self.assertEqual(book.sheetnames[:3], ["Sales", name2, name])
        self.assertEqual(book[name]["A1"].value, "Region")
        book.close()
        for sheet in ("Sales", "staff", "Notes", "products"):
            self.assertTrue(read_excel_sheet(original, sheet).equals(read_excel_sheet(second, sheet)), sheet)
        self.assertTrue(read_excel_sheet(first, name).equals(read_excel_sheet(second, name)))
        third, name3 = append_result_sheet(second, frame, "total_sales_by region", after_sheet="Sales")
        self.assertEqual(name3, "total_sales_by region_2")
        self.assertTrue(read_excel_sheet(second, name2).equals(read_excel_sheet(third, name2)))

    def test_unchanged_array_formulas_are_not_compared_by_object_identity(self):
        original = modern_formula_content()
        restored = preserve_formula_values(original, original, reject_changed=True)
        self.assertTrue(read_excel_sheet(original, "Sales").equals(read_excel_sheet(restored, "Sales")))
        proposed = local_excel_edit_plan(
            'Add a column "Reviewed" with "Yes" for the top 3 amounts',
            {"Sales": dict.fromkeys(read_excel_sheet(original, "Sales").columns, "text")}, "Sales",
        )
        edited = apply_workbook_edit(original, proposed)
        self.assertEqual(read_excel_sheet(edited.content, "Sales")["Reviewed"].eq("Yes").sum(), 3)
        self.assertEqual(read_excel_sheet(edited.content, "Sales")["product"].iloc[0], "Widget")
        book = load_workbook(BytesIO(edited.content))
        self.assertEqual(book["Sales"]["D2"].value.text, '="Widget"')
        self.assertEqual(book["Sales"]["D2"].value.ref, "D2")
        book.close()

    def test_static_named_ranges_allow_unrelated_column_but_reject_changed_inputs(self):
        original = modern_formula_content(named_ranges=True)
        edited = apply_workbook_edit(original, plan({
            "op": "add_column", "name": "Reviewed", "expression": {"literal": "Yes"},
        }))
        self.assertEqual(read_excel_sheet(edited.content, "Sales")["Average sales"].iloc[0], 2100)
        with self.assertRaisesRegex(ValueError, "changes inputs"):
            apply_workbook_edit(original, plan({
                "op": "add_rows", "rows": [{"Salesperson": "New", "amount": 4000}],
            }))

    def test_top_amount_marking_preserves_formula_cells_and_original_row_order(self):
        original = formula_content()
        frame = read_excel_sheet(original, "Sales")
        question = (
            'Add a column \u201cReviewed\u201d with \u201cYes\u201d for the top 3 amounts, '
            'and save as a new file in sales sheet'
        )
        proposed = local_excel_edit_plan(question, {"Sales": dict.fromkeys(frame.columns, "text")}, "staff")
        self.assertIsNotNone(proposed)
        self.assertNotIn("sort", [step["op"] for step in proposed["steps"]])
        self.assertNotIn("lookup", [step["op"] for step in proposed["steps"]])
        edited = apply_workbook_edit(original, proposed)
        actual = read_excel_sheet(edited.content, "Sales")
        self.assertTrue(actual[frame.columns].equals(frame))
        self.assertEqual(actual["Reviewed"].iloc[[0, 1, 4]].tolist(), ["Yes"] * 3)
        self.assertTrue(actual["Reviewed"].iloc[[2, 3]].isna().all())
        self.assertEqual(actual.columns.tolist(), frame.columns.tolist() + ["Reviewed"])
        self.assertTrue(read_excel_sheet(original, "Notes").equals(read_excel_sheet(edited.content, "Notes")))

    def test_native_top_marking_does_not_swallow_extra_conditions(self):
        schemas = {"Sales": {"amount": "number", "Rank": "number", "ReviewRank": "number"}}
        simple = 'Add a column "Reviewed" with "Yes" for the top 3 amounts'
        proposed = local_excel_edit_plan(simple, schemas, "Sales")
        self.assertIsNotNone(proposed)
        self.assertNotIn(proposed["steps"][0]["name"], schemas["Sales"])
        for suffix in (" and sort highest first", " using who joined first", " above 1500"):
            with self.subTest(suffix=suffix):
                self.assertIsNone(local_excel_edit_plan(simple + suffix, schemas, "Sales"))

    def test_wide_results_use_concise_summary_instead_of_repeated_cell_bullets(self):
        frame = read_excel_sheet(positional_content(), "Sales")
        for index in range(5):
            frame[f"Extra {index}"] = None
        answer = spreadsheet_answer(frame)
        self.assertIn("8 columns", answer)
        self.assertIn("data table", answer)
        self.assertNotIn("not available", answer)
        self.assertNotIn("Salesperson:", answer)

    def test_formula_cache_loss_is_reproduced_and_unchanged_values_survive_edits(self):
        original = formula_content()
        book = load_workbook(BytesIO(original))
        raw = BytesIO()
        book.save(raw)
        book.close()
        lost = load_workbook(BytesIO(raw.getvalue()), data_only=True)
        self.assertIsNone(lost["Sales"]["D2"].value)
        lost.close()
        with self.assertRaisesRegex(ValueError, "formulas without saved calculation results"):
            read_excel_sheet(raw.getvalue(), "Sales")
        restored = preserve_formula_values(original, raw.getvalue())
        self.assertEqual(read_excel_sheet(restored, "Sales").iloc[0]["product"], "Widget")
        edited = apply_workbook_edit(original, plan({
            "op": "add_column", "name": "Status", "expression": {"literal": "Yes"}, "row_numbers": [3],
        }))
        for sheet in ("Sales", "staff", "Notes"):
            actual = read_excel_sheet(edited.content, sheet)
            expected = read_excel_sheet(original, sheet)
            self.assertTrue(actual[expected.columns].equals(expected), sheet)
        book = load_workbook(BytesIO(edited.content))
        self.assertEqual(book["Sales"]["F2"].value, "=AVERAGE(B2:B100)")
        book.close()
        deleted = apply_workbook_edit(edited.content, plan({"op": "delete_columns", "columns": ["Status"]}))
        self.assertTrue(read_excel_sheet(deleted.content, "Sales").equals(read_excel_sheet(original, "Sales")))

    def test_changed_formula_inputs_and_transitive_results_are_not_reused(self):
        original = formula_content()
        book = load_workbook(BytesIO(original))
        book["Sales"]["B2"] = 1
        buffer = BytesIO()
        book.save(buffer)
        book.close()
        restored = preserve_formula_values(original, buffer.getvalue())
        cached = load_workbook(BytesIO(restored), data_only=True)
        self.assertEqual(cached["Sales"]["D2"].value, "Widget")
        self.assertIsNone(cached["Sales"]["E2"].value)
        self.assertIsNone(cached["Sales"]["F2"].value)
        self.assertIsNone(cached["Notes"]["A2"].value)
        cached.close()
        with self.assertRaisesRegex(ValueError, "changes inputs"):
            preserve_formula_values(original, buffer.getvalue(), reject_changed=True)

    def test_adding_rows_inside_formula_ranges_does_not_reuse_stale_totals(self):
        original = formula_content()
        with self.assertRaisesRegex(ValueError, "changes inputs"):
            apply_workbook_edit(original, plan({
                "op": "add_rows", "rows": [{"Salesperson": "New", "amount": 4000}],
            }))

    def test_add_column_at_explicit_rows_preserves_positions_and_other_cells(self):
        original = positional_content()
        edited = apply_workbook_edit(original, plan({
            "op": "add_column", "name": "Reviewed", "expression": {"literal": "Yes"},
            "row_numbers": [11, 3, 9],
        }))
        before = load_workbook(BytesIO(original))
        after = load_workbook(BytesIO(edited.content))
        self.assertEqual(after["Sales"].max_row, 12)
        for number in range(1, 13):
            self.assertEqual([cell.value for cell in after["Sales"][number][:3]],
                             [cell.value for cell in before["Sales"][number]])
            self.assertEqual(after["Sales"].cell(number, 4).value,
                             "Reviewed" if number == 1 else "Yes" if number in {11, 3, 9} else None)
        self.assertEqual(list(after["staff"].values), list(before["staff"].values))
        before.close()
        after.close()
        deleted = apply_workbook_edit(edited.content, plan({"op": "delete_columns", "columns": ["Reviewed"]}))
        self.assertTrue(read_excel_sheet(original, "Sales").equals(read_excel_sheet(deleted.content, "Sales")))

    def test_row_scoped_column_rejects_header_invalid_and_duplicate_rows(self):
        original = positional_content()
        for numbers in ([1], [0], [13], [True], [2, 2], [], ["3"]):
            with self.subTest(numbers=numbers), self.assertRaisesRegex(ValueError, "row numbers"):
                apply_workbook_edit(original, plan({
                    "op": "add_column", "name": "Reviewed", "expression": {"literal": "Yes"},
                    "row_numbers": numbers,
                }))

    def test_prompt_distinguishes_explicit_row_marking_from_top_n_sorting(self):
        prompt = excel_edit_prompt([], "Sales")
        self.assertIn("row_numbers", prompt)
        self.assertIn("Explicit row numbers take precedence", prompt)
        self.assertIn("without sorting", prompt)

    def test_edit_prompt_uses_real_json_keys_and_builtin_name_normalization(self):
        prompt = excel_edit_prompt([], "Sales")
        self.assertNotIn("?:", prompt)
        self.assertIn("both sides", prompt)
        self.assertIn("Staff ID is not required", prompt)
        self.assertIn('"left_on": ["Salesperson"]', prompt)
        self.assertIn('"right_on": ["Name"]', prompt)

    def test_staff_name_lookup_normalizes_without_editing_staff_or_requiring_staff_id(self):
        book = load_workbook(BytesIO(content()))
        book["Sales"]["A2"] = " AMY "
        book["Sales"]["A4"] = "cara"
        book["staff"]["A2"] = "amy"
        book["staff"]["A3"] = " BOB "
        book["staff"]["A4"] = "CARA "
        data = BytesIO()
        book.save(data)
        book.close()
        original = data.getvalue()
        edited = apply_workbook_edit(original, plan(
            {"op": "filter", "where": fn("gt", col("amount"), {"literal": 1500})},
            {"op": "lookup", "sheet": "staff", "left_on": ["Salesperson"], "right_on": ["Name"],
             "columns": [{"column": "Join date", "name": "Joined"}]},
            {"op": "sort", "by": [{"column": "amount", "direction": "desc"}, {"column": "Joined", "direction": "asc"}]},
            {"op": "rank", "column": "amount", "name": "ReviewRank", "direction": "desc", "method": "first"},
            {"op": "add_column", "name": "Reviewed", "expression": fn(
                "if", fn("lte", col("ReviewRank"), {"literal": 2}), {"literal": "Yes"}, {"literal": None},
            )},
            {"op": "delete_columns", "columns": ["Joined", "ReviewRank"]},
        ))
        self.assertEqual(edited.frame["Salesperson"].tolist(), ["Bob", "cara", " AMY "])
        self.assertEqual(edited.frame["Reviewed"].iloc[:2].tolist(), ["Yes", "Yes"])
        self.assertTrue(edited.frame["Reviewed"].isna().iloc[-1])
        self.assertTrue(read_excel_sheet(original, "staff").equals(read_excel_sheet(edited.content, "staff")))
        self.assertEqual(read_excel_sheet(edited.content, "staff").columns.tolist(), ["Name", "Join date"])

    def test_lookup_schema_errors_identify_invalid_optional_key(self):
        with self.assertRaisesRegex(ValueError, r"keep\?"):
            apply_workbook_edit(content(), plan(
                {"op": "lookup", "sheet": "staff", "left_on": ["Salesperson"], "right_on": ["Name"],
                 "columns": [{"column": "Join date", "name": "Joined"}], "keep?": "first"},
                {"op": "add_column", "name": "Reviewed", "expression": {"literal": "Yes"}},
            ))

    def test_add_delete_columns_and_keep_other_sheets_styles_and_original(self):
        original = content()
        edited = apply_workbook_edit(original, plan(
            {"op": "add_column", "name": "Reviewed", "expression": {"literal": "Yes"}},
            {"op": "delete_columns", "columns": ["Code"]},
        ))
        self.assertEqual(edited.frame.columns.tolist(), ["Salesperson", "amount", "Reviewed"])
        self.assertEqual(edited.frame["Reviewed"].tolist(), ["Yes"] * 4)
        self.assertEqual(read_excel_sheet(original, "Sales").columns.tolist(), ["Salesperson", "amount", "Code"])
        book = load_workbook(BytesIO(edited.content))
        self.assertEqual(book.sheetnames, ["Sales", "staff"])
        self.assertEqual(book["staff"]["A2"].value, "Amy")
        self.assertTrue(book["Sales"]["B2"].font.bold)
        self.assertEqual(book["Sales"]["B2"].number_format, "#,##0.00")
        self.assertEqual(book["Sales"].column_dimensions["B"].width, 22)
        self.assertEqual(book["Sales"].freeze_panes, "A2")
        book.close()

    def test_insert_rows_and_delete_by_condition_or_excel_row_number(self):
        edited = apply_workbook_edit(content(), plan(
            {"op": "add_rows", "before": 3, "rows": [{"Salesperson": "Eve", "amount": 5000, "Code": "=literal"}]},
            {"op": "delete_rows", "row_numbers": [2]},
            {"op": "delete_rows", "where": fn("lt", col("amount"), {"literal": 2000})},
        ))
        self.assertEqual(edited.frame["Salesperson"].tolist(), ["Eve", "Bob", "Cara"])
        book = load_workbook(BytesIO(edited.content), data_only=False)
        self.assertEqual(book["Sales"]["C2"].value, "=literal")
        self.assertEqual(book["Sales"]["C2"].data_type, "s")
        book.close()
        second = apply_workbook_edit(edited.content, plan({"op": "delete_columns", "columns": ["Salesperson"]}))
        book = load_workbook(BytesIO(second.content), data_only=False)
        self.assertEqual(book["Sales"]["B2"].value, "=literal")
        self.assertEqual(book["Sales"]["B2"].data_type, "s")
        book.close()

    def test_filter_sort_join_date_mark_top_three_and_save_full_workbook(self):
        book = load_workbook(BytesIO(content()))
        book["Sales"].append(["Eve", 2500, "E"])
        book["staff"].append(["Eve", "2022-01-01"])
        data = BytesIO()
        book.save(data)
        book.close()
        edited = apply_workbook_edit(data.getvalue(), plan(
            {"op": "filter", "where": fn("gt", col("amount"), {"literal": 1500})},
            {"op": "lookup", "sheet": "staff", "left_on": ["Salesperson"], "right_on": ["Name"],
             "columns": [{"column": "Join date", "name": "Joined"}]},
            {"op": "sort", "by": [{"column": "amount", "direction": "desc"}, {"column": "Joined", "direction": "asc"}]},
            {"op": "rank", "column": "amount", "name": "ReviewRank", "direction": "desc", "method": "first"},
            {"op": "add_column", "name": "Reviewed", "expression": fn(
                "if", fn("lte", col("ReviewRank"), {"literal": 3}), {"literal": "Yes"}, {"literal": None},
            )},
            {"op": "delete_columns", "columns": ["Joined", "ReviewRank"]},
        ))
        self.assertEqual(edited.frame["Salesperson"].tolist(), ["Bob", "Eve", "Cara", "Amy"])
        self.assertEqual(edited.frame["Reviewed"].iloc[:3].tolist(), ["Yes", "Yes", "Yes"])
        self.assertTrue(edited.frame["Reviewed"].isna().iloc[-1])
        self.assertEqual(read_excel_sheet(edited.content, "Sales")["Salesperson"].tolist(), ["Bob", "Eve", "Cara", "Amy"])
        self.assertEqual(read_excel_sheet(edited.content, "staff").shape, (4, 2))

    def test_failed_edit_is_atomic_and_header_and_existing_columns_are_protected(self):
        original = content()
        for steps in (
            [{"op": "delete_rows", "row_numbers": [1]}],
            [{"op": "delete_columns", "columns": ["Salesperson", "amount", "Code"]}],
            [{"op": "add_column", "name": "AMOUNT", "expression": {"literal": 1}}],
            [{"op": "add_rows", "rows": [{"not a header": 1}]}],
            [{"op": "add_column", "name": "New", "expression": {"literal": 1}},
             {"op": "delete_columns", "columns": ["Missing"]}],
        ):
            with self.subTest(steps=steps), self.assertRaises(ValueError):
                apply_workbook_edit(original, plan(*steps))
        self.assertEqual(read_excel_sheet(original, "Sales").shape, (4, 3))

    def test_deleting_all_rows_retains_headers_and_can_accept_new_rows(self):
        empty = apply_workbook_edit(content(), plan(
            {"op": "delete_rows", "where": fn("gte", col("amount"), {"literal": 0})},
        ))
        self.assertEqual(read_excel_sheet(empty.content, "Sales").shape, (0, 3))
        added = apply_workbook_edit(empty.content, plan(
            {"op": "add_rows", "rows": [{"Salesperson": "New", "amount": 1}]},
        ))
        self.assertEqual(read_excel_sheet(added.content, "Sales")["Salesperson"].tolist(), ["New"])

    def test_delete_then_readd_column_uses_new_values_not_original_cells(self):
        edited = apply_workbook_edit(content(), plan(
            {"op": "delete_columns", "columns": ["amount"]},
            {"op": "add_column", "name": "amount", "expression": {"literal": 99}},
        ))
        self.assertEqual(read_excel_sheet(edited.content, "Sales")["amount"].tolist(), [99] * 4)

    def test_insert_column_and_physical_blank_row_positions(self):
        book = load_workbook(BytesIO(content()))
        book["Sales"].insert_rows(3)
        output = BytesIO()
        book.save(output)
        book.close()
        edited = apply_workbook_edit(output.getvalue(), plan(
            {"op": "delete_rows", "row_numbers": [3]},
            {"op": "add_column", "name": "Reviewed", "expression": {"literal": None}, "before": "amount"},
        ))
        self.assertEqual(edited.frame["Salesperson"].tolist(), ["Amy", "Bob", "Cara", "Dan"])
        self.assertEqual(edited.frame.columns.tolist(), ["Salesperson", "Reviewed", "amount", "Code"])
        self.assertTrue(edited.frame["Reviewed"].isna().all())

    def test_zero_match_and_omitted_edits_fail_explicitly(self):
        with self.assertRaisesRegex(ValueError, "No rows matched"):
            apply_workbook_edit(content(), plan(
                {"op": "delete_rows", "where": fn("gt", col("amount"), {"literal": 999999})},
            ))
        with self.assertRaisesRegex(ValueError, "omitted"):
            apply_workbook_edit(content(), plan({"op": "limit", "count": 2}))

    def test_missing_join_date_in_tied_rows_is_not_guessed(self):
        book = load_workbook(BytesIO(content()))
        book["staff"]["B2"] = None
        output = BytesIO()
        book.save(output)
        book.close()
        with self.assertRaisesRegex(ValueError, "missing tie-breaker"):
            apply_workbook_edit(output.getvalue(), plan(
                {"op": "lookup", "sheet": "staff", "left_on": ["Salesperson"], "right_on": ["Name"],
                 "columns": [{"column": "Join date", "name": "Joined"}]},
                {"op": "sort", "by": [{"column": "amount", "direction": "desc"}, {"column": "Joined", "direction": "asc"}]},
                {"op": "add_column", "name": "Reviewed", "expression": {"literal": "Yes"}},
            ))

    def test_lookup_and_rank_scratch_columns_are_not_saved(self):
        edited = apply_workbook_edit(content(), plan(
            {"op": "lookup", "sheet": "staff", "left_on": ["Salesperson"], "right_on": ["Name"],
             "columns": [{"column": "Join date", "name": "Joined"}]},
            {"op": "rank", "column": "amount", "name": "ScratchRank", "direction": "desc", "method": "first"},
            {"op": "add_column", "name": "Reviewed", "expression": {"literal": "Yes"}},
        ))
        self.assertEqual(edited.frame.columns.tolist(), ["Salesperson", "amount", "Code", "Reviewed"])
        self.assertEqual(read_excel_sheet(edited.content, "Sales").columns.tolist(), edited.frame.columns.tolist())

    def test_unrelated_sheet_formulas_do_not_prevent_edits(self):
        book = load_workbook(BytesIO(content()))
        book["staff"]["C1"] = "Constant"
        book["staff"]["C2"] = "=1+1"
        output = BytesIO()
        book.save(output)
        book.close()
        edited = apply_workbook_edit(output.getvalue(), plan({"op": "delete_columns", "columns": ["Code"]}))
        book = load_workbook(BytesIO(edited.content), data_only=False)
        self.assertEqual(book["staff"]["C2"].value, "=1+1")
        book.close()

    def test_formula_dependencies_are_not_silently_corrupted(self):
        workbook = load_workbook(BytesIO(content()))
        workbook["staff"]["D1"] = "=Sales!B2"
        output = BytesIO()
        workbook.save(output)
        workbook.close()
        original = output.getvalue()
        with self.assertRaisesRegex(ValueError, "invalidate workbook formulas"):
            apply_workbook_edit(original, plan({"op": "delete_columns", "columns": ["Salesperson"]}))
        edited = apply_workbook_edit(original, plan(
            {"op": "add_column", "name": "Reviewed", "expression": {"literal": "Yes"}},
        ))
        workbook = load_workbook(BytesIO(edited.content), data_only=False)
        self.assertEqual(workbook["staff"]["D1"].value, "=Sales!B2")
        workbook.close()

    def test_edit_intent_and_simple_local_plans(self):
        for question in (
            "Add a Reviewed column", "Delete column Code", "Delete row 3",
            "From Sales show amount above 1500, add a Reviewed column with Yes for the top 3 and save a new file",
        ):
            with self.subTest(question=question):
                self.assertTrue(is_excel_edit_request(question))
        for question in (
            "Show rows with amount above 1500", "How do I delete a column?", "How to add a row",
            "Don't delete any rows, only show them",
        ):
            self.assertFalse(is_excel_edit_request(question))
        schemas = {"Sales": {"Salesperson": "text", "amount": "number", "Code": "text"}}
        self.assertEqual(local_excel_edit_plan("Delete column Code from Sales", schemas, "Sales")["steps"],
                         [{"op": "delete_columns", "columns": ["Code"]}])
        self.assertIsNone(local_excel_edit_plan("Add column Reviewed with Yes", schemas, "Sales"))


class WorkbookEditChatTests(unittest.TestCase):
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
        self.enterContext(patch.object(self.chatbot, "get_engine", side_effect=AssertionError("SQL write/read attempted")))
        self.client = Mock()
        self.enterContext(patch.object(self.chatbot, "get_azure_openai_client", return_value=self.client))
        self.enterContext(patch("openai.AzureOpenAI", return_value=self.client))
        self.enterContext(patch("sqlalchemy.create_engine", side_effect=AssertionError("SQL access attempted")))
        self.chatbot.st.cache_resource.clear()
        self.addCleanup(self.chatbot.st.cache_resource.clear)
        self.upload = BytesIO(content())
        self.upload.name = "sales.xlsx"
        self.enterContext(patch_chat_upload(return_value=self.upload))
        self.app = AppTest.from_file(str(Path(__file__).resolve().parents[1] / "employee_sql_chatbot.py"))

    def test_edits_survive_reruns_mode_changes_following_queries_and_keep_original(self):
        self.app.run(timeout=30)
        for question in ("Add column Reviewed to Sales", "Delete row 2 from Sales", "Delete column Reviewed from Sales"):
            self.app.chat_input[0].set_value(question).run(timeout=30)
            self.assertEqual(len(self.app.exception), 0)
            self.assertNotIn("error", self.app.session_state["chat_history"][-1])
        cache = self.app.session_state["excel_workbook"]
        edited = cache["content"]
        self.assertNotEqual(edited, self.upload.getvalue())
        self.assertEqual(cache["original_content"], self.upload.getvalue())
        self.assertEqual(read_excel_sheet(edited, "Sales")["Salesperson"].tolist(), ["Bob", "Cara", "Dan"])
        self.app.run(timeout=30)
        self.app.radio(key="data_source").set_value("Employee database").run(timeout=30)
        self.app.radio(key="data_source").set_value("Excel workbook").run(timeout=30)
        self.app.chat_input[0].set_value("How many rows in Sales").run(timeout=30)
        self.assertEqual(len(self.app.exception), 0)
        self.assertEqual(self.app.session_state["excel_workbook"]["content"], edited)
        self.assertEqual(self.app.session_state["chat_history"][-1]["result"].iloc[0, 0], 3)
        self.client.chat.completions.create.assert_not_called()
        self.assertTrue(any(button.label == "Download updated workbook" for button in self.app.get("download_button")))

    def test_invalid_edit_does_not_commit_partial_data_or_delete_header(self):
        self.app.run(timeout=30)
        original = self.app.session_state["excel_workbook"]["content"]
        self.app.chat_input[0].set_value("Delete row 1 from Sales").run(timeout=30)
        self.assertEqual(len(self.app.exception), 0)
        self.assertIn("header", self.app.session_state["chat_history"][-1]["error"])
        self.assertEqual(self.app.session_state["excel_workbook"]["content"], original)

    def test_regional_answer_is_saved_to_unique_adjacent_sheet_in_same_working_copy(self):
        original = regional_content()
        self.upload = BytesIO(original)
        self.upload.name = "regions.xlsx"
        question = "give sum of sales for all regions and save in other sheet in same workbook"
        with patch_chat_upload(return_value=self.upload):
            self.app.run(timeout=30)
            self.app.selectbox[0].select("products").run(timeout=30)
            snapshots = []
            for title in ("total_sales_by region", "total_sales_by region_1"):
                if title.endswith("_1"):
                    self.app.selectbox[0].select("Sales").run(timeout=30)
                self.app.chat_input[0].set_value(question).run(timeout=30)
                self.assertEqual(len(self.app.exception), 0)
                message = self.app.session_state["chat_history"][-1]
                self.assertNotIn("error", message)
                self.assertEqual(message["kind"], "excel_sheet_result")
                self.assertEqual(message["edited_sheet"], title)
                self.assertEqual(message["source_sheets"], ["Sales"])
                cache = self.app.session_state["excel_workbook"]
                saved = read_excel_sheet(cache["content"], title)
                self.assertEqual(saved.columns.tolist(), ["Region", "Total Sales"])
                self.assertEqual(saved.set_index("Region")["Total Sales"].to_dict(),
                                 {"North": 4000, "South": 5500, "East": 1000})
                book = load_workbook(BytesIO(cache["content"]))
                self.assertEqual(book.sheetnames[1], title)
                count = len(book.sheetnames)
                book.close()
                for sheet in ("Sales", "staff", "Notes", "products"):
                    self.assertTrue(read_excel_sheet(original, sheet).equals(read_excel_sheet(cache["content"], sheet)), sheet)
                self.assertEqual(cache["original_content"], original)
                snapshots.append(message["edited_workbook"])
                self.app.run(timeout=30)
                book = load_workbook(BytesIO(self.app.session_state["excel_workbook"]["content"]))
                self.assertEqual(len(book.sheetnames), count)
                book.close()
            self.assertNotEqual(snapshots[0], snapshots[1])
            self.assertEqual(self.app.session_state["chat_history"][0]["edited_workbook"], snapshots[0])
            self.client.chat.completions.create.assert_not_called()
            self.assertTrue(any(button.label == "Download updated workbook" for button in self.app.get("download_button")))

    def test_selected_sheet_preview_saves_to_requested_name_without_overwriting(self):
        original = regional_content()
        self.upload = BytesIO(original)
        self.upload.name = "regions.xlsx"
        with patch_chat_upload(return_value=self.upload):
            self.app.run(timeout=30)
            self.app.selectbox[0].select("products").run(timeout=30)
            self.app.chat_input[0].set_value('preview and save in another sheet named "Selected data"').run(timeout=30)
            self.assertEqual(len(self.app.exception), 0)
            message = self.app.session_state["chat_history"][-1]
            self.assertNotIn("error", message)
            self.assertEqual(message["edited_sheet"], "Selected data")
            self.assertEqual(message["source_sheets"], ["products"])
            saved = message["edited_workbook"]
            self.assertTrue(read_excel_sheet(saved, "Selected data").equals(read_excel_sheet(original, "products")))
            book = load_workbook(BytesIO(saved))
            self.assertEqual(book.sheetnames[-2:], ["products", "Selected data"])
            book.close()
            self.client.chat.completions.create.assert_not_called()

    def test_saving_result_without_upload_fails_without_database_access(self):
        with patch_chat_upload(return_value=None):
            self.app.run(timeout=30)
            self.app.chat_input[0].set_value(
                "give sum of sales for all regions and save in other sheet in same workbook",
            ).run(timeout=30)
            self.assertEqual(len(self.app.exception), 0)
            message = self.app.session_state["chat_history"][-1]
            self.assertIn("Upload an Excel workbook", message["error"])
            self.assertNotIn("edited_workbook", message)
            self.assertNotIn("excel_workbook", self.app.session_state)
            self.client.chat.completions.create.assert_not_called()

    def test_top_amount_request_updates_formula_working_copy_without_ai_or_staff_lookup(self):
        original = modern_formula_content(named_ranges=True)
        self.upload = BytesIO(original)
        self.upload.name = "formulas.xlsx"
        with patch_chat_upload(return_value=self.upload):
            self.app.run(timeout=30)
            self.app.chat_input[0].set_value(
                'Add a column \u201cReviewed\u201d with \u201cYes\u201d for the top 3 amounts, '
                'and save as a new file in sales sheet'
            ).run(timeout=30)
            self.assertEqual(len(self.app.exception), 0)
            message = self.app.session_state["chat_history"][-1]
            self.assertNotIn("error", message)
            cache = self.app.session_state["excel_workbook"]
            actual = read_excel_sheet(cache["content"], "Sales")
            expected = read_excel_sheet(original, "Sales")
            self.assertTrue(actual[expected.columns].equals(expected))
            self.assertEqual(actual["Reviewed"].eq("Yes").sum(), 3)
            self.assertEqual(cache["original_content"], original)
            self.assertTrue(any(button.label == "Download updated workbook" for button in self.app.get("download_button")))
            self.client.chat.completions.create.assert_not_called()

    def test_ai_top_marking_repair_removes_unrequested_sort_for_formula_sheet(self):
        original = formula_content()
        self.upload = BytesIO(original)
        self.upload.name = "formulas.xlsx"
        rank = {"op": "rank", "column": "amount", "name": "ReviewRank", "direction": "desc", "method": "first"}
        addition = {"op": "add_column", "name": "Reviewed", "expression": fn(
            "if", fn("lte", col("ReviewRank"), {"literal": 3}), {"literal": "Yes"}, {"literal": None},
        )}
        wrong = plan({"op": "sort", "by": [{"column": "amount", "direction": "desc"}]}, rank, addition)
        correct = plan(rank, addition)
        self.client.chat.completions.create.side_effect = [
            SimpleNamespace(choices=[SimpleNamespace(finish_reason="stop", message=SimpleNamespace(
                content=json.dumps(value),
            ))]) for value in (wrong, correct)
        ]
        with patch_chat_upload(return_value=self.upload), patch.dict(os.environ, {
            "AZURE_OPENAI_DEPLOYMENT_MINI": "test", "AZURE_OPENAI_ENDPOINT": "https://unit-test.invalid",
            "AZURE_OPENAI_API_KEY": "test-only", "AZURE_OPENAI_API_VERSION": "2024-02-01",
        }):
            self.app.run(timeout=30)
            self.app.chat_input[0].set_value(
                'Please add a column "Reviewed" with "Yes" for the top 3 amounts in Sales sheet',
            ).run(timeout=30)
        self.assertEqual(len(self.app.exception), 0)
        self.assertNotIn("error", self.app.session_state["chat_history"][-1])
        self.assertEqual(self.client.chat.completions.create.call_count, 2)
        repair = self.client.chat.completions.create.call_args.kwargs["messages"][-1]["content"]
        self.assertIn("Top N marking does not require sorting", repair)
        expected = read_excel_sheet(original, "Sales")
        actual = read_excel_sheet(self.app.session_state["excel_workbook"]["content"], "Sales")
        self.assertTrue(actual[expected.columns].equals(expected))
        self.assertEqual(actual["Reviewed"].eq("Yes").sum(), 3)

    def test_previous_working_copy_recovers_formula_results_from_original_upload(self):
        original = modern_formula_content(named_ranges=True)
        self.upload = BytesIO(original)
        self.upload.name = "formulas.xlsx"
        book = load_workbook(BytesIO(original))
        book["Sales"]["G1"] = "Status"
        book["Sales"]["G3"] = "Yes"
        raw = BytesIO()
        book.save(raw)
        book.close()
        broken = raw.getvalue()
        with patch_chat_upload(return_value=self.upload):
            self.app.session_state["excel_workbook"] = {
                "content": broken, "file_name": self.upload.name, "digest": sha256(broken).hexdigest(),
                "upload_digest": sha256(original).hexdigest(), "edited": True, "original_content": original,
                "formula_cache_preserved": True,
            }
            self.app.run(timeout=30)
            self.assertEqual(len(self.app.exception), 0)
            cache = self.app.session_state["excel_workbook"]
            self.assertTrue(cache["formula_cache_preserved"])
            self.assertEqual(cache["formula_cache_version"], 2)
            self.assertEqual(cache["frame"]["product"].iloc[0], "Widget")
            self.assertEqual(cache["frame"]["Rating"].iloc[0], "High")
            self.assertEqual(cache["frame"]["Average sales"].iloc[0], 2100)
            self.assertEqual(cache["frame"]["Status"].iloc[1], "Yes")
            self.assertEqual(cache["original_content"], original)
            self.assertNotIn("Create region summary", [expander.label for expander in self.app.expander])
            self.app.radio(key="data_source").set_value("Excel workbook").run(timeout=30)
            self.app.chat_input[0].set_value("preview").run(timeout=30)
            self.assertNotIn("error", self.app.session_state["chat_history"][-1])
            self.assertEqual(self.app.session_state["chat_history"][-1]["result"]["product"].iloc[0], "Widget")

    def test_explicit_row_column_edit_and_delete_persist_only_in_working_copy(self):
        self.upload = BytesIO(positional_content())
        self.upload.name = "sales.xlsx"
        valid = plan({
            "op": "add_column", "name": "Reviewed", "expression": {"literal": "Yes"},
            "row_numbers": [11, 3, 9],
        })
        self.client.chat.completions.create.return_value = SimpleNamespace(choices=[
            SimpleNamespace(finish_reason="stop", message=SimpleNamespace(content=json.dumps(valid))),
        ])
        with patch_chat_upload(return_value=self.upload), patch.dict(os.environ, {
            "AZURE_OPENAI_DEPLOYMENT_MINI": "test", "AZURE_OPENAI_ENDPOINT": "https://unit-test.invalid",
            "AZURE_OPENAI_API_KEY": "test-only", "AZURE_OPENAI_API_VERSION": "2024-02-01",
        }):
            self.app.run(timeout=30)
            self.app.chat_input[0].set_value(
                'In the Sales sheet, add a new column "Reviewed" and put "Yes" in rows 11, 3 and 9 '
                '(the top 3 amounts above 1500). Leave the other rows blank. '
                'Do not change the header row or row 12. Save as a new file.'
            ).run(timeout=30)
            self.assertEqual(len(self.app.exception), 0)
            message = self.app.session_state["chat_history"][-1]
            self.assertNotIn("error", message)
            cache = self.app.session_state["excel_workbook"]
            self.assertEqual(cache["original_content"], self.upload.getvalue())
            book = load_workbook(BytesIO(cache["content"]))
            self.assertEqual([number for number in range(2, 13) if book["Sales"].cell(number, 4).value == "Yes"],
                             [3, 9, 11])
            book.close()
            snapshot = cache["content"]
            self.app.run(timeout=30)
            self.assertEqual(self.app.session_state["excel_workbook"]["content"], snapshot)
            self.app.chat_input[0].set_value("Delete column Reviewed from Sales").run(timeout=30)
            self.assertNotIn("error", self.app.session_state["chat_history"][-1])
            self.assertTrue(read_excel_sheet(self.upload.getvalue(), "Sales").equals(
                read_excel_sheet(self.app.session_state["excel_workbook"]["content"], "Sales"),
            ))
            self.assertEqual(message["edited_workbook"], snapshot)
            self.assertEqual(self.client.chat.completions.create.call_count, 1)

    def test_malformed_edit_plan_is_repaired_once_before_any_commit(self):
        valid = plan({"op": "add_column", "name": "Reviewed", "expression": {"literal": "Yes"}})
        self.client.chat.completions.create.side_effect = [
            SimpleNamespace(choices=[SimpleNamespace(finish_reason="stop", message=SimpleNamespace(
                content=json.dumps({**valid, "save_as": "new.xlsx"}),
            ))]),
            SimpleNamespace(choices=[SimpleNamespace(finish_reason="stop", message=SimpleNamespace(content=json.dumps(valid)))]),
        ]
        with patch.dict(os.environ, {
            "AZURE_OPENAI_DEPLOYMENT_MINI": "test", "AZURE_OPENAI_ENDPOINT": "https://unit-test.invalid",
            "AZURE_OPENAI_API_KEY": "test-only", "AZURE_OPENAI_API_VERSION": "2024-02-01",
        }):
            self.app.run(timeout=30)
            self.app.chat_input[0].set_value("Add a Reviewed column with Yes to Sales").run(timeout=30)
        self.assertEqual(len(self.app.exception), 0)
        self.assertEqual(self.client.chat.completions.create.call_count, 2)
        message = self.app.session_state["chat_history"][-1]
        self.assertNotIn("error", message)
        self.assertEqual(read_excel_sheet(self.app.session_state["excel_workbook"]["content"], "Sales")["Reviewed"].tolist(),
                         ["Yes"] * 4)
        self.assertEqual(message["original_workbook"], self.upload.getvalue())

    def test_false_missing_staff_id_refusal_gets_one_schema_grounded_recheck(self):
        refused = {
            "understood": False, "unhandled_conditions": [
                "Cannot resolve ties: Sales has no Staff ID and staff cannot be edited to normalize Name.",
            ], "sheet": "Sales", "steps": [],
        }
        valid = plan(
            {"op": "lookup", "sheet": "staff", "left_on": ["Salesperson"], "right_on": ["Name"],
             "columns": [{"column": "Join date", "name": "Joined"}]},
            {"op": "add_column", "name": "Reviewed", "expression": {"literal": "Yes"}},
        )
        for final, expected_error in ((valid, False), (refused, True)):
            with self.subTest(expected_error=expected_error), patch.dict(os.environ, {
                "AZURE_OPENAI_DEPLOYMENT_MINI": "test", "AZURE_OPENAI_ENDPOINT": "https://unit-test.invalid",
                "AZURE_OPENAI_API_KEY": "test-only", "AZURE_OPENAI_API_VERSION": "2024-02-01",
            }):
                self.client.reset_mock()
                self.client.chat.completions.create.side_effect = [
                    SimpleNamespace(choices=[SimpleNamespace(finish_reason="stop", message=SimpleNamespace(
                        content=json.dumps(value),
                    ))]) for value in (refused, final)
                ]
                self.app = AppTest.from_file(str(Path(__file__).resolve().parents[1] / "employee_sql_chatbot.py"))
                self.app.run(timeout=30)
                self.app.chat_input[0].set_value(
                    "Add a Reviewed column with Yes to Sales using staff joining dates",
                ).run(timeout=30)
                self.assertEqual(len(self.app.exception), 0)
                self.assertEqual(self.client.chat.completions.create.call_count, 2)
                message = self.app.session_state["chat_history"][-1]
                self.assertEqual("error" in message, expected_error)
                if expected_error:
                    self.assertEqual(self.app.session_state["excel_workbook"]["content"], self.upload.getvalue())
                    self.assertNotIn("['", message["error"])
                else:
                    self.assertEqual(message["result"].columns.tolist(), ["Salesperson", "amount", "Code", "Reviewed"])
                    self.assertTrue(read_excel_sheet(self.upload.getvalue(), "staff").equals(
                        read_excel_sheet(message["edited_workbook"], "staff"),
                    ))

    def test_database_mode_does_not_edit_excel_or_sql_implicitly(self):
        self.app.run(timeout=30)
        original = self.app.session_state["excel_workbook"]["content"]
        self.app.radio(key="data_source").set_value("Employee database").run(timeout=30)
        self.app.chat_input[0].set_value("Delete row 2").run(timeout=30)
        self.assertEqual(len(self.app.exception), 0)
        self.assertIn("read-only", self.app.session_state["chat_history"][-1]["error"])
        self.assertEqual(self.app.session_state["excel_workbook"]["content"], original)
        self.client.chat.completions.create.assert_not_called()


if __name__ == "__main__":
    unittest.main()
