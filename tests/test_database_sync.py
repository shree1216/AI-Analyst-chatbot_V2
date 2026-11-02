from datetime import date, datetime
from decimal import Decimal
import json
from pathlib import Path
import tempfile
import unittest
from unittest.mock import Mock, patch

import pandas as pd

from demo_tables import REQUIRED_EMPLOYEE_COLUMNS, DemoTable, demo_data_version, load_demo_tables, write_demo_tables
import sync_local_database as sync


def local_table(name="sql_products", rows=(("P1", "Pen", Decimal("1.50")),), schema="dbo"):
    columns = (
        sync.Column("code", "varchar", 10, 0, 0, False),
        sync.Column("product", "nvarchar", 100, 0, 0, True),
        sync.Column("price", "decimal", 9, 10, 2, True),
    )
    return sync.LocalTable(schema, name, columns, ("code",), tuple(rows))


def employees_table():
    columns = tuple(
        sync.Column(name, "bigint" if name in {"Age", "Salary"} else "datetime" if name == "Joining Date" else "varchar",
                    -1, 0, 0, True)
        for name in REQUIRED_EMPLOYEE_COLUMNS
    )
    row = ("E1", "Ann", 30, "Female", "IT", "Pune", "Active", "Analyst", 50000, datetime(2024, 1, 2),
           "ann@example.test", "555")
    return sync.LocalTable("dbo", "Employees", columns, (), (row,))


class DemoTablesTests(unittest.TestCase):
    def setUp(self):
        self.directory = Path(self.enterContext(tempfile.TemporaryDirectory()))

    def test_round_trip_preserves_types_nulls_and_text(self):
        table = DemoTable("Mixed", (
            ("id", "integer"), ("price", "decimal"), ("active", "boolean"), ("joined", "datetime"),
            ("born", "date"), ("phone", "text"),
        ), (
            (1, Decimal("10.50"), True, datetime(2024, 1, 2), date(1990, 5, 6), "007"),
            (None, None, False, datetime(2024, 1, 2, 9, 30), None, ""),
        ))
        self.assertTrue(write_demo_tables([table], self.directory))
        frame = load_demo_tables(self.directory)["Mixed"]
        self.assertEqual(str(frame["id"].dtype), "Int64")
        self.assertEqual(frame["id"].tolist()[0], 1)
        self.assertTrue(pd.isna(frame["id"].iloc[1]))
        self.assertEqual(frame["price"].tolist()[0], 10.5)
        self.assertEqual(frame["active"].tolist(), [True, False])
        self.assertEqual(frame["joined"].tolist(), ["2024-01-02", "2024-01-02 09:30:00"])
        self.assertEqual(frame["born"].tolist(), ["1990-05-06", None])
        self.assertEqual(frame["phone"].tolist()[0], "007")

    def test_unchanged_data_is_not_rewritten_and_removed_tables_are_deleted(self):
        first = DemoTable("A", (("x", "integer"),), ((1,),))
        second = DemoTable("B", (("y", "text"),), (("v",),))
        self.assertTrue(write_demo_tables([first, second], self.directory))
        version = demo_data_version(self.directory)
        self.assertFalse(write_demo_tables([second, first], self.directory))
        self.assertEqual(demo_data_version(self.directory), version)
        self.assertTrue(write_demo_tables([second], self.directory))
        self.assertFalse((self.directory / "A.csv").exists())
        self.assertNotEqual(demo_data_version(self.directory), version)
        self.assertEqual(list(load_demo_tables(self.directory)), ["B"])

    def test_csv_that_does_not_match_schema_is_rejected(self):
        write_demo_tables([DemoTable("A", (("x", "integer"),), ((1,),))], self.directory)
        (self.directory / "A.csv").write_text("other\n1\n", encoding="utf-8")
        with self.assertRaisesRegex(ValueError, "does not match its schema"):
            load_demo_tables(self.directory)

    def test_bundled_demo_data_has_required_employee_columns(self):
        tables = load_demo_tables()
        self.assertTrue(set(REQUIRED_EMPLOYEE_COLUMNS) <= set(tables["Employees"].columns))
        self.assertIn("sql_products", tables)


class SyncHelperTests(unittest.TestCase):
    def test_column_ddl_preserves_sql_server_types(self):
        cases = {
            sync.Column("a", "varchar", -1, 0, 0, True): "[a] varchar(max) NULL",
            sync.Column("b", "nvarchar", 100, 0, 0, False): "[b] nvarchar(50) NOT NULL",
            sync.Column("c", "decimal", 9, 10, 2, True): "[c] decimal(10,2) NULL",
            sync.Column("d", "datetime2", 8, 27, 7, True): "[d] datetime2(7) NULL",
            sync.Column("e]x", "xml", -1, 0, 0, True): "[e]]x] nvarchar(max) NULL",
            sync.Column("f", "bigint", 8, 19, 0, True): "[f] bigint NULL",
        }
        for column, expected in cases.items():
            self.assertEqual(column.ddl(), expected)
        with self.assertRaisesRegex(sync.SyncError, "unsupported type"):
            sync.Column("g", "geography", -1, 0, 0, True).ddl()

    def test_fingerprint_ignores_row_order_and_detects_changes(self):
        rows = (("P1", "Pen", Decimal("1.50")), ("P2", "Ink", Decimal("2.00")))
        self.assertEqual(local_table(rows=rows).fingerprint(), local_table(rows=rows[::-1]).fingerprint())
        self.assertNotEqual(local_table(rows=rows).fingerprint(),
                            local_table(rows=rows + (("P3", "Cap", None),)).fingerprint())

    def test_demo_export_requires_valid_employees_and_names_non_dbo_tables(self):
        with self.assertRaisesRegex(sync.SyncError, "no dbo.Employees"):
            sync.demo_tables([local_table()])
        exported = sync.demo_tables([employees_table(), local_table(schema="sales")])
        self.assertEqual([table.name for table in exported], ["Employees", "sales_sql_products"])
        self.assertEqual(dict(exported[1].columns), {"code": "text", "product": "text", "price": "decimal"})
        broken = employees_table()
        broken = sync.LocalTable("dbo", "Employees", broken.columns[1:], (), tuple(row[1:] for row in broken.rows))
        with self.assertRaisesRegex(sync.SyncError, "Employee ID"):
            sync.demo_tables([broken])


class CloudSyncTests(unittest.TestCase):
    def setUp(self):
        self.state_dir = Path(self.enterContext(tempfile.TemporaryDirectory()))
        self.enterContext(patch.object(sync, "STATE_DIR", self.state_dir))
        self.enterContext(patch.dict("os.environ", {
            "SYNC_CLOUD_SQL_SERVER": "unit.database.windows.net", "SYNC_CLOUD_READER_USER": "reader",
        }))
        self.cursor = Mock()
        self.cursor.execute.return_value = self.cursor
        self.connection = Mock()
        self.connection.cursor.return_value = self.cursor
        self.connection.__enter__ = Mock(return_value=self.connection)
        self.connection.__exit__ = Mock(return_value=False)
        self.connect = self.enterContext(patch.object(sync, "connect_cloud", return_value=self.connection))

    def statements(self):
        return [call.args[0] for call in self.cursor.execute.call_args_list]

    def test_unchanged_tables_make_no_cloud_connection(self):
        table = local_table()
        sync.sync_cloud([table], {"cloud": {table.qualified: table.fingerprint()}}, force=False)
        self.connect.assert_not_called()

    def test_changed_table_is_recreated_granted_and_removed_tables_dropped(self):
        table = local_table(rows=(("P1", "Pen", Decimal("1.50")), ("P2", "Ink", None)))
        self.cursor.fetchone.return_value = (1,)
        self.cursor.fetchval.return_value = 2
        state = {"cloud": {table.qualified: "old", "dbo.retired": "old"}}
        sync.sync_cloud([table], state, force=False)
        statements = self.statements()
        self.assertIn("DROP TABLE IF EXISTS [dbo].[retired]", statements)
        self.assertIn("DROP TABLE IF EXISTS [dbo].[sql_products]", statements)
        self.assertIn(
            "CREATE TABLE [dbo].[sql_products] ([code] varchar(10) NOT NULL, [product] nvarchar(50) NULL, "
            "[price] decimal(10,2) NULL, PRIMARY KEY ([code]))", statements,
        )
        self.assertIn("GRANT SELECT ON OBJECT::[dbo].[sql_products] TO [reader]", statements)
        self.assertFalse(any("employees_messy" in statement for statement in statements))
        self.cursor.executemany.assert_called_once()
        self.connection.commit.assert_called_once()
        self.assertEqual(state["cloud"], {table.qualified: table.fingerprint()})
        saved = json.loads((self.state_dir / "state.json").read_text(encoding="utf-8"))
        self.assertEqual(saved["cloud"], state["cloud"])

    def test_row_count_mismatch_rolls_back_and_keeps_previous_state(self):
        table = local_table()
        self.cursor.fetchone.return_value = None
        self.cursor.fetchval.return_value = 0
        state = {"cloud": {table.qualified: "old"}}
        with self.assertRaisesRegex(sync.SyncError, "expected 1 rows"):
            sync.sync_cloud([table], state, force=False)
        self.connection.rollback.assert_called_once()
        self.connection.commit.assert_not_called()
        self.assertEqual(state["cloud"], {table.qualified: "old"})

    def test_unchanged_demo_data_is_not_pushed(self):
        tables = sync.demo_tables([employees_table()])
        state = {"demo": ""}
        with patch.object(sync, "sync_clone") as clone, patch.object(sync, "git") as git, patch.object(
            sync, "write_demo_tables", return_value=True,
        ):
            clone.return_value = self.state_dir
            git.return_value = Mock(returncode=0)
            sync.push_demo_data(tables, state, force=False)
            pushed = [call.args[1:] for call in git.call_args_list if call.args[1] == "push"]
            self.assertEqual(len(pushed), 1)
            git.reset_mock()
            sync.push_demo_data(tables, state, force=False)
            git.assert_not_called()


if __name__ == "__main__":
    unittest.main()
