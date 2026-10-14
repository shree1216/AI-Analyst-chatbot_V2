from pathlib import Path

import pandas as pd
from sqlalchemy import create_engine, text
from sqlalchemy.engine import URL

SERVER = r"localhost\SQLEXPRESS"
DATABASE = "EmployeeDB"
TABLE_NAME = "employees_data"
EXCEL_FILE = Path(__file__).with_name("employees_data.xlsx")

master_url = URL.create(
    "mssql+pyodbc",
    host=SERVER,
    database="master",
    query={"driver": "ODBC Driver 17 for SQL Server", "trusted_connection": "yes"},
)
master_engine = create_engine(master_url, isolation_level="AUTOCOMMIT")
with master_engine.connect() as connection:
    connection.exec_driver_sql(
        f"IF NOT EXISTS (SELECT * FROM sys.databases WHERE name = '{DATABASE}') "
        f"CREATE DATABASE [{DATABASE}]"
    )

database_url = URL.create(
    "mssql+pyodbc",
    host=SERVER,
    database=DATABASE,
    query={"driver": "ODBC Driver 17 for SQL Server", "trusted_connection": "yes"},
)
engine = create_engine(database_url)

employees_data = pd.read_excel(EXCEL_FILE)
employees_data.to_sql(TABLE_NAME, engine, schema="dbo", if_exists="replace", index=False)

with engine.connect() as connection:
    imported_rows = connection.execute(
        text(f"SELECT COUNT(*) FROM dbo.[{TABLE_NAME}]")
    ).scalar_one()

print(f"Imported {imported_rows} rows into [{DATABASE}].[dbo].[{TABLE_NAME}]")