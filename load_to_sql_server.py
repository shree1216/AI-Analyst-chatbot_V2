import pandas as pd
from sqlalchemy import create_engine, text
from sqlalchemy.engine import URL

SERVER = r"localhost\SQLEXPRESS"
DATABASE = "EmployeeDB"
EMPLOYEE_TABLE_NAME = "employee"
DEPARTMENTS_TABLE_NAME = "departments"

# connect to master first to create the target database if it doesn't exist
master_url = URL.create(
    "mssql+pyodbc",
    host=SERVER,
    database="master",
    query={"driver": "ODBC Driver 17 for SQL Server", "trusted_connection": "yes"},
)
master_engine = create_engine(master_url, isolation_level="AUTOCOMMIT")
with master_engine.connect() as conn:
    conn.exec_driver_sql(
        f"IF NOT EXISTS (SELECT * FROM sys.databases WHERE name = '{DATABASE}') "
        f"CREATE DATABASE [{DATABASE}]"
    )

db_url = URL.create(
    "mssql+pyodbc",
    host=SERVER,
    database=DATABASE,
    query={"driver": "ODBC Driver 17 for SQL Server", "trusted_connection": "yes"},
)
engine = create_engine(db_url)

with engine.connect() as conn:
    employee_rows = conn.execute(
        text(f"SELECT COUNT(*) FROM dbo.[{EMPLOYEE_TABLE_NAME}]")
    ).scalar_one()
    department_report = pd.read_sql(
        text(
            f"""
            SELECT
                employee.[employee_id] AS EmployeeID,
                employee.[name] AS EmployeeName,
                employee.[dept_name] AS EmployeeDepartment,
                department.[dept_name] AS ReferenceDepartment,
                CASE
                    WHEN employee.[dept_name] IS NULL
                      OR department.[dept_name] IS NULL
                    THEN 'Data Cleaning Required'
                    ELSE 'No Data Cleaning Required'
                END AS DataCleaningStatus
            FROM dbo.[{EMPLOYEE_TABLE_NAME}] AS employee
            FULL OUTER JOIN dbo.[{DEPARTMENTS_TABLE_NAME}] AS department
                ON employee.[dept_name] = department.[dept_name]
            ORDER BY
                COALESCE(employee.[dept_name], department.[dept_name]),
                employee.[employee_id]
            """
        ),
        conn,
    )

print(f"Read {employee_rows} rows from [{DATABASE}].[dbo].[{EMPLOYEE_TABLE_NAME}]")
print("\nFull Outer Join Department Cleaning Report")
print(department_report.to_string(index=False))
