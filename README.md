# Workforce Compass

Workforce Compass is a Streamlit-based AI analyst chatbot for exploring employee
data stored in Microsoft SQL Server. It converts natural-language questions into
read-only workforce queries and can produce summaries, charts, business reports,
and downloadable PDF reports.

## Features

- Ask natural-language questions about employees, departments, salaries, cities,
  employment status, and joining dates.
- Generate workforce breakdowns and interactive charts.
- Create business reports with findings and recommendations.
- Download reports as PDF files.
- Look up employee details such as department, city, email, and phone number.
- Optional Azure OpenAI support for generated SQL and response summaries.
- Optional Azure Speech support for voice input and spoken responses.
- Optional SMTP support for emailing PDF reports.
- Analyze uploaded JPG, PNG, and WebP images when Azure OpenAI is configured.

## Technology

- Python
- Streamlit
- pandas and Plotly
- SQLAlchemy and pyodbc
- Microsoft SQL Server Express
- Azure OpenAI and Azure AI Speech
- ReportLab

## Project Structure

| Path | Purpose |
| --- | --- |
| `employee_sql_chatbot.py` | Main Streamlit application |
| `employee_dataset.xlsx` | Synthetic 150-employee sample dataset |
| `generate_employee_dataset.py` | Recreates the sample Excel dataset |
| `import_employees_data.py` | Creates `EmployeeDB` and imports `dbo.Employees` |
| `.env.example` | Template for optional Azure and SMTP settings |
| `requirements.txt` | Python dependencies |

## Prerequisites

Install the following before starting:

1. Python 3.11 or newer
2. Microsoft SQL Server Express with the `SQLEXPRESS` instance
3. Microsoft ODBC Driver 17 for SQL Server
4. Git

The application currently connects to:

```text
Server: localhost\SQLEXPRESS
Database: EmployeeDB
Table: dbo.Employees
```

Update the database settings near the top of `employee_sql_chatbot.py` and
`import_employees_data.py` if your SQL Server instance uses a different name.

## Local Setup on Windows

### 1. Clone the repository

```powershell
git clone https://github.com/shree1216/AI-Analyst-chatbot_V2.git
Set-Location AI-Analyst-chatbot_V2
```

### 2. Create and activate a virtual environment

```powershell
py -m venv .venv
.\.venv\Scripts\Activate.ps1
```

If PowerShell blocks activation, run this once in the current terminal:

```powershell
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
.\.venv\Scripts\Activate.ps1
```

### 3. Install dependencies

```powershell
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
```

### 4. Create the SQL Server database and import the sample data

Make sure the `SQLEXPRESS` service is running, then execute:

```powershell
python import_employees_data.py
```

The script creates `EmployeeDB` when needed and replaces `dbo.Employees` with
the 150 rows from `employee_dataset.xlsx`.

Expected output:

```text
Imported 150 rows into [EmployeeDB].[dbo].[Employees]
```

### 5. Configure optional integrations

Copy the environment template:

```powershell
Copy-Item .env.example .env
```

Add only the credentials for the integrations you want to use:

| Variable | Purpose |
| --- | --- |
| `AZURE_OPENAI_ENDPOINT` | Azure OpenAI resource endpoint |
| `AZURE_OPENAI_API_KEY` | Azure OpenAI API key |
| `AZURE_OPENAI_API_VERSION` | Azure OpenAI API version |
| `AZURE_OPENAI_DEPLOYMENT_MINI` | Chat model deployment name |
| `AZURE_SPEECH_KEY` | Azure AI Speech key |
| `AZURE_SPEECH_REGION` | Azure AI Speech region |
| `SMTP_HOST` | SMTP server hostname |
| `SMTP_PORT` | SMTP server port |
| `SMTP_USERNAME` | SMTP login username |
| `SMTP_APP_PASSWORD` | SMTP app password |
| `SMTP_FROM_EMAIL` | Sender email address |

Never commit `.env`, API keys, passwords, or
`.streamlit/secrets.toml`. These files are excluded by `.gitignore`.

### 6. Run Workforce Compass

```powershell
python -m streamlit run employee_sql_chatbot.py
```

Open [http://localhost:8501](http://localhost:8501) in a browser.

## Example Questions

```text
How many employees are in each department?
Show the average salary by department.
Create a workforce report showing the breakdown by employment status.
Who are the top 3 highest-paid employees in each department?
What is the email address for Shree?
Create a PDF report comparing department salaries.
```

## Recreating the Sample Dataset

The committed dataset is synthetic. To regenerate it:

```powershell
python generate_employee_dataset.py
python import_employees_data.py
```

Do not commit real employee names, email addresses, phone numbers, salaries, or
other confidential information.

## Deployment Note

The default `localhost\SQLEXPRESS` connection works only on the computer running
SQL Server. It will not connect from Streamlit Community Cloud or another hosted
environment.

Before deploying publicly, move the employee data to a secured hosted database,
configure database credentials through the hosting platform's secret manager,
and restrict access to sensitive workforce information.

## Troubleshooting

### ODBC driver error

Install Microsoft ODBC Driver 17 for SQL Server and verify that its name matches
the driver configured in the Python files.

### Cannot connect to SQL Server

- Confirm that the `SQL Server (SQLEXPRESS)` service is running.
- Confirm that the instance name is `localhost\SQLEXPRESS`.
- Run `python import_employees_data.py` before starting the chatbot.

### Streamlit command not found

Activate the virtual environment and use:

```powershell
python -m streamlit run employee_sql_chatbot.py
```

## Repository

<https://github.com/shree1216/AI-Analyst-chatbot_V2>
