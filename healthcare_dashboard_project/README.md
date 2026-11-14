# Healthcare Reporting Dashboard – VBA

An automated monthly reporting workbook for a healthcare organization. One **Refresh Report** button:

- cleans the raw visit records
- calculates appointments, revenue and department performance
- rebuilds the dashboard and its charts
- creates a report sheet for every month

All patient data is synthetic.

## Files

| File | Purpose |
| --- | --- |
| `Healthcare_Dashboard.xlsm` | Main workbook (Dashboard, monthly reports, Clean Data, Raw Data) |
| `HealthcareReport.bas` | The VBA module |
| `healthcare_visits_raw_100.csv` | The 100-row raw dataset |
| `generate_visits.py` | Regenerates the dataset (`python generate_visits.py`) |
| `build_workbook.ps1` | Rebuilds the workbook from the CSV and `.bas` file |
| `Test-Dashboard.ps1` | Checks the results in Excel and writes `validation_results.txt` |

## Dataset (100 raw visits, Aug–Oct 2026)

`Visit ID | Patient ID | Patient Name | Department | Doctor | Appointment Date | Visit Type | Status | Billing Amount`

- 5 departments: Cardiology, Orthopedics, Pediatrics, Dermatology, General Medicine (2 doctors each)
- Status: Completed, Cancelled or No-Show. Only completed visits are billed.
- The raw extract has realistic problems:
  - 4 duplicate visits
  - 3 incomplete records: a missing department, a missing date, and a completed visit with no bill
  - inconsistent text, such as `ORTHOPEDICS`, `general medicine`, `no-show` and names with extra spaces

## How to use

1. Open `Healthcare_Dashboard.xlsm` and click **Enable Content**.
2. On **Dashboard**, click **Refresh Report**, or press **Alt+F8** and run `RefreshReport`.
3. For the next month, add or paste the new visits into **Raw Data** and click **Refresh Report**. A new
   monthly report sheet is created automatically.

## What the macro produces

| Sheet | Contents |
| --- | --- |
| **Dashboard** | KPI tiles (Total Appointments, Completion Rate, Total Revenue, Average Bill), a Department Performance table ranked by revenue, a Monthly Trend table, a Data Quality table, and 3 charts: Revenue by Department, Monthly Appointments and Monthly Revenue |
| **Report Aug 2026** (one per month) | The same KPI tiles, a department table and a chart for that month only |
| **Clean Data** | Standardized, de-duplicated, complete records |
| **Raw Data** | The original extract (never changed by the macro) |

**Cleaning rules:**

- Text is trimmed and given consistent capitalization.
- Status is standardized to Completed, Cancelled or No-Show.
- Duplicate Visit IDs are removed.
- A record is dropped as incomplete if it is missing a Visit ID, department or valid date, or if it is a completed visit without a bill.

**Calculations:**

- **Revenue** is the total billed amount.
- **Average Bill** = Revenue ÷ completed visits.
- **Completion Rate** = completed visits ÷ appointments.

## Results for the sample data

| Metric | Value |
| --- | --- |
| Raw / clean records | 100 / 93 (4 duplicates, 3 incomplete) |
| Total appointments | 93 |
| Completion rate | 81.7% |
| Total revenue | 19,755.00 |
| Average bill | 259.93 |
| Top department | Orthopedics, 6,655.00 (33.7% of revenue) |
| Monthly appointments | Aug 26 → Sep 29 → Oct 38 |

## Validation

`Test-Dashboard.ps1` ran in desktop Excel, and all 12 checks pass (see `validation_results.txt`). The checks compare the
workbook with figures calculated directly from the CSV:

- the clean, duplicate and incomplete counts
- the Clean Data contents
- every KPI value
- the department table and its sorting
- the monthly trend table and each monthly report
- the charts and the button

They also confirm that refreshing is repeatable and that adding a November visit creates a **Report Nov 2026** sheet.
