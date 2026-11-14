# Employee Management System – VBA UserForm

An Excel application where HR staff add, search, update and delete employee records through a
custom form instead of editing worksheet cells. All employee data is synthetic.

## Files

| File | Purpose |
| --- | --- |
| `Employee_Management_System.xlsm` | The application: **Dashboard** and **Employees** sheets, form and code |
| `frmEmployee.frm` / `frmEmployee.frx` | The UserForm (layout and code), ready to import into another workbook |
| `modEmployees.bas` | Data module: dropdown lists, search, validation, save and delete |
| `frmEmployee_code.txt` | Source of the form's code (used by the build) |
| `employees_100.csv` | The 100-employee dataset |
| `generate_employees.py` | Regenerates the dataset (`python generate_employees.py`) |
| `build_workbook.ps1`, `BuildForm.bas` | Rebuild the workbook. `BuildForm.bas` lays out the form during the build and is not kept in the workbook. |
| `Test-EmployeeApp.ps1` | Checks the application in Excel and writes `validation_results.txt` |

## Employee database (100 rows)

The data is stored in the Excel table `tblEmployees` on the **Employees** sheet:

`Employee ID | Full Name | Gender | Department | Job Title | Email | Phone | Hire Date | Annual Salary | Status`

- 6 departments: Finance, Human Resources, IT, Marketing, Operations, Sales
- Status: Active, On Leave or Resigned

## How to use

1. Open `Employee_Management_System.xlsm` and click **Enable Content**.
2. On **Dashboard**, click **Open Employee Manager**.

| Action | How |
| --- | --- |
| **Add** | Fill in the fields and click **Add**. The next Employee ID (e.g. EMP101) is assigned automatically. |
| **Search** | Type an Employee ID (`EMP005`), a full name or part of a name, then press **Enter** or click **Search**. |
| **Update** | Search for the employee, change the fields and click **Update**. You will be asked to confirm. |
| **Delete** | Search for the employee and click **Delete**. You will be asked to confirm; the default answer is *No*. |
| **Clear / Close** | **Clear** starts a new entry. **Close** (or Esc) closes the form. |

The form switches between *New employee* mode, where only Add is enabled, and *Editing* mode, where
Update and Delete are enabled. This prevents accidental duplicates.

The Dashboard tiles and the department table use live formulas, so they update as soon as a record changes.

## Form controls

- **Text boxes:** Search, Employee ID (read-only), Full Name, Job Title, Email, Phone, Hire Date, Annual Salary
- **Dropdown lists:** Gender, Department and Status. You can only pick values from the list.
- **Command buttons:** Search, Add, Update, Delete, Clear, Close
- Each field shows a tooltip with the expected format.

## Validation (before every save)

| Field | Rule |
| --- | --- |
| Full Name | Required; letters, spaces and `. ' -` only |
| Gender, Department, Status | Must be selected from the list |
| Job Title | Required |
| Email | Valid format (`name@company.com`) and not used by another employee |
| Phone | Format `555-123-4567` |
| Hire Date | A valid date, not in the future |
| Annual Salary | A number between 10,000 and 1,000,000 |

All problems are listed together in one error message, and nothing is saved until they are fixed.
Saved values are tidied up: extra spaces are removed, names are put in Proper Case, and email addresses
in lower case.

## Pop-up messages

- **Errors:** *Invalid Input*, which lists every problem; *No employee found*; a reminder to enter a search term
- **Confirmations:** *Confirm Update* and *Confirm Delete* (Yes/No)
- **Success:** *Employee Added*, *Employee Updated*, *Employee Deleted*

## Validation results

`Test-EmployeeApp.ps1` ran in desktop Excel, and all 14 checks pass (see `validation_results.txt`). The checks cover:

- the Dashboard figures and the button
- the form's controls and dropdown lists
- search by ID, full name and partial name
- valid and invalid input (all 9 rules)
- the duplicate-email rule
- Add through the form, Update and Delete, including the Dashboard updating after each change

The pop-up messages themselves can't be clicked in an automated test. The checks cover all the logic
behind them.
