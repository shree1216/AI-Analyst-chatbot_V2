# HR Employee Management (Excel VBA UserForm)

Synthetic HR data only. Open `HR_Employee_Manager.xlsm`, allow macros for this
file if your organization permits, and click **Open Employee Manager** on the
Dashboard.

## Workbook

| Sheet | Contents |
| --- | --- |
| Dashboard | Button that opens the form, headcount/status KPIs, average active salary, headcount by department |
| Employees | `tblEmployees` with 25 synthetic employees; protected (no password) so records are edited through the form |
| Lists | `tblDepartments` and `tblStatuses`, which feed the form's dropdowns |

## Form features

- **Text boxes:** First Name, Last Name, Email, Phone, Job Title, Hire Date, Annual Salary.
- **Dropdowns:** Department and Status (list values only).
- **Add New:** validates input, then assigns the next `EMP-####` ID.
- **Search:** finds by ID, name, email, department, job title, or status; **Show All** lists everyone.
- **Update / Delete:** select an employee first; both ask for confirmation.
- **Pop-ups:** success confirmations, Yes/No confirmations, and specific validation and error messages.

Validation requires every field. Names may contain letters, spaces, hyphens, apostrophes, and periods.
Email must be valid and unique. Phone must contain 7–15 digits. Department and status must come from the lists.
Hire Date must use `yyyy-mm-dd`, be a real date, and fall between 1950 and today.
Salary must be greater than 0, no more than 10,000,000, and have at most two decimal places.
Text that starts with `=`, `+`, or `@` is rejected or stored as plain text, never as a formula.

## Source files

- [EmployeeManager.bas](EmployeeManager.bas): data access, validation, add/search/update/delete, dashboard refresh.
- [EmployeeForm.frm](EmployeeForm.frm) + `EmployeeForm.frx`: UserForm. Its controls are created in code when
  the form opens.
- [EmployeeAppTests.bas](EmployeeAppTests.bas): in-workbook tests. Run `RunEmployeeAppTestsWithMessage`
  from **Alt+F8**. These **84 checks** exercise the actual form: add, search, update, and delete;
  confirmation and cancellation paths; and 20 invalid inputs. They also verify protection and dashboard
  totals. Test records are deleted afterward, and existing rows are verified unchanged.

To add a department or status, add a row to the corresponding table on **Lists**.
