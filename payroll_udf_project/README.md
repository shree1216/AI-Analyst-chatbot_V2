# Employee Payroll – Custom VBA Functions (UDFs)

A payroll workbook that calculates employee bonuses and deductions with two reusable VBA
User-Defined Functions. All employee data is synthetic (monthly salaries in USD).

## Files

| File | Purpose |
| --- | --- |
| `Employee_Payroll_UDF.xlsm` | Main workbook: **Payroll**, **Company Rules** and **Validation** sheets |
| `PayrollFunctions.bas` | The VBA module with the two UDFs |
| `employee_payroll_50.csv` | The 50-employee dataset |
| `generate_payroll.py` | Regenerates the dataset (`python generate_payroll.py`) |
| `build_workbook.ps1` | Rebuilds the workbook from the CSV and `.bas` file |
| `Test-Payroll.ps1` | Checks every result in Excel and writes `validation_results.txt` |

## Dataset (50 employees)

`Employee ID | Employee Name | Department | Job Title | Monthly Salary | Performance Rating`

- 5 departments: Finance, Human Resources, IT, Sales, Operations
- Monthly salaries from 3,200 to 10,100
- Ratings from 1 (Unsatisfactory) to 5 (Outstanding)

## Custom functions

```text
=CalculateBonus(Salary, Rating)
=CalculateDeductions(Salary)
```

**Bonus rule** – percentage of monthly salary by performance rating:

| Rating | Meaning | Bonus |
| --- | --- | --- |
| 5 | Outstanding | 20% |
| 4 | Exceeds Expectations | 15% |
| 3 | Meets Expectations | 10% |
| 2 | Needs Improvement | 5% |
| 1 | Unsatisfactory | 0% |

**Deduction rules** – total deductions are the sum of:

| Deduction | Rule |
| --- | --- |
| Provident Fund | 12% of monthly salary |
| Health Insurance | Flat 150 per month |
| Income Tax | 10% of the salary above 5,000 |

**Net Pay** = Monthly Salary + Bonus − Deductions (a normal Excel formula).

If the salary is zero, negative or text, or the rating is not 1–5, the functions return `#VALUE!`.
This makes bad data easy to spot. The rules are stored as constants at the top of
`PayrollFunctions.bas`, so you only need to change them in one place.

## How to use

1. Open `Employee_Payroll_UDF.xlsm` and click **Enable Content**.
2. **Payroll** sheet: columns G–I hold the formulas, and the TOTAL row is at the bottom. When you change
   a salary or rating, the results update automatically. A rating cell accepts only whole numbers from
   1 to 5.
3. To use the functions on new records, type `=CalculateBonus(E53,F53)` and `=CalculateDeductions(E53)`,
   or fill the existing formulas down.
4. To reuse the functions in another workbook, open the Visual Basic Editor with **Alt+F11**, then choose
   **File › Import File** and select `PayrollFunctions.bas`.

## Validation

The **Validation** sheet compares the UDFs with 9 hand-calculated examples. Each example includes its
manual working. The examples cover:

- the tax limit boundary (exactly 5,000)
- each rating
- a bonus with cents
- an invalid rating
- an invalid salary

All 9 show **PASS**.

`Test-Payroll.ps1` also recalculates all 50 employees outside of VBA and compares every Bonus,
Deduction and Net Pay value and the totals. It also checks that the functions recalculate when an input
changes, and that a text salary returns `#VALUE!`. All 15 checks pass (see `validation_results.txt`).

| Total | Amount |
| --- | --- |
| Bonus | 38,152.50 |
| Deductions | 49,774.00 |
| Net Pay | 285,578.50 |
