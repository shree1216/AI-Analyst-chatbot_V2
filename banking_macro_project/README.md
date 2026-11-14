# Banking Transaction Report – Excel Macro Troubleshooting

A simple project that formats a daily banking transaction report with a VBA macro. It shows a
recorded macro failing and the corrected, reusable version that fixes it. All data is synthetic (USD).

## Files

| File | Purpose |
| --- | --- |
| `Banking_Transactions_Report.xlsm` | Main workbook: **Transactions** sheet (100 raw rows) with a **Format Report** button, plus a **Troubleshooting Log** sheet |
| `banking_transactions_100.csv` | The 100-row dataset |
| `RecordedMacro.bas` | **Before** – the macro as recorded (contains 5 faults) |
| `BankingReport.bas` | **After** – the corrected, reusable macro |
| `generate_transactions.py` | Regenerates the dataset (`python generate_transactions.py`) |
| `build_workbook.ps1` | Rebuilds the workbook from the CSV and `.bas` files |
| `Test-Macro.ps1` | Runs both macros on a temporary copy and writes `validation_results.txt` |

## Dataset (100 rows)

`Transaction ID | Date | Account Number | Customer Name | Branch | Transaction Type | Debit | Credit`

- 12 customers across 6 branches, 10 business days (02–13 Nov 2026)
- Transaction types: Cash Deposit, Salary Credit, Transfer In, Interest Credit, Cash Withdrawal,
  Bill Payment, Loan EMI, Transfer Out, Card Purchase
- Each row has either a Debit or a Credit amount
- Totals: Debit **69,162.19**, Credit **78,981.53**, Net **9,819.34**

## How to use

1. Open `Banking_Transactions_Report.xlsm` and click **Enable Content**.
2. On **Transactions**, click **Format Report**, or press **Alt+F8** and run `FormatTransactionReport`.
3. The macro then:
   - styles the header row
   - formats dates as `dd-mmm-yyyy` and amounts as `#,##0.00`, with negative values in red
   - adds **Net Amount = Credit − Debit**
   - adds a **TOTAL** row
   - adds borders, auto-fits column widths, and freezes the header row

The macro works on whichever sheet is active, so you can paste a new day's data onto any sheet and run it.

## Troubleshooting summary

| # | Recorded code | Problem | Fix |
| --- | --- | --- | --- |
| 1 | `Sheets("Sheet1").Select` | Error 9 *Subscript out of range* after the sheet was renamed | Uses the active sheet |
| 2 | `Range("B2:B101")`, AutoFill to `I101` | Rows after 101 are ignored | Last row found dynamically (`End(xlUp)`) |
| 3 | `NumberFormat = "0.00%"` | Amounts shown as percentages (325166.00%) | `#,##0.00` |
| 4 | `=RC[-2]+RC[-1]` | Net adds Debit + Credit; breaks if columns move | Credit − Debit, columns located by header name |
| 5 | Totals in fixed row 102 | Overwrites new data; rerun duplicates totals | Totals placed after last row; old totals removed on rerun |

The corrected macro also checks the data before changing anything. For example, if Debit or Credit
contains text, it reports the row (such as *Row 10: Debit 'abc' is not a number.*) and leaves the
sheet unchanged.

## Validation

`Test-Macro.ps1` ran in desktop Excel, and all 21 checks pass (see `validation_results.txt`). The checks cover:

- the recorded macro's faults
- correct totals and number formats
- safe reruns
- added and deleted rows
- a changed layout: title rows above the data, a table starting in column B, and Credit before Debit
- rejection of invalid data

`build_workbook.ps1` and `Test-Macro.ps1` temporarily enable *Trust access to the VBA project object
model* to import code. They restore the original setting when Excel closes.

## Record your own macro (optional)

Use **Developer › Record Macro** on a copy of the Transactions sheet and apply some formatting. Then compare
the generated code with `RecordedMacro.bas` and `BankingReport.bas`.
