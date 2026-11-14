# Automated banking transaction report

This is a separate, synthetic USD demonstration. Existing Excel workbooks and
Excel security settings are not modified. This demo is not connected to a bank.

## Files and status

- `Banking_Transactions_Demo.xlsm`: final macro-enabled workbook, with the corrected
  VBA already imported and both valid reports formatted and calculated.
- `Banking_Transactions_Demo.xlsx`: sample data, changed-layout sample, deliberately
  invalid sample, instructions, and Excel calculation checks. An `.xlsx` cannot
  contain VBA.
- [BankingReport.bas](BankingReport.bas): corrected, importable VBA module.
- [RecordedMacro_Broken.bas](RecordedMacro_Broken.bas): intentionally faulty
  **recorded-style illustration**, not a macro actually recorded in your session.
- `sample_transactions.csv`: 20 fictional transactions.
- [Test-BankingMacro.ps1](Test-BankingMacro.ps1): actual Excel/VBA integration tests
  to run after importing the module and saving an `.xlsm`.
- `macro-validation-results.json`: **19 passing actual desktop Excel/VBA tests**.

Desktop Excel is available, but programmatic VBA insertion is blocked. The
module was imported through Excel's native VBA editor instead, without changing
security settings. Actual VBA execution passed all 19 integration checks, and the
corrected macro was run on both final report worksheets before saving the `.xlsm`.
The original `.xlsx` source checks alone do not establish VBA execution.

## Run the delivered workbook

Open `Banking_Transactions_Demo.xlsm`, select **Transactions** or **Changed
Layout**, press **Alt+F8**, and run **BankingReport.FormatBankingReport** (Excel may
display just `FormatBankingReport`). The module is already imported. Allow macros
only for this reviewed workbook if your organization's policy permits.

## Rebuild from the raw workbook (no VBA-project trust setting required)

1. Open `Banking_Transactions_Demo.xlsx` in desktop Excel.
2. Press **Alt+F11**. In the Project Explorer select this demo workbook, not
   another open workbook.
3. Choose **File > Import File**, select `BankingReport.bas`, and verify that a
   `BankingReport` module appears under this demo's project.
4. Choose **Debug > Compile VBAProject**. Save as
   `Banking_Transactions_Demo.xlsm` using **Excel Macro-Enabled Workbook**.
5. If prompted, allow macros only for this reviewed demo according to your
   organization's policy. Do not globally enable all macros.
6. Return to Excel, select **Transactions**, press **Alt+F8**, and run
   **FormatBankingReport**.
7. Repeat on **Changed Layout**. **Invalid Data** should show an explicit
   duplicate-ID error without creating a formatted table.
8. Save the demo and keep it open before running the integration test:

   ```powershell
   powershell.exe -NoProfile -ExecutionPolicy RemoteSigned -File .\Test-BankingMacro.ps1 -WorkbookPath .\Banking_Transactions_Demo.xlsm -UseOpenWorkbook
   ```

   The script tests disposable worksheet copies inside the open demo and removes
   them afterward. It does not change the source worksheets. This uses the
   workbook's existing permission to run its reviewed macros, without changing
   any security setting. Without `-UseOpenWorkbook`, it tests a temporary file
   copy, which Excel may block even when the original file is allowed.
   Macro execution remains subject to Excel security policy. No security policy
   or `AccessVBOM` setting is changed.
   `RemoteSigned` applies only to this PowerShell process, not a persistent
   execution-policy setting.
   For intentional invalid-input tests, Excel can display native VBA runtime
   error dialogs. Choose **End**, not **Debug**, to let the runner check the
   rejection and continue. The 19 recorded checks were completed with those
   expected dialogs explicitly dismissed. Normal use through
   `FormatBankingReport` displays a friendly error message instead.

## Record a macro yourself

Native Excel recording was attempted on a scratch worksheet, but Excel displayed
**"Cannot run Visual Basic macro because of a syntax error."** A completed genuine
recording was not validated; the scratch worksheet was removed. The corrected
`BankingReport` macro continued to execute successfully after this warning.
To satisfy an assignment requiring a genuine recording, make a scratch copy of
the **Transactions** sheet, enable the **Developer** tab, choose **Record Macro**,
name it `RecordedFormatting`, store it in **This Workbook**, and record bold
headers, numeric Debit/Credit formatting, and an initial net formula. Choose
**Stop Recording**, then inspect the module with **Alt+F11**. Keep that recording
as the before version; use `BankingReport` as the corrected version.

## Troubleshooting the illustrative broken macro

| Fault | Consequence | Correction |
| --- | --- | --- |
| `Worksheets("Sheet1").Select` | Error 9 after a worksheet rename | Pass the selected worksheet explicitly; no hardcoded sheet name |
| Fixed `A1:G11` and row 1 | Title rows, new columns, and added records are missed | Discover the header row in rows 1-20 and find source columns by name |
| Whole selection made bold | Every transaction is bold | Apply a table style and bold only its header |
| Debit/Credit formatted as `0%` | Amounts appear as percentages | Two-decimal monetary display plus explicit USD Currency |
| Debit plus Credit | Wrong net movement | Credit minus Debit |
| `Select` and `Selection` | Depends on focus and active workbook | Fully qualified worksheet/range references |
| No checks or error handling | Bad data can look like success | Validate first, explain the row error, restore application state |

## Reusable behavior and limits

Required, case-insensitive, unique headers are **Transaction ID**, **Transaction
Date**, **Account ID**, **Description**, **Debit**, **Credit**, and **Currency**.
All must be on the same row within rows 1-20. Their order and worksheet name can
change; the table can start in a later column. Keep the transaction area
contiguous, without merged cells, duplicate/blank headers, unrelated tables, or
content immediately below it.

Dates must be real Excel dates, not locale-dependent text. Amounts must be
nonnegative numeric cells (a blank opposite-side amount is zero), at most two
decimal places, with exactly one positive Debit or Credit per transaction.
IDs must be nonblank and unique, and account and description must be populated.
This demo accepts **USD only** and refuses other/mixed currencies. It does not
perform currency conversion.

The macro creates/reuses an Excel table with filtering, banded rows, date and
amount formats, Credit-minus-Debit **Net Amount**, cumulative **Running Net**,
visible-row totals, and landscape print settings. **Running Net is a report
movement total in worksheet order, starting at zero, not an account balance.**
It does not assume an opening balance or sort transactions. Filtered totals use
Excel's `SUBTOTAL`; Running Net still includes all rows.

On subsequent runs, calculated columns are reused, not duplicated. Add or delete
transactions **inside the Excel table, above its totals row**, then rerun. Data
pasted below totals is refused explicitly. Calculated fields are owned by the
macro and overwritten on each run. Other input values and columns are preserved.
No protected worksheet is unprotected automatically. Validation errors are
reported before formatting; an unexpected Excel error during formatting is
reported but is not transactionally rolled back.

## Expected figures

| Dataset | Transactions | Debit (USD) | Credit (USD) | Net / final Running Net (USD) |
| --- | ---: | ---: | ---: | ---: |
| Transactions | 20 | 2,837.75 | 9,300.00 | 6,462.25 |
| Changed Layout | 23 | 3,087.75 | 10,140.00 | 7,052.25 |

The changed layout starts at B8, reorders columns, and adds three transactions.
The integration tests also cover reruns, row insertion/deletion, invalid IDs,
missing headers, invalid amounts/dates/currency, protected sheets, and restoring
Excel's event/screen-update state.
