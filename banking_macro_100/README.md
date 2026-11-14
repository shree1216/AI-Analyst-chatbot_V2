# Automated Excel Macro Troubleshooting – Banking (100 transactions)

All data is synthetic, in USD. Open `Banking_Transactions_100.xlsm`, allow macros for this file if your
organization permits, then choose **Format This Report** on **Raw Transactions**. You can also select a
transaction sheet, press **Alt+F8**, and run **FormatBankingReport**.

## Sheets

| Sheet | Purpose |
| --- | --- |
| Start Here | Instructions and expected totals |
| Raw Transactions | 100 unformatted transactions (input) with a **Format This Report** button |
| Transactions | The same data after the corrected macro ran |
| Changed Layout | Renamed sheet, reordered columns, extra title rows, table starting in column B, 110 rows; formatted by the same macro |
| Invalid Data | Row 50 repeats TXN-0001; the macro refuses to format it |
| Troubleshooting Log | Each fault in the recorded macro, the evidence seen when it ran, and the fix |
| Validation | 17 automated checks run in desktop Excel VBA, all PASS |

Expected totals: 100 rows have Debit **66,727.91**, Credit **125,550.56**, and Net **58,822.65**.
110 rows net to **53,400.94**.

## Macros

- [RecordedMacro_Broken.bas](RecordedMacro_Broken.bas): the "before" macro, in Excel's macro-recorder
  style. It hard-codes `Sheets("Daily Report")`, fixed ranges `A5:G5` and `E6:F105`, and column H.
  It formats amounts as `0%`, computes Debit **+** Credit, and uses `Select`/`Selection`. The
  Troubleshooting Log records what happened when it ran:
  - error 9 after a sheet rename
  - a wrong net of 192,278.47
  - an overwritten Channel column
  - a data cell overwritten in the changed layout
- [BankingReport.bas](BankingReport.bas): the corrected, reusable macro. It finds headers by name in
  rows 1–20, in any order and starting in any column, and keeps extra columns such as Channel. It
  validates every row before changing anything. It then builds an Excel table with date and amount
  formats, **Net Amount = Credit − Debit**, a **Running Net** column, a totals row, and print setup.
  Rerunning it reuses the same columns. Rows added or deleted inside the table are picked up.
- [MacroHarness.bas](MacroHarness.bas): test helper that returns errors as text, so validation never
  stops on a dialog.

The recorded macro is a faithful reconstruction of recorder output, not a live recording. To record your
own, use **Developer > Record Macro** on a copy of Raw Transactions, then compare its code with
`BankingReport`.
