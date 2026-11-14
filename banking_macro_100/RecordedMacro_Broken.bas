Attribute VB_Name = "RecordedMacro_Broken"
Option Explicit

' BEFORE: macro as captured by Excel's macro recorder on the original
' "Daily Report" sheet (header in row 5, exactly 100 rows, 7 columns).
' Kept unchanged so the faults can be demonstrated. Do not use on real reports.
Public Sub RecordedFormatting_Broken()
    Sheets("Daily Report").Select
    Range("A5:G5").Select
    Selection.Font.Bold = True
    Selection.Interior.Color = 12611584
    Range("B6:B105").Select
    Selection.NumberFormat = "m/d/yyyy"
    Range("E6:F105").Select
    Selection.NumberFormat = "0%"
    Range("H5").Select
    ActiveCell.FormulaR1C1 = "Net Amount"
    Range("H6").Select
    ActiveCell.FormulaR1C1 = "=RC[-3]+RC[-2]"
    Range("H6").Select
    Selection.AutoFill Destination:=Range("H6:H105")
    Range("H106").Select
    ActiveCell.FormulaR1C1 = "=SUM(R[-100]C:R[-1]C)"
    Range("A5:H106").Select
    Selection.Columns.AutoFit
End Sub
