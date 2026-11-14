Attribute VB_Name = "RecordedMacro"
Option Explicit

' BEFORE: code produced by Developer > Record Macro on the original sheet "Sheet1"
' (exactly 100 transactions). Kept for troubleshooting reference - do not use.
Sub FormatReport_Recorded()
    Sheets("Sheet1").Select                         ' Issue 1: hard-coded sheet name
    Range("A1:H1").Select
    Selection.Font.Bold = True
    With Selection.Interior
        .Color = 7884319
    End With
    Range("B2:B101").Select                         ' Issue 2: fixed range (100 rows only)
    Selection.NumberFormat = "dd-mmm-yyyy"
    Range("G2:H101").Select
    Selection.NumberFormat = "0.00%"                ' Issue 3: amounts formatted as percent
    Range("I1").Select
    ActiveCell.FormulaR1C1 = "Net Amount"
    Range("I2").Select
    ActiveCell.FormulaR1C1 = "=RC[-2]+RC[-1]"       ' Issue 4: Debit + Credit (should be Credit - Debit)
    Selection.AutoFill Destination:=Range("I2:I101")
    Range("A102").Select                            ' Issue 5: totals written to fixed row 102
    ActiveCell.FormulaR1C1 = "TOTAL"
    Range("G102").Select
    ActiveCell.FormulaR1C1 = "=SUM(R[-100]C:R[-1]C)"
    Selection.AutoFill Destination:=Range("G102:I102"), Type:=xlFillDefault
    Columns("A:I").EntireColumn.AutoFit
End Sub
