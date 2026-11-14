Attribute VB_Name = "RecordedMacro_Broken"
Option Explicit

' Illustrative recorded-style code, not a recording from your Excel session.
' Do not run this on a real report; see README.md for the faults.
Public Sub RecordedFormatting_Broken()
    Worksheets("Sheet1").Select
    Range("A1:G11").Select
    Selection.Font.Bold = True
    Range("E2:F11").NumberFormat = "0%"
    Range("H2:H11").FormulaR1C1 = "=RC[-3]+RC[-2]"
    Range("A1:H11").Columns.AutoFit
End Sub
