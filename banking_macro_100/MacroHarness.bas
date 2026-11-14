Attribute VB_Name = "MacroHarness"
Option Explicit

' Returns "OK" or the error text, so automated validation never stops on a VBA dialog.
Public Function TryFormat(ByVal sheetName As String) As String
    On Error GoTo Failed
    BankingReport.FormatBankingWorksheet ThisWorkbook.Worksheets(sheetName)
    TryFormat = "OK"
    Exit Function
Failed:
    TryFormat = "ERROR " & Err.Number & ": " & Err.Description
End Function

Public Function TryRecorded() As String
    On Error GoTo Failed
    RecordedMacro_Broken.RecordedFormatting_Broken
    TryRecorded = "OK"
    Exit Function
Failed:
    TryRecorded = "ERROR " & Err.Number & ": " & Err.Description
End Function
