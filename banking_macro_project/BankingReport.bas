Attribute VB_Name = "BankingReport"
Option Explicit

' AFTER: corrected, reusable version of the recorded macro.
' Works on the active worksheet, finds columns by header name and counts rows
' dynamically, so renamed sheets, reordered columns and added/removed rows are handled.

Private Const TOTAL_LABEL As String = "TOTAL"
Private Const AMOUNT_FORMAT As String = "#,##0.00;[Red]-#,##0.00"

Public Sub FormatTransactionReport()
    Dim rowCount As Long
    On Error GoTo Failed
    rowCount = FormatReport(ActiveSheet)
    MsgBox rowCount & " transactions formatted on '" & ActiveSheet.Name & "'.", _
           vbInformation, "Banking Report"
    Exit Sub
Failed:
    Application.ScreenUpdating = True
    MsgBox Err.Description, vbExclamation, "Banking Report - not formatted"
End Sub

' Formats one transaction sheet and returns the number of transactions processed.
Public Function FormatReport(ByVal ws As Worksheet) As Long
    Dim headerRow As Long, lastRow As Long, firstCol As Long, lastCol As Long, r As Long
    Dim idCol As Long, dateCol As Long, debitCol As Long, creditCol As Long, netCol As Long
    Dim totalRow As Long

    headerRow = FindHeaderRow(ws)
    idCol = FindColumn(ws, headerRow, "Transaction ID")
    dateCol = FindColumn(ws, headerRow, "Date")
    debitCol = FindColumn(ws, headerRow, "Debit")
    creditCol = FindColumn(ws, headerRow, "Credit")
    netCol = FindColumn(ws, headerRow, "Net Amount", False)

    ' Remove the totals row left by a previous run so the macro can be rerun safely.
    lastRow = ws.Cells(ws.Rows.Count, idCol).End(xlUp).Row
    For r = lastRow To headerRow + 1 Step -1
        If UCase$(Trim$(CStr(ws.Cells(r, idCol).Value))) = TOTAL_LABEL Then ws.Rows(r).Delete
    Next r
    lastRow = ws.Cells(ws.Rows.Count, idCol).End(xlUp).Row
    If lastRow <= headerRow Then Err.Raise vbObjectError + 1, , "No transactions found below the header row."

    ' Validate before changing anything.
    For r = headerRow + 1 To lastRow
        If VarType(ws.Cells(r, dateCol).Value) <> vbDate Then _
            Err.Raise vbObjectError + 2, , "Row " & r & ": Date '" & ws.Cells(r, dateCol).Text & "' is not a valid date."
        CheckAmount ws.Cells(r, debitCol), "Debit"
        CheckAmount ws.Cells(r, creditCol), "Credit"
    Next r

    Application.ScreenUpdating = False
    firstCol = 1
    If IsEmpty(ws.Cells(headerRow, 1).Value) Then firstCol = ws.Cells(headerRow, 1).End(xlToRight).Column
    lastCol = ws.Cells(headerRow, ws.Columns.Count).End(xlToLeft).Column
    If netCol = 0 Then
        netCol = lastCol + 1
        lastCol = netCol
        ws.Cells(headerRow, netCol).Value = "Net Amount"
    End If
    totalRow = lastRow + 1

    ' Header style
    With ws.Range(ws.Cells(headerRow, firstCol), ws.Cells(headerRow, lastCol))
        .Font.Bold = True
        .Font.Color = vbWhite
        .Interior.Color = RGB(31, 78, 121)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With

    ' Data formats and calculation
    With ws.Range(ws.Cells(headerRow + 1, firstCol), ws.Cells(lastRow, lastCol))
        .Font.Bold = False
        .Interior.Pattern = xlNone
    End With
    ws.Range(ws.Cells(headerRow + 1, dateCol), ws.Cells(lastRow, dateCol)).NumberFormat = "dd-mmm-yyyy"
    ws.Range(ws.Cells(headerRow + 1, netCol), ws.Cells(lastRow, netCol)).FormulaR1C1 = _
        "=N(RC" & creditCol & ")-N(RC" & debitCol & ")"
    FormatAmountColumn ws, headerRow + 1, totalRow, debitCol
    FormatAmountColumn ws, headerRow + 1, totalRow, creditCol
    FormatAmountColumn ws, headerRow + 1, totalRow, netCol

    ' Totals row
    ws.Cells(totalRow, idCol).Value = TOTAL_LABEL
    ws.Cells(totalRow, debitCol).FormulaR1C1 = "=SUM(R" & headerRow + 1 & "C:R[-1]C)"
    ws.Cells(totalRow, creditCol).FormulaR1C1 = "=SUM(R" & headerRow + 1 & "C:R[-1]C)"
    ws.Cells(totalRow, netCol).FormulaR1C1 = "=SUM(R" & headerRow + 1 & "C:R[-1]C)"
    With ws.Range(ws.Cells(totalRow, firstCol), ws.Cells(totalRow, lastCol))
        .Font.Bold = True
        .Interior.Color = RGB(221, 235, 247)
        .Borders(xlEdgeTop).LineStyle = xlDouble
    End With

    ' Borders, widths, frozen header
    With ws.Range(ws.Cells(headerRow, firstCol), ws.Cells(lastRow, lastCol)).Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
        .Color = RGB(191, 191, 191)
    End With
    ws.Range(ws.Cells(headerRow, firstCol), ws.Cells(totalRow, lastCol)).Columns.AutoFit
    If ws Is ActiveSheet Then
        ActiveWindow.FreezePanes = False
        ActiveWindow.SplitColumn = 0
        ActiveWindow.SplitRow = headerRow
        ActiveWindow.FreezePanes = True
    End If
    Application.ScreenUpdating = True

    FormatReport = lastRow - headerRow
End Function

Private Function FindHeaderRow(ByVal ws As Worksheet) As Long
    Dim found As Range
    Set found = ws.Range("A1:Z10").Find(What:="Transaction ID", LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False)
    If found Is Nothing Then Err.Raise vbObjectError + 3, , _
        "Header 'Transaction ID' was not found in the first 10 rows of '" & ws.Name & "'."
    FindHeaderRow = found.Row
End Function

Private Function FindColumn(ByVal ws As Worksheet, ByVal headerRow As Long, ByVal header As String, _
                            Optional ByVal required As Boolean = True) As Long
    Dim found As Range
    Set found = ws.Rows(headerRow).Find(What:=header, LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False)
    If Not found Is Nothing Then
        FindColumn = found.Column
    ElseIf required Then
        Err.Raise vbObjectError + 4, , "Required column '" & header & "' was not found."
    End If
End Function

Private Sub CheckAmount(ByVal cell As Range, ByVal header As String)
    If IsEmpty(cell.Value) Then Exit Sub
    If VarType(cell.Value) <> vbDouble And VarType(cell.Value) <> vbCurrency Then
        Err.Raise vbObjectError + 5, , "Row " & cell.Row & ": " & header & " '" & cell.Text & "' is not a number."
    End If
    If cell.Value < 0 Then Err.Raise vbObjectError + 6, , "Row " & cell.Row & ": " & header & " cannot be negative."
End Sub

Private Sub FormatAmountColumn(ByVal ws As Worksheet, ByVal firstRow As Long, ByVal lastRow As Long, ByVal col As Long)
    With ws.Range(ws.Cells(firstRow, col), ws.Cells(lastRow, col))
        .NumberFormat = AMOUNT_FORMAT
        .HorizontalAlignment = xlRight
    End With
End Sub
