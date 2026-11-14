Attribute VB_Name = "BankingReport"
Option Explicit

Private Const ERROR_BASE As Long = vbObjectError + 2400

Public Sub FormatBankingReport()
    On Error GoTo Failed
    If Not TypeOf ActiveSheet Is Worksheet Then
        Err.Raise ERROR_BASE, "BankingReport", "Select a transaction worksheet first."
    End If
    FormatBankingWorksheet ActiveSheet
    MsgBox "Banking report formatted and calculated successfully.", vbInformation, "Banking report"
    Exit Sub
Failed:
    MsgBox Err.Description, vbExclamation, "Banking report was not completed"
End Sub

Public Sub FormatBankingWorksheet(ByVal ws As Worksheet)
    Dim required As Variant, columns(0 To 6) As Long
    Dim headerRow As Long, firstColumn As Long, lastColumn As Long
    Dim lastRow As Long, totalsRow As Long, rowNumber As Long, i As Long
    Dim netColumn As Long, runningColumn As Long, candidateRow As Long
    Dim table As ListObject, otherTable As ListObject, reportRange As Range
    Dim identifiers As Object, identifier As String, value As Variant
    Dim debit As Double, credit As Double, debitTotal As Double, creditTotal As Double
    Dim previousEvents As Boolean, previousScreen As Boolean
    Dim changedApplication As Boolean, errorNumber As Long, errorText As String

    On Error GoTo Failed
    required = Array("Transaction ID", "Transaction Date", "Account ID", _
                     "Description", "Debit", "Credit", "Currency")
    If ws.ProtectContents Then
        Err.Raise ERROR_BASE + 1, "BankingReport", "Unprotect the worksheet before formatting."
    End If
    headerRow = LocateHeaderRow(ws, required)
    lastColumn = ws.Cells(headerRow, ws.Columns.Count).End(xlToLeft).Column
    firstColumn = lastColumn
    For i = 0 To UBound(required)
        columns(i) = HeaderColumn(ws, headerRow, lastColumn, CStr(required(i)))
        If columns(i) < firstColumn Then firstColumn = columns(i)
    Next i
    netColumn = HeaderColumn(ws, headerRow, lastColumn, "Net Amount", False)
    runningColumn = HeaderColumn(ws, headerRow, lastColumn, "Running Net", False)
    If netColumn = 0 Then
        lastColumn = lastColumn + 1
        netColumn = lastColumn
    End If
    If runningColumn = 0 Then
        lastColumn = lastColumn + 1
        runningColumn = lastColumn
    End If
    If netColumn < firstColumn Then firstColumn = netColumn
    If runningColumn < firstColumn Then firstColumn = runningColumn
    If lastColumn > ws.Columns.Count Then
        Err.Raise ERROR_BASE + 2, "BankingReport", "No space for calculated columns."
    End If
    For Each otherTable In ws.ListObjects
        If otherTable.HeaderRowRange.Row = headerRow Then
            If Not Intersect(otherTable.HeaderRowRange, ws.Cells(headerRow, columns(0))) Is Nothing Then
                Set table = otherTable
                If table.ShowTotals Then totalsRow = table.TotalsRowRange.Row
                Exit For
            End If
        End If
    Next otherTable
    lastRow = headerRow
    For i = 0 To UBound(required)
        candidateRow = ws.Cells(ws.Rows.Count, columns(i)).End(xlUp).Row
        If candidateRow = totalsRow Then candidateRow = totalsRow - 1
        If candidateRow > lastRow Then lastRow = candidateRow
    Next i
    If lastRow <= headerRow Then
        Err.Raise ERROR_BASE + 3, "BankingReport", "The report has no transactions."
    End If
    If totalsRow > 0 And lastRow >= totalsRow Then
        Err.Raise ERROR_BASE + 4, "BankingReport", _
                  "Data was entered below the totals row. Insert transactions above totals, then rerun."
    End If
    If lastRow >= ws.Rows.Count Then
        Err.Raise ERROR_BASE + 5, "BankingReport", "Leave one worksheet row available for totals."
    End If
    Set reportRange = ws.Range(ws.Cells(headerRow, firstColumn), ws.Cells(lastRow, lastColumn))
    If Not IsNull(reportRange.MergeCells) Then
        If reportRange.MergeCells Then
            Err.Raise ERROR_BASE + 6, "BankingReport", "Unmerge the report cells first."
        End If
    Else
        Err.Raise ERROR_BASE + 6, "BankingReport", "Unmerge the report cells first."
    End If
    Set identifiers = CreateObject("Scripting.Dictionary")
    identifiers.CompareMode = vbTextCompare
    For i = firstColumn To lastColumn
        If i <> netColumn And i <> runningColumn Then
            value = ws.Cells(headerRow, i).Value2
            If IsError(value) Then
                Err.Raise ERROR_BASE + 7, "BankingReport", "A report header contains an Excel error."
            End If
            identifier = Trim$(CStr(value))
            If Len(identifier) = 0 Or identifiers.Exists(identifier) Then
                Err.Raise ERROR_BASE + 7, "BankingReport", "Report headers must be nonblank and unique."
            End If
            identifiers.Add identifier, True
        End If
    Next i
    Set identifiers = CreateObject("Scripting.Dictionary")
    identifiers.CompareMode = vbTextCompare
    For rowNumber = headerRow + 1 To lastRow
        For i = 0 To UBound(required)
            value = ws.Cells(rowNumber, columns(i)).Value2
            If IsError(value) Then
                InvalidRow rowNumber, CStr(required(i)) & " contains an Excel error."
            End If
        Next i
        identifier = Trim$(CStr(ws.Cells(rowNumber, columns(0)).Value2))
        If Len(identifier) = 0 Then InvalidRow rowNumber, "Transaction ID is missing."
        If identifiers.Exists(identifier) Then InvalidRow rowNumber, "Transaction ID is duplicated."
        identifiers.Add identifier, True
        If Len(Trim$(CStr(ws.Cells(rowNumber, columns(2)).Value2))) = 0 Then
            InvalidRow rowNumber, "Account ID is missing."
        End If
        If Len(Trim$(CStr(ws.Cells(rowNumber, columns(3)).Value2))) = 0 Then
            InvalidRow rowNumber, "Description is missing."
        End If
        value = ws.Cells(rowNumber, columns(1)).Value2
        If Not IsNumericCell(value) Then
            InvalidRow rowNumber, "Transaction Date must be an Excel date, not text."
        End If
        If CDbl(value) < 1 Or CDbl(value) >= 2958466# Then
            InvalidRow rowNumber, "Transaction Date is outside Excel's supported date range."
        End If
        debit = ReadAmount(ws.Cells(rowNumber, columns(4)).Value2, rowNumber, "Debit")
        credit = ReadAmount(ws.Cells(rowNumber, columns(5)).Value2, rowNumber, "Credit")
        If (debit > 0 And credit > 0) Or (debit = 0 And credit = 0) Then
            InvalidRow rowNumber, "Enter a positive Debit OR Credit, not both or neither."
        End If
        If UCase$(Trim$(CStr(ws.Cells(rowNumber, columns(6)).Value2))) <> "USD" Then
            InvalidRow rowNumber, "This demo supports USD only; do not combine currencies."
        End If
        debitTotal = debitTotal + debit
        creditTotal = creditTotal + credit
    Next rowNumber
    For Each otherTable In ws.ListObjects
        If table Is Nothing Then
            If Not Intersect(otherTable.Range, reportRange) Is Nothing Then
                Err.Raise ERROR_BASE + 8, "BankingReport", "The report overlaps another Excel table."
            End If
        ElseIf otherTable.Name <> table.Name Then
            If Not Intersect(otherTable.Range, reportRange) Is Nothing Then
                Err.Raise ERROR_BASE + 8, "BankingReport", "The report overlaps another Excel table."
            End If
        End If
    Next otherTable
    If table Is Nothing Or totalsRow = 0 Then
        If Application.CountA(ws.Range(ws.Cells(lastRow + 1, firstColumn), _
                                       ws.Cells(lastRow + 1, lastColumn))) > 0 Then
            Err.Raise ERROR_BASE + 9, "BankingReport", "Move content below the report to make room for totals."
        End If
    End If

    previousEvents = Application.EnableEvents
    previousScreen = Application.ScreenUpdating
    changedApplication = True
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    ws.Cells(headerRow, netColumn).Value2 = "Net Amount"
    ws.Cells(headerRow, runningColumn).Value2 = "Running Net"
    If table Is Nothing Then
        Set table = ws.ListObjects.Add(xlSrcRange, reportRange, , xlYes)
    Else
        table.ShowTotals = False
        table.Resize reportRange
    End If
    table.TableStyle = "TableStyleMedium2"
    table.ShowTableStyleRowStripes = True
    ws.Range(ws.Cells(headerRow + 1, netColumn), ws.Cells(lastRow, netColumn)).FormulaR1C1 = _
        "=RC[" & CStr(columns(5) - netColumn) & "]-RC[" & CStr(columns(4) - netColumn) & "]"
    ws.Range(ws.Cells(headerRow + 1, runningColumn), ws.Cells(lastRow, runningColumn)).FormulaR1C1 = _
        "=SUM(R" & CStr(headerRow + 1) & "C" & CStr(netColumn) & ":RC" & CStr(netColumn) & ")"
    ws.Range(ws.Cells(headerRow + 1, columns(1)), ws.Cells(lastRow, columns(1))).NumberFormat = "yyyy-mm-dd"
    For i = 4 To 5
        ws.Range(ws.Cells(headerRow + 1, columns(i)), ws.Cells(lastRow, columns(i))).NumberFormat = _
            "#,##0.00;[Red](#,##0.00);0.00"
    Next i
    ws.Range(ws.Cells(headerRow + 1, netColumn), ws.Cells(lastRow, netColumn)).NumberFormat = _
        "#,##0.00;[Red](#,##0.00);0.00"
    ws.Range(ws.Cells(headerRow + 1, runningColumn), ws.Cells(lastRow, runningColumn)).NumberFormat = _
        "#,##0.00;[Red](#,##0.00);0.00"
    table.ShowTotals = True
    table.ListColumns(columns(3) - firstColumn + 1).TotalsCalculation = xlTotalsCalculationNone
    table.TotalsRowRange.Cells(1, columns(3) - firstColumn + 1).Value2 = "VISIBLE TOTAL (USD)"
    For i = 4 To 5
        table.ListColumns(columns(i) - firstColumn + 1).TotalsCalculation = xlTotalsCalculationSum
        table.TotalsRowRange.Cells(1, columns(i) - firstColumn + 1).NumberFormat = _
            "#,##0.00;[Red](#,##0.00);0.00"
    Next i
    table.ListColumns(netColumn - firstColumn + 1).TotalsCalculation = xlTotalsCalculationSum
    table.TotalsRowRange.Cells(1, netColumn - firstColumn + 1).NumberFormat = _
        "#,##0.00;[Red](#,##0.00);0.00"
    table.ListColumns(runningColumn - firstColumn + 1).TotalsCalculation = xlTotalsCalculationNone
    table.TotalsRowRange.Cells(1, runningColumn - firstColumn + 1).ClearContents
    reportRange.Columns.AutoFit
    For i = firstColumn To lastColumn
        If ws.Columns(i).ColumnWidth > 45 Then ws.Columns(i).ColumnWidth = 45
    Next i
    table.HeaderRowRange.Font.Bold = True
    table.HeaderRowRange.WrapText = True
    table.HeaderRowRange.Rows.AutoFit
    ws.Calculate
    If Abs(Application.Sum(ws.Range(ws.Cells(headerRow + 1, netColumn), _
                                     ws.Cells(lastRow, netColumn))) - (creditTotal - debitTotal)) > 0.005 Then
        Err.Raise ERROR_BASE + 10, "BankingReport", "Net Amount reconciliation failed."
    End If
    If Abs(CDbl(ws.Cells(lastRow, runningColumn).Value2) - (creditTotal - debitTotal)) > 0.005 Then
        Err.Raise ERROR_BASE + 11, "BankingReport", "Running Net reconciliation failed."
    End If
    ws.PageSetup.Orientation = xlLandscape
    ws.PageSetup.Zoom = False
    ws.PageSetup.FitToPagesWide = 1
    ws.PageSetup.FitToPagesTall = False
    ws.PageSetup.PrintTitleRows = "$" & CStr(headerRow) & ":$" & CStr(headerRow)
    ws.PageSetup.PrintArea = table.Range.Address
    Application.EnableEvents = previousEvents
    Application.ScreenUpdating = previousScreen
    Exit Sub
Failed:
    errorNumber = Err.Number
    errorText = Err.Description
    If changedApplication Then
        Application.EnableEvents = previousEvents
        Application.ScreenUpdating = previousScreen
    End If
    Err.Raise errorNumber, "BankingReport.FormatBankingWorksheet", errorText
End Sub

Private Function LocateHeaderRow(ByVal ws As Worksheet, ByVal required As Variant) As Long
    Dim rowNumber As Long, lastColumn As Long, i As Long, complete As Boolean
    For rowNumber = 1 To 20
        lastColumn = ws.Cells(rowNumber, ws.Columns.Count).End(xlToLeft).Column
        complete = True
        For i = 0 To UBound(required)
            If HeaderColumn(ws, rowNumber, lastColumn, CStr(required(i)), False) = 0 Then
                complete = False
                Exit For
            End If
        Next i
        If complete Then
            LocateHeaderRow = rowNumber
            Exit Function
        End If
    Next rowNumber
    Err.Raise ERROR_BASE + 12, "BankingReport", _
              "Required headers were not found together in rows 1-20: " & Join(required, ", ") & "."
End Function

Private Function HeaderColumn(ByVal ws As Worksheet, ByVal headerRow As Long, _
                              ByVal lastColumn As Long, ByVal title As String, _
                              Optional ByVal mandatory As Boolean = True) As Long
    Dim columnNumber As Long, value As Variant
    For columnNumber = 1 To lastColumn
        value = ws.Cells(headerRow, columnNumber).Value2
        If Not IsError(value) Then
            If StrComp(Trim$(CStr(value)), title, vbTextCompare) = 0 Then
                If HeaderColumn <> 0 Then
                    Err.Raise ERROR_BASE + 13, "BankingReport", "Duplicate header: " & title & "."
                End If
                HeaderColumn = columnNumber
            End If
        End If
    Next columnNumber
    If mandatory And HeaderColumn = 0 Then
        Err.Raise ERROR_BASE + 14, "BankingReport", "Missing header: " & title & "."
    End If
End Function

Private Function IsNumericCell(ByVal value As Variant) As Boolean
    Select Case VarType(value)
        Case vbByte, vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, vbDecimal
            IsNumericCell = True
    End Select
End Function

Private Function ReadAmount(ByVal value As Variant, ByVal rowNumber As Long, ByVal title As String) As Double
    If IsEmpty(value) Then Exit Function
    If Not IsNumericCell(value) Then InvalidRow rowNumber, title & " must be numeric or blank."
    ReadAmount = CDbl(value)
    If ReadAmount < 0 Then InvalidRow rowNumber, title & " must not be negative."
    If Abs(ReadAmount - Application.Round(ReadAmount, 2)) > 0.000001 Then
        InvalidRow rowNumber, title & " must have at most two decimal places."
    End If
End Function

Private Sub InvalidRow(ByVal rowNumber As Long, ByVal explanation As String)
    Err.Raise ERROR_BASE + 15, "BankingReport", "Row " & CStr(rowNumber) & ": " & explanation
End Sub
