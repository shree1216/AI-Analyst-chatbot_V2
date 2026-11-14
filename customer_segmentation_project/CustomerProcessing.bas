Attribute VB_Name = "CustomerProcessing"
Option Explicit

' Cleans the customer records on the "Raw Data" sheet and distributes them into
' category worksheets. The raw sheet is never changed, so the macro can be rerun at any time.
' Raw Data columns: Customer ID | Customer Name | Age | City | Phone | Email |
'                   Account Type | Loan Type | Loan Amount

Private Const RAW_SHEET As String = "Raw Data"
Private Const COLS As Long = 9
Private Const SENIOR_AGE As Long = 60

' Button / Alt+F8 entry point.
Public Sub ProcessCustomers()
    On Error GoTo Failed
    MsgBox ProcessCustomerData(), vbInformation, "Customer Segmentation"
    Exit Sub
Failed:
    Application.ScreenUpdating = True
    MsgBox Err.Description, vbExclamation, "Customer Segmentation"
End Sub

' Does the work and returns a short summary message.
Public Function ProcessCustomerData() As String
    Dim raw As Worksheet, data As Variant, headers As Variant, rec(1 To COLS) As Variant
    Dim clean As Worksheet, seniors As Worksheet, homeLoans As Worksheet, carLoans As Worksheet
    Dim duplicates As Worksheet, incomplete As Worksheet, summary As Worksheet
    Dim seen As Object, sheet As Variant, lastRow As Long, r As Long, c As Long, issue As String
    Dim nClean As Long, nSenior As Long, nHome As Long, nCar As Long, nDuplicate As Long, nIncomplete As Long

    Set raw = ThisWorkbook.Worksheets(RAW_SHEET)
    lastRow = LastUsedRow(raw)
    If lastRow < 2 Then Err.Raise vbObjectError + 1, , "No customer records found on '" & RAW_SHEET & "'."
    headers = raw.Range("A1").Resize(1, COLS).Value
    data = raw.Range("A2").Resize(lastRow - 1, COLS).Value

    Application.ScreenUpdating = False
    Set summary = PrepareSheet("Summary", Empty)
    Set clean = PrepareSheet("Clean Data", headers)
    Set seniors = PrepareSheet("Senior Citizens", headers)
    Set homeLoans = PrepareSheet("Home Loan", headers)
    Set carLoans = PrepareSheet("Car Loan", headers)
    Set duplicates = PrepareSheet("Duplicates", headers, True)
    Set incomplete = PrepareSheet("Incomplete", headers, True)
    Set seen = CreateObject("Scripting.Dictionary")

    For r = 1 To UBound(data, 1)
        For c = 1 To COLS
            rec(c) = data(r, c)
        Next c
        StandardizeRecord rec
        issue = FindIssues(rec)

        If Len(issue) > 0 Then
            AddRow incomplete, rec, "Missing or invalid: " & issue
            nIncomplete = nIncomplete + 1
        ElseIf seen.Exists(rec(1)) Then
            AddRow duplicates, rec, "Duplicate of Raw Data row " & seen(rec(1))
            nDuplicate = nDuplicate + 1
        Else
            seen.Add rec(1), r + 1
            AddRow clean, rec
            nClean = nClean + 1
            If rec(3) >= SENIOR_AGE Then
                AddRow seniors, rec
                nSenior = nSenior + 1
            End If
            If rec(8) = "Home Loan" Then
                AddRow homeLoans, rec
                nHome = nHome + 1
            End If
            If rec(8) = "Car Loan" Then
                AddRow carLoans, rec
                nCar = nCar + 1
            End If
        End If
    Next r

    For Each sheet In Array(clean, seniors, homeLoans, carLoans, duplicates, incomplete)
        FormatSheet sheet
    Next sheet
    WriteSummary summary, UBound(data, 1), nDuplicate, nIncomplete, nClean, nSenior, nHome, nCar
    Application.ScreenUpdating = True

    ProcessCustomerData = UBound(data, 1) & " records processed." & vbNewLine & vbNewLine & _
        "Clean customers: " & nClean & vbNewLine & _
        "Duplicates: " & nDuplicate & vbNewLine & _
        "Incomplete: " & nIncomplete & vbNewLine & vbNewLine & _
        "Senior citizens: " & nSenior & vbNewLine & _
        "Home loan holders: " & nHome & vbNewLine & _
        "Car loan holders: " & nCar
End Function

' ---------- Standardization ----------------------------------------------------------------

Private Sub StandardizeRecord(rec() As Variant)
    rec(1) = UCase$(CleanText(rec(1)))                         ' Customer ID
    rec(2) = StrConv(CleanText(rec(2)), vbProperCase)          ' Customer Name
    rec(3) = CleanText(rec(3))                                 ' Age
    If IsNumeric(rec(3)) Then rec(3) = CLng(rec(3))
    rec(4) = StrConv(CleanText(rec(4)), vbProperCase)          ' City
    rec(5) = FormatPhone(rec(5))                               ' Phone
    rec(6) = LCase$(CleanText(rec(6)))                         ' Email
    rec(7) = StrConv(CleanText(rec(7)), vbProperCase)          ' Account Type
    rec(8) = StandardLoanType(rec(8))                          ' Loan Type
    If IsNumeric(rec(9)) And Not IsEmpty(rec(9)) Then rec(9) = CDbl(rec(9)) Else rec(9) = Empty
End Sub

' Trims the text and collapses repeated spaces.
Private Function CleanText(ByVal value As Variant) As String
    CleanText = Application.WorksheetFunction.Trim(CStr(value))
End Function

' Returns phone numbers as 555-123-4567. Invalid numbers are returned unchanged.
Private Function FormatPhone(ByVal value As Variant) As String
    Dim digits As String, i As Long, ch As String
    For i = 1 To Len(CStr(value))
        ch = Mid$(CStr(value), i, 1)
        If ch Like "#" Then digits = digits & ch
    Next i
    If Len(digits) = 11 And Left$(digits, 1) = "1" Then digits = Mid$(digits, 2)
    If Len(digits) = 10 Then
        FormatPhone = Left$(digits, 3) & "-" & Mid$(digits, 4, 3) & "-" & Right$(digits, 4)
    Else
        FormatPhone = CleanText(value)
    End If
End Function

Private Function StandardLoanType(ByVal value As Variant) As String
    Dim text As String
    text = LCase$(CleanText(value))
    If text Like "*home*" Or text Like "*housing*" Then
        StandardLoanType = "Home Loan"
    ElseIf text Like "*car*" Or text Like "*auto*" Then
        StandardLoanType = "Car Loan"
    ElseIf text Like "*personal*" Then
        StandardLoanType = "Personal Loan"
    ElseIf text = "" Or text = "none" Or text = "no loan" Then
        StandardLoanType = "None"
    Else
        StandardLoanType = StrConv(text, vbProperCase)
    End If
End Function

' Lists the required fields that are missing or invalid (empty string = record is complete).
Private Function FindIssues(rec() As Variant) As String
    Dim issues As String
    If rec(1) = "" Then issues = issues & ", Customer ID"
    If rec(2) = "" Then issues = issues & ", Name"
    If VarType(rec(3)) <> vbLong Then
        issues = issues & ", Age"
    ElseIf rec(3) < 18 Or rec(3) > 120 Then
        issues = issues & ", Age"
    End If
    If rec(4) = "" Then issues = issues & ", City"
    If Not (rec(5) Like "###-###-####") Then issues = issues & ", Phone"
    If Len(issues) > 0 Then FindIssues = Mid$(issues, 3)
End Function

' ---------- Worksheet output ---------------------------------------------------------------

' Returns an empty sheet with the given name (created at the end if it does not exist).
Private Function PrepareSheet(ByVal name As String, ByVal headers As Variant, _
                              Optional ByVal withIssue As Boolean = False) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(name)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = name
    End If
    ws.Cells.Clear
    ws.Columns(5).NumberFormat = "@"    ' keep phone numbers as text
    If Not IsEmpty(headers) Then
        ws.Range("A1").Resize(1, COLS).Value = headers
        If withIssue Then ws.Cells(1, COLS + 1).Value = "Issue"
    End If
    Set PrepareSheet = ws
End Function

Private Sub AddRow(ByVal ws As Worksheet, rec() As Variant, Optional ByVal issue As String = "")
    Dim nextRow As Long
    nextRow = LastUsedRow(ws) + 1
    ws.Cells(nextRow, 1).Resize(1, COLS).Value = rec
    If Len(issue) > 0 Then ws.Cells(nextRow, COLS + 1).Value = issue
End Sub

Private Function LastUsedRow(ByVal ws As Worksheet) As Long
    Dim found As Range
    Set found = ws.Cells.Find(What:="*", LookIn:=xlFormulas, SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
    If Not found Is Nothing Then LastUsedRow = found.Row
End Function

Private Sub FormatSheet(ByVal ws As Worksheet)
    Dim table As Range
    Set table = ws.Range("A1").CurrentRegion
    StyleHeader table.Rows(1)
    With table.Borders
        .LineStyle = xlContinuous
        .Color = RGB(191, 191, 191)
    End With
    ws.Range("I2:I" & table.Rows.Count).NumberFormat = "#,##0"
    ws.Range("C2:C" & table.Rows.Count).HorizontalAlignment = xlCenter
    table.Columns.AutoFit
End Sub

Private Sub StyleHeader(ByVal header As Range)
    With header
        .Font.Bold = True
        .Font.Color = vbWhite
        .Interior.Color = RGB(31, 78, 121)
        .HorizontalAlignment = xlCenter
    End With
End Sub

Private Sub WriteSummary(ByVal ws As Worksheet, ByVal total As Long, ByVal nDuplicate As Long, _
                         ByVal nIncomplete As Long, ByVal nClean As Long, ByVal nSenior As Long, _
                         ByVal nHome As Long, ByVal nCar As Long)
    Dim chartBox As ChartObject

    ws.Range("A1").Value = "Customer Segmentation Summary"
    ws.Range("A1").Font.Size = 14
    ws.Range("A1").Font.Bold = True
    ws.Range("A2").Value = "Processed on " & Format$(Now, "dd-mmm-yyyy hh:nn")
    ws.Range("A2").Font.Italic = True

    ws.Range("A4:B4").Value = Array("Category", "Customers")
    ws.Range("A5:B5").Value = Array("Total records received", total)
    ws.Range("A6:B6").Value = Array("Duplicate records", nDuplicate)
    ws.Range("A7:B7").Value = Array("Incomplete records", nIncomplete)
    ws.Range("A8:B8").Value = Array("Clean customers", nClean)
    ws.Range("A9:B9").Value = Array("Senior citizens (" & SENIOR_AGE & "+)", nSenior)
    ws.Range("A10:B10").Value = Array("Home loan holders", nHome)
    ws.Range("A11:B11").Value = Array("Car loan holders", nCar)

    StyleHeader ws.Range("A4:B4")
    With ws.Range("A4:B11").Borders
        .LineStyle = xlContinuous
        .Color = RGB(191, 191, 191)
    End With
    ws.Range("A8:B8").Font.Bold = True
    ws.Range("B5:B11").HorizontalAlignment = xlCenter
    ws.Columns("A").ColumnWidth = 30
    ws.Columns("B").ColumnWidth = 14

    For Each chartBox In ws.ChartObjects
        chartBox.Delete
    Next chartBox
    Set chartBox = ws.ChartObjects.Add(ws.Range("D4").Left, ws.Range("D4").Top, 420, 230)
    With chartBox.Chart
        .ChartType = xlBarClustered
        .SetSourceData ws.Range("A9:B11")
        .HasTitle = True
        .ChartTitle.Text = "Customers by Category"
        .HasLegend = False
        .SeriesCollection(1).HasDataLabels = True
        .SeriesCollection(1).Format.Fill.ForeColor.RGB = RGB(31, 78, 121)
    End With
End Sub
