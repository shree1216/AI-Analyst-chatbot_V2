Attribute VB_Name = "HealthcareReport"
Option Explicit

' Builds the healthcare dashboard and one report sheet per month from the "Raw Data" sheet.
' Raw Data columns: Visit ID | Patient ID | Patient Name | Department | Doctor |
'                   Appointment Date | Visit Type | Status | Billing Amount

Private Const RAW_SHEET As String = "Raw Data"
Private Const CLEAN_SHEET As String = "Clean Data"
Private Const DASHBOARD_SHEET As String = "Dashboard"
Private Const REPORT_PREFIX As String = "Report "
Private Const COLS As Long = 9
Private Const C_DEPT As Long = 4
Private Const C_DATE As Long = 6
Private Const C_STATUS As Long = 8
Private Const C_BILL As Long = 9

Private Const NAVY As Long = 7949855         ' RGB(31, 78, 121)
Private Const LIGHT_BLUE As Long = 16247773  ' RGB(221, 235, 247)
Private Const GREY As Long = 12566463        ' RGB(191, 191, 191)
Private Const MONEY As String = "#,##0.00"

' Button entry point: refreshes the clean data, dashboard, monthly reports and charts.
Public Sub RefreshReport()
    On Error GoTo Failed
    MsgBox BuildReports(), vbInformation, "Healthcare Dashboard"
    Exit Sub
Failed:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    MsgBox Err.Description, vbExclamation, "Healthcare Dashboard"
End Sub

Public Function BuildReports() As String
    Dim data As Variant, k As Long, rawCount As Long, dupCount As Long, badCount As Long
    Dim months As Variant, i As Long, ws As Worksheet, lastSheet As Worksheet

    Application.ScreenUpdating = False
    data = CleanRecords(k, rawCount, dupCount, badCount)
    If k = 0 Then Err.Raise vbObjectError + 1, , "No valid records found on '" & RAW_SHEET & "'."
    months = MonthKeys(data, k)

    BuildDashboard data, k, months, rawCount, dupCount, badCount

    ' Monthly reports: remove the old ones and create one sheet per month.
    Application.DisplayAlerts = False
    For i = ThisWorkbook.Worksheets.Count To 1 Step -1
        If ThisWorkbook.Worksheets(i).Name Like REPORT_PREFIX & "*" Then ThisWorkbook.Worksheets(i).Delete
    Next i
    Application.DisplayAlerts = True
    Set lastSheet = ThisWorkbook.Worksheets(DASHBOARD_SHEET)
    For i = 0 To UBound(months)
        Set ws = ThisWorkbook.Worksheets.Add(After:=lastSheet)
        ws.Name = REPORT_PREFIX & MonthLabel(months(i))
        BuildMonthlyReport ws, data, k, months(i)
        Set lastSheet = ws
    Next i

    ThisWorkbook.Worksheets(DASHBOARD_SHEET).Activate
    ThisWorkbook.Worksheets(DASHBOARD_SHEET).Range("A1").Select
    Application.ScreenUpdating = True

    BuildReports = "Report refreshed." & vbNewLine & vbNewLine & _
                   "Raw records: " & rawCount & vbNewLine & _
                   "Duplicates removed: " & dupCount & vbNewLine & _
                   "Incomplete removed: " & badCount & vbNewLine & _
                   "Clean records: " & k & vbNewLine & _
                   "Monthly reports created: " & UBound(months) + 1
End Function

' ---------- 1. Cleaning ---------------------------------------------------------------------

' Standardizes the raw records, drops duplicates and incomplete rows, writes "Clean Data"
' and returns the clean records (only the first k rows are used).
Private Function CleanRecords(ByRef k As Long, ByRef rawCount As Long, ByRef dupCount As Long, _
                              ByRef badCount As Long) As Variant
    Dim raw As Worksheet, ws As Worksheet, src As Variant, out() As Variant, seen As Object
    Dim lastRow As Long, r As Long, id As String, dept As String, status As String, bill As Variant
    Dim valid As Boolean

    Set raw = ThisWorkbook.Worksheets(RAW_SHEET)
    lastRow = raw.Cells(raw.Rows.Count, 1).End(xlUp).Row
    If lastRow < 2 Then Err.Raise vbObjectError + 2, , "No records found on '" & RAW_SHEET & "'."
    src = raw.Range("A2").Resize(lastRow - 1, COLS).Value
    rawCount = UBound(src, 1)
    ReDim out(1 To rawCount, 1 To COLS)
    Set seen = CreateObject("Scripting.Dictionary")

    For r = 1 To rawCount
        id = UCase$(CleanText(src(r, 1)))
        dept = StrConv(CleanText(src(r, C_DEPT)), vbProperCase)
        status = StandardStatus(src(r, C_STATUS))
        bill = src(r, C_BILL)
        If CleanText(bill) = "" Then bill = 0

        If id = "" Or dept = "" Or Not IsDate(src(r, C_DATE)) Or Not IsNumeric(bill) Then
            valid = False
        Else
            valid = Not (status = "Completed" And CDbl(bill) <= 0)   ' a completed visit must be billed
        End If

        If Not valid Then
            badCount = badCount + 1
        ElseIf seen.Exists(id) Then
            dupCount = dupCount + 1
        Else
            seen.Add id, True
            k = k + 1
            out(k, 1) = id
            out(k, 2) = UCase$(CleanText(src(r, 2)))
            out(k, 3) = StrConv(CleanText(src(r, 3)), vbProperCase)
            out(k, 4) = dept
            out(k, 5) = CleanText(src(r, 5))
            out(k, 6) = CDate(src(r, C_DATE))
            out(k, 7) = StrConv(CleanText(src(r, 7)), vbProperCase)
            out(k, 8) = status
            out(k, 9) = CDbl(bill)
        End If
    Next r

    Set ws = GetSheet(CLEAN_SHEET)
    ws.Range("A1").Resize(1, COLS).Value = raw.Range("A1").Resize(1, COLS).Value
    StyleTable ws.Range("A1").Resize(k + 1, COLS)
    If k > 0 Then
        ws.Range("A2").Resize(k, COLS).Value = out
        ws.Range("F2").Resize(k, 1).NumberFormat = "dd-mmm-yyyy"
        ws.Range("I2").Resize(k, 1).NumberFormat = MONEY
    End If
    ws.Columns("A:I").AutoFit
    CleanRecords = out
End Function

Private Function CleanText(ByVal value As Variant) As String
    CleanText = Application.WorksheetFunction.Trim(CStr(value))
End Function

Private Function StandardStatus(ByVal value As Variant) As String
    Select Case LCase$(Replace(CleanText(value), " ", "-"))
        Case "completed": StandardStatus = "Completed"
        Case "cancelled", "canceled": StandardStatus = "Cancelled"
        Case "no-show", "noshow": StandardStatus = "No-Show"
        Case Else: StandardStatus = StrConv(CleanText(value), vbProperCase)
    End Select
End Function

' ---------- 2. Dashboard --------------------------------------------------------------------

Private Sub BuildDashboard(ByVal data As Variant, ByVal k As Long, ByVal months As Variant, _
                           ByVal rawCount As Long, ByVal dupCount As Long, ByVal badCount As Long)
    Dim ws As Worksheet, deptEnd As Long, top As Long, i As Long, r As Long
    Dim appts As Long, completed As Long, revenue As Double

    Set ws = GetSheet(DASHBOARD_SHEET)
    SetupPage ws, "Healthcare Performance Dashboard", _
              "Period: " & MonthLabel(months(0)) & " to " & MonthLabel(months(UBound(months)))
    WriteKpis ws, 4, data, k, ""
    deptEnd = WriteDepartmentTable(ws, 7, data, k, "")

    ' Monthly trend table
    top = deptEnd + 3
    SectionTitle ws.Cells(top, 2), "Monthly Trend"
    ws.Cells(top + 1, 2).Resize(1, 5).Value = Array("Month", "Appointments", "Completed", "Revenue", "Completion Rate")
    For i = 0 To UBound(months)
        r = top + 2 + i
        Totals data, k, months(i), "", appts, completed, revenue
        ws.Cells(r, 2).NumberFormat = "@"     ' keep "Aug 2026" as text, not a date
        ws.Cells(r, 2).Resize(1, 5).Value = Array(MonthLabel(months(i)), appts, completed, revenue, completed / appts)
    Next i
    r = top + 2 + UBound(months)
    ws.Range(ws.Cells(top + 2, 5), ws.Cells(r, 5)).NumberFormat = MONEY
    ws.Range(ws.Cells(top + 2, 6), ws.Cells(r, 6)).NumberFormat = "0.0%"
    StyleTable ws.Range(ws.Cells(top + 1, 2), ws.Cells(r, 6))

    ' Data quality table
    SectionTitle ws.Cells(top, 8), "Data Quality"
    ws.Cells(top + 1, 8).Resize(1, 2).Value = Array("Check", "Records")
    ws.Cells(top + 2, 8).Resize(1, 2).Value = Array("Raw records", rawCount)
    ws.Cells(top + 3, 8).Resize(1, 2).Value = Array("Duplicates", dupCount)
    ws.Cells(top + 4, 8).Resize(1, 2).Value = Array("Incomplete", badCount)
    ws.Cells(top + 5, 8).Resize(1, 2).Value = Array("Clean records", k)
    StyleTable ws.Range(ws.Cells(top + 1, 8), ws.Cells(top + 5, 9))
    ws.Range(ws.Cells(top + 5, 8), ws.Cells(top + 5, 9)).Font.Bold = True

    ' Charts
    AddChart ws, "Revenue by Department", ws.Range(ws.Cells(9, 2), ws.Cells(deptEnd - 1, 2)), _
             ws.Range(ws.Cells(9, 5), ws.Cells(deptEnd - 1, 5)), ws.Cells(7, 11), xlBarClustered
    AddChart ws, "Monthly Appointments", ws.Range(ws.Cells(top + 2, 2), ws.Cells(r, 2)), _
             ws.Range(ws.Cells(top + 2, 3), ws.Cells(r, 3)), ws.Cells(r + 3, 2), xlColumnClustered
    AddChart ws, "Monthly Revenue", ws.Range(ws.Cells(top + 2, 2), ws.Cells(r, 2)), _
             ws.Range(ws.Cells(top + 2, 5), ws.Cells(r, 5)), ws.Cells(r + 3, 6), xlLineMarkers
End Sub

' ---------- 3. Monthly report ---------------------------------------------------------------

Private Sub BuildMonthlyReport(ByVal ws As Worksheet, ByVal data As Variant, ByVal k As Long, ByVal monthKey As String)
    Dim deptEnd As Long
    SetupPage ws, "Monthly Report - " & Format$(MonthStart(monthKey), "mmmm yyyy"), _
              "Department performance for " & Format$(MonthStart(monthKey), "mmmm yyyy")
    WriteKpis ws, 4, data, k, monthKey
    deptEnd = WriteDepartmentTable(ws, 7, data, k, monthKey)
    AddChart ws, "Revenue by Department", ws.Range(ws.Cells(9, 2), ws.Cells(deptEnd - 1, 2)), _
             ws.Range(ws.Cells(9, 5), ws.Cells(deptEnd - 1, 5)), ws.Cells(7, 11), xlBarClustered
End Sub

' ---------- Shared building blocks ----------------------------------------------------------

' Totals for one month ("" = all months) and one department ("" = all departments).
Private Sub Totals(ByVal data As Variant, ByVal k As Long, ByVal monthKey As String, ByVal dept As String, _
                   ByRef appts As Long, ByRef completed As Long, ByRef revenue As Double)
    Dim r As Long
    appts = 0: completed = 0: revenue = 0
    For r = 1 To k
        If (monthKey = "" Or Format$(data(r, C_DATE), "yyyy-mm") = monthKey) And _
           (dept = "" Or data(r, C_DEPT) = dept) Then
            appts = appts + 1
            If data(r, C_STATUS) = "Completed" Then completed = completed + 1
            revenue = revenue + data(r, C_BILL)
        End If
    Next r
End Sub

Private Sub WriteKpis(ByVal ws As Worksheet, ByVal topRow As Long, ByVal data As Variant, ByVal k As Long, _
                      ByVal monthKey As String)
    Dim appts As Long, completed As Long, revenue As Double
    Totals data, k, monthKey, "", appts, completed, revenue
    KpiTile ws.Cells(topRow, 2), "Total Appointments", appts, "#,##0"
    KpiTile ws.Cells(topRow, 4), "Completion Rate", completed / appts, "0.0%"
    KpiTile ws.Cells(topRow, 6), "Total Revenue", revenue, MONEY
    KpiTile ws.Cells(topRow, 8), "Average Bill", IIf(completed > 0, revenue / completed, 0), MONEY
End Sub

' Writes the department table (sorted by revenue) and returns the row of its Total line.
Private Function WriteDepartmentTable(ByVal ws As Worksheet, ByVal top As Long, ByVal data As Variant, _
                                      ByVal k As Long, ByVal monthKey As String) As Long
    Dim depts As Object, dept As Variant, r As Long, appts As Long, completed As Long, revenue As Double
    Dim totalAppts As Long, totalCompleted As Long, totalRevenue As Double

    Set depts = CreateObject("Scripting.Dictionary")
    For r = 1 To k
        depts(data(r, C_DEPT)) = True
    Next r
    Totals data, k, monthKey, "", totalAppts, totalCompleted, totalRevenue

    SectionTitle ws.Cells(top, 2), "Department Performance"
    ws.Cells(top + 1, 2).Resize(1, 6).Value = Array("Department", "Appointments", "Completed", "Revenue", "Average Bill", "% of Revenue")
    r = top + 2
    For Each dept In depts.Keys
        Totals data, k, monthKey, CStr(dept), appts, completed, revenue
        If appts > 0 Then
            ws.Cells(r, 2).Resize(1, 6).Value = Array(dept, appts, completed, revenue, _
                IIf(completed > 0, revenue / completed, 0), IIf(totalRevenue > 0, revenue / totalRevenue, 0))
            r = r + 1
        End If
    Next dept
    ws.Range(ws.Cells(top + 2, 2), ws.Cells(r - 1, 7)).Sort Key1:=ws.Cells(top + 2, 5), Order1:=xlDescending, Header:=xlNo

    ws.Cells(r, 2).Resize(1, 6).Value = Array("Total", totalAppts, totalCompleted, totalRevenue, _
        IIf(totalCompleted > 0, totalRevenue / totalCompleted, 0), 1)
    ws.Range(ws.Cells(top + 2, 5), ws.Cells(r, 6)).NumberFormat = MONEY
    ws.Range(ws.Cells(top + 2, 7), ws.Cells(r, 7)).NumberFormat = "0.0%"
    StyleTable ws.Range(ws.Cells(top + 1, 2), ws.Cells(r, 7))
    With ws.Range(ws.Cells(r, 2), ws.Cells(r, 7))
        .Font.Bold = True
        .Interior.Color = LIGHT_BLUE
    End With
    WriteDepartmentTable = r
End Function

' Sorted list of months ("yyyy-mm") found in the data.
Private Function MonthKeys(ByVal data As Variant, ByVal k As Long) As Variant
    Dim found As Object, keys As Variant, r As Long, i As Long, j As Long, temp As Variant
    Set found = CreateObject("Scripting.Dictionary")
    For r = 1 To k
        found(Format$(data(r, C_DATE), "yyyy-mm")) = True
    Next r
    keys = found.Keys
    For i = 0 To UBound(keys) - 1
        For j = i + 1 To UBound(keys)
            If keys(j) < keys(i) Then
                temp = keys(i)
                keys(i) = keys(j)
                keys(j) = temp
            End If
        Next j
    Next i
    MonthKeys = keys
End Function

Private Function MonthStart(ByVal monthKey As String) As Date
    MonthStart = DateSerial(CLng(Left$(monthKey, 4)), CLng(Right$(monthKey, 2)), 1)
End Function

Private Function MonthLabel(ByVal monthKey As String) As String
    MonthLabel = Format$(MonthStart(monthKey), "mmm yyyy")
End Function

' ---------- Formatting helpers --------------------------------------------------------------

' Returns the named sheet, emptied (charts removed, buttons kept). Creates it if needed.
Private Function GetSheet(ByVal name As String) As Worksheet
    Dim ws As Worksheet, chartBox As ChartObject
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(name)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = name
    End If
    ws.Cells.Clear
    For Each chartBox In ws.ChartObjects
        chartBox.Delete
    Next chartBox
    Set GetSheet = ws
End Function

Private Sub SetupPage(ByVal ws As Worksheet, ByVal title As String, ByVal subtitle As String)
    ws.Cells.Clear
    ws.Columns("A").ColumnWidth = 2
    ws.Columns("B").ColumnWidth = 20
    ws.Columns("C:I").ColumnWidth = 15
    ws.Columns("J").ColumnWidth = 3
    With ws.Range("B1")
        .Value = title
        .Font.Size = 18
        .Font.Bold = True
        .Font.Color = NAVY
    End With
    With ws.Range("B2")
        .Value = subtitle & "     |     Last refreshed: " & Format$(Now, "dd-mmm-yyyy hh:nn")
        .Font.Italic = True
        .Font.Color = RGB(89, 89, 89)
    End With
    ws.Activate
    ActiveWindow.DisplayGridlines = False
End Sub

Private Sub KpiTile(ByVal cell As Range, ByVal label As String, ByVal value As Variant, ByVal numberFormat As String)
    With cell.Resize(1, 2)
        .Merge
        .Value = label
        .Font.Color = RGB(89, 89, 89)
        .Font.Bold = True
    End With
    With cell.Offset(1, 0).Resize(1, 2)
        .Merge
        .Value = value
        .NumberFormat = numberFormat
        .Font.Size = 18
        .Font.Bold = True
        .Font.Color = NAVY
    End With
    With cell.Resize(2, 2)
        .Interior.Color = LIGHT_BLUE
        .HorizontalAlignment = xlCenter
        .BorderAround xlContinuous, xlThin, , NAVY
    End With
    cell.Offset(1, 0).RowHeight = 30
End Sub

Private Sub SectionTitle(ByVal cell As Range, ByVal title As String)
    cell.Value = title
    cell.Font.Bold = True
    cell.Font.Size = 12
    cell.Font.Color = NAVY
End Sub

Private Sub StyleTable(ByVal table As Range)
    With table.Rows(1)
        .Font.Bold = True
        .Font.Color = vbWhite
        .Interior.Color = NAVY
        .HorizontalAlignment = xlCenter
    End With
    With table.Borders
        .LineStyle = xlContinuous
        .Color = GREY
    End With
End Sub

Private Sub AddChart(ByVal ws As Worksheet, ByVal title As String, ByVal categories As Range, _
                     ByVal values As Range, ByVal position As Range, ByVal chartType As XlChartType)
    With ws.ChartObjects.Add(position.Left, position.Top, 380, 230).Chart
        .ChartType = chartType
        With .SeriesCollection.NewSeries
            .Name = title
            .XValues = categories
            .Values = values
            .HasDataLabels = True
            .DataLabels.NumberFormat = "#,##0"
            If chartType = xlLineMarkers Then
                .Format.Line.ForeColor.RGB = RGB(31, 78, 121)
                .MarkerBackgroundColor = RGB(31, 78, 121)
                .DataLabels.Position = xlLabelPositionAbove
            Else
                .Format.Fill.ForeColor.RGB = RGB(31, 78, 121)
            End If
        End With
        .HasTitle = True
        .ChartTitle.Text = title
        .HasLegend = False
        .Axes(xlValue).TickLabels.NumberFormat = "#,##0"
        If chartType = xlBarClustered Then .Axes(xlCategory).ReversePlotOrder = True
    End With
End Sub
