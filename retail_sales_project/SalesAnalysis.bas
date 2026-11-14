Attribute VB_Name = "SalesAnalysis"
Option Explicit

' Analyzes the "Sales Data" sheet and writes metrics, rankings and insights to "Sales Analysis".
' Sales Data columns: Order ID | Order Date | Product | Category | Region | Quantity | Unit Price | Sales Amount

Private Const DATA_SHEET As String = "Sales Data"
Private Const REPORT_SHEET As String = "Sales Analysis"
Private Const HIGH_VALUE As Double = 500      ' high-value: 500 or more
Private Const LOW_VALUE As Double = 100       ' low-value: below 100

Private Const COL_DATE As Long = 2
Private Const COL_PRODUCT As Long = 3
Private Const COL_REGION As Long = 5
Private Const COL_QUANTITY As Long = 6
Private Const COL_SALES As Long = 8
Private Const LAST_COL As Long = 8

Private Const MONEY As String = "#,##0.00"

' Button / Alt+F8 entry point.
Public Sub RunSalesAnalysis()
    On Error GoTo Failed
    MsgBox AnalyzeSales(), vbInformation, "Sales Analysis"
    Exit Sub
Failed:
    Application.ScreenUpdating = True
    MsgBox Err.Description, vbExclamation, "Sales Analysis"
End Sub

' Does the analysis and returns a short summary message.
Public Function AnalyzeSales() As String
    Dim ws As Worksheet, report As Worksheet, data As Variant
    Dim lastRow As Long, n As Long, r As Long, i As Long, outRow As Long
    Dim amount As Double, totalSales As Double, totalUnits As Double
    Dim highest As Double, lowest As Double, highSales As Double
    Dim highCount As Long, lowCount As Long, firstDate As Date, lastDate As Date
    Dim productSales As Object, productUnits As Object, regionSales As Object, regionCount As Object
    Dim products As Variant, regions As Variant, topCount As Long, topSales As Double
    Dim insights(1 To 5) As String, regionEnd As Long, insightRow As Long

    Set ws = ThisWorkbook.Worksheets(DATA_SHEET)
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRow < 2 Then Err.Raise vbObjectError + 1, , "No sales records found on '" & DATA_SHEET & "'."
    data = ws.Range("A2").Resize(lastRow - 1, LAST_COL).Value
    n = UBound(data, 1)

    Set productSales = CreateObject("Scripting.Dictionary")
    Set productUnits = CreateObject("Scripting.Dictionary")
    Set regionSales = CreateObject("Scripting.Dictionary")
    Set regionCount = CreateObject("Scripting.Dictionary")

    ' 1. Validate and calculate (nothing is changed until every row is valid)
    For r = 1 To n
        If Not IsDate(data(r, COL_DATE)) Then Err.Raise vbObjectError + 2, , "Row " & r + 1 & ": Order Date is not a valid date."
        If IsEmpty(data(r, COL_SALES)) Or Not IsNumeric(data(r, COL_SALES)) Then _
            Err.Raise vbObjectError + 3, , "Row " & r + 1 & ": Sales Amount is not a number."
        If IsEmpty(data(r, COL_QUANTITY)) Or Not IsNumeric(data(r, COL_QUANTITY)) Then _
            Err.Raise vbObjectError + 4, , "Row " & r + 1 & ": Quantity is not a number."

        amount = data(r, COL_SALES)
        totalSales = totalSales + amount
        totalUnits = totalUnits + data(r, COL_QUANTITY)
        If r = 1 Or amount > highest Then highest = amount
        If r = 1 Or amount < lowest Then lowest = amount
        If r = 1 Or data(r, COL_DATE) < firstDate Then firstDate = data(r, COL_DATE)
        If r = 1 Or data(r, COL_DATE) > lastDate Then lastDate = data(r, COL_DATE)
        If amount >= HIGH_VALUE Then
            highCount = highCount + 1
            highSales = highSales + amount
        ElseIf amount < LOW_VALUE Then
            lowCount = lowCount + 1
        End If

        AddTo productSales, data(r, COL_PRODUCT), amount
        AddTo productUnits, data(r, COL_PRODUCT), data(r, COL_QUANTITY)
        AddTo regionSales, data(r, COL_REGION), amount
        AddTo regionCount, data(r, COL_REGION), 1
    Next r

    products = SortedKeys(productSales)
    regions = SortedKeys(regionSales)
    topCount = Application.WorksheetFunction.Min(5, productSales.Count)

    Application.ScreenUpdating = False

    ' 2. Format the data sheet and highlight high / low value transactions
    With ws.Range(ws.Cells(2, 1), ws.Cells(lastRow, LAST_COL))
        .Interior.Pattern = xlNone
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(191, 191, 191)
    End With
    ws.Range(ws.Cells(2, COL_DATE), ws.Cells(lastRow, COL_DATE)).NumberFormat = "dd-mmm-yyyy"
    ws.Range(ws.Cells(2, COL_SALES - 1), ws.Cells(lastRow, COL_SALES)).NumberFormat = MONEY
    For r = 1 To n
        If data(r, COL_SALES) >= HIGH_VALUE Then
            ws.Range(ws.Cells(r + 1, 1), ws.Cells(r + 1, LAST_COL)).Interior.Color = RGB(198, 239, 206)
        ElseIf data(r, COL_SALES) < LOW_VALUE Then
            ws.Range(ws.Cells(r + 1, 1), ws.Cells(r + 1, LAST_COL)).Interior.Color = RGB(255, 199, 206)
        End If
    Next r

    ' 3. Write the report
    Set report = PrepareReportSheet()
    report.Range("A1").Value = "Retail Sales Analysis"
    report.Range("A2").Value = "Period: " & Format$(firstDate, "dd-mmm-yyyy") & " to " & Format$(lastDate, "dd-mmm-yyyy") & _
                               "     |     Generated: " & Format$(Now, "dd-mmm-yyyy hh:nn")

    ' Key metrics
    SectionTitle report.Range("A4"), "Key Metrics"
    report.Range("A5:B5").Value = Array("Metric", "Value")
    report.Range("A6:B6").Value = Array("Total Sales", totalSales)
    report.Range("A7:B7").Value = Array("Number of Transactions", n)
    report.Range("A8:B8").Value = Array("Average Sale per Transaction", totalSales / n)
    report.Range("A9:B9").Value = Array("Total Units Sold", totalUnits)
    report.Range("A10:B10").Value = Array("Highest Transaction", highest)
    report.Range("A11:B11").Value = Array("Lowest Transaction", lowest)
    report.Range("A12:B12").Value = Array("High-Value Transactions (>= " & HIGH_VALUE & ")", highCount)
    report.Range("A13:B13").Value = Array("Low-Value Transactions (< " & LOW_VALUE & ")", lowCount)
    report.Range("B6,B8,B10,B11").NumberFormat = MONEY
    report.Range("B7,B9,B12,B13").NumberFormat = "#,##0"
    StyleTable report.Range("A5:B13")
    report.Range("A12:B12").Interior.Color = RGB(198, 239, 206)
    report.Range("A13:B13").Interior.Color = RGB(255, 199, 206)

    ' Top 5 products
    SectionTitle report.Range("D4"), "Top 5 Products by Revenue"
    report.Range("D5:H5").Value = Array("Rank", "Product", "Units Sold", "Revenue", "% of Sales")
    For i = 1 To topCount
        outRow = 5 + i
        report.Cells(outRow, 4).Value = i
        report.Cells(outRow, 5).Value = products(i - 1)
        report.Cells(outRow, 6).Value = productUnits(products(i - 1))
        report.Cells(outRow, 7).Value = productSales(products(i - 1))
        report.Cells(outRow, 8).Value = productSales(products(i - 1)) / totalSales
        topSales = topSales + productSales(products(i - 1))
    Next i
    report.Range("G6:G10").NumberFormat = MONEY
    report.Range("H6:H10").NumberFormat = "0.0%"
    report.Range("D6:D10").HorizontalAlignment = xlCenter
    StyleTable report.Range(report.Cells(5, 4), report.Cells(5 + topCount, 8))

    ' Sales by region
    SectionTitle report.Range("D12"), "Sales by Region"
    report.Range("D13:H13").Value = Array("Region", "Transactions", "Revenue", "Average Sale", "% of Sales")
    For i = 0 To UBound(regions)
        outRow = 14 + i
        report.Cells(outRow, 4).Value = regions(i)
        report.Cells(outRow, 5).Value = regionCount(regions(i))
        report.Cells(outRow, 6).Value = regionSales(regions(i))
        report.Cells(outRow, 7).Value = regionSales(regions(i)) / regionCount(regions(i))
        report.Cells(outRow, 8).Value = regionSales(regions(i)) / totalSales
    Next i
    regionEnd = 14 + UBound(regions)
    report.Range("F14:G" & regionEnd).NumberFormat = MONEY
    report.Range("H14:H" & regionEnd).NumberFormat = "0.0%"
    StyleTable report.Range("D13:H" & regionEnd)

    ' Insights
    insights(1) = "Total sales of " & Format$(totalSales, MONEY) & " from " & n & " transactions, an average of " & _
                  Format$(totalSales / n, MONEY) & " per transaction."
    insights(2) = "Best-selling product: " & products(0) & " with " & Format$(productSales(products(0)), MONEY) & _
                  " (" & Format$(productSales(products(0)) / totalSales, "0.0%") & " of sales)."
    insights(3) = "The top " & topCount & " products generated " & Format$(topSales / totalSales, "0.0%") & " of total sales."
    insights(4) = "Strongest region: " & regions(0) & " (" & Format$(regionSales(regions(0)) / totalSales, "0.0%") & _
                  "); weakest region: " & regions(UBound(regions)) & " (" & _
                  Format$(regionSales(regions(UBound(regions))) / totalSales, "0.0%") & ")."
    insights(5) = highCount & " high-value transactions (>= " & HIGH_VALUE & ") contributed " & _
                  Format$(highSales / totalSales, "0.0%") & " of sales; " & lowCount & _
                  " transactions were low-value (< " & LOW_VALUE & ")."
    insightRow = Application.WorksheetFunction.Max(16, regionEnd + 2)
    SectionTitle report.Cells(insightRow, 1), "Key Insights"
    For i = 1 To 5
        report.Cells(insightRow + i, 1).Value = ChrW(8226) & "  " & insights(i)
    Next i

    ' Charts
    AddChart report, "Top 5 Products by Revenue", report.Range(report.Cells(6, 5), report.Cells(5 + topCount, 5)), _
             report.Range(report.Cells(6, 7), report.Cells(5 + topCount, 7)), report.Cells(insightRow + 7, 1), xlBarClustered
    AddChart report, "Sales by Region", report.Range("D14:D" & regionEnd), report.Range("F14:F" & regionEnd), _
             report.Cells(insightRow + 7, 5), xlColumnClustered

    report.Activate
    report.Range("A1").Select
    Application.ScreenUpdating = True

    AnalyzeSales = "Sales analysis complete." & vbNewLine & vbNewLine & _
                   "Total sales: " & Format$(totalSales, MONEY) & vbNewLine & _
                   "Transactions: " & n & vbNewLine & _
                   "Top product: " & products(0) & vbNewLine & _
                   "Top region: " & regions(0)
End Function

' ---------- Helpers ------------------------------------------------------------------------

Private Sub AddTo(ByVal dict As Object, ByVal key As Variant, ByVal value As Double)
    dict(key) = dict(key) + value
End Sub

' Returns the dictionary keys sorted by value, highest first.
Private Function SortedKeys(ByVal dict As Object) As Variant
    Dim keys As Variant, i As Long, j As Long, temp As Variant
    keys = dict.Keys
    For i = 0 To UBound(keys) - 1
        For j = i + 1 To UBound(keys)
            If dict(keys(j)) > dict(keys(i)) Then
                temp = keys(i)
                keys(i) = keys(j)
                keys(j) = temp
            End If
        Next j
    Next i
    SortedKeys = keys
End Function

Private Function PrepareReportSheet() As Worksheet
    Dim ws As Worksheet, chartBox As ChartObject
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(REPORT_SHEET)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(Before:=ThisWorkbook.Worksheets(1))
        ws.Name = REPORT_SHEET
    End If
    ws.Cells.Clear
    For Each chartBox In ws.ChartObjects
        chartBox.Delete
    Next chartBox
    ws.Range("A1").Font.Size = 16
    ws.Range("A1").Font.Bold = True
    ws.Range("A1").Font.Color = RGB(31, 78, 121)
    ws.Range("A2").Font.Italic = True
    ws.Range("A2").Font.Color = RGB(89, 89, 89)
    ws.Columns("A").ColumnWidth = 34
    ws.Columns("B").ColumnWidth = 14
    ws.Columns("C").ColumnWidth = 3
    ws.Columns("D").ColumnWidth = 10
    ws.Columns("E").ColumnWidth = 22
    ws.Columns("F:H").ColumnWidth = 14
    Set PrepareReportSheet = ws
End Function

Private Sub SectionTitle(ByVal cell As Range, ByVal title As String)
    cell.Value = title
    cell.Font.Bold = True
    cell.Font.Size = 12
    cell.Font.Color = RGB(31, 78, 121)
End Sub

Private Sub StyleTable(ByVal table As Range)
    With table.Rows(1)
        .Font.Bold = True
        .Font.Color = vbWhite
        .Interior.Color = RGB(31, 78, 121)
        .HorizontalAlignment = xlCenter
    End With
    With table.Borders
        .LineStyle = xlContinuous
        .Color = RGB(191, 191, 191)
    End With
End Sub

Private Sub AddChart(ByVal ws As Worksheet, ByVal title As String, ByVal categories As Range, _
                     ByVal values As Range, ByVal position As Range, ByVal chartType As XlChartType)
    With ws.ChartObjects.Add(position.Left, position.Top, 360, 220).Chart
        .ChartType = chartType
        With .SeriesCollection.NewSeries
            .Name = title
            .XValues = categories
            .Values = values
            .HasDataLabels = True
            .DataLabels.NumberFormat = "#,##0"
            .Format.Fill.ForeColor.RGB = RGB(31, 78, 121)
        End With
        .HasTitle = True
        .ChartTitle.Text = title
        .HasLegend = False
        .Axes(xlValue).TickLabels.NumberFormat = "#,##0"
        If chartType = xlBarClustered Then .Axes(xlCategory).ReversePlotOrder = True
    End With
End Sub
