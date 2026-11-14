# Builds HR_Employee_Management.xlsm from employees.csv and the VBA source files in this folder.
# Requires Microsoft Excel (Windows). Run:  powershell -ExecutionPolicy Bypass -File build_workbook.ps1
#
# Creating VBA code from a script needs "Trust access to the VBA project object model".
# The script turns it on only for the build and restores the previous setting afterwards.
# Excel saves that setting when it exits, so the build runs in a child process and the
# setting is restored only after the build's own Excel instance has closed.
param([switch]$Inner)

$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $MyInvocation.MyCommand.Path
$output = Join-Path $root "HR_Employee_Management.xlsm"

if (-not $Inner) {
    $securityKey = "HKCU:\Software\Microsoft\Office\16.0\Excel\Security"
    if (-not (Test-Path $securityKey)) { New-Item -Path $securityKey -Force | Out-Null }
    $previousAccess = (Get-ItemProperty $securityKey -Name AccessVBOM -ErrorAction SilentlyContinue).AccessVBOM
    Set-ItemProperty $securityKey -Name AccessVBOM -Value 1 -Type DWord
    $exitCode = 1
    try {
        & powershell -NoProfile -ExecutionPolicy Bypass -File $MyInvocation.MyCommand.Path -Inner
        $exitCode = $LASTEXITCODE
    }
    finally {
        if ($null -eq $previousAccess) { Remove-ItemProperty $securityKey -Name AccessVBOM -ErrorAction SilentlyContinue }
        else { Set-ItemProperty $securityKey -Name AccessVBOM -Value $previousAccess -Type DWord }
    }
    exit $exitCode
}

# Colours (VBA/OLE colours are B*65536 + G*256 + R)
function Rgb($r, $g, $b) { return [int]($r + $g * 256 + $b * 65536) }
$WINE = Rgb 123 45 67
$WINE_SOFT = Rgb 246 233 237
$INK = Rgb 31 29 27
$MUTED = Rgb 107 101 96
$CANVAS = Rgb 250 248 246
$BAND = Rgb 243 239 235
$LINE = Rgb 231 226 221
$WHITE = Rgb 255 255 255
$OCHRE = Rgb 192 138 62

$excel = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $missing = [System.Reflection.Missing]::Value

    $wb = $excel.Workbooks.Add()
    while ($wb.Worksheets.Count -lt 3) { $null = $wb.Worksheets.Add($missing, $wb.Worksheets.Item($wb.Worksheets.Count)) }
    $dash = $wb.Worksheets.Item(1); $dash.Name = "Dashboard"
    $emp = $wb.Worksheets.Item(2); $emp.Name = "Employees"
    $lists = $wb.Worksheets.Item(3); $lists.Name = "Lists"
    foreach ($sheet in @($dash, $emp, $lists)) { $sheet.Cells.Font.Name = "Segoe UI"; $sheet.Cells.Font.Size = 10 }

    # ---------------------------------------------------------------- Lists
    $listData = @{
        "A" = @("Departments", "Finance", "Human Resources", "IT", "Marketing", "Operations", "Sales")
        "C" = @("Gender", "Female", "Male")
        "E" = @("Status", "Active", "On Leave", "Resigned")
    }
    foreach ($column in $listData.Keys) {
        $values = $listData[$column]
        for ($i = 0; $i -lt $values.Count; $i++) { $lists.Range("$column$($i + 1)").Value2 = $values[$i] }
        $header = $lists.Range("${column}1")
        $header.Font.Bold = $true; $header.Font.Color = $WHITE; $header.Interior.Color = $WINE
        $lists.Columns($column).ColumnWidth = 20
    }
    $null = $wb.Names.Add("DepartmentList", "=Lists!`$A`$2:`$A`$7")
    $null = $wb.Names.Add("GenderList", "=Lists!`$C`$2:`$C`$3")
    $null = $wb.Names.Add("StatusList", "=Lists!`$E`$2:`$E`$4")
    $lists.Range("G1").Value2 = "These lists feed the dropdowns on the Employee Manager form."
    $lists.Range("G1").Font.Color = $MUTED
    $lists.Tab.Color = $MUTED

    # ---------------------------------------------------------------- Employees
    $rows = @(Import-Csv (Join-Path $root "employees.csv"))
    $headers = @("Employee ID", "First Name", "Last Name", "Gender", "Department", "Job Title", "Email",
                 "Phone", "Hire Date", "Annual Salary", "Status")
    $emp.Range("H:H").NumberFormat = "@"
    for ($c = 0; $c -lt $headers.Count; $c++) { $emp.Cells(4, $c + 1).Value2 = [string]$headers[$c] }
    for ($r = 0; $r -lt $rows.Count; $r++) {
        for ($c = 0; $c -lt $headers.Count; $c++) {
            # One assignment line per value type: PowerShell caches the type used at each COM call site.
            $cell = $emp.Cells($r + 5, $c + 1)
            $text = [string]$rows[$r].($headers[$c])
            if ($c -eq 8) { $cell.Value2 = [double]([datetime]::ParseExact($text, "yyyy-MM-dd", $null)).ToOADate() }
            elseif ($c -eq 9) { $cell.Value2 = [double]$text }
            else { $cell.Value2 = $text }
        }
    }
    $tableRange = $emp.Range($emp.Cells(4, 1), $emp.Cells($rows.Count + 4, $headers.Count))
    $table = $emp.ListObjects.Add(1, $tableRange, $missing, 1)
    $table.Name = "tblEmployees"
    $table.TableStyle = "TableStyleLight1"
    $table.HeaderRowRange.Interior.Color = $WINE
    $table.HeaderRowRange.Font.Color = $WHITE
    $table.HeaderRowRange.Font.Bold = $true
    $table.HeaderRowRange.RowHeight = 22
    $table.ListColumns.Item(9).DataBodyRange.NumberFormat = "dd-mmm-yyyy"
    $table.ListColumns.Item(10).DataBodyRange.NumberFormat = "#,##0"
    $emp.Range("A1").Value2 = "Employee Records"
    $emp.Range("A1").Font.Size = 16; $emp.Range("A1").Font.Bold = $true; $emp.Range("A1").Font.Color = $WINE
    $emp.Range("A2").Value2 = "Records are added, updated and deleted through the Employee Manager form (Dashboard > Open Employee Manager)."
    $emp.Range("A2").Font.Color = $MUTED
    $null = $emp.Columns("A:K").AutoFit()
    $emp.Columns("A").ColumnWidth = 13
    $emp.Tab.Color = $OCHRE
    $emp.Activate()
    $excel.ActiveWindow.DisplayGridlines = $false
    $excel.ActiveWindow.SplitRow = 4
    $excel.ActiveWindow.FreezePanes = $true

    # ---------------------------------------------------------------- Dashboard
    $dash.Activate()
    $excel.ActiveWindow.DisplayGridlines = $false
    $dash.Columns("A").ColumnWidth = 2
    $dash.Columns("B:M").ColumnWidth = 12
    $dash.Cells.Interior.Color = $CANVAS
    $dash.Tab.Color = $WINE

    $banner = $dash.Range("B2:M3")
    $null = $banner.Merge()
    $banner.Value2 = "HR Employee Management Dashboard"
    $banner.Interior.Color = $WINE
    $banner.Font.Color = $WHITE; $banner.Font.Size = 20; $banner.Font.Bold = $true
    $banner.VerticalAlignment = -4108; $banner.IndentLevel = 1
    $dash.Rows("2:3").RowHeight = 24
    $dash.Range("B4").Value2 = "Live summary of the Employees table. Use the button to add, search, update or delete employees."
    $dash.Range("B4").Font.Color = $MUTED; $dash.Range("B4").Font.Italic = $true

    # KPI cards
    $cards = @(
        @("B", "D", "Total Employees", "=COUNTA(tblEmployees[Employee ID])", "#,##0"),
        @("E", "G", "Active", "=COUNTIF(tblEmployees[Status],""Active"")", "#,##0"),
        @("H", "J", "On Leave", "=COUNTIF(tblEmployees[Status],""On Leave"")", "#,##0"),
        @("K", "M", "Average Salary", "=IFERROR(AVERAGE(tblEmployees[Annual Salary]),0)", "#,##0")
    )
    foreach ($card in $cards) {
        $label = $dash.Range("$($card[0])9:$($card[1])9"); $null = $label.Merge()
        $value = $dash.Range("$($card[0])10:$($card[1])11"); $null = $value.Merge()
        $label.Value2 = $card[2].ToUpper()
        $label.Font.Size = 9; $label.Font.Bold = $true; $label.Font.Color = $MUTED; $label.IndentLevel = 1
        $value.Formula = $card[3]
        $value.NumberFormat = $card[4]
        $value.Font.Size = 22; $value.Font.Bold = $true; $value.Font.Color = $WINE
        $value.HorizontalAlignment = -4131; $value.VerticalAlignment = -4108; $value.IndentLevel = 1
        $box = $dash.Range("$($card[0])9:$($card[1])11")
        $box.Interior.Color = $WHITE
        $box.BorderAround(1, 2, $missing, $LINE) | Out-Null
    }

    # Department summary table
    $dash.Range("B13").Value2 = "Headcount by Department"
    $dash.Range("B13").Font.Size = 12; $dash.Range("B13").Font.Bold = $true; $dash.Range("B13").Font.Color = $WINE
    $summaryHeaders = @("Department", "", "Headcount", "Active", "Avg Salary")
    for ($c = 0; $c -lt $summaryHeaders.Count; $c++) { $dash.Cells(14, 2 + $c).Value2 = $summaryHeaders[$c] }
    $null = $dash.Range("B14:C14").Merge()
    $head = $dash.Range("B14:F14")
    $head.Interior.Color = $WINE; $head.Font.Color = $WHITE; $head.Font.Bold = $true
    $departments = $listData["A"][1..6]
    for ($i = 0; $i -lt $departments.Count; $i++) {
        $row = 15 + $i
        $null = $dash.Range("B${row}:C${row}").Merge()
        $dash.Range("B$row").Value2 = $departments[$i]
        $dash.Range("D$row").Formula = "=COUNTIFS(tblEmployees[Department],`$B$row)"
        $dash.Range("E$row").Formula = "=COUNTIFS(tblEmployees[Department],`$B$row,tblEmployees[Status],""Active"")"
        $dash.Range("F$row").Formula = "=IFERROR(AVERAGEIFS(tblEmployees[Annual Salary],tblEmployees[Department],`$B$row),0)"
        $fill = if ($i % 2 -eq 0) { $WHITE } else { $BAND }
        $dash.Range("B${row}:F${row}").Interior.Color = $fill
    }
    $total = 21
    $null = $dash.Range("B${total}:C${total}").Merge()
    $dash.Range("B$total").Value2 = "Total"
    $dash.Range("D$total").Formula = "=SUM(D15:D20)"
    $dash.Range("E$total").Formula = "=SUM(E15:E20)"
    $dash.Range("F$total").Formula = "=IFERROR(AVERAGE(tblEmployees[Annual Salary]),0)"
    $totalRow = $dash.Range("B${total}:F${total}")
    $totalRow.Font.Bold = $true; $totalRow.Interior.Color = $WINE_SOFT
    $dash.Range("D15:E21").NumberFormat = "#,##0"; $dash.Range("D15:E21").HorizontalAlignment = -4108
    $dash.Range("F15:F21").NumberFormat = "#,##0"
    $dash.Range("D14:F14").HorizontalAlignment = -4108

    # Status summary
    $dash.Range("B23").Value2 = "Employees by Status"
    $dash.Range("B23").Font.Size = 12; $dash.Range("B23").Font.Bold = $true; $dash.Range("B23").Font.Color = $WINE
    $dash.Range("B24").Value2 = "Status"; $dash.Range("D24").Value2 = "Employees"; $dash.Range("E24").Value2 = "Share"
    $null = $dash.Range("B24:C24").Merge()
    $statusHead = $dash.Range("B24:E24")
    $statusHead.Interior.Color = $WINE; $statusHead.Font.Color = $WHITE; $statusHead.Font.Bold = $true
    $statuses = $listData["E"][1..3]
    for ($i = 0; $i -lt $statuses.Count; $i++) {
        $row = 25 + $i
        $null = $dash.Range("B${row}:C${row}").Merge()
        $dash.Range("B$row").Value2 = $statuses[$i]
        $dash.Range("D$row").Formula = "=COUNTIF(tblEmployees[Status],`$B$row)"
        $dash.Range("E$row").Formula = "=IFERROR(D$row/COUNTA(tblEmployees[Employee ID]),0)"
        $fill = if ($i % 2 -eq 0) { $WHITE } else { $BAND }
        $dash.Range("B${row}:E${row}").Interior.Color = $fill
    }
    $dash.Range("D25:D27").HorizontalAlignment = -4108
    $dash.Range("E25:E27").NumberFormat = "0.0%"; $dash.Range("E25:E27").HorizontalAlignment = -4108

    # Headcount chart
    $chartArea = $dash.Range("H13:M27")
    $chartObject = $dash.ChartObjects().Add($chartArea.Left, $chartArea.Top, $chartArea.Width, $chartArea.Height)
    $chart = $chartObject.Chart
    $chart.ChartType = 57   # clustered bar
    $null = $chart.SetSourceData($dash.Range("B15:B20,D15:D20"))
    $chart.HasTitle = $true
    $chart.ChartTitle.Text = "Headcount by Department"
    $chart.ChartTitle.Font.Size = 12; $chart.ChartTitle.Font.Bold = $true; $chart.ChartTitle.Font.Color = $INK
    $chart.HasLegend = $false
    $chart.SeriesCollection(1).Format.Fill.ForeColor.RGB = $WINE
    $chart.SeriesCollection(1).HasDataLabels = $true
    $chart.Axes(1).ReversePlotOrder = $true
    $chart.Axes(2).HasMajorGridlines = $false
    $chartObject.Chart.ChartArea.Format.Line.ForeColor.RGB = $LINE

    # Button that opens the form
    $anchor = $dash.Range("B6")
    $button = $dash.Shapes.AddShape(5, $anchor.Left, $anchor.Top, 230, 34)   # rounded rectangle
    $button.Name = "btnOpenEmployeeManager"
    $button.Fill.ForeColor.RGB = $WINE
    $button.Line.Visible = 0
    $button.TextFrame2.TextRange.Text = "Open Employee Manager"
    $button.TextFrame2.TextRange.Font.Size = 12
    $button.TextFrame2.TextRange.Font.Bold = -1
    $button.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = $WHITE
    $button.TextFrame2.TextRange.Font.Name = "Segoe UI"
    $button.TextFrame2.VerticalAnchor = 3
    $button.TextFrame2.TextRange.ParagraphFormat.Alignment = 2
    $dash.Rows("6:7").RowHeight = 18

    # ---------------------------------------------------------------- VBA
    $project = $wb.VBProject
    $null = $project.VBComponents.Import((Join-Path $root "modEmployees.bas"))
    $workbookModule = $project.VBComponents.Item("ThisWorkbook").CodeModule
    if ($workbookModule.CountOfLines -gt 0) { $workbookModule.DeleteLines(1, $workbookModule.CountOfLines) }
    $workbookModule.AddFromString([IO.File]::ReadAllText((Join-Path $root "ThisWorkbook_code.txt")))
    $button.OnAction = "OpenEmployeeManager"

    # UserForm: a build-only VBA module designs the controls (VBA handles MSForms objects natively),
    # then it is removed so the finished workbook contains only modEmployees, frmEmployee and ThisWorkbook.
    $builder = $project.VBComponents.Import((Join-Path $root "modBuildForm.bas"))
    $excel.Run("BuildEmployeeForm", (Join-Path $root "frmEmployee_code.txt"))
    $project.VBComponents.Remove($builder)
    $form = $project.VBComponents.Item("frmEmployee")

    # Lock the data sheet, start on the Dashboard, save
    $excel.Run("LockEmployees")
    $tests = $project.VBComponents.Import((Join-Path $root "EmployeeWorkbookTests.bas"))
    Write-Host ($excel.Run("RunEmployeeWorkbookTests"))
    $project.VBComponents.Remove($tests)
    $dash.Activate()
    $null = $dash.Range("A1").Select()
    $wb.SaveAs($output, 52)
    $form.Export((Join-Path $root "frmEmployee.frm"))
    $wb.Close($false)
    Write-Host "Built $output"
}
finally {
    if ($excel) {
        try { foreach ($open in @($excel.Workbooks)) { $open.Close($false) } }
        finally {
            $excel.Quit()
            $null = [Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel)
            $excel = $null
            [GC]::Collect()
            [GC]::WaitForPendingFinalizers()
        }
    }
}
