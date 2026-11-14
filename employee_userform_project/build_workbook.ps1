# Builds Employee_Management_System.xlsm: Employees table, Dashboard, modEmployees and the frmEmployee UserForm.
# Requires desktop Excel. VBA project access is enabled only for the build, then restored.
$ErrorActionPreference = 'Stop'
$folder = $PSScriptRoot
$output = Join-Path $folder 'Employee_Management_System.xlsm'
Add-Type -Namespace Native -Name Win -MemberDefinition '[DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint pid);'
$securityKey = 'HKCU:\Software\Microsoft\Office\16.0\Excel\Security'
$originalAccess = (Get-ItemProperty $securityKey -Name AccessVBOM -ErrorAction SilentlyContinue).AccessVBOM

$navy = 7949855          # RGB(31, 78, 121)
$lightBlue = 16247773    # RGB(221, 235, 247)
$grey = 5855577          # RGB(89, 89, 89)

$rows = @(Import-Csv (Join-Path $folder 'employees_100.csv'))
$headers = $rows[0].PSObject.Properties.Name
$grid = New-Object 'object[,]' ($rows.Count + 1), $headers.Count
for ($c = 0; $c -lt $headers.Count; $c++) { $grid[0, $c] = [string]$headers[$c] }
for ($r = 0; $r -lt $rows.Count; $r++) {
    for ($c = 0; $c -lt $headers.Count; $c++) {
        $value = [string]$rows[$r].($headers[$c])
        if ($headers[$c] -eq 'Hire Date') { $value = ([datetime]::ParseExact($value, 'yyyy-MM-dd', $null)).ToOADate() }
        elseif ($headers[$c] -eq 'Annual Salary') { $value = [double]$value }
        $grid[($r + 1), $c] = $value
    }
}

$excel = $null
try {
    Set-ItemProperty $securityKey -Name AccessVBOM -Value 1 -Type DWord
    $excel = New-Object -ComObject Excel.Application
    $excelPid = [uint32]0
    [void][Native.Win]::GetWindowThreadProcessId([IntPtr]$excel.Hwnd, [ref]$excelPid)
    $excel.DisplayAlerts = $false
    $book = $excel.Workbooks.Add()
    while ($book.Worksheets.Count -gt 1) { $book.Worksheets.Item(2).Delete() }

    # --- Employees sheet (Excel table) --------------------------------------------------------
    $dash = $book.Worksheets.Item(1)
    $dash.Name = 'Dashboard'
    $ws = $book.Worksheets.Add([Type]::Missing, $dash)
    $ws.Name = 'Employees'
    $ws.Range('A:G').NumberFormat = '@'
    $ws.Range('J:J').NumberFormat = '@'
    $ws.Range($ws.Cells(1, 1), $ws.Cells($rows.Count + 1, $headers.Count)).Value2 = $grid
    $table = $ws.ListObjects.Add(1, $ws.Range($ws.Cells(1, 1), $ws.Cells($rows.Count + 1, $headers.Count)), $null, 1)
    $table.Name = 'tblEmployees'
    $table.TableStyle = 'TableStyleMedium2'
    $ws.Range('H:H').NumberFormat = 'dd-mmm-yyyy'
    $ws.Range('I:I').NumberFormat = '#,##0'
    $ws.Range('A:A').HorizontalAlignment = -4108
    $ws.Columns('A:J').AutoFit() | Out-Null
    $ws.Columns('H:I').ColumnWidth = 14
    $ws.Activate()
    $excel.ActiveWindow.SplitRow = 1
    $excel.ActiveWindow.FreezePanes = $true

    # --- VBA: module + UserForm ---------------------------------------------------------------
    $project = $book.VBProject
    [void]$project.VBComponents.Import((Join-Path $folder 'modEmployees.bas'))
    $form = $project.VBComponents.Add(3)
    $form.Name = 'frmEmployee'
    # Lay out the form controls with a temporary VBA helper, then remove the helper.
    $helper = $project.VBComponents.Import((Join-Path $folder 'BuildForm.bas'))
    [void]$excel.Run("'$($book.Name)'!BuildForm.BuildEmployeeFormLayout")
    $project.VBComponents.Remove($helper)

    $code = $form.CodeModule
    if ($code.CountOfLines -gt 0) { $code.DeleteLines(1, $code.CountOfLines) }
    $code.AddFromString([IO.File]::ReadAllText((Join-Path $folder 'frmEmployee_code.txt')))
    $form.Export((Join-Path $folder 'frmEmployee.frm'))

    # --- Dashboard ----------------------------------------------------------------------------
    $dash.Activate()
    $excel.ActiveWindow.DisplayGridlines = $false
    $dash.Columns('A').ColumnWidth = 2
    $dash.Columns('B').ColumnWidth = 22
    $dash.Columns('C:I').ColumnWidth = 15
    $dash.Range('B2').Value2 = 'Employee Management System'
    $dash.Range('B2').Font.Size = 20; $dash.Range('B2').Font.Bold = $true; $dash.Range('B2').Font.Color = $navy
    $dash.Range('B3').Value2 = 'Human Resources  |  Add, search, update and delete employee records through the Employee Manager form'
    $dash.Range('B3').Font.Italic = $true; $dash.Range('B3').Font.Color = $grey
    $open = $dash.Buttons().Add($dash.Range('B5').Left, $dash.Range('B5').Top, 230, 36)
    $open.Caption = 'Open Employee Manager'
    $open.OnAction = 'OpenEmployeeForm'
    $open.Font.Bold = $true; $open.Font.Size = 12
    $open.Placement = 3

    $tiles = @(
        @('B', 'Total Employees', '=COUNTA(tblEmployees[Employee ID])', '#,##0'),
        @('D', 'Active', '=COUNTIF(tblEmployees[Status],"Active")', '#,##0'),
        @('F', 'On Leave', '=COUNTIF(tblEmployees[Status],"On Leave")', '#,##0'),
        @('H', 'Average Salary', '=IFERROR(AVERAGE(tblEmployees[Annual Salary]),0)', '#,##0')
    )
    foreach ($tile in $tiles) {
        $label = $dash.Range("$($tile[0])9").Resize(1, 2)
        $value = $dash.Range("$($tile[0])10").Resize(1, 2)
        $label.Merge(); $label.Cells.Item(1, 1).Value2 = $tile[1]
        $label.Font.Bold = $true; $label.Font.Color = $grey
        $value.Merge(); $value.Cells.Item(1, 1).Formula = $tile[2]
        $value.NumberFormat = $tile[3]; $value.Font.Size = 20; $value.Font.Bold = $true; $value.Font.Color = $navy
        $box = $dash.Range("$($tile[0])9").Resize(2, 2)
        $box.Interior.Color = $lightBlue; $box.HorizontalAlignment = -4108
        [void]$box.BorderAround(1, 2, [Type]::Missing, $navy)
    }
    $dash.Rows(10).RowHeight = 32

    $dash.Range('B13').Value2 = 'Headcount by Department'
    $dash.Range('B13').Font.Bold = $true; $dash.Range('B13').Font.Size = 12; $dash.Range('B13').Font.Color = $navy
    $dash.Range('B14:E14').Value2 = [object[]]@('Department', 'Employees', 'Active', 'Average Salary')
    $departments = 'Finance', 'Human Resources', 'IT', 'Marketing', 'Operations', 'Sales'
    for ($i = 0; $i -lt $departments.Count; $i++) {
        $r = 15 + $i
        $dash.Range("B$r").Value2 = $departments[$i]
        $dash.Range("C$r").Formula = "=COUNTIF(tblEmployees[Department],B$r)"
        $dash.Range("D$r").Formula = "=COUNTIFS(tblEmployees[Department],B$r,tblEmployees[Status],""Active"")"
        $dash.Range("E$r").Formula = "=IFERROR(AVERAGEIF(tblEmployees[Department],B$r,tblEmployees[Annual Salary]),0)"
    }
    $dash.Range('B21').Value2 = 'Total'
    $dash.Range('C21').Formula = '=SUM(C15:C20)'
    $dash.Range('D21').Formula = '=SUM(D15:D20)'
    $dash.Range('E21').Formula = '=IFERROR(AVERAGE(tblEmployees[Annual Salary]),0)'
    $dash.Range('E15:E21').NumberFormat = '#,##0'
    $header = $dash.Range('B14:E14')
    $header.Font.Bold = $true; $header.Font.Color = 16777215; $header.Interior.Color = $navy; $header.HorizontalAlignment = -4108
    $dash.Range('B14:E21').Borders.LineStyle = 1
    $dash.Range('B14:E21').Borders.Color = 12566463
    $dash.Range('B21:E21').Font.Bold = $true
    $dash.Range('B21:E21').Interior.Color = $lightBlue

    $dash.Range('G13').Value2 = 'How to use'
    $dash.Range('G13').Font.Bold = $true; $dash.Range('G13').Font.Size = 12; $dash.Range('G13').Font.Color = $navy
    $steps = @(
        '1. Click Open Employee Manager.',
        '2. Add: fill in the fields and click Add.',
        '3. Search: type an ID (EMP001) or name and press Enter.',
        '4. Update / Delete: search first, then edit or delete.',
        '5. Clear starts a new entry; Close exits the form.',
        'Fields marked * are required and validated before saving.'
    )
    for ($i = 0; $i -lt $steps.Count; $i++) { $dash.Range("G$(14 + $i)").Value2 = $steps[$i] }
    $dash.Range('G19').Font.Italic = $true; $dash.Range('G19').Font.Color = $grey
    [void]$dash.Hyperlinks.Add($dash.Range('B23'), '', "'Employees'!A1", 'Open the employee list', 'View all employee records >')
    $dash.Range('A1').Select() | Out-Null

    if (Test-Path $output) { Remove-Item $output }
    $book.SaveAs($output, 52)
    $book.Close($false)
    Write-Output "Saved $output"
}
finally {
    if ($excel) {
        $excel.Quit()
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
        $excel = $null
        [GC]::Collect(); [GC]::WaitForPendingFinalizers()
        # Excel saves Trust Center settings on exit, so restore only after it has closed.
        Wait-Process -Id $excelPid -Timeout 15 -ErrorAction SilentlyContinue
        Stop-Process -Id $excelPid -ErrorAction SilentlyContinue
    }
    if ($null -eq $originalAccess) { Remove-ItemProperty $securityKey -Name AccessVBOM -ErrorAction SilentlyContinue }
    else { Set-ItemProperty $securityKey -Name AccessVBOM -Value $originalAccess -Type DWord }
}
