# Builds Employee_Payroll_UDF.xlsm from employee_payroll_50.csv and PayrollFunctions.bas.
# Requires desktop Excel. VBA project access is enabled only for the build, then restored.
$ErrorActionPreference = 'Stop'
$folder = $PSScriptRoot
$output = Join-Path $folder 'Employee_Payroll_UDF.xlsm'
Add-Type -Namespace Native -Name Win -MemberDefinition '[DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint pid);'
$securityKey = 'HKCU:\Software\Microsoft\Office\16.0\Excel\Security'
$originalAccess = (Get-ItemProperty $securityKey -Name AccessVBOM -ErrorAction SilentlyContinue).AccessVBOM

$navy = 7949855          # RGB(31, 78, 121)
$lightBlue = 16247773    # RGB(221, 235, 247)
$money = '#,##0.00'

function To-Grid($rows) {
    $grid = New-Object 'object[,]' $rows.Count, $rows[0].Count
    for ($r = 0; $r -lt $rows.Count; $r++) {
        for ($c = 0; $c -lt $rows[0].Count; $c++) { $grid[$r, $c] = $rows[$r][$c] }
    }
    return , $grid
}
function Set-Header($range) {
    $range.Font.Bold = $true
    $range.Font.Color = 16777215
    $range.Interior.Color = $navy
    $range.HorizontalAlignment = -4108
    $range.WrapText = $true
}
function Set-Borders($range) {
    $range.Borders.LineStyle = 1
    $range.Borders.Color = 12566463   # RGB(191, 191, 191)
}

$employees = @(Import-Csv (Join-Path $folder 'employee_payroll_50.csv'))
$count = $employees.Count
$lastRow = $count + 1
$totalRow = $lastRow + 1
$payroll = @(, @('Employee ID', 'Employee Name', 'Department', 'Job Title', 'Monthly Salary', 'Performance Rating'))
foreach ($e in $employees) {
    $payroll += , @([string]$e.'Employee ID', [string]$e.'Employee Name', [string]$e.Department,
                    [string]$e.'Job Title', [double]$e.'Monthly Salary', [int]$e.'Performance Rating')
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
    [void]$book.VBProject.VBComponents.Import((Join-Path $folder 'PayrollFunctions.bas'))

    # --- Payroll sheet ---------------------------------------------------------------------
    $ws = $book.Worksheets.Item(1)
    $ws.Name = 'Payroll'
    $ws.Range("A1:F$lastRow").Value2 = To-Grid $payroll
    $ws.Range('G1:I1').Value2 = To-Grid @(, @('Bonus', 'Deductions', 'Net Pay'))
    $ws.Range("G2:G$lastRow").Formula = '=CalculateBonus(E2,F2)'
    $ws.Range("H2:H$lastRow").Formula = '=CalculateDeductions(E2)'
    $ws.Range("I2:I$lastRow").Formula = '=E2+G2-H2'
    $ws.Range("A$totalRow").Value2 = 'TOTAL'
    foreach ($col in 'E', 'G', 'H', 'I') { $ws.Range("$col$totalRow").Formula = "=SUM(${col}2:$col$lastRow)" }
    Set-Header $ws.Range('A1:I1')
    $ws.Rows(1).RowHeight = 30
    $ws.Range("E2:E$totalRow").NumberFormat = $money
    $ws.Range("G2:I$totalRow").NumberFormat = $money
    $ws.Range("F2:F$lastRow").HorizontalAlignment = -4108
    $ws.Range("A2:A$lastRow").HorizontalAlignment = -4108
    $rating = $ws.Range("F2:F$lastRow").Validation
    $rating.Delete()
    $rating.Add(1, 1, 1, '1', '5')    # whole number between 1 and 5
    $rating.ErrorMessage = 'Performance Rating must be a whole number from 1 to 5.'
    Set-Borders $ws.Range("A1:I$lastRow")
    $ws.Range("A${totalRow}:I$totalRow").Font.Bold = $true
    $ws.Range("A${totalRow}:I$totalRow").Interior.Color = $lightBlue
    $ws.Range("A${totalRow}:I$totalRow").Borders.Item(8).LineStyle = -4119   # double top border
    $ws.Columns('A:I').AutoFit() | Out-Null
    $ws.Columns('E:I').ColumnWidth = 14

    # --- Company Rules sheet ---------------------------------------------------------------
    $rules = $book.Worksheets.Add([Type]::Missing, $ws)
    $rules.Name = 'Company Rules'
    $rules.Range('A1').Value2 = 'Performance Bonus  -  CalculateBonus(Salary, Rating)'
    $rules.Range('A2:C7').Value2 = To-Grid @(
        @('Rating', 'Meaning', 'Bonus (% of monthly salary)'),
        @(5, 'Outstanding', 0.2), @(4, 'Exceeds Expectations', 0.15), @(3, 'Meets Expectations', 0.1),
        @(2, 'Needs Improvement', 0.05), @(1, 'Unsatisfactory', 0))
    $rules.Range('C3:C7').NumberFormat = '0%'
    $rules.Range('A9').Value2 = 'Monthly Deductions  -  CalculateDeductions(Salary)'
    $rules.Range('A10:C14').Value2 = To-Grid @(
        @('Deduction', 'Rule', 'Example: salary 6,000'),
        @('Provident Fund', '12% of monthly salary', 720),
        @('Health Insurance', 'Flat 150 per month', 150),
        @('Income Tax', '10% of salary above 5,000', 100),
        @('Total Deductions', 'Sum of the above', 970))
    $rules.Range('A16').Value2 = 'Net Pay = Monthly Salary + Bonus - Deductions.  Invalid salary or rating returns #VALUE!.'
    $rules.Range('C11:C14').NumberFormat = $money
    foreach ($title in 'A1', 'A9') { $rules.Range($title).Font.Bold = $true; $rules.Range($title).Font.Size = 12 }
    foreach ($header in 'A2:C2', 'A10:C10') { Set-Header $rules.Range($header) }
    Set-Borders $rules.Range('A2:C7')
    Set-Borders $rules.Range('A10:C14')
    $rules.Range('A14:C14').Font.Bold = $true
    $rules.Range('A3:A7').HorizontalAlignment = -4108
    $rules.Columns('A:C').ColumnWidth = 28

    # --- Validation sheet (manually verified examples) -------------------------------------
    $check = $book.Worksheets.Add([Type]::Missing, $rules)
    $check.Name = 'Validation'
    $cases = @(
        @('Case', 'Scenario', 'Salary', 'Rating', 'Expected Bonus', 'UDF Bonus', 'Expected Deductions', 'UDF Deductions', 'Result', 'Manual working'),
        @(1, 'Top rating, below tax limit', 4000, 5, 800, $null, 630, $null, $null, 'Bonus 4,000 x 20% = 800 | 480 + 150 + 0 = 630'),
        @(2, 'Salary exactly at tax limit', 5000, 3, 500, $null, 750, $null, $null, 'Bonus 5,000 x 10% = 500 | 600 + 150 + 0 = 750'),
        @(3, 'Salary above tax limit', 6000, 4, 900, $null, 970, $null, $null, 'Bonus 6,000 x 15% = 900 | 720 + 150 + 100 = 970'),
        @(4, 'Needs improvement', 10000, 2, 500, $null, 1850, $null, $null, 'Bonus 10,000 x 5% = 500 | 1,200 + 150 + 500 = 1,850'),
        @(5, 'Unsatisfactory - no bonus', 8500, 1, 0, $null, 1520, $null, $null, 'Bonus 8,500 x 0% = 0 | 1,020 + 150 + 350 = 1,520'),
        @(6, 'Outstanding, high salary', 12000, 5, 2400, $null, 2290, $null, $null, 'Bonus 12,000 x 20% = 2,400 | 1,440 + 150 + 700 = 2,290'),
        @(7, 'Bonus with cents', 3250, 4, 487.5, $null, 540, $null, $null, 'Bonus 3,250 x 15% = 487.50 | 390 + 150 + 0 = 540'),
        @(8, 'Invalid rating (6)', 7000, 6, "'#VALUE!", $null, 1190, $null, $null, 'Rating outside 1-5 | 840 + 150 + 200 = 1,190'),
        @(9, 'Invalid salary (0)', 0, 3, "'#VALUE!", $null, "'#VALUE!", $null, $null, 'Salary must be greater than 0')
    )
    $check.Range('A1:J10').Value2 = To-Grid $cases
    $check.Range('F2:F10').Formula = '=CalculateBonus(C2,D2)'
    $check.Range('H2:H10').Formula = '=CalculateDeductions(C2)'
    $check.Range('I2:I10').Formula = '=IF(AND(IF(ISERROR(F2),AND(E2="#VALUE!",ERROR.TYPE(F2)=3),F2=E2),' +
                                     'IF(ISERROR(H2),AND(G2="#VALUE!",ERROR.TYPE(H2)=3),H2=G2)),"PASS","FAIL")'
    $check.Range('A12').Formula = '=COUNTIF(I2:I10,"PASS")&" of "&COUNT(A2:A10)&" manually verified cases passed"'
    $check.Range('A12').Font.Bold = $true
    Set-Header $check.Range('A1:J1')
    Set-Borders $check.Range('A1:J10')
    $check.Range('C2:C10').NumberFormat = $money
    $check.Range('E2:H10').NumberFormat = $money
    $check.Range('E2:H10').HorizontalAlignment = -4152
    $check.Range('A2:A10').HorizontalAlignment = -4108
    $check.Range('D2:D10').HorizontalAlignment = -4108
    $check.Range('I2:I10').HorizontalAlignment = -4108
    $check.Range('I2:I10').Font.Bold = $true
    $pass = $check.Range('I2:I10').FormatConditions.Add(1, 3, '="PASS"')
    $pass.Font.Color = 24832          # dark green
    $pass.Interior.Color = 13561798   # light green
    $fail = $check.Range('I2:I10').FormatConditions.Add(1, 3, '="FAIL"')
    $fail.Font.Color = 393372         # dark red
    $fail.Interior.Color = 13551615   # light red
    $check.Columns('A:J').AutoFit() | Out-Null
    $check.Columns('E:H').ColumnWidth = 13
    $check.Rows(1).RowHeight = 30

    $excel.CalculateFull()
    foreach ($sheet in $ws, $rules, $check) {
        $sheet.Activate()
        $excel.ActiveWindow.DisplayGridlines = $false
        $sheet.Range('A1').Select() | Out-Null
    }
    $ws.Activate()
    $excel.ActiveWindow.SplitRow = 1
    $excel.ActiveWindow.FreezePanes = $true
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
