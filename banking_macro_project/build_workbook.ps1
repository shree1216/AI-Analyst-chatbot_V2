# Builds Banking_Transactions_Report.xlsm from banking_transactions_100.csv and the .bas modules.
# Requires desktop Excel. VBA project access is enabled only for the build, then restored.
$ErrorActionPreference = 'Stop'
$folder = $PSScriptRoot
$output = Join-Path $folder 'Banking_Transactions_Report.xlsm'
Add-Type -Namespace Native -Name Win -MemberDefinition '[DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint pid);'
$securityKey = 'HKCU:\Software\Microsoft\Office\16.0\Excel\Security'
$originalAccess = (Get-ItemProperty $securityKey -Name AccessVBOM -ErrorAction SilentlyContinue).AccessVBOM

function Set-Values($range, $values) { $range.Value2 = $values }

$rows = @(Import-Csv (Join-Path $folder 'banking_transactions_100.csv'))
$headers = $rows[0].PSObject.Properties.Name
$data = New-Object 'object[,]' ($rows.Count + 1), $headers.Count
for ($c = 0; $c -lt $headers.Count; $c++) { $data[0, $c] = [string]$headers[$c] }
for ($r = 0; $r -lt $rows.Count; $r++) {
    for ($c = 0; $c -lt $headers.Count; $c++) {
        $value = [string]$rows[$r].($headers[$c])
        switch ($headers[$c]) {
            'Date' { $value = ([datetime]::ParseExact($value, 'yyyy-MM-dd', $null)).ToOADate() }
            { $_ -in 'Debit', 'Credit' } { if ($value) { $value = [double]$value } else { $value = $null } }
        }
        $data[($r + 1), $c] = $value
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

    $sheet = $book.Worksheets.Item(1)
    $sheet.Name = 'Transactions'
    $sheet.Range('C:C').NumberFormat = '@'
    $sheet.Range('B:B').NumberFormat = 'yyyy-mm-dd'
    Set-Values $sheet.Range($sheet.Cells(1, 1), $sheet.Cells($rows.Count + 1, $headers.Count)) $data
    $sheet.Range('A:H').EntireColumn.AutoFit() | Out-Null
    $button = $sheet.Buttons().Add($sheet.Range('K2').Left, $sheet.Range('K2').Top, 130, 30)
    $button.Caption = 'Format Report'
    $button.OnAction = 'FormatTransactionReport'
    $button.Placement = 3

    $log = $book.Worksheets.Add([Type]::Missing, $sheet)
    $log.Name = 'Troubleshooting Log'
    $issues = New-Object 'object[,]' 6, 4
    $logRows = @(
        @('#', 'Recorded code (problem)', 'What went wrong', 'Fix in BankingReport macro'),
        @(1, 'Sheets("Sheet1").Select', 'Run-time error 9 (Subscript out of range) after the sheet was renamed', 'Works on ActiveSheet - any sheet name'),
        @(2, 'Range("B2:B101"), AutoFill to I101', 'New rows beyond row 101 were not formatted or calculated', 'Last row found with End(xlUp)'),
        @(3, 'NumberFormat = "0.00%"', 'Amounts displayed as percentages (e.g. 325166.00%)', 'Currency style "#,##0.00", negatives in red'),
        @(4, '=RC[-2]+RC[-1]', 'Net Amount added Debit and Credit; also broke when columns moved', 'Net = Credit - Debit, columns found by header name'),
        @(5, 'Range("A102") ... SUM(R[-100]C:R[-1]C)', 'Totals overwrote data when rows were added; rerun duplicated totals', 'Totals placed below last row; old totals row removed on rerun')
    )
    for ($r = 0; $r -lt 6; $r++) { for ($c = 0; $c -lt 4; $c++) { $issues[$r, $c] = $logRows[$r][$c] -as [object] } }
    Set-Values $log.Range('A1:D6') $issues
    $log.Range('A1:D1').Font.Bold = $true
    $log.Range('A1:D1').Font.Color = 16777215
    $log.Range('A1:D1').Interior.Color = 7949855
    $log.Range('A1:D6').Borders.LineStyle = 1
    $log.Range('A1:A6').HorizontalAlignment = -4108
    $log.Range('A:D').EntireColumn.AutoFit() | Out-Null

    foreach ($module in 'BankingReport.bas', 'RecordedMacro.bas') {
        [void]$book.VBProject.VBComponents.Import((Join-Path $folder $module))
    }
    $sheet.Activate()
    $sheet.Range('A1').Select() | Out-Null
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
