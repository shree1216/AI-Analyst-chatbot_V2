# Builds Customer_Segmentation.xlsm from bank_customers_raw_100.csv and CustomerProcessing.bas.
# Requires desktop Excel. VBA project access is enabled only for the build, then restored.
$ErrorActionPreference = 'Stop'
$folder = $PSScriptRoot
$output = Join-Path $folder 'Customer_Segmentation.xlsm'
Add-Type -Namespace Native -Name Win -MemberDefinition '[DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint pid);'
$securityKey = 'HKCU:\Software\Microsoft\Office\16.0\Excel\Security'
$originalAccess = (Get-ItemProperty $securityKey -Name AccessVBOM -ErrorAction SilentlyContinue).AccessVBOM

$rows = @(Import-Csv (Join-Path $folder 'bank_customers_raw_100.csv'))
$headers = $rows[0].PSObject.Properties.Name
$grid = New-Object 'object[,]' ($rows.Count + 1), $headers.Count
for ($c = 0; $c -lt $headers.Count; $c++) { $grid[0, $c] = [string]$headers[$c] }
for ($r = 0; $r -lt $rows.Count; $r++) {
    for ($c = 0; $c -lt $headers.Count; $c++) {
        $value = [string]$rows[$r].($headers[$c])
        $number = 0.0
        if ($headers[$c] -in 'Age', 'Loan Amount' -and [double]::TryParse($value, [ref]$number)) { $grid[($r + 1), $c] = $number }
        else { $grid[($r + 1), $c] = $value }
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
    [void]$book.VBProject.VBComponents.Import((Join-Path $folder 'CustomerProcessing.bas'))

    $raw = $book.Worksheets.Item(1)
    $raw.Name = 'Raw Data'
    foreach ($col in 'A', 'B', 'D', 'E', 'F', 'G', 'H') { $raw.Columns($col).NumberFormat = '@' }
    $raw.Range($raw.Cells(1, 1), $raw.Cells($rows.Count + 1, $headers.Count)).Value2 = $grid
    $raw.Range('A1:I1').Font.Bold = $true
    $raw.Range('A1:I1').Interior.Color = 14277081   # light grey: raw extract, not yet processed
    $raw.Columns('A:I').AutoFit() | Out-Null
    $button = $raw.Buttons().Add($raw.Range('K2').Left, $raw.Range('K2').Top, 150, 32)
    $button.Caption = 'Process Customers'
    $button.OnAction = 'ProcessCustomers'
    $button.Placement = 3

    # Run once so the delivered workbook already shows the results.
    [void]$excel.Run("'$($book.Name)'!CustomerProcessing.ProcessCustomerData")
    $book.Worksheets.Item('Summary').Move($raw)
    foreach ($sheet in $book.Worksheets) {
        $sheet.Activate()
        $excel.ActiveWindow.DisplayGridlines = ($sheet.Name -eq 'Raw Data')
        if ($sheet.Name -ne 'Summary') {
            $excel.ActiveWindow.SplitRow = 1
            $excel.ActiveWindow.FreezePanes = $true
        }
        $sheet.Range('A1').Select() | Out-Null
    }
    $book.Worksheets.Item('Summary').Activate()
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
