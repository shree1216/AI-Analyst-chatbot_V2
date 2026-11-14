# Builds Retail_Sales_Analysis.xlsm from retail_sales_oct_2026.csv and SalesAnalysis.bas.
# Requires desktop Excel. VBA project access is enabled only for the build, then restored.
$ErrorActionPreference = 'Stop'
$folder = $PSScriptRoot
$output = Join-Path $folder 'Retail_Sales_Analysis.xlsm'
Add-Type -Namespace Native -Name Win -MemberDefinition '[DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint pid);'
$securityKey = 'HKCU:\Software\Microsoft\Office\16.0\Excel\Security'
$originalAccess = (Get-ItemProperty $securityKey -Name AccessVBOM -ErrorAction SilentlyContinue).AccessVBOM

$rows = @(Import-Csv (Join-Path $folder 'retail_sales_oct_2026.csv'))
$headers = $rows[0].PSObject.Properties.Name
$grid = New-Object 'object[,]' ($rows.Count + 1), $headers.Count
for ($c = 0; $c -lt $headers.Count; $c++) { $grid[0, $c] = [string]$headers[$c] }
for ($r = 0; $r -lt $rows.Count; $r++) {
    for ($c = 0; $c -lt $headers.Count; $c++) {
        $value = [string]$rows[$r].($headers[$c])
        switch ($headers[$c]) {
            'Order Date' { $value = ([datetime]::ParseExact($value, 'yyyy-MM-dd', $null)).ToOADate() }
            { $_ -in 'Quantity', 'Unit Price', 'Sales Amount' } { $value = [double]$value }
        }
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
    [void]$book.VBProject.VBComponents.Import((Join-Path $folder 'SalesAnalysis.bas'))

    $ws = $book.Worksheets.Item(1)
    $ws.Name = 'Sales Data'
    $ws.Range($ws.Cells(1, 1), $ws.Cells($rows.Count + 1, $headers.Count)).Value2 = $grid
    $header = $ws.Range('A1:H1')
    $header.Font.Bold = $true
    $header.Font.Color = 16777215
    $header.Interior.Color = 7949855      # RGB(31, 78, 121)
    $header.HorizontalAlignment = -4108
    $ws.Range('F:F').HorizontalAlignment = -4108
    $ws.Range('B:B').NumberFormat = 'dd-mmm-yyyy'
    $ws.Range('G:H').NumberFormat = '#,##0.00'
    $ws.Columns('A:H').AutoFit() | Out-Null
    $ws.Columns('A:H').ColumnWidth = 16
    $ws.Columns('C').ColumnWidth = 22
    $ws.Range('J4').Value2 = 'Legend'
    $ws.Range('J4').Font.Bold = $true
    $ws.Range('J5').Value2 = 'High-value (>= 500)'
    $ws.Range('J5').Interior.Color = 13561798   # RGB(198, 239, 206)
    $ws.Range('J6').Value2 = 'Low-value (< 100)'
    $ws.Range('J6').Interior.Color = 13551615   # RGB(255, 199, 206)
    $ws.Columns('J').ColumnWidth = 22
    $button = $ws.Buttons().Add($ws.Range('J1').Left, $ws.Range('J1').Top + 2, 160, 32)
    $button.Caption = 'Run Sales Analysis'
    $button.OnAction = 'RunSalesAnalysis'
    $button.Placement = 3
    $ws.Activate()
    $excel.ActiveWindow.SplitRow = 1
    $excel.ActiveWindow.FreezePanes = $true

    # Run once so the delivered workbook already shows the analysis.
    [void]$excel.Run("'$($book.Name)'!SalesAnalysis.AnalyzeSales")
    $report = $book.Worksheets.Item('Sales Analysis')
    $report.Activate()
    $excel.ActiveWindow.DisplayGridlines = $false
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
