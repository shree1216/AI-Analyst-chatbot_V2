# Builds Healthcare_Dashboard.xlsm from healthcare_visits_raw_100.csv and HealthcareReport.bas.
# Requires desktop Excel. VBA project access is enabled only for the build, then restored.
$ErrorActionPreference = 'Stop'
$folder = $PSScriptRoot
$output = Join-Path $folder 'Healthcare_Dashboard.xlsm'
Add-Type -Namespace Native -Name Win -MemberDefinition '[DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint pid);'
$securityKey = 'HKCU:\Software\Microsoft\Office\16.0\Excel\Security'
$originalAccess = (Get-ItemProperty $securityKey -Name AccessVBOM -ErrorAction SilentlyContinue).AccessVBOM

$rows = @(Import-Csv (Join-Path $folder 'healthcare_visits_raw_100.csv'))
$headers = $rows[0].PSObject.Properties.Name
$grid = New-Object 'object[,]' ($rows.Count + 1), $headers.Count
for ($c = 0; $c -lt $headers.Count; $c++) { $grid[0, $c] = [string]$headers[$c] }
for ($r = 0; $r -lt $rows.Count; $r++) {
    for ($c = 0; $c -lt $headers.Count; $c++) {
        $value = [string]$rows[$r].($headers[$c])
        if ($value -and $headers[$c] -eq 'Appointment Date') { $value = ([datetime]::ParseExact($value, 'yyyy-MM-dd', $null)).ToOADate() }
        elseif ($value -and $headers[$c] -eq 'Billing Amount') { $value = [double]$value }
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
    [void]$book.VBProject.VBComponents.Import((Join-Path $folder 'HealthcareReport.bas'))

    $raw = $book.Worksheets.Item(1)
    $raw.Name = 'Raw Data'
    $raw.Range('A:E').NumberFormat = '@'
    $raw.Range('G:H').NumberFormat = '@'
    $raw.Range('F:F').NumberFormat = 'yyyy-mm-dd'
    $raw.Range($raw.Cells(1, 1), $raw.Cells($rows.Count + 1, $headers.Count)).Value2 = $grid
    $raw.Range('A1:I1').Font.Bold = $true
    $raw.Range('A1:I1').Interior.Color = 14277081   # light grey: raw extract
    $raw.Columns('A:I').AutoFit() | Out-Null

    # Run once, then put the dashboard first and add the refresh button to it.
    [void]$excel.Run("'$($book.Name)'!HealthcareReport.BuildReports")
    $dashboard = $book.Worksheets.Item('Dashboard')
    $dashboard.Move($book.Worksheets.Item(1))
    $book.Worksheets.Item('Clean Data').Move([Type]::Missing, $book.Worksheets.Item($book.Worksheets.Count))
    $raw.Move([Type]::Missing, $book.Worksheets.Item($book.Worksheets.Count))
    $anchor = $dashboard.Range('H1')
    $button = $dashboard.Buttons().Add($anchor.Left, $anchor.Top + 4, 170, 30)
    $button.Caption = 'Refresh Report'
    $button.OnAction = 'RefreshReport'
    $button.Font.Bold = $true
    $button.Placement = 3
    foreach ($name in 'Clean Data', 'Raw Data') {
        $book.Worksheets.Item($name).Activate()
        $excel.ActiveWindow.SplitRow = 1
        $excel.ActiveWindow.FreezePanes = $true
        $book.Worksheets.Item($name).Range('A1').Select() | Out-Null
    }
    $dashboard.Activate()
    $dashboard.Range('A1').Select() | Out-Null
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
