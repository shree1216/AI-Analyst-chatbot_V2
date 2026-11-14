# Validates Healthcare_Dashboard.xlsm against an independent calculation from the raw CSV.
$ErrorActionPreference = 'Stop'
$folder = $PSScriptRoot
Add-Type -Namespace Native -Name Win -MemberDefinition '[DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint pid);'
$copy = Join-Path $env:TEMP ('HealthTest_' + [guid]::NewGuid().ToString('N') + '.xlsm')
Copy-Item (Join-Path $folder 'Healthcare_Dashboard.xlsm') $copy
$results = New-Object System.Collections.Generic.List[string]

function Check([string]$name, [bool]$condition, [string]$detail = '') {
    $status = if ($condition) { 'PASS' } else { 'FAIL' }
    $results.Add("$status  $name $detail".TrimEnd())
}
function Near($a, $b) { $null -ne $a -and [math]::Abs([double]$a - [double]$b) -lt 0.005 }
function Refresh { [void]$excel.Run("'$($book.Name)'!HealthcareReport.BuildReports") }
function Summarize($visits) {
    $completed = @($visits | Where-Object Completed)
    [pscustomobject]@{ Appts = @($visits).Count; Completed = $completed.Count
                       Revenue = [double](($visits | Measure-Object Bill -Sum).Sum) }
}

# Independent cleaning of the CSV
$clean = @(); $seen = @{}; $duplicates = 0; $incomplete = 0
foreach ($row in Import-Csv (Join-Path $folder 'healthcare_visits_raw_100.csv')) {
    $id = $row.'Visit ID'.Trim().ToUpper()
    $dept = (Get-Culture).TextInfo.ToTitleCase($row.Department.Trim().ToLower())
    $done = $row.Status.Trim() -eq 'completed'
    $bill = if ($row.'Billing Amount'.Trim()) { [double]$row.'Billing Amount' } else { 0 }
    if (-not $id -or -not $dept -or -not $row.'Appointment Date' -or ($done -and $bill -le 0)) { $incomplete++; continue }
    if ($seen.ContainsKey($id)) { $duplicates++; continue }
    $seen[$id] = $true
    $clean += [pscustomobject]@{ Dept = $dept; Month = ([datetime]$row.'Appointment Date').ToString('MMM yyyy', [Globalization.CultureInfo]::InvariantCulture)
                                 Completed = $done; Bill = $bill }
}
$all = Summarize $clean

$excel = New-Object -ComObject Excel.Application
$excelPid = [uint32]0
[void][Native.Win]::GetWindowThreadProcessId([IntPtr]$excel.Hwnd, [ref]$excelPid)
$excel.DisplayAlerts = $false
try {
    $book = $excel.Workbooks.Open($copy)
    Refresh
    $dash = $book.Worksheets.Item('Dashboard')

    Check 'Sheets: Dashboard, 3 monthly reports, Clean Data, Raw Data' ((($book.Worksheets | ForEach-Object Name) -join '|') -eq
        'Dashboard|Report Aug 2026|Report Sep 2026|Report Oct 2026|Clean Data|Raw Data')
    Check 'Clean records / duplicates / incomplete' ($dash.Range('I19').Value2 -eq 100 -and $dash.Range('I20').Value2 -eq $duplicates -and
        $dash.Range('I21').Value2 -eq $incomplete -and $dash.Range('I22').Value2 -eq $clean.Count) "[clean=$($clean.Count), dup=$duplicates, incomplete=$incomplete]"
    $cleanSheet = $book.Worksheets.Item('Clean Data')
    $cleanRows = [int]$excel.WorksheetFunction.CountA($cleanSheet.Range('A:A')) - 1
    $ids = @($cleanSheet.Range("A2:A$($cleanRows + 1)").Value2 | ForEach-Object { $_ })
    $depts = @($cleanSheet.Range("D2:D$($cleanRows + 1)").Value2 | ForEach-Object { $_ } | Sort-Object -Unique)
    $statuses = @($cleanSheet.Range("H2:H$($cleanRows + 1)").Value2 | ForEach-Object { $_ } | Sort-Object -Unique)
    Check 'Clean Data: unique IDs, standard departments and statuses' ($cleanRows -eq $clean.Count -and ($ids | Sort-Object -Unique).Count -eq $ids.Count -and
        $depts.Count -eq 5 -and ($statuses -join ',') -eq 'Cancelled,Completed,No-Show') "[$($depts -join ', ') | $($statuses -join ', ')]"

    Check 'KPI Total Appointments' ($dash.Range('B5').Value2 -eq $all.Appts) "[$($dash.Range('B5').Text)]"
    Check 'KPI Completion Rate' (Near $dash.Range('D5').Value2 ($all.Completed / $all.Appts)) "[$($dash.Range('D5').Text)]"
    Check 'KPI Total Revenue' (Near $dash.Range('F5').Value2 $all.Revenue) "[$($dash.Range('F5').Text)]"
    Check 'KPI Average Bill' (Near $dash.Range('H5').Value2 ($all.Revenue / $all.Completed)) "[$($dash.Range('H5').Text)]"

    $ok = $true; $previous = [double]::MaxValue
    for ($r = 9; $r -le 13; $r++) {
        $expected = Summarize ($clean | Where-Object Dept -eq $dash.Cells($r, 2).Text)
        $revenue = $dash.Cells($r, 5).Value2
        $ok = $ok -and $dash.Cells($r, 3).Value2 -eq $expected.Appts -and $dash.Cells($r, 4).Value2 -eq $expected.Completed -and
              (Near $revenue $expected.Revenue) -and $revenue -le $previous
        $previous = $revenue
    }
    Check 'Department performance (values, sorted by revenue, total row)' ($ok -and $dash.Range('B14').Text -eq 'Total' -and (Near $dash.Range('E14').Value2 $all.Revenue))

    $ok = $true
    $months = @('Aug 2026', 'Sep 2026', 'Oct 2026')
    for ($i = 0; $i -lt 3; $i++) {
        $expected = Summarize ($clean | Where-Object Month -eq $months[$i])
        $ok = $ok -and $dash.Cells(19 + $i, 2).Text -eq $months[$i] -and $dash.Cells(19 + $i, 3).Value2 -eq $expected.Appts -and
              (Near $dash.Cells(19 + $i, 5).Value2 $expected.Revenue)
        $report = $book.Worksheets.Item("Report $($months[$i])")
        $ok = $ok -and $report.Range('B5').Value2 -eq $expected.Appts -and (Near $report.Range('F5').Value2 $expected.Revenue) -and
              $report.ChartObjects().Count -eq 1
    }
    Check 'Monthly trend table and monthly report sheets' $ok
    Check 'Dashboard has 3 charts and the Refresh button' ($dash.ChartObjects().Count -eq 3 -and $dash.Buttons().Count -eq 1 -and
        $dash.Buttons(1).OnAction -like '*RefreshReport')

    Refresh
    Check 'Refresh is repeatable (same sheets, charts and totals)' ($book.Worksheets.Count -eq 6 -and $dash.ChartObjects().Count -eq 3 -and
        $dash.Buttons().Count -eq 1 -and (Near $dash.Range('F5').Value2 $all.Revenue))

    $raw = $book.Worksheets.Item('Raw Data')
    $raw.Range('A102:I102').Value2 = [object[]]@('v0999', 'PT1001', '  jane DOE ', 'cardiology', 'Dr. Alan Reed', ([datetime]'2026-11-05').ToOADate(), 'new', 'COMPLETED', 500)
    Refresh
    Check 'New month in raw data creates a new monthly report' ($book.Worksheets.Count -eq 7 -and $book.Worksheets.Item(5).Name -eq 'Report Nov 2026' -and
        $dash.Range('B5').Value2 -eq $all.Appts + 1 -and (Near $dash.Range('F5').Value2 ($all.Revenue + 500)) -and $dash.Range('B22').Text -eq 'Nov 2026')
    $book.Close($false)
}
finally {
    $excel.Quit()
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
    $excel = $null
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    Wait-Process -Id $excelPid -Timeout 15 -ErrorAction SilentlyContinue
    Stop-Process -Id $excelPid -ErrorAction SilentlyContinue
    Remove-Item $copy -ErrorAction SilentlyContinue
}
$results | Set-Content (Join-Path $folder 'validation_results.txt') -Encoding UTF8
$results
