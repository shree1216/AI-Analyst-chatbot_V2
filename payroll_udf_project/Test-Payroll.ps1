# Checks every UDF result in Employee_Payroll_UDF.xlsm against an independent calculation.
$ErrorActionPreference = 'Stop'
$folder = $PSScriptRoot
Add-Type -Namespace Native -Name Win -MemberDefinition '[DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint pid);'
$copy = Join-Path $env:TEMP ('PayrollTest_' + [guid]::NewGuid().ToString('N') + '.xlsm')
Copy-Item (Join-Path $folder 'Employee_Payroll_UDF.xlsm') $copy
$results = New-Object System.Collections.Generic.List[string]

function Check([string]$name, [bool]$condition, [string]$detail = '') {
    $status = if ($condition) { 'PASS' } else { 'FAIL' }
    $results.Add("$status  $name $detail".TrimEnd())
}
function Near($a, $b) { $a -is [double] -and [math]::Abs($a - $b) -lt 0.005 }
function Expected-Bonus([double]$salary, [int]$rating) {
    $rate = @{ 5 = 0.20; 4 = 0.15; 3 = 0.10; 2 = 0.05; 1 = 0 }[$rating]
    return [math]::Round($salary * $rate, 2)
}
function Expected-Deductions([double]$salary) {
    return [math]::Round($salary * 0.12 + 150 + [math]::Max(0, $salary - 5000) * 0.10, 2)
}

$excel = New-Object -ComObject Excel.Application
$excelPid = [uint32]0
[void][Native.Win]::GetWindowThreadProcessId([IntPtr]$excel.Hwnd, [ref]$excelPid)
$excel.DisplayAlerts = $false
try {
    $book = $excel.Workbooks.Open($copy)
    $excel.CalculateFull()
    $ws = $book.Worksheets.Item('Payroll')
    $data = $ws.Range('A2:I51').Value2
    $bad = 0; $sumBonus = 0; $sumDeductions = 0; $sumNet = 0
    for ($r = 1; $r -le 50; $r++) {
        $salary = $data[$r, 5]; $rating = [int]$data[$r, 6]
        $bonus = Expected-Bonus $salary $rating
        $deductions = Expected-Deductions $salary
        $net = $salary + $bonus - $deductions
        if (-not ((Near $data[$r, 7] $bonus) -and (Near $data[$r, 8] $deductions) -and (Near $data[$r, 9] $net))) {
            $bad++; $results.Add("      mismatch $($data[$r, 1]): got $($data[$r, 7]) / $($data[$r, 8]) / $($data[$r, 9])")
        }
        $sumBonus += $bonus; $sumDeductions += $deductions; $sumNet += $net
    }
    Check 'All 50 employees: Bonus, Deductions and Net Pay match independent calculation' ($bad -eq 0) "[$bad mismatches]"
    Check 'Total Bonus' (Near $ws.Range('G52').Value2 $sumBonus) "[$($ws.Range('G52').Text)]"
    Check 'Total Deductions' (Near $ws.Range('H52').Value2 $sumDeductions) "[$($ws.Range('H52').Text)]"
    Check 'Total Net Pay' (Near $ws.Range('I52').Value2 $sumNet) "[$($ws.Range('I52').Text)]"

    $v = $book.Worksheets.Item('Validation')
    for ($r = 2; $r -le 10; $r++) {
        Check "Manual case $($v.Cells($r, 1).Text): $($v.Cells($r, 2).Text)" ($v.Cells($r, 9).Text -eq 'PASS') "[bonus $($v.Cells($r, 6).Text), deductions $($v.Cells($r, 8).Text)]"
    }

    $salary = $ws.Range('E2').Value2
    $ws.Range('F2').Value2 = 5
    Check 'UDFs recalculate when rating changes' (Near $ws.Range('G2').Value2 ([math]::Round($salary * 0.2, 2)))
    $ws.Range('E2').Value2 = 'abc'
    Check 'Text salary returns #VALUE!' ($ws.Range('G2').Text -eq '#VALUE!' -and $ws.Range('H2').Text -eq '#VALUE!')
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
