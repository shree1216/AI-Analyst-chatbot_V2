# Validates the recorded (broken) and corrected macros on a temporary copy of the workbook.
$ErrorActionPreference = 'Stop'
$folder = $PSScriptRoot
$source = Join-Path $folder 'Banking_Transactions_Report.xlsm'
$copy = Join-Path $env:TEMP ('BankingMacroTest_' + [guid]::NewGuid().ToString('N') + '.xlsm')
Copy-Item $source $copy
$results = New-Object System.Collections.Generic.List[string]

function Check([string]$name, [bool]$condition, [string]$detail = '') {
    $status = if ($condition) { 'PASS' } else { 'FAIL' }
    $results.Add("$status  $name $detail".TrimEnd())
}
function Near($a, $b) { [math]::Abs([double]$a - [double]$b) -lt 0.005 }
function Run-Macro([string]$macro, $sheet) {
    $text = if ($macro -like 'RecordedMacro*') { $excel.Run("'$($book.Name)'!TestHarness.TryRecorded") }
            else { $excel.Run("'$($book.Name)'!TestHarness.TryFormat", $sheet) }
    if ($text -like 'OK*') { return @{ ok = $true; value = [int]($text -replace '^OK:?', '0') } }
    return @{ ok = $false; value = $text }
}
function New-TestSheet([string]$name) {
    $base.Copy([Type]::Missing, $book.Worksheets.Item($book.Worksheets.Count))
    $sheet = $book.Worksheets.Item($book.Worksheets.Count)
    $sheet.Name = $name
    $sheet.Activate()
    return $sheet
}
function Count-Total($sheet) { [int]$excel.WorksheetFunction.CountIf($sheet.Range('A:A'), 'TOTAL') }

$rows = @(Import-Csv (Join-Path $folder 'banking_transactions_100.csv'))
$expectedDebit = ($rows | Where-Object Debit | Measure-Object -Property Debit -Sum).Sum
$expectedCredit = ($rows | Where-Object Credit | Measure-Object -Property Credit -Sum).Sum
$expectedNet = $expectedCredit - $expectedDebit

Add-Type -Namespace Native -Name Win -MemberDefinition '[DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint pid);'
$securityKey = 'HKCU:\Software\Microsoft\Office\16.0\Excel\Security'
$originalAccess = (Get-ItemProperty $securityKey -Name AccessVBOM -ErrorAction SilentlyContinue).AccessVBOM
# Temporary harness (added to the test copy only) traps VBA errors so no dialog blocks automation.
$harness = @'
Public Function TryFormat(ByVal ws As Worksheet) As String
    On Error GoTo Failed
    TryFormat = "OK:" & BankingReport.FormatReport(ws)
    Exit Function
Failed:
    TryFormat = "ERROR " & Err.Number & ": " & Err.Description
End Function

Public Function TryRecorded() As String
    On Error GoTo Failed
    RecordedMacro.FormatReport_Recorded
    TryRecorded = "OK"
    Exit Function
Failed:
    Application.ScreenUpdating = True
    TryRecorded = "ERROR " & Err.Number & ": " & Err.Description
End Function
'@

Set-ItemProperty $securityKey -Name AccessVBOM -Value 1 -Type DWord
$excel = New-Object -ComObject Excel.Application
    $excelPid = [uint32]0
    [void][Native.Win]::GetWindowThreadProcessId([IntPtr]$excel.Hwnd, [ref]$excelPid)
$excel.DisplayAlerts = $false
try {
    $book = $excel.Workbooks.Open($copy)
    $module = $book.VBProject.VBComponents.Add(1)
    $module.Name = 'TestHarness'
    $module.CodeModule.AddFromString($harness)
    $base = $book.Worksheets.Item('Transactions')

    # --- Recorded macro: reproduce the faults -------------------------------------------
    $base.Activate()
    $run = Run-Macro 'RecordedMacro.FormatReport_Recorded' $null
    Check 'Recorded macro fails on renamed sheet (Sheet1 -> Transactions)' (-not $run.ok) "[$($run.value)]"

    $s = New-TestSheet 'Sheet1'
    $s.Range('A102:H106').Value2 = $s.Range('A97:H101').Value2   # 5 new transactions
    $id102 = $s.Range('A102').Text
    [void](Run-Macro 'RecordedMacro.FormatReport_Recorded' $null)
    Check 'Recorded macro overwrites new data in row 102' ($s.Range('A102').Text -eq 'TOTAL') "[$id102 replaced by TOTAL]"
    Check 'Recorded macro formats amounts as percent' ($s.Range('G2').NumberFormat -eq '0.00%') "[$($s.Range('G2').Text)]"
    Check 'Recorded macro computes Debit + Credit (wrong sign)' (Near $s.Range('I2').Value2 ([double]$s.Range('G2').Value2 + [double]$s.Range('H2').Value2))
    Check 'Recorded macro leaves rows 103-106 uncalculated' ($null -eq $s.Range('I103').Value2)

    # --- Corrected macro ------------------------------------------------------------------
    $s = New-TestSheet 'Daily Report'
    $run = Run-Macro 'BankingReport.FormatReport' $s
    Check 'Corrected macro runs on any sheet name' ($run.ok -and $run.value -eq 100) "[rows=$($run.value)]"
    Check 'Net Amount column added in I' ($s.Range('I1').Text -eq 'Net Amount')
    Check 'Net Amount = Credit - Debit' ($s.Range('I2').Formula -eq '=N($H2)-N($G2)') "[$($s.Range('I2').Formula)]"
    Check 'Totals row at row 102' ($s.Range('A102').Text -eq 'TOTAL')
    Check 'Debit total' (Near $s.Range('G102').Value2 $expectedDebit) "[$($s.Range('G102').Text)]"
    Check 'Credit total' (Near $s.Range('H102').Value2 $expectedCredit) "[$($s.Range('H102').Text)]"
    Check 'Net total' (Near $s.Range('I102').Value2 $expectedNet) "[$($s.Range('I102').Text)]"
    Check 'Date format dd-mmm-yyyy' ($s.Range('B2').NumberFormat -eq 'dd-mmm-yyyy') "[$($s.Range('B2').Text)]"
    Check 'Amount format #,##0.00' ($s.Range('G2').NumberFormat -like '#,##0.00*') "[$($s.Range('G2').Text)]"
    Check 'Header styled bold' ([bool]$s.Range('A1').Font.Bold -and [bool]$s.Range('I1').Font.Bold)
    Check 'Header row frozen' ([bool]$excel.ActiveWindow.FreezePanes)

    $run = Run-Macro 'BankingReport.FormatReport' $s
    Check 'Rerun is safe (one TOTAL row, no extra Net column)' ($run.value -eq 100 -and (Count-Total $s) -eq 1 -and $s.Range('J1').Text -eq '')

    $s.Range('A103:H107').Value2 = $s.Range('A97:H101').Value2     # rows typed below old totals
    $run = Run-Macro 'BankingReport.FormatReport' $s
    $addedNet = $expectedNet + $excel.WorksheetFunction.Sum($s.Range('I97:I101'))
    Check 'Added rows are included' ($run.value -eq 105 -and (Count-Total $s) -eq 1 -and $s.Range('A107').Text -eq 'TOTAL' -and (Near $s.Range('I107').Value2 $addedNet)) "[rows=$($run.value), net=$($s.Range('I107').Text)]"

    $s.Range('A2:A11').EntireRow.Delete() | Out-Null
    $run = Run-Macro 'BankingReport.FormatReport' $s
    Check 'Deleted rows are handled' ($run.value -eq 95 -and $s.Range('A97').Text -eq 'TOTAL') "[rows=$($run.value)]"

    $s = New-TestSheet 'Changed Layout'
    $s.Cells.Clear() | Out-Null
    # Title rows above the data, table starts in column B, Credit placed before Debit
    $s.Range('B1').Value2 = 'Daily Transaction Report'
    $order = 'A', 'B', 'C', 'D', 'E', 'F', 'H', 'G'
    for ($i = 0; $i -lt $order.Count; $i++) {
        $s.Range($s.Cells(3, $i + 2), $s.Cells(103, $i + 2)).Value2 = $base.Range("$($order[$i])1:$($order[$i])101").Value2
    }
    $s.Range('C4:C103').NumberFormat = 'yyyy-mm-dd'
    $run = Run-Macro 'BankingReport.FormatReport' $s
    Check 'Changed layout (title rows, column B start, Credit before Debit)' ($run.value -eq 100 -and (Near $s.Range('J104').Value2 $expectedNet)) "[net=$($s.Range('J104').Text)]"

    $s = New-TestSheet 'Invalid Data'
    $s.Range('G10').Value2 = 'abc'
    $run = Run-Macro 'BankingReport.FormatReport' $s
    Check 'Invalid amount is reported and sheet is left unchanged' ((-not $run.ok) -and $run.value -like '*Row 10*' -and $s.Range('I1').Text -eq '') "[$($run.value)]"

    $book.Close($false)
}
finally {
    $excel.Quit()
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
    $excel = $null
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
        # Excel saves Trust Center settings on exit, so restore only after it has closed.
        Wait-Process -Id $excelPid -Timeout 15 -ErrorAction SilentlyContinue
        Stop-Process -Id $excelPid -ErrorAction SilentlyContinue
    if ($null -eq $originalAccess) { Remove-ItemProperty $securityKey -Name AccessVBOM -ErrorAction SilentlyContinue }
    else { Set-ItemProperty $securityKey -Name AccessVBOM -Value $originalAccess -Type DWord }
    Remove-Item $copy -ErrorAction SilentlyContinue
}
$results | Set-Content (Join-Path $folder 'validation_results.txt') -Encoding UTF8
$results
