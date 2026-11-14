param(
    [Parameter(Mandatory = $true)]
    [string]$WorkbookPath,
    [switch]$UseOpenWorkbook
)
$ErrorActionPreference = 'Stop'
$source = (Resolve-Path -LiteralPath $WorkbookPath).Path
if ([IO.Path]::GetExtension($source) -ne '.xlsm') { throw 'Import BankingReport.bas and save as .xlsm first.' }
$temporary = Join-Path ([IO.Path]::GetDirectoryName($source)) ('BankingTest_' + [guid]::NewGuid().ToString('N') + '.xlsm')
$excel = $null
$book = $null
$results = [Collections.Generic.List[object]]::new()
$createdSheets = [Collections.Generic.List[object]]::new()
$originalAlerts = $true
$originalEvents = $true
$originalScreen = $true

function Set-CellValue($cell, $value) {
    # IDispatch avoids Windows PowerShell's cached string-only Value2 setter.
    [void]$cell.GetType().InvokeMember('Value2',[Reflection.BindingFlags]::SetProperty,$null,$cell,[object[]]@($value))
}
function Assert-Equal($actual, $expected, [string]$message) {
    if ($actual -ne $expected) { throw "$message Expected '$expected'; got '$actual'." }
}
function Assert-Amount($actual, [double]$expected, [string]$message) {
    if ($null -eq $actual -or [math]::Abs([double]$actual - $expected) -gt 0.005) {
        throw "$message Expected $expected; got '$actual'."
    }
}
function Invoke-Formatter($sheet) {
    $null = $excel.Run("'" + $book.Name.Replace("'","''") + "'!BankingReport.FormatBankingWorksheet", $sheet)
}
function Add-Pass([string]$name) {
    $results.Add([pscustomobject]@{test=$name;result='PASS'})
    Write-Output "PASS: $name"
}
function Copy-TestSheet($sourceSheet, [string]$name) {
    $sourceSheet.Copy([Type]::Missing, $book.Worksheets.Item($book.Worksheets.Count))
    $copy = $book.Worksheets.Item($book.Worksheets.Count)
    $copy.Name = $name
    $createdSheets.Add($copy)
    return $copy
}
function Assert-Rejected($sheet, [string]$name) {
    $before = $sheet.UsedRange.Value2
    $beforeAddress = $sheet.UsedRange.Address()
    $tablesBefore = $sheet.ListObjects.Count
    $eventsBefore = $excel.EnableEvents
    $screenBefore = $excel.ScreenUpdating
    $rejected = $false
    try { Invoke-Formatter $sheet } catch {
        $rejected = $true
        Write-Output ('Expected rejection: ' + $_.Exception.Message)
    }
    if (-not $rejected) { throw "Invalid input accepted: $name" }
    Assert-Equal $sheet.ListObjects.Count $tablesBefore "$name table count changed."
    Assert-Equal $sheet.UsedRange.Address() $beforeAddress "$name range changed."
    $after = $sheet.UsedRange.Value2
    for ($r=$before.GetLowerBound(0); $r -le $before.GetUpperBound(0); $r++) {
        for ($c=$before.GetLowerBound(1); $c -le $before.GetUpperBound(1); $c++) {
            Assert-Equal $after[$r,$c] $before[$r,$c] "$name changed input at $r,$c."
        }
    }
    Assert-Equal $excel.EnableEvents $eventsBefore "$name did not restore events."
    Assert-Equal $excel.ScreenUpdating $screenBefore "$name did not restore screen updating."
    Add-Pass $name
}
function Assert-Report($sheet, [int]$rowCount, [double]$debit, [double]$credit) {
    Assert-Equal $sheet.ListObjects.Count 1 'Expected exactly one report table.'
    $table = $sheet.ListObjects.Item(1)
    Assert-Equal $table.ListRows.Count $rowCount 'Incorrect transaction count.'
    Assert-Equal $table.ListColumns.Count 9 'Calculated columns missing or duplicated.'
    Assert-Equal $table.TableStyle.Name 'TableStyleMedium2' 'Incorrect formatting.'
    Assert-Equal $table.ShowTotals $true 'Totals row missing.'
    $debitColumn = $table.ListColumns.Item('Debit').Index
    $creditColumn = $table.ListColumns.Item('Credit').Index
    $netColumn = $table.ListColumns.Item('Net Amount').Index
    $runningColumn = $table.ListColumns.Item('Running Net').Index
    Assert-Amount $table.TotalsRowRange.Cells.Item(1,$debitColumn).Value2 $debit 'Debit total mismatch.'
    Assert-Amount $table.TotalsRowRange.Cells.Item(1,$creditColumn).Value2 $credit 'Credit total mismatch.'
    Assert-Amount $table.TotalsRowRange.Cells.Item(1,$netColumn).Value2 ($credit-$debit) 'Net total mismatch.'
    Assert-Amount $table.DataBodyRange.Cells.Item($rowCount,$runningColumn).Value2 ($credit-$debit) 'Running Net mismatch.'
    Assert-Equal $table.ListColumns.Item('Transaction Date').DataBodyRange.Cells.Item(1,1).NumberFormat 'yyyy-mm-dd' 'Date format mismatch.'
    Assert-Equal $sheet.PageSetup.Orientation 2 'Expected landscape printing.'
    for ($r=1; $r -le $rowCount; $r++) {
        $expected = [double]$table.DataBodyRange.Cells.Item($r,$creditColumn).Value2 - [double]$table.DataBodyRange.Cells.Item($r,$debitColumn).Value2
        Assert-Amount $table.DataBodyRange.Cells.Item($r,$netColumn).Value2 $expected "Net mismatch in row $r."
        Assert-Equal $table.DataBodyRange.Cells.Item($r,$netColumn).HasFormula $true "Missing Net formula in row $r."
        Assert-Equal $table.DataBodyRange.Cells.Item($r,$runningColumn).HasFormula $true "Missing Running Net formula in row $r."
    }
}
try {
    if ($UseOpenWorkbook) {
        $excel = [Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application')
        $book = $excel.Workbooks.Item([IO.Path]::GetFileName($source))
        Assert-Equal $book.FileFormat 52 'Expected the open macro-enabled demo workbook.'
    } else {
        Copy-Item -LiteralPath $source -Destination $temporary
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $true
        $excel.AutomationSecurity = 2
        $book = $excel.Workbooks.Open($temporary,0,$false)
    }
    if ($null -eq $book) { throw 'Excel did not open the test copy.' }
    $originalAlerts = $excel.DisplayAlerts
    $originalEvents = $excel.EnableEvents
    $originalScreen = $excel.ScreenUpdating
    $excel.DisplayAlerts = $true
    if (-not $book.HasVBProject) { throw 'The workbook has no VBA project. Import BankingReport.bas first.' }
    $transactions = Copy-TestSheet ($book.Worksheets.Item('Transactions')) 'Test Transactions'
    $changed = Copy-TestSheet ($book.Worksheets.Item('Changed Layout')) 'Test Changed'
    $raw = Copy-TestSheet $transactions 'Test Raw'
    if ($raw.ListObjects.Count -gt 0) {
        $raw.ListObjects.Item(1).ShowTotals = $false
        $raw.ListObjects.Item(1).Unlist()
    }
    $raw.Range('H:I').Clear()
    $snapshot = $transactions.Range('A6:G25').Value2
    $excel.EnableEvents = $false
    $excel.ScreenUpdating = $false
    Invoke-Formatter $transactions
    Assert-Equal $excel.EnableEvents $false 'Success changed the original event state.'
    Assert-Equal $excel.ScreenUpdating $false 'Success changed the original screen state.'
    $excel.EnableEvents = $true
    $excel.ScreenUpdating = $true
    Assert-Report $transactions 20 2837.75 9300
    $after = $transactions.Range('A6:G25').Value2
    for ($r=1; $r -le 20; $r++) {
        for ($c=1; $c -le 7; $c++) { Assert-Equal $after[$r,$c] $snapshot[$r,$c] 'Source input changed.' }
    }
    Add-Pass 'Baseline calculations, formatting, input preservation, and state restoration'
    Invoke-Formatter $transactions
    Assert-Report $transactions 20 2837.75 9300
    Add-Pass 'Rerun without duplicated output columns or totals'
    Invoke-Formatter $changed
    Assert-Report $changed 23 3087.75 10140
    Assert-Equal $changed.ListObjects.Item(1).HeaderRowRange.Row 8 'Shifted header was not found.'
    Assert-Equal $changed.ListObjects.Item(1).HeaderRowRange.Column 2 'Shifted columns were not found.'
    Add-Pass 'Renamed worksheet, reordered columns, title rows, and added data'
    $table = $transactions.ListObjects.Item(1)
    $newRow = $table.ListRows.Add()
    $values = @('TXN-TEST-0021',[datetime]::new(2026,11,7).ToOADate(),'ACCT-DEMO-001','Integration test deposit',0.0,100.0,'USD')
    for ($c=0; $c -lt 7; $c++) { Set-CellValue ($newRow.Range.Cells.Item(1,$c+1)) $values[$c] }
    Invoke-Formatter $transactions
    Assert-Report $transactions 21 2837.75 9400
    Add-Pass 'Inserted table row included in formulas and totals'
    $table.ListRows.Item(21).Delete()
    Invoke-Formatter $transactions
    Assert-Report $transactions 20 2837.75 9300
    Add-Pass 'Deleted table row and totals correctly resized'
    Assert-Rejected (Copy-TestSheet ($book.Worksheets.Item('Invalid Data')) 'Test Invalid') 'Duplicate transaction ID refused before formatting'
    $cases = @(
        @{name='Missing ID';row=6;column=1;value=''},
        @{name='Missing header';row=5;column=5;value='Amount'},
        @{name='Duplicate header';row=5;column=7;value='Debit'},
        @{name='Negative debit';row=7;column=5;value=-1.0},
        @{name='Text amount';row=7;column=5;value='not a number'},
        @{name='Fractional cent';row=7;column=5;value=1.001},
        @{name='Both debit and credit';row=6;column=5;value=1.0},
        @{name='Neither debit nor credit';row=6;column=6;value=0.0},
        @{name='Invalid date';row=6;column=2;value='not a date'},
        @{name='Mixed currency';row=6;column=7;value='EUR'}
    )
    $n=0
    foreach ($case in $cases) {
        $n++
        $sheet = Copy-TestSheet $raw "Reject $n"
        Set-CellValue ($sheet.Cells.Item($case.row,$case.column)) $case.value
        Assert-Rejected $sheet $case.name
    }
    $protected = Copy-TestSheet $raw 'Protected'
    $protected.Protect()
    Assert-Rejected $protected 'Protected sheet refused without unprotecting'
    $empty = Copy-TestSheet $raw 'Empty'
    $empty.Range('A6:G25').ClearContents()
    Assert-Rejected $empty 'Empty dataset refused explicitly'
    $below = Copy-TestSheet $transactions 'Below Totals'
    $below.Cells.Item($below.ListObjects.Item(1).TotalsRowRange.Row+1,1).Value2 = 'TXN-BELOW'
    Assert-Rejected $below 'Data below totals refused explicitly'
    $reportPath = Join-Path $PSScriptRoot 'macro-validation-results.json'
    [pscustomobject]@{validatedAt=[datetime]::UtcNow.ToString('o');workbook=$source;engine='Desktop Excel VBA';worksheetIsolation=$true;tests=$results.ToArray()} |
        ConvertTo-Json -Depth 5 | Set-Content -LiteralPath $reportPath -Encoding UTF8
    Write-Output ("SUCCESS: {0} actual VBA integration checks passed. Results: {1}" -f $results.Count,$reportPath)
} finally {
    if ($null -ne $book) {
        if ($UseOpenWorkbook) {
            $excel.DisplayAlerts = $false
            foreach ($sheet in $createdSheets) { $sheet.Delete() }
            $excel.DisplayAlerts = $originalAlerts
            $excel.EnableEvents = $originalEvents
            $excel.ScreenUpdating = $originalScreen
        } else {
            $book.Close($false)
        }
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($book)
    }
    if ($null -ne $excel) {
        if (-not $UseOpenWorkbook) { $excel.Quit() }
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    if (Test-Path -LiteralPath $temporary) { Remove-Item -LiteralPath $temporary }
}
