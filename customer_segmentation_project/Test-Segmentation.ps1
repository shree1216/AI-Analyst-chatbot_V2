# Validates Customer_Segmentation.xlsm against an independent calculation from the raw CSV.
$ErrorActionPreference = 'Stop'
$folder = $PSScriptRoot
Add-Type -Namespace Native -Name Win -MemberDefinition '[DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint pid);'
$copy = Join-Path $env:TEMP ('SegmentationTest_' + [guid]::NewGuid().ToString('N') + '.xlsm')
Copy-Item (Join-Path $folder 'Customer_Segmentation.xlsm') $copy
$results = New-Object System.Collections.Generic.List[string]

function Check([string]$name, [bool]$condition, [string]$detail = '') {
    $status = if ($condition) { 'PASS' } else { 'FAIL' }
    $results.Add("$status  $name $detail".TrimEnd())
}
function Data-Rows($sheet) { [int]$excel.WorksheetFunction.CountA($sheet.Range('A:A')) - 1 }
function Column-Values($sheet, [string]$col) {
    $n = Data-Rows $sheet
    if ($n -lt 1) { return @() }
    return @($sheet.Range("${col}2:$col$($n + 1)").Value2 | ForEach-Object { $_ })
}

# Independent expectation from the CSV
$expected = @{ Duplicates = 0; Incomplete = 0; Clean = 0; Senior = 0; Home = 0; Car = 0 }
$seen = @{}
foreach ($row in Import-Csv (Join-Path $folder 'bank_customers_raw_100.csv')) {
    $digits = ($row.Phone -replace '\D', '') -replace '^1(\d{10})$', '$1'
    $age = 0
    $complete = $row.'Customer ID'.Trim() -and $row.'Customer Name'.Trim() -and $row.City.Trim() -and
                [int]::TryParse($row.Age, [ref]$age) -and $age -ge 18 -and $digits.Length -eq 10
    $id = $row.'Customer ID'.Trim().ToUpper()
    if (-not $complete) { $expected.Incomplete++ }
    elseif ($seen.ContainsKey($id)) { $expected.Duplicates++ }
    else {
        $seen[$id] = $true; $expected.Clean++
        if ($age -ge 60) { $expected.Senior++ }
        if ($row.'Loan Type' -match 'home|housing') { $expected.Home++ }
        if ($row.'Loan Type' -match 'car|auto') { $expected.Car++ }
    }
}

$excel = New-Object -ComObject Excel.Application
$excelPid = [uint32]0
[void][Native.Win]::GetWindowThreadProcessId([IntPtr]$excel.Hwnd, [ref]$excelPid)
$excel.DisplayAlerts = $false
try {
    $book = $excel.Workbooks.Open($copy)
    [void]$excel.Run("'$($book.Name)'!CustomerProcessing.ProcessCustomerData")
    $ws = @{}
    foreach ($name in 'Summary', 'Raw Data', 'Clean Data', 'Senior Citizens', 'Home Loan', 'Car Loan', 'Duplicates', 'Incomplete') {
        $ws[$name] = $book.Worksheets.Item($name)
    }
    $counts = @{
        Clean = Data-Rows $ws['Clean Data']; Senior = Data-Rows $ws['Senior Citizens']; Home = Data-Rows $ws['Home Loan']
        Car = Data-Rows $ws['Car Loan']; Duplicates = Data-Rows $ws['Duplicates']; Incomplete = Data-Rows $ws['Incomplete']
    }
    foreach ($key in 'Duplicates', 'Incomplete', 'Clean', 'Senior', 'Home', 'Car') {
        Check "$key count matches independent calculation" ($counts[$key] -eq $expected[$key]) "[$($counts[$key])]"
    }
    Check 'Every raw record accounted for (clean + duplicates + incomplete = 100)' (($counts.Clean + $counts.Duplicates + $counts.Incomplete) -eq 100)
    $summary = $ws['Summary'].Range('B5:B11').Value2
    Check 'Summary sheet matches category sheets' ($summary[1, 1] -eq 100 -and $summary[2, 1] -eq $counts.Duplicates -and
        $summary[3, 1] -eq $counts.Incomplete -and $summary[4, 1] -eq $counts.Clean -and $summary[5, 1] -eq $counts.Senior -and
        $summary[6, 1] -eq $counts.Home -and $summary[7, 1] -eq $counts.Car)
    Check 'Summary chart present' ($ws['Summary'].ChartObjects().Count -eq 1)

    $clean = $ws['Clean Data']
    $ids = Column-Values $clean 'A'
    Check 'No duplicate Customer IDs in Clean Data' (($ids | Select-Object -Unique).Count -eq $ids.Count)
    $names = Column-Values $clean 'B'
    Check 'Names trimmed and in Proper Case' (@($names | Where-Object { $_ -cne (Get-Culture).TextInfo.ToTitleCase($_.ToLower()) -or $_ -match '^\s|\s$|\s\s' }).Count -eq 0)
    Check 'Phones standardized to 555-123-4567' (@(Column-Values $clean 'E' | Where-Object { $_ -notmatch '^\d{3}-\d{3}-\d{4}$' }).Count -eq 0)
    Check 'Emails lower case' (@(Column-Values $clean 'F' | Where-Object { $_ -cne $_.ToLower() }).Count -eq 0)
    $loanTypes = @(Column-Values $clean 'H' | Select-Object -Unique | Sort-Object)
    Check 'Loan types standardized' (@($loanTypes | Where-Object { $_ -notin 'Home Loan', 'Car Loan', 'Personal Loan', 'None' }).Count -eq 0) "[$($loanTypes -join ', ')]"
    Check 'Senior sheet only has age 60+' (@(Column-Values $ws['Senior Citizens'] 'C' | Where-Object { $_ -lt 60 }).Count -eq 0)
    Check 'Home Loan sheet only has home loans' (@(Column-Values $ws['Home Loan'] 'H' | Where-Object { $_ -ne 'Home Loan' }).Count -eq 0)
    Check 'Car Loan sheet only has car loans' (@(Column-Values $ws['Car Loan'] 'H' | Where-Object { $_ -ne 'Car Loan' }).Count -eq 0)
    $issues = @(Column-Values $ws['Incomplete'] 'J') + @(Column-Values $ws['Duplicates'] 'J')
    Check 'Every incomplete/duplicate row has an Issue' (@($issues | Where-Object { -not $_ }).Count -eq 0) "[$($issues[0]) ...]"
    Check 'Raw Data left unchanged' ($ws['Raw Data'].Range('B2').Text -eq (Import-Csv (Join-Path $folder 'bank_customers_raw_100.csv'))[0].'Customer Name')

    [void]$excel.Run("'$($book.Name)'!CustomerProcessing.ProcessCustomerData")
    Check 'Rerun gives the same result (no doubled rows)' ((Data-Rows $ws['Clean Data']) -eq $counts.Clean -and (Data-Rows $ws['Duplicates']) -eq $counts.Duplicates)

    $raw = $ws['Raw Data']
    $raw.Range('A102:I102').Value2 = [object[]]@('cust9001', '  mary   SMITH ', 72, 'dallas', '(555) 777-1234', 'MARY.SMITH@EXAMPLE.COM', 'savings', 'AUTO LOAN', 25000)
    [void]$excel.Run("'$($book.Name)'!CustomerProcessing.ProcessCustomerData")
    $car = $ws['Car Loan']
    $last = Data-Rows $car
    Check 'New raw record is cleaned and placed in Senior + Car Loan' ((Data-Rows $ws['Senior Citizens']) -eq $counts.Senior + 1 -and
        $last -eq $counts.Car + 1 -and $car.Range("B$($last + 1)").Text -eq 'Mary Smith' -and $car.Range("E$($last + 1)").Text -eq '555-777-1234' -and
        $car.Range("A$($last + 1)").Text -eq 'CUST9001') "[$($car.Range("A$($last + 1)").Text) | $($car.Range("B$($last + 1)").Text) | $($car.Range("E$($last + 1)").Text)]"
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
