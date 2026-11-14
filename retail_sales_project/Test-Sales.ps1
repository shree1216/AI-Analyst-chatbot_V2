# Validates Retail_Sales_Analysis.xlsm against an independent calculation from the CSV.
$ErrorActionPreference = 'Stop'
$folder = $PSScriptRoot
Add-Type -Namespace Native -Name Win -MemberDefinition '[DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint pid);'
$securityKey = 'HKCU:\Software\Microsoft\Office\16.0\Excel\Security'
$originalAccess = (Get-ItemProperty $securityKey -Name AccessVBOM -ErrorAction SilentlyContinue).AccessVBOM
$copy = Join-Path $env:TEMP ('SalesTest_' + [guid]::NewGuid().ToString('N') + '.xlsm')
Copy-Item (Join-Path $folder 'Retail_Sales_Analysis.xlsm') $copy
$results = New-Object System.Collections.Generic.List[string]

function Check([string]$name, [bool]$condition, [string]$detail = '') {
    $status = if ($condition) { 'PASS' } else { 'FAIL' }
    $results.Add("$status  $name $detail".TrimEnd())
}
function Near($a, $b) { $null -ne $a -and [math]::Abs([double]$a - [double]$b) -lt 0.005 }
function Analyze { $excel.Run("'$($book.Name)'!TestHarness.TryAnalyze") }

# Independent expectation
$rows = @(Import-Csv (Join-Path $folder 'retail_sales_oct_2026.csv'))
$amounts = @($rows | ForEach-Object { [double]$_.'Sales Amount' })
$total = ($amounts | Measure-Object -Sum).Sum
$units = ($rows | ForEach-Object { [int]$_.Quantity } | Measure-Object -Sum).Sum
$high = @($amounts | Where-Object { $_ -ge 500 }).Count
$low = @($amounts | Where-Object { $_ -lt 100 }).Count
$top5 = @($rows | Group-Object Product | ForEach-Object {
        [pscustomobject]@{ Name = $_.Name; Revenue = ($_.Group | ForEach-Object { [double]$_.'Sales Amount' } | Measure-Object -Sum).Sum }
    } | Sort-Object Revenue -Descending | Select-Object -First 5)
$regions = @($rows | Group-Object Region | ForEach-Object {
        [pscustomobject]@{ Name = $_.Name; Count = $_.Count; Revenue = ($_.Group | ForEach-Object { [double]$_.'Sales Amount' } | Measure-Object -Sum).Sum }
    } | Sort-Object Revenue -Descending)

$harness = @'
Public Function TryAnalyze() As String
    On Error GoTo Failed
    SalesAnalysis.AnalyzeSales
    TryAnalyze = "OK"
    Exit Function
Failed:
    Application.ScreenUpdating = True
    TryAnalyze = "ERROR: " & Err.Description
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
    Check 'Analysis runs' ((Analyze) -eq 'OK')
    $rep = $book.Worksheets.Item('Sales Analysis')
    $data = $book.Worksheets.Item('Sales Data')
    $m = $rep.Range('B6:B13').Value2

    Check 'Total Sales' (Near $m[1, 1] $total) "[$($rep.Range('B6').Text)]"
    Check 'Number of Transactions' ($m[2, 1] -eq 100)
    Check 'Average Sale' (Near $m[3, 1] ($total / 100)) "[$($rep.Range('B8').Text)]"
    Check 'Total Units Sold' ($m[4, 1] -eq $units) "[$units]"
    Check 'Highest / Lowest transaction' ((Near $m[5, 1] ($amounts | Measure-Object -Maximum).Maximum) -and (Near $m[6, 1] ($amounts | Measure-Object -Minimum).Minimum))
    Check 'High-value / low-value counts' ($m[7, 1] -eq $high -and $m[8, 1] -eq $low) "[high=$high, low=$low]"

    $ok = $true
    for ($i = 0; $i -lt 5; $i++) {
        $ok = $ok -and $rep.Cells(6 + $i, 5).Text -eq $top5[$i].Name -and (Near $rep.Cells(6 + $i, 7).Value2 $top5[$i].Revenue)
    }
    Check 'Top 5 products by revenue (names, order, revenue)' $ok "[1. $($top5[0].Name) $($top5[0].Revenue)]"

    $ok = $true
    for ($i = 0; $i -lt $regions.Count; $i++) {
        $ok = $ok -and $rep.Cells(14 + $i, 4).Text -eq $regions[$i].Name -and $rep.Cells(14 + $i, 5).Value2 -eq $regions[$i].Count -and
              (Near $rep.Cells(14 + $i, 6).Value2 $regions[$i].Revenue) -and (Near $rep.Cells(14 + $i, 8).Value2 ($regions[$i].Revenue / $total))
    }
    Check 'Sales by region (order, counts, revenue, share)' $ok "[$(($regions | ForEach-Object { "$($_.Name) $($_.Revenue)" }) -join ', ')]"

    $green = 0; $red = 0; $wrong = 0
    for ($r = 2; $r -le 101; $r++) {
        $color = $data.Cells($r, 1).Interior.Color
        $amount = $data.Cells($r, 8).Value2
        if ($color -eq 13561798) { $green++; if ($amount -lt 500) { $wrong++ } }
        elseif ($color -eq 13551615) { $red++; if ($amount -ge 100) { $wrong++ } }
    }
    Check 'High-value rows green, low-value rows red' ($green -eq $high -and $red -eq $low -and $wrong -eq 0) "[green=$green, red=$red]"
    $insights = @(19..24 | ForEach-Object { $rep.Cells($_, 1).Text } | Where-Object { $_ })
    Check 'Analytical summary written (title + 5 insights)' ($insights.Count -eq 6 -and $insights[2] -like "*$($top5[0].Name)*")
    Check 'Two charts created' ($rep.ChartObjects().Count -eq 2)

    Check 'Rerun is safe (still 2 charts, same total)' ((Analyze) -eq 'OK' -and $rep.ChartObjects().Count -eq 2 -and (Near $rep.Range('B6').Value2 $total))

    $data.Range('A102:H102').Value2 = [object[]]@('ORD-0101', 46326.0, 'Smart Watch', 'Electronics', 'East', 4, 250, 1000)
    [void](Analyze)
    Check 'New record included and highlighted' ($rep.Range('B7').Value2 -eq 101 -and (Near $rep.Range('B6').Value2 ($total + 1000)) -and
        $data.Range('A102').Interior.Color -eq 13561798 -and (Near $rep.Range('B10').Value2 1000))

    $data.Range('H50').Value2 = 'n/a'
    $before = $rep.Range('B6').Value2
    $message = Analyze
    Check 'Invalid amount reported, report unchanged' ($message -like '*Row 50*' -and (Near $rep.Range('B6').Value2 $before)) "[$message]"
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
