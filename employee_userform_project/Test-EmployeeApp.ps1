# Validates Employee_Management_System.xlsm: form, validation, Add/Search/Update/Delete and dashboard.
$ErrorActionPreference = 'Stop'
$folder = $PSScriptRoot
Add-Type -Namespace Native -Name Win -MemberDefinition '[DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint pid);'
$securityKey = 'HKCU:\Software\Microsoft\Office\16.0\Excel\Security'
$originalAccess = (Get-ItemProperty $securityKey -Name AccessVBOM -ErrorAction SilentlyContinue).AccessVBOM
$copy = Join-Path $env:TEMP ('EmployeeTest_' + [guid]::NewGuid().ToString('N') + '.xlsm')
Copy-Item (Join-Path $folder 'Employee_Management_System.xlsm') $copy
$results = New-Object System.Collections.Generic.List[string]

function Check([string]$name, [bool]$condition, [string]$detail = '') {
    $status = if ($condition) { 'PASS' } else { 'FAIL' }
    $results.Add("$status  $name $detail".TrimEnd())
}
function Run-Vba([string]$macro, [object[]]$arguments = @()) {
    $name = "'$($book.Name)'!$macro"
    switch ($arguments.Count) {
        0 { return $excel.Run($name) }
        1 { return $excel.Run($name, $arguments[0]) }
        2 { return $excel.Run($name, $arguments[0], $arguments[1]) }
    }
}

$employees = @(Import-Csv (Join-Path $folder 'employees_100.csv'))

# Test helpers added to the throwaway copy only. They drive the real form without showing it.
$harness = @'
Public Function FormCheck() As String
    On Error GoTo Failed
    Load frmEmployee
    With frmEmployee
        FormCheck = .Controls.Count & "|" & .cboGender.ListCount & "|" & .cboDepartment.ListCount & "|" & _
                    .cboStatus.ListCount & "|" & .txtID.Text & "|" & .btnAdd.Enabled & "|" & .btnUpdate.Enabled
        .LoadEmployee FindEmployee("EMP005")
        FormCheck = FormCheck & "|" & .txtName.Text & "|" & .cboDepartment.Value & "|" & .txtEmail.Text & "|" & _
                    .btnAdd.Enabled & "|" & .btnUpdate.Enabled & "|" & .btnDelete.Enabled & "|" & .lblMode.Caption
    End With
    Unload frmEmployee
    Exit Function
Failed:
    FormCheck = "ERROR: " & Err.Description
End Function

' Fills the form like a user would, then validates and saves through the form's own FormValues.
Public Function FormAdd(ByVal values As Variant) As String
    On Error GoTo Failed
    Dim problems As String
    Load frmEmployee
    With frmEmployee
        .ClearForm
        .txtName.Text = values(1): .cboGender.Value = values(2): .cboDepartment.Value = values(3)
        .txtTitle.Text = values(4): .txtEmail.Text = values(5): .txtPhone.Text = values(6)
        .txtHireDate.Text = values(7): .txtSalary.Text = values(8): .cboStatus.Value = values(9)
        problems = ValidateEmployee(.FormValues)
        If Len(problems) = 0 Then FormAdd = "OK:" & SaveEmployee(.FormValues) Else FormAdd = "INVALID:" & problems
    End With
    Unload frmEmployee
    Exit Function
Failed:
    FormAdd = "ERROR: " & Err.Description
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
    $dash = $book.Worksheets.Item('Dashboard')
    $table = $book.Worksheets.Item('Employees').ListObjects.Item('tblEmployees')

    # Dashboard
    $active = @($employees | Where-Object Status -eq 'Active').Count
    $onLeave = @($employees | Where-Object Status -eq 'On Leave').Count
    $avg = ($employees | ForEach-Object { [double]$_.'Annual Salary' } | Measure-Object -Average).Average
    Check 'Dashboard tiles (total, active, on leave, average salary)' ($dash.Range('B10').Value2 -eq 100 -and $dash.Range('D10').Value2 -eq $active -and
        $dash.Range('F10').Value2 -eq $onLeave -and [math]::Abs($dash.Range('H10').Value2 - $avg) -lt 0.01) "[100 / $active / $onLeave / $([math]::Round($avg))]"
    $ok = $true
    for ($r = 15; $r -le 20; $r++) {
        $ok = $ok -and $dash.Range("C$r").Value2 -eq @($employees | Where-Object Department -eq $dash.Range("B$r").Text).Count
    }
    Check 'Headcount by department matches data' ($ok -and $dash.Range('C21').Value2 -eq 100)
    Check 'Dashboard button opens the form' ($dash.Buttons().Count -eq 1 -and $dash.Buttons(1).OnAction -like '*OpenEmployeeForm')

    # Form
    $form = (Run-Vba 'TestHarness.FormCheck') -split '\|'
    Check 'Form loads with 30 controls and filled dropdowns (2 genders, 6 departments, 3 statuses)' ($form[0] -eq '30' -and $form[1] -eq '2' -and $form[2] -eq '6' -and $form[3] -eq '3') "[$($form[0..3] -join ', ')]"
    Check 'New mode: next ID shown, Add enabled, Update disabled' ($form[4] -eq 'EMP101 (new)' -and $form[5] -eq 'True' -and $form[6] -eq 'False')
    Check 'Search result loads into the form (edit mode)' ($form[7] -eq $employees[4].'Full Name' -and $form[8] -eq $employees[4].Department -and
        $form[9] -eq $employees[4].Email -and $form[10] -eq 'False' -and $form[11] -eq 'True' -and $form[12] -eq 'True') "[$($form[7]) | $($form[13])]"

    # Search
    $firstName = $employees[9].'Full Name'
    $expectedRow = [array]::IndexOf(@($employees | ForEach-Object 'Full Name'), $firstName) + 1
    Check 'Search by ID (any case), full name, partial name; unknown = not found' ((Run-Vba 'modEmployees.FindEmployee' @('EMP005')) -eq 5 -and
        (Run-Vba 'modEmployees.FindEmployee' @('emp005')) -eq 5 -and (Run-Vba 'modEmployees.FindEmployee' @($firstName)) -eq $expectedRow -and
        (Run-Vba 'modEmployees.FindEmployee' @($firstName.Split(' ')[1].ToLower())) -gt 0 -and (Run-Vba 'modEmployees.FindEmployee' @('Nobody Here')) -eq 0)

    # Validation
    $good = [object[]]@('', 'Jane Doe', 'Female', 'IT', 'Software Engineer', 'jane.doe@company.com', '555-123-4567', '15-Mar-2024', '85000', 'Active')
    Check 'Valid input passes validation' ([string](Run-Vba 'modEmployees.ValidateEmployee' @($good, 0)) -eq '')
    $bad = [object[]]@('', 'J1', 'Unknown', '', '', 'not-an-email', '5551234', '01-Jan-2099', 'abc', '')
    $problems = Run-Vba 'modEmployees.ValidateEmployee' @($bad, 0)
    $lines = @($problems -split "`r`n" | Where-Object { $_ })
    Check 'Invalid input: all 9 problems reported' ($lines.Count -eq 9) "[$($lines.Count) messages, e.g. '$($lines[0])']"
    $dupe = $good.Clone(); $dupe[5] = $employees[0].Email.ToUpper()
    Check 'Duplicate email rejected, but allowed for the same employee on update' ((Run-Vba 'modEmployees.ValidateEmployee' @($dupe, 0)) -like '*already used*' -and
        [string](Run-Vba 'modEmployees.ValidateEmployee' @($dupe, 1)) -eq '')

    # Add through the form
    $input = [object[]]@('', '  jane   doe ', 'Female', 'IT', 'Software Engineer', 'Jane.Doe@Company.com', '555-123-4567', '15-Mar-2024', '85,000', 'Active')
    $added = Run-Vba 'TestHarness.FormAdd' @(, $input)
    $row = $table.ListRows.Item($table.ListRows.Count).Range
    Check 'Add: saved as EMP101 with clean values' ($added -eq 'OK:EMP101' -and $table.ListRows.Count -eq 101 -and $row.Cells.Item(1, 2).Text -eq 'Jane Doe' -and
        $row.Cells.Item(1, 6).Text -eq 'jane.doe@company.com' -and $row.Cells.Item(1, 8).Value2 -is [double] -and $row.Cells.Item(1, 9).Value2 -eq 85000 -and
        $dash.Range('B10').Value2 -eq 101) "[$added | $($row.Cells.Item(1, 2).Text) | $($row.Cells.Item(1, 8).Text)]"
    $rejected = Run-Vba 'TestHarness.FormAdd' @(, [object[]]@('', 'Bad Phone', 'Male', 'Sales', 'Sales Executive', 'bad.phone@company.com', '12345', '01-Jan-2020', '50000', 'Active'))
    Check 'Add with invalid input is refused and nothing is saved' ($rejected -like 'INVALID:*Phone*' -and $table.ListRows.Count -eq 101)

    # Update
    $update = $good.Clone(); $update[0] = 'EMP101'; $update[8] = '90000'; $update[9] = 'On Leave'
    [void](Run-Vba 'modEmployees.SaveEmployee' @($update, 101))
    Check 'Update: row changed in place, dashboard refreshed' ($table.ListRows.Count -eq 101 -and $row.Cells.Item(1, 9).Value2 -eq 90000 -and
        $row.Cells.Item(1, 10).Text -eq 'On Leave' -and $dash.Range('F10').Value2 -eq $onLeave + 1)

    # Delete
    [void](Run-Vba 'modEmployees.DeleteEmployee' @(101))
    Check 'Delete: record removed and no longer found' ($table.ListRows.Count -eq 100 -and (Run-Vba 'modEmployees.FindEmployee' @('EMP101')) -eq 0 -and
        $dash.Range('B10').Value2 -eq 100)
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
