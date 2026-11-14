Attribute VB_Name = "EmployeeWorkbookTests"
Option Explicit

Private checks As Long

Public Function RunEmployeeWorkbookTests() As String
    Dim original As Variant, values As Variant, actual As Variant
    Dim originalCount As Long, newId As String, rowNumber As Long
    Dim rows As Collection, r As Long, c As Long
    Dim failureNumber As Long, failureMessage As String
    On Error GoTo Failed

    checks = 0
    originalCount = EmployeeTable.ListRows.Count
    original = EmployeeTable.DataBodyRange.Value2
    Check originalCount = 100, "The workbook must contain 100 synthetic employees."
    Check ThisWorkbook.Worksheets("Employees").ProtectContents, "Employee records must be protected."
    Check ThisWorkbook.Worksheets("Dashboard").Range("B10").Value2 = originalCount, "Dashboard headcount."

    Load frmEmployee
    Check frmEmployee.lstResults.ListCount = originalCount, "Form loads every employee."
    Check frmEmployee.cboGender.ListCount = 2, "Gender dropdown."
    Check frmEmployee.cboDepartment.ListCount = 6, "Department dropdown."
    Check frmEmployee.cboStatus.ListCount = 3, "Status dropdown."
    Check frmEmployee.cmdAdd.Enabled, "Add is enabled for a new record."
    Check Not frmEmployee.cmdUpdate.Enabled, "Update requires a selection."
    Check Not frmEmployee.cmdDelete.Enabled, "Delete requires a selection."
    Unload frmEmployee

    values = ReadEmployee(1)
    values(COL_ID) = ""
    values(COL_FIRST) = "Validation"
    values(COL_LAST) = "Example"
    values(COL_EMAIL) = "vba-validation@example.invalid"
    values(COL_PHONE) = "0123456789"
    values(COL_HIRE) = "15-Jan-2024"
    values(COL_SALARY) = "650000"
    Check ValidateEmployee(values) = "", "Valid employee is accepted."

    values(COL_PHONE) = "123"
    Check ValidateEmployee(values) <> "", "Short phone is rejected."
    values(COL_PHONE) = "0123456789"
    values(COL_HIRE) = "31-Feb-2024"
    Check ValidateEmployee(values) <> "", "Impossible date is rejected."
    values(COL_HIRE) = Format$(Date + 1, DATE_FORMAT)
    Check ValidateEmployee(values) <> "", "Future date is rejected."
    values(COL_HIRE) = "15-Jan-2024"
    values(COL_SALARY) = "-1"
    Check ValidateEmployee(values) <> "", "Negative salary is rejected."
    values(COL_SALARY) = "650000"
    values(COL_DEPT) = "Unknown"
    Check ValidateEmployee(values) <> "", "Unknown department is rejected."
    values(COL_DEPT) = ThisWorkbook.Names("DepartmentList").RefersToRange.Cells(1, 1).Value2
    values(COL_EMAIL) = original(1, COL_EMAIL)
    Check ValidateEmployee(values) <> "", "Duplicate email is rejected."
    values(COL_EMAIL) = "vba-validation@example.invalid"

    newId = AddEmployee(values)
    Check EmployeeTable.ListRows.Count = originalCount + 1, "Add inserts one row."
    rowNumber = FindEmployeeRow(newId)
    Check rowNumber > 0, "New ID can be found."
    actual = ReadEmployee(rowNumber)
    Check actual(COL_PHONE) = "0123456789", "Phone retains its leading zero."
    Check actual(COL_HIRE) = "15-Jan-2024", "Hire date round-trips."
    Check ValidateEmployee(actual, newId) = "", "Update permits the employee's own email."
    Set rows = SearchEmployees("VBA-VALIDATION@EXAMPLE.INVALID")
    Check rows.Count = 1, "Search is case-insensitive."
    Check CLng(rows(1)) = rowNumber, "Search returns the correct employee."
    ThisWorkbook.Worksheets("Dashboard").Calculate
    Check ThisWorkbook.Worksheets("Dashboard").Range("B10").Value2 = originalCount + 1, "Dashboard reflects add."

    values(COL_TITLE) = "Updated Test Role"
    UpdateEmployee rowNumber, values
    actual = ReadEmployee(rowNumber)
    Check actual(COL_TITLE) = "Updated Test Role", "Update persists the changed title."
    Check actual(COL_ID) = newId, "Update preserves the employee ID."
    Check ThisWorkbook.Worksheets("Employees").ProtectContents, "Update restores protection."

    DeleteEmployee rowNumber
    Check FindEmployeeRow(newId) = 0, "Deleted employee is no longer found."
    newId = ""
    Check EmployeeTable.ListRows.Count = originalCount, "Delete restores the original row count."
    Check ThisWorkbook.Worksheets("Employees").ProtectContents, "Delete restores protection."
    ThisWorkbook.Worksheets("Dashboard").Calculate
    Check ThisWorkbook.Worksheets("Dashboard").Range("B10").Value2 = originalCount, "Dashboard reflects delete."

    actual = EmployeeTable.DataBodyRange.Value2
    For r = 1 To originalCount
        For c = 1 To FIELD_COUNT
            If actual(r, c) <> original(r, c) Then
                Err.Raise vbObjectError + 1101, , "Tests changed an existing employee."
            End If
        Next c
    Next r
    Check True, "Existing employee records remain unchanged."
    RunEmployeeWorkbookTests = checks & " checks passed; original employees unchanged."
    Exit Function

Failed:
    failureNumber = Err.Number
    failureMessage = Err.Description
    If newId <> "" Then
        rowNumber = FindEmployeeRow(newId)
        If rowNumber > 0 Then DeleteEmployee rowNumber
    End If
    Unload frmEmployee
    LockEmployees
    Err.Raise failureNumber, "RunEmployeeWorkbookTests", failureMessage
End Function

Private Sub Check(ByVal passed As Boolean, ByVal description As String)
    If Not passed Then Err.Raise vbObjectError + 1100, , description
    checks = checks + 1
End Sub
