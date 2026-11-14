Attribute VB_Name = "EmployeeAppTests"
Option Explicit

Private checksPassed As Long

' Run from Alt+F8 to verify the application. Test records are removed afterwards.
Public Sub RunEmployeeAppTestsWithMessage()
    Dim result As String
    result = RunEmployeeAppTests()
    MsgBox result, IIf(Left$(result, 4) = "PASS", vbInformation, vbCritical), "Employee application tests"
End Sub

Public Function RunEmployeeAppTests() As String
    Dim f As EmployeeForm, tbl As ListObject, ws As Worksheet, dashboard As Worksheet
    Dim baseline As Long, snapshot As Variant, newID As String, expectedID As String, record As Variant
    Dim cases As Variant, testCase As Variant, i As Long, cleaned As Variant, fieldIndex As Long, problem As String
    Dim phoneCell As Range, existingEmail As String, errorText As String

    checksPassed = 0
    On Error GoTo Failed
    SilentMode = True
    SilentResponse = vbYes
    Set tbl = EmployeeTable()
    Set ws = tbl.Parent
    Set dashboard = ThisWorkbook.Worksheets(DASHBOARD_SHEET)
    baseline = tbl.ListRows.Count
    snapshot = tbl.Range.Formula
    existingEmail = CStr(tbl.DataBodyRange.Cells(1, ColumnIndex(tbl, "Email")).Value)
    expectedID = NextEmployeeID()

    Set f = New EmployeeForm
    Check f.Controls("cboDepartment").ListCount = UBound(ListValues("tblDepartments")) + 1, "department dropdown"
    Check f.Controls("cboStatus").ListCount = UBound(ListValues("tblStatuses")) + 1, "status dropdown"
    Check f.Controls("cboDepartment").Style = 2, "department dropdown accepts list values only"
    Check f.ResultCount = baseline, "form lists every employee on open"
    Check f.Controls("btnUpdate").Enabled = False And f.Controls("btnDelete").Enabled = False, _
          "Update and Delete disabled until a record is selected"
    Check f.Controls("txtEmployeeID").Locked, "Employee ID is read-only"

    FillForm f, Array("Testa", "O'Neil-Smith", "Testa.ONeil@Example.com", "+1 (555) 010-9999", "Finance", _
                      "Financial Analyst", "2024-03-15", "72,500.50", "Active")
    f.PerformAdd
    newID = f.LoadedEmployeeID
    Check tbl.ListRows.Count = baseline + 1, "add creates one row"
    Check newID = expectedID, "add assigns next Employee ID " & expectedID
    Check InStr(LastMessage, "added successfully") > 0, "add success confirmation"
    record = GetEmployee(newID)
    Check record(1) = "Testa" And record(2) = "O'Neil-Smith", "names saved"
    Check record(3) = "testa.oneil@example.com", "email normalized to lower case"
    Set phoneCell = tbl.DataBodyRange.Cells(FindEmployeeRow(newID), ColumnIndex(tbl, "Phone"))
    Check record(4) = "+1 (555) 010-9999" And Not phoneCell.HasFormula, "phone saved as text, not formula"
    Check VarType(record(7)) = vbDate And record(7) = DateSerial(2024, 3, 15), "hire date saved as real date"
    Check VarType(record(8)) = vbDouble And Abs(record(8) - 72500.5) < 0.001, "salary saved as number"
    Check IsDate(record(10)), "last updated timestamp"
    Check ws.ProtectContents, "employee sheet protected after add"
    Check f.Controls("btnUpdate").Enabled And f.Controls("btnDelete").Enabled, "new record loaded for editing"
    dashboard.Calculate
    Check dashboard.Range("B9").Value = baseline + 1, "dashboard total includes new employee"

    cases = Array( _
        Array(0, "", "First Name is required"), _
        Array(0, "=HYPERLINK(""x"")", "First Name must start"), _
        Array(1, "Smith2", "Last Name must start"), _
        Array(2, "", "Email is required"), _
        Array(2, "not-an-email", "Email must be a valid"), _
        Array(2, "a@b", "Email must be a valid"), _
        Array(2, existingEmail, "already used by employee"), _
        Array(3, "12345", "Phone must contain"), _
        Array(3, "555-0100 ext=2", "Phone must contain"), _
        Array(5, "@SUM(A1)", "Job Title must start"), _
        Array(6, "2026-02-30", "Hire Date must be a real date"), _
        Array(6, "15/03/2024", "Hire Date must be a real date"), _
        Array(6, Format$(Date + 1, "yyyy-mm-dd"), "cannot be in the future"), _
        Array(6, "1949-12-31", "earlier than 1950"), _
        Array(7, "", "Annual Salary is required"), _
        Array(7, "-5", "Annual Salary must be"), _
        Array(7, "abc", "Annual Salary must be"), _
        Array(7, "100.555", "Annual Salary must be"), _
        Array(7, "0", "Annual Salary must be"), _
        Array(7, "10000000.01", "Annual Salary must be"))
    For Each testCase In cases
        i = i + 1
        f.ClearFields
        FillForm f, ValidValues("qa.case" & i & "@example.com")
        f.Controls(FormFieldName(CLng(testCase(0)))).Value = testCase(1)
        f.PerformAdd
        Check tbl.ListRows.Count = baseline + 1, "rejected input not saved: " & testCase(2)
        Check InStr(1, LastMessage, CStr(testCase(2)), vbTextCompare) > 0, _
              "error message '" & testCase(2) & "' (got: " & LastMessage & ")"
    Next testCase

    problem = ValidateEmployee(Array("Ann", "Lee", "ann.lee@example.com", "555 010 1234", "Unknown Dept", _
                                     "Analyst", "2024-01-01", "50000", "Active"), "", cleaned, fieldIndex)
    Check InStr(problem, "Department") > 0 And fieldIndex = 4, "department must come from the list"
    problem = ValidateEmployee(Array("Ann", "Lee", "ann.lee@example.com", "555 010 1234", "finance", _
                                     "Analyst", "2024-01-01", "50000", "retired"), "", cleaned, fieldIndex)
    Check InStr(problem, "Status") > 0 And fieldIndex = 8, "status must come from the list"
    problem = ValidateEmployee(Array("Ann", "Lee", "ann.lee@example.com", "555 010 1234", "finance", _
                                     "Analyst", "2024-01-01", "50000", "active"), "", cleaned, fieldIndex)
    Check Len(problem) = 0 And cleaned(4) = "Finance" And cleaned(8) = "Active", "list values matched to canonical case"

    f.PerformSearch newID
    Check f.ResultCount = 1 And f.LoadedEmployeeID = newID, "search by Employee ID loads record"
    Check f.Controls("txtFirstName").Value = "Testa" And f.Controls("txtHireDate").Value = "2024-03-15" And _
          f.Controls("txtSalary").Value = "72500.5" And f.Controls("cboDepartment").Value = "Finance", _
          "search fills form fields"
    f.PerformSearch "o'neil"
    Check f.ResultCount >= 1, "partial name search"
    f.PerformSearch "zz-no-such-employee"
    Check f.ResultCount = 0 And InStr(LastMessage, "No employees match") > 0, "no-match message"
    f.PerformSearch ""
    Check f.ResultCount = baseline + 1, "blank search shows all employees"

    f.LoadEmployeeByID newID
    f.Controls("txtJobTitle").Value = "Senior Financial Analyst"
    SilentResponse = vbNo
    f.PerformUpdate
    Check InStr(LastMessage, "Save changes to employee " & newID) > 0, "update asks for confirmation"
    Check GetEmployee(newID)(6) = "Financial Analyst", "declined update leaves record unchanged"
    SilentResponse = vbYes
    f.PerformUpdate
    Check GetEmployee(newID)(6) = "Senior Financial Analyst", "confirmed update saves changes"
    Check tbl.ListRows.Count = baseline + 1, "update does not add rows"
    Check InStr(LastMessage, "updated successfully") > 0, "update success confirmation"
    Check GetEmployee(newID)(3) = "testa.oneil@example.com", "own email is not treated as a duplicate"
    f.Controls("txtSalary").Value = "0"
    f.PerformUpdate
    Check Abs(GetEmployee(newID)(8) - 72500.5) < 0.001 And InStr(LastMessage, "Annual Salary") > 0, _
          "invalid update rejected"
    Check ws.ProtectContents, "employee sheet protected after update"

    f.ClearFields
    f.PerformUpdate
    Check InStr(LastMessage, "before choosing Update") > 0, "update without selection rejected"
    f.PerformDelete
    Check InStr(LastMessage, "before choosing Delete") > 0, "delete without selection rejected"

    f.LoadEmployeeByID newID
    SilentResponse = vbNo
    f.PerformDelete
    Check InStr(LastMessage, "Permanently delete employee " & newID) > 0, "delete asks for confirmation"
    Check FindEmployeeRow(newID) > 0, "declined delete keeps employee"
    SilentResponse = vbYes
    f.PerformDelete
    Check FindEmployeeRow(newID) = 0 And tbl.ListRows.Count = baseline, "confirmed delete removes employee"
    Check InStr(LastMessage, "was deleted") > 0, "delete success confirmation"
    Check f.LoadedEmployeeID = "" And f.Controls("txtFirstName").Value = "", "form cleared after delete"
    Check ws.ProtectContents, "employee sheet protected after delete"
    dashboard.Calculate
    Check dashboard.Range("B9").Value = baseline, "dashboard total restored"
    Check RangesMatch(snapshot, tbl.Range.Formula), "existing employee records unchanged"

    Unload f
    Set f = Nothing
    SilentMode = False
    RunEmployeeAppTests = "PASS: " & checksPassed & " checks passed."
    Exit Function
Failed:
    errorText = Err.Description
    Resume CleanUp
CleanUp:
    On Error Resume Next
    If Len(newID) > 0 Then
        If FindEmployeeRow(newID) > 0 Then DeleteEmployee newID
    End If
    If Not f Is Nothing Then Unload f
    SilentMode = False
    RunEmployeeAppTests = "FAIL after " & checksPassed & " passing checks: " & errorText
End Function

Public Sub PreviewEmployeeForm()
    EmployeeForm.Show vbModeless
End Sub

Public Sub ClosePreviewEmployeeForm()
    Unload EmployeeForm
End Sub

Private Sub Check(ByVal condition As Boolean, ByVal description As String)
    If Not condition Then Err.Raise vbObjectError + 3900, "EmployeeAppTests", "Check failed: " & description
    checksPassed = checksPassed + 1
End Sub

Private Function ValidValues(ByVal email As String) As Variant
    ValidValues = Array("Quinn", "Tester", email, "555-010-2000", "Operations", "Coordinator", _
                        "2023-06-01", "48000", "Active")
End Function

Private Function FormFieldName(ByVal index As Long) As String
    FormFieldName = Array("txtFirstName", "txtLastName", "txtEmail", "txtPhone", "cboDepartment", _
                          "txtJobTitle", "txtHireDate", "txtSalary", "cboStatus")(index)
End Function

Private Sub FillForm(ByVal f As Object, ByVal values As Variant)
    Dim i As Long
    For i = 0 To 8
        f.Controls(FormFieldName(i)).Value = values(i)
    Next i
End Sub

Private Function RangesMatch(ByVal expected As Variant, ByVal actual As Variant) As Boolean
    Dim r As Long, c As Long
    If UBound(expected, 1) <> UBound(actual, 1) Or UBound(expected, 2) <> UBound(actual, 2) Then Exit Function
    For r = 1 To UBound(expected, 1)
        For c = 1 To UBound(expected, 2)
            If CStr(expected(r, c)) <> CStr(actual(r, c)) Then Exit Function
        Next c
    Next r
    RangesMatch = True
End Function
