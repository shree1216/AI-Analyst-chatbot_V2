VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmEmployee 
   Caption         =   "HR Employee Manager"
   ClientHeight    =   8832.001
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   14268
   OleObjectBlob   =   "frmEmployee.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmEmployee"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Row of the employee selected in the results list (0 = adding a new employee)
Private selectedRow As Long

Private Sub UserForm_Initialize()
    cboGender.List = ListValues("GenderList")
    cboDepartment.List = ListValues("DepartmentList")
    cboStatus.List = ListValues("StatusList")
    lstResults.ColumnCount = 5
    lstResults.ColumnWidths = "50;104;98;56;0"
    ShowResults SearchEmployees("")
    ClearForm
End Sub

' ----- Buttons ---------------------------------------------------------------

Private Sub cmdAdd_Click()
    Dim values As Variant, problem As String, newId As String
    values = FormValues()
    problem = ValidateEmployee(values)
    If problem <> "" Then
        MsgBox problem, vbExclamation, "Please check the details"
        Exit Sub
    End If
    If MsgBox("Add " & values(COL_FIRST) & " " & values(COL_LAST) & " as a new employee?", _
              vbQuestion + vbYesNo, "Confirm new employee") = vbNo Then Exit Sub

    newId = AddEmployee(values)
    MsgBox values(COL_FIRST) & " " & values(COL_LAST) & " was added as " & newId & ".", vbInformation, "Employee added"
    txtSearch.value = ""
    ShowResults SearchEmployees("")
    ClearForm
End Sub

Private Sub cmdUpdate_Click()
    Dim values As Variant, problem As String
    If selectedRow = 0 Then
        MsgBox "Select an employee from the list first.", vbExclamation, "No employee selected"
        Exit Sub
    End If
    values = FormValues()
    problem = ValidateEmployee(values, txtEmployeeID.value)
    If problem <> "" Then
        MsgBox problem, vbExclamation, "Please check the details"
        Exit Sub
    End If
    If MsgBox("Save changes to " & txtEmployeeID.value & "?", vbQuestion + vbYesNo, "Confirm update") = vbNo Then Exit Sub

    UpdateEmployee selectedRow, values
    MsgBox "Employee " & txtEmployeeID.value & " was updated.", vbInformation, "Employee updated"
    ShowResults SearchEmployees(txtSearch.value)
    ClearForm
End Sub

Private Sub cmdDelete_Click()
    If selectedRow = 0 Then
        MsgBox "Select an employee from the list first.", vbExclamation, "No employee selected"
        Exit Sub
    End If
    If MsgBox("Delete " & txtEmployeeID.value & " (" & txtFirstName.value & " " & txtLastName.value & ")?" & _
              vbCrLf & vbCrLf & "This cannot be undone.", vbExclamation + vbYesNo + vbDefaultButton2, _
              "Confirm delete") = vbNo Then Exit Sub

    DeleteEmployee selectedRow
    MsgBox "Employee " & txtEmployeeID.value & " was deleted.", vbInformation, "Employee deleted"
    ShowResults SearchEmployees(txtSearch.value)
    ClearForm
End Sub

Private Sub cmdSearch_Click()
    Dim results As Collection
    Set results = SearchEmployees(txtSearch.value)
    If results.Count = 0 Then
        MsgBox "No employees match """ & txtSearch.value & """.", vbInformation, "Search"
    End If
    ShowResults results
End Sub

Private Sub cmdShowAll_Click()
    txtSearch.value = ""
    ShowResults SearchEmployees("")
End Sub

Private Sub cmdClear_Click()
    ClearForm
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub txtSearch_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyReturn Then cmdSearch_Click
End Sub

' Selecting an employee in the list loads their details for Update / Delete
Private Sub lstResults_Click()
    Dim values As Variant
    If lstResults.ListIndex < 0 Then Exit Sub
    selectedRow = CLng(lstResults.List(lstResults.ListIndex, 4))
    values = ReadEmployee(selectedRow)
    txtEmployeeID.value = values(COL_ID)
    txtFirstName.value = values(COL_FIRST)
    txtLastName.value = values(COL_LAST)
    cboGender.value = values(COL_GENDER)
    cboDepartment.value = values(COL_DEPT)
    txtJobTitle.value = values(COL_TITLE)
    txtEmail.value = values(COL_EMAIL)
    txtPhone.value = values(COL_PHONE)
    txtHireDate.value = values(COL_HIRE)
    txtSalary.value = values(COL_SALARY)
    cboStatus.value = values(COL_STATUS)
    SetMode True
End Sub

' ----- Helpers ---------------------------------------------------------------

Private Function FormValues() As Variant
    Dim values(1 To FIELD_COUNT) As Variant
    values(COL_ID) = txtEmployeeID.value
    values(COL_FIRST) = Trim$(txtFirstName.value)
    values(COL_LAST) = Trim$(txtLastName.value)
    values(COL_GENDER) = cboGender.value
    values(COL_DEPT) = cboDepartment.value
    values(COL_TITLE) = Trim$(txtJobTitle.value)
    values(COL_EMAIL) = Trim$(txtEmail.value)
    values(COL_PHONE) = Trim$(txtPhone.value)
    values(COL_HIRE) = Trim$(txtHireDate.value)
    values(COL_SALARY) = Trim$(txtSalary.value)
    values(COL_STATUS) = cboStatus.value
    FormValues = values
End Function

Private Sub ShowResults(rows As Collection)
    Dim rowNumber As Variant, values As Variant, item As Long
    lstResults.Clear
    For Each rowNumber In rows
        values = ReadEmployee(rowNumber)
        lstResults.AddItem values(COL_ID)
        item = lstResults.ListCount - 1
        lstResults.List(item, 1) = values(COL_FIRST) & " " & values(COL_LAST)
        lstResults.List(item, 2) = values(COL_DEPT)
        lstResults.List(item, 3) = values(COL_STATUS)
        lstResults.List(item, 4) = rowNumber
    Next rowNumber
    lblCount.caption = rows.Count & " employee(s) shown"
End Sub

Private Sub ClearForm()
    txtEmployeeID.value = NextEmployeeId() & "  (new)"
    txtFirstName.value = ""
    txtLastName.value = ""
    cboGender.ListIndex = -1
    cboDepartment.ListIndex = -1
    txtJobTitle.value = ""
    txtEmail.value = ""
    txtPhone.value = ""
    txtHireDate.value = Format$(Date, DATE_FORMAT)
    txtSalary.value = ""
    cboStatus.value = "Active"
    lstResults.ListIndex = -1
    SetMode False
    If Me.Visible Then txtFirstName.SetFocus
End Sub

' Add is available for a new record; Update/Delete only after selecting one
Private Sub SetMode(ByVal editing As Boolean)
    If Not editing Then selectedRow = 0
    cmdAdd.Enabled = Not editing
    cmdUpdate.Enabled = editing
    cmdDelete.Enabled = editing
End Sub

Private Function ListValues(ByVal listName As String) As Variant
    ListValues = Application.WorksheetFunction.Transpose(ThisWorkbook.Names(listName).RefersToRange.value)
End Function


