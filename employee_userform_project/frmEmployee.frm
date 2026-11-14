VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmEmployee 
   Caption         =   "Employee Manager"
   ClientHeight    =   8472.001
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   8388.001
   OleObjectBlob   =   "frmEmployee.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmEmployee"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Employee Manager form. Data rules and storage are in modEmployees.

Private currentRow As Long   ' table row of the employee shown (0 = new employee)

Private Sub UserForm_Initialize()
    cboGender.list = GenderList
    cboDepartment.list = DepartmentList
    cboStatus.list = StatusList
    ClearForm
End Sub

' ---------- Buttons ----------------------------------------------------------------------------

Private Sub btnSearch_Click()
    Dim r As Long
    If Len(Trim$(txtSearch.text)) = 0 Then
        MsgBox "Enter an Employee ID or name to search.", vbExclamation, "Search"
        Exit Sub
    End If
    r = FindEmployee(txtSearch.text)
    If r = 0 Then
        MsgBox "No employee found for '" & txtSearch.text & "'.", vbExclamation, "Search"
    Else
        LoadEmployee r
    End If
End Sub

Private Sub btnAdd_Click()
    Dim problems As String, newID As String
    problems = ValidateEmployee(FormValues)
    If Len(problems) > 0 Then
        MsgBox "Please correct the following:" & vbNewLine & vbNewLine & problems, vbExclamation, "Invalid Input"
        Exit Sub
    End If
    newID = SaveEmployee(FormValues)
    MsgBox "Employee " & newID & " (" & Trim$(txtName.text) & ") was added successfully.", vbInformation, "Employee Added"
    ClearForm
End Sub

Private Sub btnUpdate_Click()
    Dim problems As String
    problems = ValidateEmployee(FormValues, currentRow)
    If Len(problems) > 0 Then
        MsgBox "Please correct the following:" & vbNewLine & vbNewLine & problems, vbExclamation, "Invalid Input"
        Exit Sub
    End If
    If MsgBox("Save the changes to employee " & txtID.text & "?", vbQuestion + vbYesNo, "Confirm Update") = vbNo Then Exit Sub
    SaveEmployee FormValues, currentRow
    MsgBox "Employee " & txtID.text & " was updated successfully.", vbInformation, "Employee Updated"
    LoadEmployee currentRow
End Sub

Private Sub btnDelete_Click()
    If MsgBox("Delete employee " & txtID.text & " - " & txtName.text & "?" & vbNewLine & vbNewLine & _
              "This cannot be undone.", vbExclamation + vbYesNo + vbDefaultButton2, "Confirm Delete") = vbNo Then Exit Sub
    DeleteEmployee currentRow
    MsgBox "Employee " & txtID.text & " was deleted.", vbInformation, "Employee Deleted"
    ClearForm
End Sub

Private Sub btnClear_Click()
    ClearForm
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub

Private Sub txtSearch_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyReturn Then
        btnSearch_Click
        KeyCode = 0
    End If
End Sub

' ---------- Form helpers -----------------------------------------------------------------------

' Shows the employee in table row r and switches the form to edit mode.
Public Sub LoadEmployee(ByVal r As Long)
    Dim v As Variant
    v = EmployeeTable.ListRows(r).Range.value
    txtID.text = v(1, F_ID + 1)
    txtName.text = v(1, F_NAME + 1)
    cboGender.value = v(1, F_GENDER + 1)
    cboDepartment.value = v(1, F_DEPARTMENT + 1)
    txtTitle.text = v(1, F_TITLE + 1)
    txtEmail.text = v(1, F_EMAIL + 1)
    txtPhone.text = v(1, F_PHONE + 1)
    txtHireDate.text = Format$(v(1, F_HIRE_DATE + 1), "dd-mmm-yyyy")
    txtSalary.text = Format$(v(1, F_SALARY + 1), "0")
    cboStatus.value = v(1, F_STATUS + 1)
    currentRow = r
    SetMode "Editing employee " & txtID.text, True
End Sub

' Empties the form for a new employee.
Public Sub ClearForm()
    txtSearch.text = ""
    txtID.text = NextEmployeeID & " (new)"
    txtName.text = ""
    cboGender.value = Null
    cboDepartment.value = Null
    txtTitle.text = ""
    txtEmail.text = ""
    txtPhone.text = ""
    txtHireDate.text = Format$(Date, "dd-mmm-yyyy")
    txtSalary.text = ""
    cboStatus.value = "Active"
    currentRow = 0
    SetMode "New employee - fill in the details and click Add", False
End Sub

Private Sub SetMode(ByVal message As String, ByVal editing As Boolean)
    lblMode.caption = message
    btnAdd.Enabled = Not editing
    btnUpdate.Enabled = editing
    btnDelete.Enabled = editing
End Sub

' Collects the form inputs in the order used by modEmployees (F_ID ... F_STATUS).
Public Function FormValues() As Variant
    FormValues = Array(txtID.text, Trim$(txtName.text), cboGender.text, cboDepartment.text, Trim$(txtTitle.text), _
                       Trim$(txtEmail.text), Trim$(txtPhone.text), Trim$(txtHireDate.text), _
                       Replace(Trim$(txtSalary.text), ",", ""), cboStatus.text)
End Function


