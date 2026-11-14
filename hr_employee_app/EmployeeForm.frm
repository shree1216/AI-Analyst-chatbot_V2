VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} EmployeeForm 
   Caption         =   "HR Employee Management"
   ClientHeight    =   3036
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   4584
   OleObjectBlob   =   "EmployeeForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "EmployeeForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Controls are created in UserForm_Initialize so the form can be imported without a binary .frx file.
Private WithEvents txtSearch As MSForms.TextBox
Attribute txtSearch.VB_VarHelpID = -1
Private WithEvents btnSearch As MSForms.CommandButton
Attribute btnSearch.VB_VarHelpID = -1
Private WithEvents btnShowAll As MSForms.CommandButton
Attribute btnShowAll.VB_VarHelpID = -1
Private WithEvents lstResults As MSForms.ListBox
Attribute lstResults.VB_VarHelpID = -1
Private WithEvents btnAdd As MSForms.CommandButton
Attribute btnAdd.VB_VarHelpID = -1
Private WithEvents btnUpdate As MSForms.CommandButton
Attribute btnUpdate.VB_VarHelpID = -1
Private WithEvents btnDelete As MSForms.CommandButton
Attribute btnDelete.VB_VarHelpID = -1
Private WithEvents btnClear As MSForms.CommandButton
Attribute btnClear.VB_VarHelpID = -1
Private WithEvents btnClose As MSForms.CommandButton
Attribute btnClose.VB_VarHelpID = -1

Private fieldControls As Variant
Private loadedID As String
Private lastSearch As String
Private suppressListEvents As Boolean

Public Property Get LoadedEmployeeID() As String
    LoadedEmployeeID = loadedID
End Property

Public Property Get ResultCount() As Long
    ResultCount = lstResults.ListCount
End Property

Private Sub UserForm_Initialize()
    On Error GoTo Failed
    fieldControls = Array("txtFirstName", "txtLastName", "txtEmail", "txtPhone", "cboDepartment", _
                          "txtJobTitle", "txtHireDate", "txtSalary", "cboStatus")
    Me.Caption = APP_TITLE
    Me.Width = 584
    Me.Height = 470
    Me.BackColor = RGB(248, 249, 251)
    BuildControls
    LoadDropdowns
    ClearFields
    RefreshResults ""
    Exit Sub
Failed:
    Notify "The employee form could not be prepared." & vbCrLf & vbCrLf & Err.Description, vbCritical
End Sub

Private Sub BuildControls()
    Dim leftTitles As Variant, rightTitles As Variant, i As Long, rowTop As Single, box As Object

    AddLabel "lblTitle", "Employee Management", 12, 6, 400, 24, 15, True
    AddLabel "lblSubtitle", "Search and select an employee to update or delete, or enter new details and choose " & _
             "Add New.  * = required", 12, 30, 556, 14, 8, False

    AddLabel "lblSearch", "Search", 12, 54, 50, 16, 9, True
    Set txtSearch = AddControl("Forms.TextBox.1", "txtSearch", 62, 51, 300, 20)
    txtSearch.ControlTipText = "Employee ID, name, email, department, job title, or status. Press Enter to search."
    Set btnSearch = AddButton("btnSearch", "Search", 370, 49, 90)
    Set btnShowAll = AddButton("btnShowAll", "Show All", 468, 49, 92)

    AddLabel "lblColumnID", "Employee ID", 16, 78, 76, 14, 8, True
    AddLabel "lblColumnName", "Name", 94, 78, 170, 14, 8, True
    AddLabel "lblColumnDepartment", "Department", 264, 78, 138, 14, 8, True
    AddLabel "lblColumnStatus", "Status", 402, 78, 100, 14, 8, True
    Set lstResults = AddControl("Forms.ListBox.1", "lstResults", 12, 93, 548, 106)
    lstResults.ColumnCount = 4
    lstResults.ColumnWidths = "78;170;138;100"
    lstResults.ControlTipText = "Click an employee to load the record."

    AddLabel "lblDetails", "Employee Details", 12, 206, 300, 18, 10, True
    leftTitles = Array("Employee ID", "First Name *", "Last Name *", "Email *", "Phone *")
    rightTitles = Array("Department *", "Job Title *", "Hire Date *", "Annual Salary *", "Status *")
    For i = 0 To 4
        rowTop = 228 + i * 26
        AddLabel "lblLeft" & i, CStr(leftTitles(i)), 12, rowTop + 2, 86, 16, 9, False
        AddLabel "lblRight" & i, CStr(rightTitles(i)), 294, rowTop + 2, 88, 16, 9, False
        If i = 0 Then
            Set box = AddControl("Forms.TextBox.1", "txtEmployeeID", 100, rowTop, 176, 20)
            box.Locked = True
            box.TabStop = False
            box.BackColor = RGB(235, 237, 240)
            box.ControlTipText = "Assigned automatically when an employee is added."
        Else
            Set box = AddControl("Forms.TextBox.1", CStr(fieldControls(i - 1)), 100, rowTop, 176, 20)
            box.MaxLength = IIf(i = 3, 100, 50)
        End If
        If i = 0 Or i = 4 Then
            Set box = AddControl("Forms.ComboBox.1", CStr(fieldControls(4 + i)), 384, rowTop, 176, 20)
            box.Style = 2
        Else
            Set box = AddControl("Forms.TextBox.1", CStr(fieldControls(4 + i)), 384, rowTop, 176, 20)
            box.MaxLength = 60
        End If
    Next i
    Me.Controls("txtPhone").ControlTipText = "Example: +1 (555) 010-1234"
    Me.Controls("txtHireDate").ControlTipText = "Use yyyy-mm-dd, for example 2024-03-15."
    Me.Controls("txtSalary").ControlTipText = "Numbers only, for example 65000 or 65000.50."
    AddLabel "lblHint", "Hire Date: yyyy-mm-dd (example 2024-03-15).   Annual Salary: numbers only " & _
             "(example 65000 or 65000.50).", 12, 360, 548, 14, 8, False

    Set btnAdd = AddButton("btnAdd", "Add New", 12, 382, 84)
    Set btnUpdate = AddButton("btnUpdate", "Update", 104, 382, 84)
    Set btnDelete = AddButton("btnDelete", "Delete", 196, 382, 84)
    Set btnClear = AddButton("btnClear", "Clear", 288, 382, 84)
    Set btnClose = AddButton("btnClose", "Close", 476, 382, 84)
    btnClose.Cancel = True
    AddLabel "lblStatus", "", 12, 414, 548, 16, 8, False
    Me.Controls("lblStatus").ForeColor = RGB(31, 78, 121)
End Sub

Private Function AddControl(ByVal progId As String, ByVal controlName As String, ByVal leftPos As Single, _
                            ByVal topPos As Single, ByVal controlWidth As Single, _
                            ByVal controlHeight As Single) As Object
    Dim item As Object
    Set item = Me.Controls.Add(progId, controlName, True)
    item.Left = leftPos
    item.Top = topPos
    item.Width = controlWidth
    item.Height = controlHeight
    item.Font.Name = "Segoe UI"
    item.Font.Size = 9
    Set AddControl = item
End Function

Private Sub AddLabel(ByVal controlName As String, ByVal caption As String, ByVal leftPos As Single, _
                     ByVal topPos As Single, ByVal controlWidth As Single, ByVal controlHeight As Single, _
                     ByVal fontSize As Single, ByVal isBold As Boolean)
    Dim item As Object
    Set item = AddControl("Forms.Label.1", controlName, leftPos, topPos, controlWidth, controlHeight)
    item.Caption = caption
    item.BackStyle = 0
    item.Font.Size = fontSize
    item.Font.Bold = isBold
End Sub

Private Function AddButton(ByVal controlName As String, ByVal caption As String, ByVal leftPos As Single, _
                           ByVal topPos As Single, ByVal controlWidth As Single) As Object
    Dim item As Object
    Set item = AddControl("Forms.CommandButton.1", controlName, leftPos, topPos, controlWidth, 26)
    item.Caption = caption
    item.TakeFocusOnClick = False
    Set AddButton = item
End Function

Private Sub LoadDropdowns()
    FillCombo Me.Controls("cboDepartment"), ListValues("tblDepartments")
    FillCombo Me.Controls("cboStatus"), ListValues("tblStatuses")
End Sub

Private Sub FillCombo(ByVal box As Object, ByVal items As Variant)
    Dim i As Long
    box.Clear
    For i = LBound(items) To UBound(items)
        box.AddItem CStr(items(i))
    Next i
End Sub

Private Sub SetComboValue(ByVal box As Object, ByVal value As String)
    Dim i As Long
    For i = 0 To box.ListCount - 1
        If StrComp(box.List(i), value, vbTextCompare) = 0 Then
            box.ListIndex = i
            Exit Sub
        End If
    Next i
    ' Keep retired list values visible so the user can choose a current value before saving.
    If Len(value) > 0 Then
        box.AddItem value
        box.ListIndex = box.ListCount - 1
    Else
        box.ListIndex = -1
    End If
End Sub

Private Function ReadInputs() As Variant
    Dim values(0 To 8) As Variant, i As Long, value As Variant
    For i = 0 To 8
        value = Me.Controls(fieldControls(i)).Value
        If IsNull(value) Then value = ""
        values(i) = CStr(value)
    Next i
    ReadInputs = values
End Function

Public Sub ClearFields()
    Dim i As Long
    LoadDropdowns
    For i = 0 To 8
        If Left$(fieldControls(i), 3) = "cbo" Then
            Me.Controls(fieldControls(i)).ListIndex = -1
        Else
            Me.Controls(fieldControls(i)).Value = ""
        End If
    Next i
    SetComboValue Me.Controls("cboStatus"), "Active"
    Me.Controls("txtEmployeeID").Value = ""
    loadedID = ""
    suppressListEvents = True
    lstResults.ListIndex = -1
    suppressListEvents = False
    UpdateButtons
    SetStatus "New employee: complete all required fields, then choose Add New."
End Sub

Public Sub PerformSearch(Optional ByVal term As Variant)
    On Error GoTo Failed
    If Not IsMissing(term) Then txtSearch.Value = CStr(term)
    lastSearch = Trim$(txtSearch.Value)
    RefreshResults lastSearch
    If lstResults.ListCount = 0 Then
        SetStatus "No matching employees."
        Notify "No employees match """ & lastSearch & """." & vbCrLf & _
               "Search by Employee ID, name, email, department, job title, or status.", vbInformation
    ElseIf lstResults.ListCount = 1 Then
        SelectResult 0
    Else
        SetStatus lstResults.ListCount & " employees found. Click a row to load it."
    End If
    Exit Sub
Failed:
    Notify "The search could not be completed." & vbCrLf & vbCrLf & Err.Description, vbCritical
End Sub

Public Sub LoadEmployeeByID(ByVal employeeID As String)
    Dim record As Variant, i As Long, value As Variant, box As Object
    record = GetEmployee(employeeID)
    Me.Controls("txtEmployeeID").Value = record(0)
    For i = 0 To 8
        value = record(i + 1)
        Set box = Me.Controls(fieldControls(i))
        If i = 6 And IsDate(value) Then
            box.Value = Format$(value, "yyyy-mm-dd")
        ElseIf i = 7 And IsNumeric(value) And Not IsEmpty(value) Then
            box.Value = Trim$(Str$(Round(CDbl(value), 2)))
        ElseIf Left$(fieldControls(i), 3) = "cbo" Then
            SetComboValue box, CStr(value)
        Else
            box.Value = CStr(value)
        End If
    Next i
    loadedID = CStr(record(0))
    UpdateButtons
    SetStatus "Editing " & loadedID & ". Change details and choose Update, or choose Delete."
End Sub

Public Sub PerformAdd()
    Dim values As Variant, cleaned As Variant, fieldIndex As Long, problem As String, newID As String
    On Error GoTo Failed
    values = ReadInputs()
    problem = ValidateEmployee(values, "", cleaned, fieldIndex)
    If Len(problem) > 0 Then
        ShowValidationProblem problem, fieldIndex
        Exit Sub
    End If
    newID = AddEmployee(values)
    txtSearch.Value = ""
    lastSearch = ""
    RefreshResults ""
    SelectByID newID
    Notify "Employee " & newID & " (" & cleaned(0) & " " & cleaned(1) & ") was added successfully.", _
           vbInformation, "Employee added"
    Exit Sub
Failed:
    Notify "The employee could not be added." & vbCrLf & vbCrLf & Err.Description, vbCritical
End Sub

Public Sub PerformUpdate()
    Dim values As Variant, cleaned As Variant, fieldIndex As Long, problem As String, employeeID As String
    On Error GoTo Failed
    If Len(loadedID) = 0 Then
        Notify "Search for and select an employee before choosing Update.", vbExclamation, "No employee selected"
        Exit Sub
    End If
    employeeID = loadedID
    values = ReadInputs()
    problem = ValidateEmployee(values, employeeID, cleaned, fieldIndex)
    If Len(problem) > 0 Then
        ShowValidationProblem problem, fieldIndex
        Exit Sub
    End If
    If Notify("Save changes to employee " & employeeID & " (" & cleaned(0) & " " & cleaned(1) & ")?", _
              vbYesNo + vbQuestion, "Confirm update") <> vbYes Then
        SetStatus "Update cancelled. No changes were saved."
        Exit Sub
    End If
    UpdateEmployee employeeID, values
    RefreshResults lastSearch
    SelectByID employeeID
    Notify "Employee " & employeeID & " was updated successfully.", vbInformation, "Employee updated"
    Exit Sub
Failed:
    Notify "The employee could not be updated." & vbCrLf & vbCrLf & Err.Description, vbCritical
End Sub

Public Sub PerformDelete()
    Dim record As Variant, employeeID As String, fullName As String
    On Error GoTo Failed
    If Len(loadedID) = 0 Then
        Notify "Search for and select an employee before choosing Delete.", vbExclamation, "No employee selected"
        Exit Sub
    End If
    employeeID = loadedID
    record = GetEmployee(employeeID)
    fullName = CStr(record(1)) & " " & CStr(record(2))
    If Notify("Permanently delete employee " & employeeID & " (" & fullName & ")?" & vbCrLf & vbCrLf & _
              "This cannot be undone.", vbYesNo + vbExclamation + vbDefaultButton2, "Confirm delete") <> vbYes Then
        SetStatus "Delete cancelled. " & employeeID & " was not changed."
        Exit Sub
    End If
    DeleteEmployee employeeID
    ClearFields
    RefreshResults lastSearch
    Notify "Employee " & employeeID & " (" & fullName & ") was deleted.", vbInformation, "Employee deleted"
    Exit Sub
Failed:
    Notify "The employee could not be deleted." & vbCrLf & vbCrLf & Err.Description, vbCritical
End Sub

Private Sub RefreshResults(ByVal term As String)
    Dim results As Variant, i As Long, c As Long
    suppressListEvents = True
    lstResults.Clear
    results = SearchEmployees(term)
    If Not IsEmpty(results) Then
        For i = 1 To UBound(results, 1)
            lstResults.AddItem CStr(results(i, 1))
            For c = 2 To 4
                lstResults.List(lstResults.ListCount - 1, c - 1) = CStr(results(i, c))
            Next c
        Next i
    End If
    suppressListEvents = False
End Sub

Private Sub SelectResult(ByVal index As Long)
    suppressListEvents = True
    lstResults.ListIndex = index
    suppressListEvents = False
    LoadEmployeeByID CStr(lstResults.List(index, 0))
End Sub

Private Sub SelectByID(ByVal employeeID As String)
    Dim i As Long
    For i = 0 To lstResults.ListCount - 1
        If StrComp(lstResults.List(i, 0), employeeID, vbTextCompare) = 0 Then
            SelectResult i
            Exit Sub
        End If
    Next i
    LoadEmployeeByID employeeID
End Sub

Private Sub ShowValidationProblem(ByVal problem As String, ByVal fieldIndex As Long)
    SetStatus "Please correct the highlighted field."
    Notify problem, vbExclamation, "Please check the employee details"
    If fieldIndex >= 0 And fieldIndex <= 8 Then
        On Error Resume Next
        Me.Controls(fieldControls(fieldIndex)).SetFocus
        On Error GoTo 0
    End If
End Sub

Private Sub UpdateButtons()
    btnUpdate.Enabled = (Len(loadedID) > 0)
    btnDelete.Enabled = (Len(loadedID) > 0)
End Sub

Private Sub SetStatus(ByVal text As String)
    Me.Controls("lblStatus").Caption = text
End Sub

Private Sub btnSearch_Click()
    PerformSearch
End Sub

Private Sub btnShowAll_Click()
    PerformSearch ""
End Sub

Private Sub txtSearch_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = 0
        PerformSearch
    End If
End Sub

Private Sub lstResults_Click()
    If suppressListEvents Then Exit Sub
    If lstResults.ListIndex < 0 Then Exit Sub
    On Error GoTo Failed
    LoadEmployeeByID CStr(lstResults.List(lstResults.ListIndex, 0))
    Exit Sub
Failed:
    Notify "The employee could not be loaded." & vbCrLf & vbCrLf & Err.Description, vbCritical
End Sub

Private Sub btnAdd_Click()
    PerformAdd
End Sub

Private Sub btnUpdate_Click()
    PerformUpdate
End Sub

Private Sub btnDelete_Click()
    PerformDelete
End Sub

Private Sub btnClear_Click()
    txtSearch.Value = ""
    lastSearch = ""
    RefreshResults ""
    ClearFields
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub
