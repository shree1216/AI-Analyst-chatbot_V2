Attribute VB_Name = "BuildForm"
Option Explicit

' Build helper (used only by build_workbook.ps1, then removed from the workbook):
' lays out the controls of the frmEmployee UserForm.

Private Const NAVY As Long = 7949855         ' RGB(31, 78, 121)

Public Sub BuildEmployeeFormLayout()
    Dim designer As Object, ctl As Object, i As Long, y As Single
    Dim fields As Variant, buttons As Variant

    With ThisWorkbook.VBProject.VBComponents("frmEmployee")
        .Properties("Caption") = "Employee Manager"
        .Properties("Width") = 430
        .Properties("Height") = 452
        .Properties("StartUpPosition") = 1
        Set designer = .Designer
    End With

    Set ctl = AddControl(designer, "Label", "lblHeader", 0, 0, 430, 34, "  Employee Management")
    ctl.BackColor = NAVY
    ctl.ForeColor = vbWhite
    ctl.Font.Size = 15
    ctl.Font.Bold = True

    AddControl designer, "Label", "lblSearch", 12, 50, 110, 18, "Search (ID or name):"
    Set ctl = AddControl(designer, "TextBox", "txtSearch", 125, 46, 190, 20)
    ctl.ControlTipText = "Type an Employee ID (EMP001) or a name, then press Enter"
    AddControl designer, "CommandButton", "btnSearch", 322, 45, 88, 22, "Search"
    Set ctl = AddControl(designer, "Label", "lblMode", 12, 78, 398, 16)
    ctl.ForeColor = NAVY
    ctl.Font.Bold = True
    ctl.Font.Italic = True

    ' Control type | name | label | tooltip
    fields = Array( _
        Array("TextBox", "txtID", "Employee ID", "Assigned automatically"), _
        Array("TextBox", "txtName", "Full Name *", "First and last name"), _
        Array("ComboBox", "cboGender", "Gender *", "Select from the list"), _
        Array("ComboBox", "cboDepartment", "Department *", "Select from the list"), _
        Array("TextBox", "txtTitle", "Job Title *", "e.g. Financial Analyst"), _
        Array("TextBox", "txtEmail", "Email *", "name@company.com"), _
        Array("TextBox", "txtPhone", "Phone *", "Format: 555-123-4567"), _
        Array("TextBox", "txtHireDate", "Hire Date *", "e.g. 15-Mar-2024 (not in the future)"), _
        Array("TextBox", "txtSalary", "Annual Salary *", "Number between 10,000 and 1,000,000"), _
        Array("ComboBox", "cboStatus", "Status *", "Select from the list"))
    For i = 0 To UBound(fields)
        y = 102 + 26 * i
        AddControl designer, "Label", "lbl" & Mid$(fields(i)(1), 4), 12, y + 3, 110, 18, fields(i)(2)
        Set ctl = AddControl(designer, fields(i)(0), fields(i)(1), 125, y, 285, 20)
        ctl.ControlTipText = fields(i)(3)
        If fields(i)(0) = "ComboBox" Then ctl.Style = 2      ' fmStyleDropDownList: list values only
    Next i
    With designer.Controls("txtID")
        .Locked = True
        .TabStop = False
        .BackColor = RGB(242, 242, 242)
    End With

    buttons = Array("Add", "Update", "Delete", "Clear", "Close")
    For i = 0 To UBound(buttons)
        Set ctl = AddControl(designer, "CommandButton", "btn" & buttons(i), 12 + 80 * i, 372, 74, 28, CStr(buttons(i)))
        ctl.Font.Bold = True
    Next i
    designer.Controls("btnClose").Cancel = True      ' Esc closes the form
End Sub

Private Function AddControl(ByVal designer As Object, ByVal controlType As String, ByVal name As String, _
                            ByVal x As Single, ByVal y As Single, ByVal w As Single, ByVal h As Single, _
                            Optional ByVal caption As String = "") As Object
    Dim ctl As Object
    Set ctl = designer.Controls.Add("Forms." & controlType & ".1", name)
    ctl.Left = x
    ctl.Top = y
    ctl.Width = w
    ctl.Height = h
    ctl.Font.Name = "Segoe UI"
    ctl.Font.Size = 9
    If Len(caption) > 0 Then ctl.Caption = caption
    Set AddControl = ctl
End Function
