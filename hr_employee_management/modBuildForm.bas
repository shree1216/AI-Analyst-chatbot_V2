Attribute VB_Name = "modBuildForm"
Option Explicit

' Build-time only: build_workbook.ps1 imports this module, runs BuildEmployeeForm,
' then removes the module before saving. It is not part of the finished workbook.

Private Const WINE As Long = 4402555        ' RGB(123, 45, 67)
Private Const WINE_SOFT As Long = 15591926  ' RGB(246, 233, 237)
Private Const INK As Long = 1776927         ' RGB(31, 29, 27)
Private Const MUTED As Long = 6317419       ' RGB(107, 101, 96)
Private Const CANVAS As Long = 16185594     ' RGB(250, 248, 246)
Private Const BAND As Long = 15462387        ' RGB(243, 239, 235)
Private Const LINE_COLOR As Long = 14541543 ' RGB(231, 226, 221)
Private Const WHITE As Long = 16777215

Public Sub BuildEmployeeForm(ByVal codePath As String)
    Dim component As Object, designer As Object, details As Object, search As Object
    Dim fields As Variant, i As Long, topPos As Single, ctl As Object, fileNumber As Integer, code As String

    Set component = ThisWorkbook.VBProject.VBComponents.Add(3)   ' vbext_ct_MSForm
    component.Name = "frmEmployee"
    component.Properties("Caption") = "HR Employee Manager"
    component.Properties("Width") = 724
    component.Properties("Height") = 470
    component.Properties("BackColor") = CANVAS
    Set designer = component.Designer

    ' Header band
    With AddControl(designer, "Forms.Label.1", "lblHeader", 0, 0, 724, 48)
        .Caption = "": .BackStyle = 1: .BackColor = WINE
    End With
    With AddLabel(designer, "lblTitle", "Employee Management", 16, 7, 400)
        .Height = 22: .ForeColor = WHITE
    End With
    SetFont designer.Controls("lblTitle"), 14, True
    AddLabel(designer, "lblSubtitle", "Add, search, update and delete employee records", 16, 29, 400).ForeColor = WINE_SOFT

    ' Employee details
    Set details = AddFrame(designer, "fraDetails", "Employee details", 12, 58, 336, 316)
    fields = Array( _
        Array("Employee ID", "txtEmployeeID", "Forms.TextBox.1"), _
        Array("First Name *", "txtFirstName", "Forms.TextBox.1"), _
        Array("Last Name *", "txtLastName", "Forms.TextBox.1"), _
        Array("Gender *", "cboGender", "Forms.ComboBox.1"), _
        Array("Department *", "cboDepartment", "Forms.ComboBox.1"), _
        Array("Job Title *", "txtJobTitle", "Forms.TextBox.1"), _
        Array("Email *", "txtEmail", "Forms.TextBox.1"), _
        Array("Phone (10 digits) *", "txtPhone", "Forms.TextBox.1"), _
        Array("Hire Date *", "txtHireDate", "Forms.TextBox.1"), _
        Array("Annual Salary *", "txtSalary", "Forms.TextBox.1"), _
        Array("Status *", "cboStatus", "Forms.ComboBox.1"))
    For i = 0 To UBound(fields)
        topPos = 14 + i * 26
        AddLabel details, "lbl" & Mid$(fields(i)(1), 4), CStr(fields(i)(0)), 12, topPos + 3, 112
        Set ctl = AddControl(details, CStr(fields(i)(2)), CStr(fields(i)(1)), 128, topPos, 192, 20)
        ctl.BorderStyle = 1: ctl.BorderColor = LINE_COLOR
        If fields(i)(2) = "Forms.ComboBox.1" Then ctl.Style = 2   ' dropdown list only
    Next i
    With details.Controls("txtEmployeeID")
        .Locked = True: .BackColor = BAND: .TabStop = False
    End With
    details.Controls("txtHireDate").ControlTipText = "Format dd-mmm-yyyy, for example 15-Jan-2024"
    details.Controls("txtSalary").ControlTipText = "Numbers only, for example 650000"
    details.Controls("txtEmail").ControlTipText = "For example name@company.com"

    ' Search
    Set search = AddFrame(designer, "fraSearch", "Find employees", 360, 58, 352, 316)
    With AddControl(search, "Forms.TextBox.1", "txtSearch", 12, 14, 190, 20)
        .BorderStyle = 1: .BorderColor = LINE_COLOR
        .ControlTipText = "Search by ID, name, department, job title, email or status"
    End With
    AddButton search, "cmdSearch", "Search", 208, 12, 62, True
    AddButton search, "cmdShowAll", "Show All", 276, 12, 62, False
    With AddControl(search, "Forms.ListBox.1", "lstResults", 12, 46, 326, 222)
        .BorderStyle = 1: .BorderColor = LINE_COLOR
    End With
    AddLabel(search, "lblCount", "", 12, 274, 326).ForeColor = MUTED
    AddLabel search, "lblHint", "Select an employee to load it for Update or Delete.", 12, 290, 326

    ' Action buttons and footer
    AddButton designer, "cmdAdd", "Add Employee", 12, 384, 104, True
    AddButton designer, "cmdUpdate", "Update", 124, 384, 84, True
    AddButton(designer, "cmdDelete", "Delete", 216, 384, 84, False).ForeColor = WINE
    AddButton designer, "cmdClear", "Clear Form", 308, 384, 84, False
    AddButton designer, "cmdClose", "Close", 628, 384, 84, False
    AddLabel(designer, "lblFooter", "* Required. Hire date as dd-mmm-yyyy (e.g. 15-Jan-2024); " & _
             "salary as a number without commas.", 12, 420, 600).ForeColor = MUTED

    ' Form code
    fileNumber = FreeFile
    Open codePath For Input As #fileNumber
    code = Input$(LOF(fileNumber), fileNumber)
    Close #fileNumber
    With component.CodeModule
        If .CountOfLines > 0 Then .DeleteLines 1, .CountOfLines
        .AddFromString code
    End With
End Sub

Private Function AddControl(parent As Object, progId As String, controlName As String, _
                            leftPos As Single, topPos As Single, widthValue As Single, heightValue As Single) As Object
    Set AddControl = parent.Controls.Add(progId, controlName, True)
    With AddControl
        .Left = leftPos: .Top = topPos: .Width = widthValue: .Height = heightValue
    End With
    SetFont AddControl, 9, False
End Function

' Each control gets its own font object; changing the shared design-time font would restyle every control.
Private Sub SetFont(ctl As Object, ByVal size As Single, ByVal bold As Boolean)
    Dim font As New StdFont
    font.Name = "Segoe UI"
    font.Size = size
    font.Bold = bold
    Set ctl.Font = font
End Sub

Private Function AddLabel(parent As Object, controlName As String, caption As String, _
                          leftPos As Single, topPos As Single, widthValue As Single) As Object
    Set AddLabel = AddControl(parent, "Forms.Label.1", controlName, leftPos, topPos, widthValue, 16)
    AddLabel.Caption = caption: AddLabel.BackStyle = 0: AddLabel.ForeColor = INK
End Function

Private Function AddButton(parent As Object, controlName As String, caption As String, _
                           leftPos As Single, topPos As Single, widthValue As Single, primary As Boolean) As Object
    Set AddButton = AddControl(parent, "Forms.CommandButton.1", controlName, leftPos, topPos, widthValue, 28)
    AddButton.Caption = caption
    If primary Then
        AddButton.BackColor = WINE: AddButton.ForeColor = WHITE
        SetFont AddButton, 9, True
    Else
        AddButton.BackColor = WHITE: AddButton.ForeColor = INK
    End If
End Function

Private Function AddFrame(parent As Object, controlName As String, caption As String, _
                          leftPos As Single, topPos As Single, widthValue As Single, heightValue As Single) As Object
    Set AddFrame = AddControl(parent, "Forms.Frame.1", controlName, leftPos, topPos, widthValue, heightValue)
    AddFrame.Caption = caption: AddFrame.BackColor = WHITE: AddFrame.ForeColor = WINE
    SetFont AddFrame, 9, True
End Function
