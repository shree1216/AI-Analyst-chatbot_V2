Attribute VB_Name = "modEmployees"
Option Explicit

' Data functions used by the Employee Manager form (frmEmployee).
' The employee records live in the Excel table "tblEmployees" on the "Employees" sheet.

' Position of each field in the values array passed between the form and this module
' (the same order as the table columns).
Public Const F_ID As Long = 0
Public Const F_NAME As Long = 1
Public Const F_GENDER As Long = 2
Public Const F_DEPARTMENT As Long = 3
Public Const F_TITLE As Long = 4
Public Const F_EMAIL As Long = 5
Public Const F_PHONE As Long = 6
Public Const F_HIRE_DATE As Long = 7
Public Const F_SALARY As Long = 8
Public Const F_STATUS As Long = 9

Private Const MIN_SALARY As Double = 10000
Private Const MAX_SALARY As Double = 1000000

' Dashboard button
Public Sub OpenEmployeeForm()
    frmEmployee.Show
End Sub

' ---------- Lists for the dropdowns ----------------------------------------------------------

Public Function GenderList() As Variant
    GenderList = Array("Male", "Female")
End Function

Public Function DepartmentList() As Variant
    DepartmentList = Array("Finance", "Human Resources", "IT", "Marketing", "Operations", "Sales")
End Function

Public Function StatusList() As Variant
    StatusList = Array("Active", "On Leave", "Resigned")
End Function

' ---------- Data access ----------------------------------------------------------------------

Public Function EmployeeTable() As ListObject
    Set EmployeeTable = ThisWorkbook.Worksheets("Employees").ListObjects("tblEmployees")
End Function

' Returns the table row (1 = first employee) matching an Employee ID or name, or 0 if not found.
' An exact ID or name match wins; otherwise the first name containing the search text is used.
Public Function FindEmployee(ByVal searchText As String) As Long
    Dim data As Variant, r As Long, text As String
    text = LCase$(Trim$(searchText))
    If Len(text) = 0 Or EmployeeTable.ListRows.Count = 0 Then Exit Function
    data = EmployeeTable.DataBodyRange.Value

    For r = 1 To UBound(data, 1)
        If LCase$(data(r, F_ID + 1)) = text Or LCase$(data(r, F_NAME + 1)) = text Then
            FindEmployee = r
            Exit Function
        End If
    Next r
    For r = 1 To UBound(data, 1)
        If InStr(1, LCase$(data(r, F_NAME + 1)), text) > 0 Then
            FindEmployee = r
            Exit Function
        End If
    Next r
End Function

' Returns the next free Employee ID, e.g. EMP101.
Public Function NextEmployeeID() As String
    Dim data As Variant, r As Long, number As Long, highest As Long
    If EmployeeTable.ListRows.Count > 0 Then
        data = EmployeeTable.DataBodyRange.Value
        For r = 1 To UBound(data, 1)
            number = Val(Mid$(data(r, F_ID + 1), 4))
            If number > highest Then highest = number
        Next r
    End If
    NextEmployeeID = "EMP" & Format$(highest + 1, "000")
End Function

' Checks the form values. Returns "" when everything is valid, otherwise one line per problem.
' rowToIgnore is the row being updated, so its own email is not reported as a duplicate.
Public Function ValidateEmployee(ByVal values As Variant, Optional ByVal rowToIgnore As Long = 0) As String
    Dim problems As String, emailRow As Long

    If Len(values(F_NAME)) < 3 Or values(F_NAME) Like "*[!A-Za-z .'-]*" Then
        problems = problems & "- Full Name is required (letters and spaces only)." & vbNewLine
    End If
    If Not IsInList(values(F_GENDER), GenderList) Then problems = problems & "- Select a Gender." & vbNewLine
    If Not IsInList(values(F_DEPARTMENT), DepartmentList) Then problems = problems & "- Select a Department." & vbNewLine
    If Len(values(F_TITLE)) = 0 Then problems = problems & "- Job Title is required." & vbNewLine

    If Not (values(F_EMAIL) Like "?*@?*.?*") Or InStr(values(F_EMAIL), " ") > 0 Then
        problems = problems & "- Enter a valid Email (name@company.com)." & vbNewLine
    Else
        emailRow = FindValue(F_EMAIL, values(F_EMAIL))
        If emailRow > 0 And emailRow <> rowToIgnore Then
            problems = problems & "- Email is already used by another employee." & vbNewLine
        End If
    End If

    If Not (values(F_PHONE) Like "###-###-####") Then
        problems = problems & "- Phone must be in the format 555-123-4567." & vbNewLine
    End If

    If Not IsDate(values(F_HIRE_DATE)) Then
        problems = problems & "- Hire Date must be a valid date (e.g. 15-Mar-2024)." & vbNewLine
    ElseIf CDate(values(F_HIRE_DATE)) > Date Then
        problems = problems & "- Hire Date cannot be in the future." & vbNewLine
    End If

    If Not IsNumeric(values(F_SALARY)) Then
        problems = problems & "- Annual Salary must be a number." & vbNewLine
    ElseIf CDbl(values(F_SALARY)) < MIN_SALARY Or CDbl(values(F_SALARY)) > MAX_SALARY Then
        problems = problems & "- Annual Salary must be between " & Format$(MIN_SALARY, "#,##0") & _
                   " and " & Format$(MAX_SALARY, "#,##0") & "." & vbNewLine
    End If

    If Not IsInList(values(F_STATUS), StatusList) Then problems = problems & "- Select a Status." & vbNewLine
    ValidateEmployee = problems
End Function

' Adds a new employee (rowIndex = 0) or updates an existing row. Returns the Employee ID.
' Call ValidateEmployee first.
Public Function SaveEmployee(ByVal values As Variant, Optional ByVal rowIndex As Long = 0) As String
    Dim target As Range
    If rowIndex = 0 Then
        values(F_ID) = NextEmployeeID()
        Set target = EmployeeTable.ListRows.Add.Range
    Else
        Set target = EmployeeTable.ListRows(rowIndex).Range
    End If

    target.Cells(1, F_ID + 1).Value = values(F_ID)
    target.Cells(1, F_NAME + 1).Value = StrConv(CleanText(values(F_NAME)), vbProperCase)
    target.Cells(1, F_GENDER + 1).Value = values(F_GENDER)
    target.Cells(1, F_DEPARTMENT + 1).Value = values(F_DEPARTMENT)
    target.Cells(1, F_TITLE + 1).Value = CleanText(values(F_TITLE))
    target.Cells(1, F_EMAIL + 1).Value = LCase$(Trim$(values(F_EMAIL)))
    target.Cells(1, F_PHONE + 1).Value = values(F_PHONE)
    target.Cells(1, F_HIRE_DATE + 1).Value = CDate(values(F_HIRE_DATE))
    target.Cells(1, F_SALARY + 1).Value = CDbl(values(F_SALARY))
    target.Cells(1, F_STATUS + 1).Value = values(F_STATUS)
    SaveEmployee = values(F_ID)
End Function

Public Sub DeleteEmployee(ByVal rowIndex As Long)
    EmployeeTable.ListRows(rowIndex).Delete
End Sub

' ---------- Helpers --------------------------------------------------------------------------

' Trims the text and collapses repeated spaces.
Private Function CleanText(ByVal value As Variant) As String
    CleanText = Application.WorksheetFunction.Trim(CStr(value))
End Function

Private Function IsInList(ByVal value As Variant, ByVal list As Variant) As Boolean
    Dim item As Variant
    For Each item In list
        If CStr(value) = item Then IsInList = True
    Next item
End Function

' Returns the row whose field equals value (not case-sensitive), or 0.
Private Function FindValue(ByVal field As Long, ByVal value As String) As Long
    Dim data As Variant, r As Long
    If EmployeeTable.ListRows.Count = 0 Then Exit Function
    data = EmployeeTable.DataBodyRange.Value
    For r = 1 To UBound(data, 1)
        If StrComp(Trim$(data(r, field + 1)), Trim$(value), vbTextCompare) = 0 Then
            FindValue = r
            Exit Function
        End If
    Next r
End Function
