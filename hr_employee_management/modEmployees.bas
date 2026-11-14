Attribute VB_Name = "modEmployees"
Option Explicit

' ---------------------------------------------------------------------------
' Employee data logic. The form (frmEmployee) only collects input and shows
' messages; every read/write of the Employees table happens here.
' ---------------------------------------------------------------------------

' Column positions in the tblEmployees table
Public Const COL_ID As Long = 1
Public Const COL_FIRST As Long = 2
Public Const COL_LAST As Long = 3
Public Const COL_GENDER As Long = 4
Public Const COL_DEPT As Long = 5
Public Const COL_TITLE As Long = 6
Public Const COL_EMAIL As Long = 7
Public Const COL_PHONE As Long = 8
Public Const COL_HIRE As Long = 9
Public Const COL_SALARY As Long = 10
Public Const COL_STATUS As Long = 11
Public Const FIELD_COUNT As Long = 11

Public Const DATE_FORMAT As String = "dd-mmm-yyyy"

' Dashboard button
Public Sub OpenEmployeeManager()
    frmEmployee.Show
End Sub

Public Function EmployeeTable() As ListObject
    Set EmployeeTable = ThisWorkbook.Worksheets("Employees").ListObjects("tblEmployees")
End Function

' Returns the table row number of an employee, or 0 when not found
Public Function FindEmployeeRow(ByVal employeeId As String) As Long
    Dim result As Variant
    If EmployeeTable.ListRows.Count = 0 Then Exit Function
    result = Application.Match(employeeId, EmployeeTable.ListColumns(COL_ID).DataBodyRange, 0)
    If Not IsError(result) Then FindEmployeeRow = CLng(result)
End Function

' Next ID in the EMP#### sequence
Public Function NextEmployeeId() As String
    Dim cell As Range, highest As Long
    highest = 1000
    If EmployeeTable.ListRows.Count > 0 Then
        For Each cell In EmployeeTable.ListColumns(COL_ID).DataBodyRange
            If Val(Mid$(cell.Value, 4)) > highest Then highest = Val(Mid$(cell.Value, 4))
        Next cell
    End If
    NextEmployeeId = "EMP" & (highest + 1)
End Function

' True when the value appears in a named list on the Lists sheet
Public Function InList(ByVal value As String, ByVal listName As String) As Boolean
    InList = Not IsError(Application.Match(value, ThisWorkbook.Names(listName).RefersToRange, 0))
End Function

' Returns "" when the record is valid, otherwise a message about the first problem.
' currentId is the employee being updated (so their own email is not a duplicate).
Public Function ValidateEmployee(values As Variant, Optional ByVal currentId As String = "") As String
    Dim emailRow As Long, match As Variant

    If values(COL_FIRST) = "" Then ValidateEmployee = "Please enter the first name.": Exit Function
    If values(COL_FIRST) Like "*[!A-Za-z .'-]*" Then
        ValidateEmployee = "First name can only contain letters, spaces, hyphens, periods and apostrophes.": Exit Function
    End If
    If values(COL_LAST) = "" Then ValidateEmployee = "Please enter the last name.": Exit Function
    If values(COL_LAST) Like "*[!A-Za-z .'-]*" Then
        ValidateEmployee = "Last name can only contain letters, spaces, hyphens, periods and apostrophes.": Exit Function
    End If
    If Not InList(values(COL_GENDER), "GenderList") Then ValidateEmployee = "Please select a gender.": Exit Function
    If Not InList(values(COL_DEPT), "DepartmentList") Then ValidateEmployee = "Please select a department.": Exit Function
    If values(COL_TITLE) = "" Then ValidateEmployee = "Please enter the job title.": Exit Function
    If values(COL_TITLE) Like "[=+@-]*" Then
        ValidateEmployee = "Job title cannot start with =, +, - or @.": Exit Function
    End If

    If Not values(COL_EMAIL) Like "?*@?*.?*" Or InStr(values(COL_EMAIL), " ") > 0 Then
        ValidateEmployee = "Please enter a valid email address, for example name@company.com.": Exit Function
    End If
    emailRow = 0
    If EmployeeTable.ListRows.Count > 0 Then
        match = Application.Match(values(COL_EMAIL), EmployeeTable.ListColumns(COL_EMAIL).DataBodyRange, 0)
        If Not IsError(match) Then emailRow = CLng(match)
    End If
    If emailRow > 0 And emailRow <> FindEmployeeRow(currentId) Then
        ValidateEmployee = "This email address is already used by another employee.": Exit Function
    End If

    If Len(values(COL_PHONE)) <> 10 Or values(COL_PHONE) Like "*[!0-9]*" Then
        ValidateEmployee = "Phone number must be exactly 10 digits.": Exit Function
    End If

    If Not values(COL_HIRE) Like "##-[A-Za-z][A-Za-z][A-Za-z]-####" Or Not IsDate(values(COL_HIRE)) Then
        ValidateEmployee = "Please enter the hire date as dd-mmm-yyyy, for example 15-Jan-2024.": Exit Function
    End If
    If CDate(values(COL_HIRE)) > Date Or Year(CDate(values(COL_HIRE))) < 1980 Then
        ValidateEmployee = "Hire date must be between 1980 and today.": Exit Function
    End If

    If Not IsNumeric(values(COL_SALARY)) Or InStr(values(COL_SALARY), ",") > 0 Then
        ValidateEmployee = "Annual salary must be a number (without commas).": Exit Function
    End If
    If CDbl(values(COL_SALARY)) <= 0 Or CDbl(values(COL_SALARY)) > 10000000 Then
        ValidateEmployee = "Annual salary must be greater than 0 and at most 10,000,000.": Exit Function
    End If

    If Not InList(values(COL_STATUS), "StatusList") Then ValidateEmployee = "Please select a status.": Exit Function
End Function

' Reads one employee into a 1-based array (dates as dd-mmm-yyyy text)
Public Function ReadEmployee(ByVal rowNumber As Long) As Variant
    Dim values(1 To FIELD_COUNT) As Variant, col As Long
    For col = 1 To FIELD_COUNT
        values(col) = CStr(EmployeeTable.DataBodyRange.Cells(rowNumber, col).Value)
    Next col
    values(COL_HIRE) = Format$(EmployeeTable.DataBodyRange.Cells(rowNumber, COL_HIRE).Value, DATE_FORMAT)
    ReadEmployee = values
End Function

' Adds a validated record and returns its new Employee ID
Public Function AddEmployee(values As Variant) As String
    values(COL_ID) = NextEmployeeId()
    UnlockEmployees
    EmployeeTable.ListRows.Add
    WriteEmployee EmployeeTable.ListRows.Count, values
    LockEmployees
    AddEmployee = values(COL_ID)
End Function

Public Sub UpdateEmployee(ByVal rowNumber As Long, values As Variant)
    UnlockEmployees
    WriteEmployee rowNumber, values
    LockEmployees
End Sub

Public Sub DeleteEmployee(ByVal rowNumber As Long)
    UnlockEmployees
    EmployeeTable.ListRows(rowNumber).Delete
    LockEmployees
End Sub

' Returns the row numbers of employees whose ID, name, department, job title, email or status contains the text
Public Function SearchEmployees(ByVal searchText As String) As Collection
    Dim results As New Collection, rowNumber As Long, record As String, cells As Range
    searchText = Trim$(searchText)
    For rowNumber = 1 To EmployeeTable.ListRows.Count
        Set cells = EmployeeTable.DataBodyRange.Rows(rowNumber)
        record = cells.Cells(1, COL_ID).Value & "|" & cells.Cells(1, COL_FIRST).Value & " " & _
                 cells.Cells(1, COL_LAST).Value & "|" & cells.Cells(1, COL_DEPT).Value & "|" & _
                 cells.Cells(1, COL_TITLE).Value & "|" & cells.Cells(1, COL_EMAIL).Value & "|" & cells.Cells(1, COL_STATUS).Value
        If searchText = "" Or InStr(1, record, searchText, vbTextCompare) > 0 Then results.Add rowNumber
    Next rowNumber
    Set SearchEmployees = results
End Function

Private Sub WriteEmployee(ByVal rowNumber As Long, values As Variant)
    With EmployeeTable.DataBodyRange.Rows(rowNumber)
        .Cells(1, COL_ID).Value = values(COL_ID)
        .Cells(1, COL_FIRST).Value = values(COL_FIRST)
        .Cells(1, COL_LAST).Value = values(COL_LAST)
        .Cells(1, COL_GENDER).Value = values(COL_GENDER)
        .Cells(1, COL_DEPT).Value = values(COL_DEPT)
        .Cells(1, COL_TITLE).Value = values(COL_TITLE)
        .Cells(1, COL_EMAIL).Value = LCase$(values(COL_EMAIL))
        .Cells(1, COL_PHONE).NumberFormat = "@"
        .Cells(1, COL_PHONE).Value = values(COL_PHONE)
        .Cells(1, COL_HIRE).Value = CDate(values(COL_HIRE))
        .Cells(1, COL_SALARY).Value = CDbl(values(COL_SALARY))
        .Cells(1, COL_STATUS).Value = values(COL_STATUS)
    End With
End Sub

' The Employees sheet is protected so records are changed only through the form
Public Sub LockEmployees()
    ThisWorkbook.Worksheets("Employees").Protect AllowSorting:=True, AllowFiltering:=True
End Sub

Public Sub UnlockEmployees()
    ThisWorkbook.Worksheets("Employees").Unprotect
End Sub
