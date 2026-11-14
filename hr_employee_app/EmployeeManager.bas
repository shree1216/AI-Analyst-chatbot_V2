Attribute VB_Name = "EmployeeManager"
Option Explicit

Public Const APP_TITLE As String = "HR Employee Management"
Public Const EMPLOYEE_SHEET As String = "Employees"
Public Const EMPLOYEE_TABLE As String = "tblEmployees"
Public Const DASHBOARD_SHEET As String = "Dashboard"
Public Const LIST_SHEET As String = "Lists"

Private Const ERROR_BASE As Long = vbObjectError + 3100
Private Const DEPARTMENT_FIRST_ROW As Long = 14

' Used by the automated tests to answer pop-ups without user interaction.
Public SilentMode As Boolean
Public SilentResponse As VbMsgBoxResult
Public LastMessage As String

Public Sub OpenEmployeeForm()
    On Error GoTo Failed
    Dim tbl As ListObject
    Set tbl = EmployeeTable()
    TryRefreshDashboard
    EmployeeForm.Show
    Exit Sub
Failed:
    Notify "The Employee Manager could not be opened." & vbCrLf & vbCrLf & Err.Description, vbCritical
End Sub

Public Function Notify(ByVal message As String, ByVal style As VbMsgBoxStyle, _
                       Optional ByVal title As String = APP_TITLE) As VbMsgBoxResult
    LastMessage = message
    If SilentMode Then
        If (style And 7) = vbYesNo Then
            Notify = SilentResponse
        Else
            Notify = vbOK
        End If
    Else
        Notify = MsgBox(message, style, title)
    End If
End Function

Public Function FieldTitles() As Variant
    FieldTitles = Array("First Name", "Last Name", "Email", "Phone", "Department", _
                        "Job Title", "Hire Date", "Annual Salary", "Status")
End Function

Public Function EmployeeTable() As ListObject
    Dim ws As Worksheet, tbl As ListObject, title As Variant
    Set ws = RequiredSheet(EMPLOYEE_SHEET)
    For Each tbl In ws.ListObjects
        If StrComp(tbl.Name, EMPLOYEE_TABLE, vbTextCompare) = 0 Then Exit For
    Next tbl
    If tbl Is Nothing Then
        Err.Raise ERROR_BASE + 1, APP_TITLE, "The table " & EMPLOYEE_TABLE & " was not found on the " & _
                  EMPLOYEE_SHEET & " worksheet."
    End If
    For Each title In Array("Employee ID", "First Name", "Last Name", "Email", "Phone", "Department", _
                            "Job Title", "Hire Date", "Annual Salary", "Status", "Last Updated")
        If ColumnIndex(tbl, CStr(title)) = 0 Then
            Err.Raise ERROR_BASE + 2, APP_TITLE, "The employee table is missing the """ & title & """ column."
        End If
    Next title
    Set EmployeeTable = tbl
End Function

Public Function ColumnIndex(ByVal tbl As ListObject, ByVal title As String) As Long
    Dim col As ListColumn
    For Each col In tbl.ListColumns
        If StrComp(Trim$(col.Name), title, vbTextCompare) = 0 Then
            ColumnIndex = col.Index
            Exit Function
        End If
    Next col
End Function

Public Function ListValues(ByVal tableName As String) As Variant
    Dim ws As Worksheet, tbl As ListObject, cell As Range, items() As String, count As Long
    Set ws = RequiredSheet(LIST_SHEET)
    For Each tbl In ws.ListObjects
        If StrComp(tbl.Name, tableName, vbTextCompare) = 0 Then Exit For
    Next tbl
    If tbl Is Nothing Then Err.Raise ERROR_BASE + 4, APP_TITLE, "The list table " & tableName & " is missing."
    If tbl.DataBodyRange Is Nothing Then Err.Raise ERROR_BASE + 5, APP_TITLE, "The list " & tableName & " is empty."
    ReDim items(0 To tbl.DataBodyRange.Rows.Count - 1)
    For Each cell In tbl.ListColumns(1).DataBodyRange.Cells
        If Not IsError(cell.Value) Then
            If Len(Trim$(CStr(cell.Value))) > 0 Then
                items(count) = Trim$(CStr(cell.Value))
                count = count + 1
            End If
        End If
    Next cell
    If count = 0 Then Err.Raise ERROR_BASE + 5, APP_TITLE, "The list " & tableName & " is empty."
    ReDim Preserve items(0 To count - 1)
    ListValues = items
End Function

Public Function ValidateEmployee(ByVal values As Variant, ByVal currentID As String, _
                                 ByRef cleaned As Variant, ByRef fieldIndex As Long) As String
    Dim result(0 To 8) As Variant, text As String, i As Long, owner As String, matched As String
    Dim hireDate As Date, salary As Double

    fieldIndex = -1
    If Not IsArray(values) Then ValidateEmployee = "Employee details are missing.": Exit Function
    If UBound(values) - LBound(values) <> 8 Then ValidateEmployee = "Employee details are incomplete.": Exit Function
    For i = 0 To 8
        If IsNull(values(LBound(values) + i)) Or IsError(values(LBound(values) + i)) Then
            result(i) = ""
        Else
            result(i) = Application.WorksheetFunction.Trim(CStr(values(LBound(values) + i)))
        End If
    Next i

    For i = 0 To 1
        text = result(i)
        If Not IsValidName(text) Then
            fieldIndex = i
            If Len(text) = 0 Then
                ValidateEmployee = FieldTitles()(i) & " is required."
            Else
                ValidateEmployee = FieldTitles()(i) & " must start with a letter and contain only letters, " & _
                                   "spaces, hyphens, apostrophes, or periods (maximum 50 characters)."
            End If
            Exit Function
        End If
    Next i

    text = LCase$(result(2))
    If Len(text) = 0 Then fieldIndex = 2: ValidateEmployee = "Email is required.": Exit Function
    If Not IsValidEmail(text) Then
        fieldIndex = 2
        ValidateEmployee = "Email must be a valid address, for example name@example.com."
        Exit Function
    End If
    owner = EmailOwner(text, currentID)
    If Len(owner) > 0 Then
        fieldIndex = 2
        ValidateEmployee = "Email " & text & " is already used by employee " & owner & "."
        Exit Function
    End If
    result(2) = text

    If Len(result(3)) = 0 Then fieldIndex = 3: ValidateEmployee = "Phone is required.": Exit Function
    If Not IsValidPhone(CStr(result(3))) Then
        fieldIndex = 3
        ValidateEmployee = "Phone must contain 7 to 15 digits and only digits, spaces, + ( ) - or periods."
        Exit Function
    End If

    matched = MatchListValue(CStr(result(4)), ListValues("tblDepartments"))
    If Len(matched) = 0 Then
        fieldIndex = 4
        ValidateEmployee = "Department is required. Choose a department from the list."
        Exit Function
    End If
    result(4) = matched

    text = result(5)
    If Len(text) = 0 Then fieldIndex = 5: ValidateEmployee = "Job Title is required.": Exit Function
    If Len(text) > 60 Or Not StartsWithLetterOrDigit(text) Or HasControlCharacters(text) Then
        fieldIndex = 5
        ValidateEmployee = "Job Title must start with a letter or number and be 60 characters or fewer."
        Exit Function
    End If

    If Len(result(6)) = 0 Then fieldIndex = 6: ValidateEmployee = "Hire Date is required.": Exit Function
    If Not TryParseIsoDate(CStr(result(6)), hireDate) Then
        fieldIndex = 6
        ValidateEmployee = "Hire Date must be a real date in yyyy-mm-dd format, for example 2024-03-15."
        Exit Function
    End If
    If hireDate > Date Then
        fieldIndex = 6
        ValidateEmployee = "Hire Date cannot be in the future."
        Exit Function
    End If
    If hireDate < DateSerial(1950, 1, 1) Then
        fieldIndex = 6
        ValidateEmployee = "Hire Date cannot be earlier than 1950-01-01."
        Exit Function
    End If
    result(6) = hireDate

    If Len(result(7)) = 0 Then fieldIndex = 7: ValidateEmployee = "Annual Salary is required.": Exit Function
    If Not TryParseSalary(CStr(result(7)), salary) Then
        fieldIndex = 7
        ValidateEmployee = "Annual Salary must be a positive number up to 10,000,000 with at most two " & _
                           "decimal places, for example 65000 or 65000.50."
        Exit Function
    End If
    result(7) = salary

    matched = MatchListValue(CStr(result(8)), ListValues("tblStatuses"))
    If Len(matched) = 0 Then
        fieldIndex = 8
        ValidateEmployee = "Status is required. Choose a status from the list."
        Exit Function
    End If
    result(8) = matched

    cleaned = result
End Function

Public Function AddEmployee(ByVal values As Variant) As String
    Dim tbl As ListObject, ws As Worksheet, cleaned As Variant, fieldIndex As Long, problem As String
    Dim newID As String, newRow As ListRow, wasProtected As Boolean, errorNumber As Long, errorText As String

    Set tbl = EmployeeTable()
    Set ws = tbl.Parent
    problem = ValidateEmployee(values, "", cleaned, fieldIndex)
    If Len(problem) > 0 Then Err.Raise ERROR_BASE + 10, APP_TITLE, problem
    newID = NextEmployeeID()

    On Error GoTo Failed
    wasProtected = ws.ProtectContents
    If wasProtected Then ws.Unprotect
    Set newRow = tbl.ListRows.Add
    WriteRecord tbl, newRow.Index, newID, cleaned
    If wasProtected Then LockEmployeeSheet ws
    TryRefreshDashboard
    AddEmployee = newID
    Exit Function
Failed:
    errorNumber = Err.Number
    errorText = Err.Description
    Resume RollBack
RollBack:
    On Error Resume Next
    If Not newRow Is Nothing Then newRow.Delete
    If wasProtected Then LockEmployeeSheet ws
    On Error GoTo 0
    Err.Raise errorNumber, APP_TITLE, "The employee was not added: " & errorText
End Function

Public Sub UpdateEmployee(ByVal employeeID As String, ByVal values As Variant)
    Dim tbl As ListObject, ws As Worksheet, cleaned As Variant, fieldIndex As Long, problem As String
    Dim rowIndex As Long, wasProtected As Boolean, original As Variant
    Dim errorNumber As Long, errorText As String

    Set tbl = EmployeeTable()
    Set ws = tbl.Parent
    rowIndex = FindEmployeeRow(employeeID)
    If rowIndex = 0 Then Err.Raise ERROR_BASE + 11, APP_TITLE, "Employee " & employeeID & " was not found."
    problem = ValidateEmployee(values, employeeID, cleaned, fieldIndex)
    If Len(problem) > 0 Then Err.Raise ERROR_BASE + 10, APP_TITLE, problem

    On Error GoTo Failed
    original = tbl.ListRows(rowIndex).Range.Formula
    wasProtected = ws.ProtectContents
    If wasProtected Then ws.Unprotect
    WriteRecord tbl, rowIndex, employeeID, cleaned
    If wasProtected Then LockEmployeeSheet ws
    TryRefreshDashboard
    Exit Sub
Failed:
    errorNumber = Err.Number
    errorText = Err.Description
    Resume RollBack
RollBack:
    On Error Resume Next
    If Not IsEmpty(original) Then tbl.ListRows(rowIndex).Range.Formula = original
    If wasProtected Then LockEmployeeSheet ws
    On Error GoTo 0
    Err.Raise errorNumber, APP_TITLE, "The employee was not updated: " & errorText
End Sub

Public Sub DeleteEmployee(ByVal employeeID As String)
    Dim tbl As ListObject, ws As Worksheet, rowIndex As Long, wasProtected As Boolean
    Dim errorNumber As Long, errorText As String

    Set tbl = EmployeeTable()
    Set ws = tbl.Parent
    rowIndex = FindEmployeeRow(employeeID)
    If rowIndex = 0 Then Err.Raise ERROR_BASE + 11, APP_TITLE, "Employee " & employeeID & " was not found."

    On Error GoTo Failed
    wasProtected = ws.ProtectContents
    If wasProtected Then ws.Unprotect
    tbl.ListRows(rowIndex).Delete
    If wasProtected Then LockEmployeeSheet ws
    TryRefreshDashboard
    Exit Sub
Failed:
    errorNumber = Err.Number
    errorText = Err.Description
    Resume RestoreProtection
RestoreProtection:
    On Error Resume Next
    If wasProtected Then LockEmployeeSheet ws
    On Error GoTo 0
    Err.Raise errorNumber, APP_TITLE, "The employee was not deleted: " & errorText
End Sub

Public Function GetEmployee(ByVal employeeID As String) As Variant
    Dim tbl As ListObject, rowIndex As Long, record(0 To 10) As Variant, titles As Variant, i As Long
    Set tbl = EmployeeTable()
    rowIndex = FindEmployeeRow(employeeID)
    If rowIndex = 0 Then Err.Raise ERROR_BASE + 11, APP_TITLE, "Employee " & employeeID & " was not found."
    titles = FieldTitles()
    record(0) = CStr(tbl.DataBodyRange.Cells(rowIndex, ColumnIndex(tbl, "Employee ID")).Value)
    For i = 0 To 8
        record(i + 1) = tbl.DataBodyRange.Cells(rowIndex, ColumnIndex(tbl, CStr(titles(i)))).Value
    Next i
    record(10) = tbl.DataBodyRange.Cells(rowIndex, ColumnIndex(tbl, "Last Updated")).Value
    GetEmployee = record
End Function

Public Function FindEmployeeRow(ByVal employeeID As String) As Long
    Dim tbl As ListObject, ids As Variant, i As Long
    Set tbl = EmployeeTable()
    employeeID = Trim$(employeeID)
    If Len(employeeID) = 0 Or tbl.DataBodyRange Is Nothing Then Exit Function
    ids = ColumnValues(tbl, "Employee ID")
    For i = 1 To UBound(ids, 1)
        If StrComp(Trim$(CStr(ids(i, 1))), employeeID, vbTextCompare) = 0 Then
            FindEmployeeRow = i
            Exit Function
        End If
    Next i
End Function

' Returns rows of Employee ID, full name, department, and status; Empty when nothing matches.
Public Function SearchEmployees(ByVal term As String) As Variant
    Dim tbl As ListObject, ids As Variant, firsts As Variant, lasts As Variant, emails As Variant
    Dim departments As Variant, titles As Variant, statuses As Variant
    Dim matches() As Long, count As Long, i As Long, haystack As String, results() As Variant, exact As Long

    Set tbl = EmployeeTable()
    If tbl.DataBodyRange Is Nothing Then Exit Function
    term = LCase$(Application.WorksheetFunction.Trim(term))
    ids = ColumnValues(tbl, "Employee ID")
    firsts = ColumnValues(tbl, "First Name")
    lasts = ColumnValues(tbl, "Last Name")
    emails = ColumnValues(tbl, "Email")
    departments = ColumnValues(tbl, "Department")
    titles = ColumnValues(tbl, "Job Title")
    statuses = ColumnValues(tbl, "Status")
    ReDim matches(1 To UBound(ids, 1))

    If Len(term) > 0 Then
        For i = 1 To UBound(ids, 1)
            If LCase$(Trim$(CStr(ids(i, 1)))) = term Then exact = i
        Next i
    End If
    If exact > 0 Then
        count = 1
        matches(1) = exact
    Else
        For i = 1 To UBound(ids, 1)
            haystack = LCase$(CStr(ids(i, 1)) & "|" & CStr(firsts(i, 1)) & " " & CStr(lasts(i, 1)) & "|" & _
                              CStr(emails(i, 1)) & "|" & CStr(departments(i, 1)) & "|" & CStr(titles(i, 1)) & _
                              "|" & CStr(statuses(i, 1)))
            If Len(term) = 0 Or InStr(1, haystack, term, vbBinaryCompare) > 0 Then
                count = count + 1
                matches(count) = i
            End If
        Next i
    End If
    If count = 0 Then Exit Function

    ReDim results(1 To count, 1 To 4)
    For i = 1 To count
        results(i, 1) = CStr(ids(matches(i), 1))
        results(i, 2) = CStr(firsts(matches(i), 1)) & " " & CStr(lasts(matches(i), 1))
        results(i, 3) = CStr(departments(matches(i), 1))
        results(i, 4) = CStr(statuses(matches(i), 1))
    Next i
    SearchEmployees = results
End Function

Public Function NextEmployeeID() As String
    Dim tbl As ListObject, ids As Variant, i As Long, number As Long, highest As Long, text As String
    Set tbl = EmployeeTable()
    If Not tbl.DataBodyRange Is Nothing Then
        ids = ColumnValues(tbl, "Employee ID")
        For i = 1 To UBound(ids, 1)
            text = UCase$(Trim$(CStr(ids(i, 1))))
            If Len(text) >= 5 And Len(text) <= 13 And Left$(text, 4) = "EMP-" Then
                If IsAllDigits(Mid$(text, 5)) Then
                    number = CLng(Mid$(text, 5))
                    If number > highest Then highest = number
                End If
            End If
        Next i
    End If
    NextEmployeeID = "EMP-" & Format$(highest + 1, "0000")
End Function

Public Sub RefreshDashboard()
    Dim ws As Worksheet, departments As Variant, i As Long, rowNumber As Long, wasProtected As Boolean
    Dim lastRow As Long

    Set ws = RequiredSheet(DASHBOARD_SHEET)
    departments = ListValues("tblDepartments")
    wasProtected = ws.ProtectContents
    If wasProtected Then ws.Unprotect
    ws.Range(ws.Cells(DEPARTMENT_FIRST_ROW, 2), ws.Cells(DEPARTMENT_FIRST_ROW + 60, 4)).Clear
    For i = 0 To UBound(departments)
        rowNumber = DEPARTMENT_FIRST_ROW + i
        ws.Cells(rowNumber, 2).NumberFormat = "@"
        ws.Cells(rowNumber, 2).Value = departments(i)
        ws.Cells(rowNumber, 3).Formula = "=COUNTIFS(tblEmployees[Department],$B" & rowNumber & _
                                         ",tblEmployees[Status],""Active"")"
        ws.Cells(rowNumber, 4).Formula = "=COUNTIF(tblEmployees[Department],$B" & rowNumber & ")"
    Next i
    lastRow = DEPARTMENT_FIRST_ROW + UBound(departments)
    ws.Cells(lastRow + 1, 2).Value = "Total"
    ws.Cells(lastRow + 1, 3).Formula = "=SUM(C" & DEPARTMENT_FIRST_ROW & ":C" & lastRow & ")"
    ws.Cells(lastRow + 1, 4).Formula = "=SUM(D" & DEPARTMENT_FIRST_ROW & ":D" & lastRow & ")"
    ws.Range(ws.Cells(lastRow + 1, 2), ws.Cells(lastRow + 1, 4)).Font.Bold = True
    With ws.Range(ws.Cells(DEPARTMENT_FIRST_ROW, 2), ws.Cells(lastRow + 1, 4))
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(191, 191, 191)
        .Font.Name = "Segoe UI"
        .Font.Size = 10
    End With
    ws.Range(ws.Cells(DEPARTMENT_FIRST_ROW, 3), ws.Cells(lastRow + 1, 4)).HorizontalAlignment = xlCenter
    If wasProtected Then ws.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True
End Sub

Private Sub TryRefreshDashboard()
    On Error Resume Next
    RefreshDashboard
End Sub

Private Sub WriteRecord(ByVal tbl As ListObject, ByVal rowIndex As Long, ByVal employeeID As String, _
                        ByVal cleaned As Variant)
    Dim titles As Variant, i As Long, cell As Range
    titles = FieldTitles()
    SetText tbl.DataBodyRange.Cells(rowIndex, ColumnIndex(tbl, "Employee ID")), employeeID
    For i = 0 To 8
        Set cell = tbl.DataBodyRange.Cells(rowIndex, ColumnIndex(tbl, CStr(titles(i))))
        Select Case i
            Case 6
                cell.NumberFormat = "yyyy-mm-dd"
                cell.Value = cleaned(i)
            Case 7
                cell.NumberFormat = "#,##0.00"
                cell.Value = cleaned(i)
            Case Else
                SetText cell, CStr(cleaned(i))
        End Select
    Next i
    Set cell = tbl.DataBodyRange.Cells(rowIndex, ColumnIndex(tbl, "Last Updated"))
    cell.NumberFormat = "yyyy-mm-dd hh:mm"
    cell.Value = Now
End Sub

Private Sub SetText(ByVal cell As Range, ByVal text As String)
    cell.NumberFormat = "@"
    cell.Value = text
End Sub

Private Sub LockEmployeeSheet(ByVal ws As Worksheet)
    ws.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, AllowFiltering:=True
End Sub

Private Function RequiredSheet(ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If StrComp(ws.Name, sheetName, vbTextCompare) = 0 Then
            Set RequiredSheet = ws
            Exit Function
        End If
    Next ws
    Err.Raise ERROR_BASE + 3, APP_TITLE, "The """ & sheetName & """ worksheet is missing."
End Function

' Always returns a 1-based two-dimensional array, even for a one-row table.
Private Function ColumnValues(ByVal tbl As ListObject, ByVal title As String) As Variant
    Dim values As Variant, oneValue(1 To 1, 1 To 1) As Variant
    values = tbl.ListColumns(ColumnIndex(tbl, title)).DataBodyRange.Value
    If IsArray(values) Then
        ColumnValues = values
    Else
        oneValue(1, 1) = values
        ColumnValues = oneValue
    End If
End Function

Private Function EmailOwner(ByVal email As String, ByVal currentID As String) As String
    Dim tbl As ListObject, ids As Variant, emails As Variant, i As Long
    Set tbl = EmployeeTable()
    If tbl.DataBodyRange Is Nothing Then Exit Function
    ids = ColumnValues(tbl, "Employee ID")
    emails = ColumnValues(tbl, "Email")
    For i = 1 To UBound(ids, 1)
        If StrComp(Trim$(CStr(emails(i, 1))), email, vbTextCompare) = 0 And _
           StrComp(Trim$(CStr(ids(i, 1))), Trim$(currentID), vbTextCompare) <> 0 Then
            EmailOwner = CStr(ids(i, 1))
            Exit Function
        End If
    Next i
End Function

Private Function MatchListValue(ByVal value As String, ByVal items As Variant) As String
    Dim i As Long
    For i = LBound(items) To UBound(items)
        If StrComp(value, CStr(items(i)), vbTextCompare) = 0 Then
            MatchListValue = CStr(items(i))
            Exit Function
        End If
    Next i
End Function

Private Function IsLetter(ByVal ch As String) As Boolean
    IsLetter = (UCase$(ch) <> LCase$(ch))
End Function

Private Function IsValidName(ByVal text As String) As Boolean
    Dim i As Long, ch As String
    If Len(text) = 0 Or Len(text) > 50 Then Exit Function
    If Not IsLetter(Left$(text, 1)) Then Exit Function
    For i = 1 To Len(text)
        ch = Mid$(text, i, 1)
        If Not (IsLetter(ch) Or ch = " " Or ch = "-" Or ch = "'" Or ch = ".") Then Exit Function
    Next i
    IsValidName = True
End Function

Private Function IsValidEmail(ByVal text As String) As Boolean
    Dim at As Long, localPart As String, domain As String, i As Long, ch As String, labels As Variant
    If Len(text) > 100 Or InStr(text, " ") > 0 Then Exit Function
    at = InStr(text, "@")
    If at < 2 Or InStr(at + 1, text, "@") > 0 Then Exit Function
    localPart = Left$(text, at - 1)
    domain = Mid$(text, at + 1)
    If Not (Left$(localPart, 1) Like "[a-z0-9]") Then Exit Function
    If InStr(text, "..") > 0 Or Right$(localPart, 1) = "." Then Exit Function
    For i = 1 To Len(localPart)
        ch = Mid$(localPart, i, 1)
        If Not (ch Like "[a-z0-9]" Or InStr("._%+-", ch) > 0) Then Exit Function
    Next i
    If InStr(domain, ".") = 0 Then Exit Function
    labels = Split(domain, ".")
    For i = LBound(labels) To UBound(labels)
        If Len(labels(i)) = 0 Then Exit Function
        If Left$(labels(i), 1) = "-" Or Right$(labels(i), 1) = "-" Then Exit Function
        If Not IsDomainLabel(CStr(labels(i))) Then Exit Function
    Next i
    If Len(labels(UBound(labels))) < 2 Or Not (labels(UBound(labels)) Like "*[a-z]*") Then Exit Function
    IsValidEmail = True
End Function

Private Function IsDomainLabel(ByVal label As String) As Boolean
    Dim i As Long
    For i = 1 To Len(label)
        If Not (Mid$(label, i, 1) Like "[a-z0-9-]") Then Exit Function
    Next i
    IsDomainLabel = True
End Function

Private Function IsValidPhone(ByVal text As String) As Boolean
    Dim i As Long, ch As String, digits As Long
    If Len(text) > 25 Then Exit Function
    For i = 1 To Len(text)
        ch = Mid$(text, i, 1)
        If ch Like "#" Then
            digits = digits + 1
        ElseIf ch = "+" Then
            If i <> 1 Then Exit Function
        ElseIf InStr(" ()-.", ch) = 0 Then
            Exit Function
        End If
    Next i
    IsValidPhone = (digits >= 7 And digits <= 15)
End Function

Private Function StartsWithLetterOrDigit(ByVal text As String) As Boolean
    StartsWithLetterOrDigit = IsLetter(Left$(text, 1)) Or (Left$(text, 1) Like "#")
End Function

Private Function HasControlCharacters(ByVal text As String) As Boolean
    Dim i As Long
    For i = 1 To Len(text)
        If AscW(Mid$(text, i, 1)) < 32 Then HasControlCharacters = True: Exit Function
    Next i
End Function

Private Function IsAllDigits(ByVal text As String) As Boolean
    Dim i As Long
    If Len(text) = 0 Then Exit Function
    For i = 1 To Len(text)
        If Not (Mid$(text, i, 1) Like "#") Then Exit Function
    Next i
    IsAllDigits = True
End Function

Private Function TryParseIsoDate(ByVal text As String, ByRef result As Date) As Boolean
    Dim y As Long, m As Long, d As Long, candidate As Date
    If Not (text Like "####-##-##") Then Exit Function
    y = CLng(Left$(text, 4))
    m = CLng(Mid$(text, 6, 2))
    d = CLng(Right$(text, 2))
    If m < 1 Or m > 12 Or d < 1 Or d > 31 Or y < 1900 Then Exit Function
    candidate = DateSerial(y, m, d)
    If Year(candidate) <> y Or Month(candidate) <> m Or Day(candidate) <> d Then Exit Function
    result = candidate
    TryParseIsoDate = True
End Function

' Accepts digits, optional thousands commas, and a period decimal separator in every locale.
Private Function TryParseSalary(ByVal text As String, ByRef result As Double) As Boolean
    Dim plain As String, dot As Long, i As Long, ch As String
    plain = Replace(text, ",", "")
    If Len(plain) = 0 Or Len(plain) > 14 Then Exit Function
    For i = 1 To Len(plain)
        ch = Mid$(plain, i, 1)
        If ch = "." Then
            If dot > 0 Then Exit Function
            dot = i
        ElseIf Not (ch Like "#") Then
            Exit Function
        End If
    Next i
    If dot = 1 Or dot = Len(plain) Then Exit Function
    If dot > 0 Then
        If Len(plain) - dot > 2 Then Exit Function
    End If
    result = Val(plain)
    TryParseSalary = (result > 0 And result <= 10000000#)
End Function
