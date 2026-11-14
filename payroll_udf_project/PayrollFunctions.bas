Attribute VB_Name = "PayrollFunctions"
Option Explicit

' Custom payroll functions (UDFs). Use them in any cell, e.g.
'   =CalculateBonus(E2, F2)      =CalculateDeductions(E2)
' The company rules below are also listed on the "Company Rules" sheet.

Private Const PROVIDENT_FUND_RATE As Double = 0.12   ' 12% of monthly salary
Private Const HEALTH_INSURANCE As Double = 150       ' flat amount per month
Private Const TAX_FREE_LIMIT As Double = 5000        ' no income tax up to this salary
Private Const INCOME_TAX_RATE As Double = 0.1        ' 10% of salary above the limit

' Performance bonus = monthly salary x bonus % for the rating.
' Rating: 5 = 20%, 4 = 15%, 3 = 10%, 2 = 5%, 1 = 0%. Invalid input returns #VALUE!.
Public Function CalculateBonus(ByVal Salary As Double, ByVal Rating As Double) As Variant
    Dim bonusRate As Double

    If Salary <= 0 Then
        CalculateBonus = CVErr(xlErrValue)
        Exit Function
    End If

    Select Case Rating
        Case 5: bonusRate = 0.2
        Case 4: bonusRate = 0.15
        Case 3: bonusRate = 0.1
        Case 2: bonusRate = 0.05
        Case 1: bonusRate = 0
        Case Else
            CalculateBonus = CVErr(xlErrValue)
            Exit Function
    End Select

    CalculateBonus = Application.WorksheetFunction.Round(Salary * bonusRate, 2)
End Function

' Total monthly deductions = Provident Fund + Health Insurance + Income Tax.
' Invalid salary (zero, negative or blank) returns #VALUE!.
Public Function CalculateDeductions(ByVal Salary As Double) As Variant
    Dim providentFund As Double, incomeTax As Double

    If Salary <= 0 Then
        CalculateDeductions = CVErr(xlErrValue)
        Exit Function
    End If

    providentFund = Salary * PROVIDENT_FUND_RATE
    If Salary > TAX_FREE_LIMIT Then incomeTax = (Salary - TAX_FREE_LIMIT) * INCOME_TAX_RATE

    CalculateDeductions = Application.WorksheetFunction.Round(providentFund + HEALTH_INSURANCE + incomeTax, 2)
End Function
