# VBA-
人民币大写
Function DX(M_MONEY As Double)
Dim M_code As String
On Error GoTo 900
M_MONEY = Round(M_MONEY, 2)
If M_MONEY = 0 Then
    M_code = ""
    GoTo 900
End If
If M_MONEY < 0 Then
    M_code = "负"
    M_MONEY = Abs(M_MONEY)
End If
If M_MONEY < 0.1 Then
    M_code = M_code & Application.WorksheetFunction.Text(Int(M_MONEY * 100), "[dbnum2]") & "分"
    GoTo 900
End If
If M_MONEY >= 1 Then
    M_code = M_code & Application.WorksheetFunction.Text(Int(M_MONEY), "[dbnum2]") & "元"
End If
M_MONEY = Round((M_MONEY - Int(M_MONEY)) * 100, 2)
If M_MONEY = 0 Then
    M_code = M_code & "整"
    GoTo 900
End If
If M_MONEY >= 10 Then
    M_code = M_code & Application.WorksheetFunction.Text(Int(M_MONEY / 10), "[dbnum2]") & "角"
    M_MONEY = M_MONEY Mod 10
Else
    M_code = M_code & "零"
End If
If M_MONEY = 0 Then
    GoTo 900
Else
  M_code = M_code & Application.WorksheetFunction.Text(Int(M_MONEY), "[dbnum2]") & "分"
End If
900:
DX = M_code
End Function
