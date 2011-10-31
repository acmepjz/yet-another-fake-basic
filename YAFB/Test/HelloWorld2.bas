Attribute VB_Name = "HelloWorld"
Option Explicit

Private Declare Function GetStdHandle Lib "kernel32.dll" (ByVal nStdHandle As Long) As Long
Private Declare Function WriteFile Lib "kernel32.dll" (ByVal hFile As Long, ByRef lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, ByRef lpNumberOfBytesWritten As Long, ByRef lpOverlapped As Any) As Long
Private Declare Function ReadFile Lib "kernel32.dll" (ByVal hFile As Long, ByRef lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, ByRef lpNumberOfBytesRead As Long, ByRef lpOverlapped As Any) As Long
Private Const STD_ERROR_HANDLE As Long = -12&
Private Const STD_INPUT_HANDLE As Long = -10&

'private z as long,zz as variant

'public const xyz=1293+STD_ERROR_HANDLE*STD_INPUT_HANDLE
'public const wc as long=xyz^5

Private hStdErr As Long, hStdInput As Long

'public const DEAD as long=&H80000000& mod -1

Public Sub Main()
Dim i As Long, j As Long
Dim c As Long
'///
hStdInput = GetStdHandle(STD_INPUT_HANDLE)
hStdErr = GetStdHandle(STD_ERROR_HANDLE)
'///input
ReadFile hStdInput, c, 1, j, ByVal 0
'///calc factorial
PrintInteger Factorial(c And &HF&)
End Sub

Private Function Factorial(ByVal n As Long) As LongLong
Dim i As Long, j As LongLong
j = 1
For i = n To 1 Step -1
 j = j * i
Next i
Factorial = j
End Function

'Private Function Factorial(ByVal n As Long) As LongLong
'If n <= 1 Then Factorial = 1 Else Factorial = Factorial(n - 1) * n
'End Function

'Private Sub PrintInteger(ByVal n As LongLong)
'Dim i As LongLong
'If n < 0 Then
' n = -n
' WriteFile hStdErr, 45&, 1, 0&, ByVal 0
'End If
'If n = 0 Then
' WriteFile hStdErr, &H30&, 1, 0&, ByVal 0
'Else
' i = n \ 10
' If i > 0 Then PrintInteger i
' WriteFile hStdErr, &H30& Or (n Mod 10), 1, 0&, ByVal 0
'End If
'End Sub

Private Sub PrintInteger(ByVal n As LongLong)
Dim b(31) As Byte
'dim xxx(255,255,255,255,255) as byte
Dim lp As Long
'///
If n < 0 Then
 n = -n
 WriteFile hStdErr, 45&, 1, 0&, ByVal 0
End If
If n = 0 Then
 WriteFile hStdErr, &H30&, 1, 0&, ByVal 0
Else
 lp = 31
 Do
  b(lp) = &H30& Or (n Mod 10)
  n = n \ 10
  If n = 0 Then Exit Do
  lp = lp - 1
 Loop
 WriteFile hStdErr, b(lp), 32 - lp, 0&, ByVal 0
End If
End Sub

