Attribute VB_Name = "HelloWorld"
Option Explicit

Private Declare Function GetStdHandle Lib "kernel32.dll" (ByVal nStdHandle As Long) As Long
Private Declare Function WriteFile Lib "kernel32.dll" (ByVal hFile As Long, ByRef lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, ByRef lpNumberOfBytesWritten As Long, ByRef lpOverlapped As Any) As Long
Private Declare Function ReadFile Lib "kernel32.dll" (ByVal hFile As Long, ByRef lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, ByRef lpNumberOfBytesRead As Long, ByRef lpOverlapped As Any) As Long
Private Const STD_ERROR_HANDLE As Long = -12&
Private Const STD_INPUT_HANDLE As Long = -10&

Private hStdErr As Long, hStdInput As Long

Public Sub Main()
Dim i As LongLong
'///
hStdInput = GetStdHandle(STD_INPUT_HANDLE)
hStdErr = GetStdHandle(STD_ERROR_HANDLE)
Do While InputInteger(i)
 PrintInteger Factorial(i)
Loop
End Sub

Private Function Factorial(ByVal n As Long) As LongLong
Dim i As Long, j As LongLong
j = 1
For i = n To 1 Step -1
 j = j * i
Next i
Factorial = j
End Function

Private Function InputInteger(ByRef n As LongLong) As Boolean
Dim i As Long, c As Long
Dim bHasNumber As Boolean
n = 0
WriteFile hStdErr, &H203F&, 2, 0&, ByVal 0
Do
 c = 0
 ReadFile hStdInput, c, 1, i, ByVal 0
 If i <> 1 Then Exit Function
 Select Case c
 Case &H30& To &H39&
  n = n * 10 + (c And &HF&)
  bHasNumber = True
 Case &H20&, 13, 10, 9
  If bHasNumber Then Exit Do
 Case Else
  Exit Function
 End Select
Loop
InputInteger = True
End Function

Private Sub PrintInteger(ByVal n As LongLong)
Dim b(31) As Byte
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

