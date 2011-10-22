Attribute VB_Name = "HelloWorld"
Option Explicit

Private Declare Function GetStdHandle Lib "kernel32.dll" (ByVal nStdHandle As Long) As Long
Private Declare Function WriteFile Lib "kernel32.dll" (ByVal hFile As Long, ByRef lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, ByRef lpNumberOfBytesWritten As Long, ByRef lpOverlapped As Any) As Long
Private Declare Function ReadFile Lib "kernel32.dll" (ByVal hFile As Long, ByRef lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, ByRef lpNumberOfBytesRead As Long, ByRef lpOverlapped As Any) As Long
Private Const STD_ERROR_HANDLE As Long = -12&
Private Const STD_INPUT_HANDLE As Long = -10&

private z as long

public const xyz=1293+STD_ERROR_HANDLE*STD_INPUT_HANDLE
public const wc=xyz^2

Public Sub Main()
Dim h As Long
Dim i As Long, j As Long
Dim c As Long
'///input
h = GetStdHandle(-10&)
ReadFile h, c, 1, j, ByVal 0
'///calc factorial
PrintInteger Factorial
'PrintInteger Factorial(c And &HF&)
End Sub

'Private Function Factorial(ByVal n As Long) As Long
'Dim i As Long, j As Long
'j = 1
'For i = n To 1 step -1
' j = j * i
'Next i
'Factorial = j
'End Function

Private Function Factorial(optional ByVal n As Long=wc\453) As Long
'If n <= 1 Then Factorial = 1 Else Factorial = Factorial(n - 1) * n
const OXZ as long=wc\453
Factorial=OXZ-n
End Function

Private Sub PrintInteger(ByVal n As Long)
Dim h As Long, i As Long
h = GetStdHandle(-12)
If n < 0 Then
 n = -n
 WriteFile h, 45&, 1, 0&, ByVal 0
End If
If n = 0 Then
 WriteFile h, &H30&, 1, 0&, ByVal 0
Else
 i = n \ 10
 If i > 0 Then PrintInteger i
 WriteFile h, &H30& Or (n Mod 10), 1, 0&, ByVal 0
End If
End Sub
