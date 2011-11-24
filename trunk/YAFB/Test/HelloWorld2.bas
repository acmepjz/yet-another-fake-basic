Attribute VB_Name = "HelloWorld"
Option Explicit

#If Win32 Then
Private Declare Function GetStdHandle Lib "kernel32.dll" (ByVal nStdHandle As Long) As Long
Private Declare Function WriteFile Lib "kernel32.dll" (ByVal hFile As Long, ByRef lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, ByRef lpNumberOfBytesWritten As Long, ByRef lpOverlapped As Any) As Long
Private Declare Function ReadFile Lib "kernel32.dll" (ByVal hFile As Long, ByRef lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, ByRef lpNumberOfBytesRead As Long, ByRef lpOverlapped As Any) As Long
Private Const STD_ERROR_HANDLE As Long = -12&
Private Const STD_INPUT_HANDLE As Long = -10&

Private hStdErr As Long, hStdInput As Long
#Else
Private Declare Function getchar Lib "msvcrt.dll" () As Long
Private Declare Function putchar Lib "msvcrt.dll" (ByVal c As Long) As Long
Private Declare Function puts Lib "msvcrt.dll" (ByRef lp As Any) As Long
Private Declare Sub [exit] Lib "msvcrt.dll" Alias "exit" (ByVal exitcode As Long)
#End If

Public Sub Main()
Dim i As LongLong 'HAHA VB6 doesn't know LongLong
'///
#If Win32 Then
hStdInput = GetStdHandle(STD_INPUT_HANDLE)
hStdErr = GetStdHandle(STD_ERROR_HANDLE)
#End If
Do While InputInteger(i)
 PrintInteger Factorial(i)
Loop
#If Win32 Then
#Else
[exit] 0
#End If
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
#If Win32 Then
WriteFile hStdErr, &H203F&, 2, 0&, ByVal 0
#Else
putchar &H3F&
putchar &H20&
#End If
Do
 c = 0
 #If Win32 Then
 ReadFile hStdInput, c, 1, i, ByVal 0
 If i <> 1 Then Exit Function
 #Else
 c = getchar
 If c < 0 Then Exit Function
 #End If
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
Dim b(32) As Byte
Dim lp As Long
'///
Erase b
'///
If n < 0 Then
 n = -n
 #If Win32 Then
 WriteFile hStdErr, 45&, 1, 0&, ByVal 0
 #Else
 putchar 45&
 #End If
End If
If n = 0 Then
 #If Win32 Then
 WriteFile hStdErr, &H30&, 1, 0&, ByVal 0
 #Else
 putchar &H30&
 #End If
Else
 lp = 31
 Do
  b(lp) = &H30& Or (n Mod 10)
  n = n \ 10
  If n = 0 Then Exit Do
  lp = lp - 1
 Loop
 #If Win32 Then
 WriteFile hStdErr, b(lp), 32 - lp, 0&, ByVal 0
 #Else
 puts b(lp)
 #End If
End If
End Sub

