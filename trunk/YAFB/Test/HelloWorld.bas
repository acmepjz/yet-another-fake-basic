Attribute VB_Name = "HelloWorld"
Option Explicit

Private Declare Function GetStdHandle Lib "kernel32.dll" (ByVal nStdHandle As Long) As Long
Private Declare Function WriteFile Lib "kernel32.dll" (ByVal hFile As Long, ByRef lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, ByRef lpNumberOfBytesWritten As Long, ByVal lpOverlapped As Long) As Long
'Private Const STD_ERROR_HANDLE As Long = -12&

Public Sub Main()
Dim h As Long
Dim i As Long, j As Long
h = GetStdHandle(&HFFFFFFF4)
i = &H6C6C6548
WriteFile h, i, 4, j, 0
i = &H57202C6F
WriteFile h, i, 4, j, 0
i = &H646C726F
WriteFile h, i, 4, j, 0
i = &HD0A2164
WriteFile h, i, 4, j, 0
End Sub

'should be -27
Public Function Test() As Long
Test = -3 ^ Not 4 - 2 ^ 3
End Function
