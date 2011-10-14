VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ILexOutput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function AddToken(ByVal nLine As Long, ByVal nColumn As Long, ByVal nType As enumTokenType) As Long
'Puts "Debug: Token " + Format(nType, "@@@@@") + vbCrLf
End Function

Public Function AddTokenWithString(ByVal nLine As Long, ByVal nColumn As Long, ByVal nType As enumTokenType, ByVal s As String) As Long
'Puts "Debug: Token " + Format(nType, "@@@@@") + " " + s + vbCrLf
End Function

