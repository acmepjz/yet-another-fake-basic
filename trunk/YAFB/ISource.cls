VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ISource"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function GetCh() As Long
'
End Function

Public Sub UnGetCh(Optional ByVal nCount As Long = 1)
'
End Sub

Public Sub Rewind()
'
End Sub

Public Property Get FileName() As String
'
End Property

Public Property Get Line() As Long
'
End Property

Public Property Get Column() As Long
'
End Property
