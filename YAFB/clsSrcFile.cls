VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSrcFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)

#Const StrConv = True

Implements ISource

#If StrConv Then
Private m_bData() As Integer
#Else
Private m_bData() As Byte
#End If
Private m_nSize As Long
Private m_nPointer As Long

Private m_sFileName As String
Private m_nLine As Long
Private m_nColumn As Long

Private Const m_nOldLineMax As Long = 32

Private m_nOldLine(m_nOldLineMax - 1) As Long
Private m_nOldColumn(m_nOldLineMax - 1) As Long

Private m_nOldLineCount As Long
Private m_nOldLineStart As Long

Private Property Get ISource_Column() As Long
ISource_Column = m_nColumn
End Property

Private Property Get ISource_FileName() As String
ISource_FileName = m_sFileName
End Property

Private Function ISource_GetCh() As Long
'///
m_nOldLineCount = m_nOldLineCount + 1
If m_nOldLineCount > m_nOldLineMax Then m_nOldLineCount = m_nOldLineMax
m_nOldLine(m_nOldLineStart) = m_nLine
m_nOldColumn(m_nOldLineStart) = m_nColumn
m_nOldLineStart = (m_nOldLineStart + 1) And (m_nOldLineMax - 1)
'///
If m_nPointer < m_nSize Then
 ISource_GetCh = m_bData(m_nPointer) And &HFFFF&
 '///
 Select Case ISource_GetCh
 Case ["\t"]
  m_nColumn = ((m_nColumn + 3) And &HFFFFFFFC) + 1
 Case ["\n"]
'  If m_nPointer > 0 Then
'   If m_bData(m_nPointer - 1) <> ["\r"] Then
'    m_nLine = m_nLine + 1
'    m_nColumn = 1
'   End If
'  Else
   m_nLine = m_nLine + 1
   m_nColumn = 1
'  End If
' Case ["\r"]
'  If m_nPointer > 0 Then
'   If m_bData(m_nPointer - 1) <> ["\n"] Then
'    m_nLine = m_nLine + 1
'    m_nColumn = 1
'   End If
'  Else
'   m_nLine = m_nLine + 1
'   m_nColumn = 1
'  End If
 Case Else
  m_nColumn = m_nColumn + 1
 End Select
Else
 ISource_GetCh = -1
End If
'///
m_nPointer = m_nPointer + 1
End Function

Private Property Get ISource_Line() As Long
ISource_Line = m_nLine
End Property

Private Sub ISource_Rewind()
m_nPointer = 0
m_nLine = 1
m_nColumn = 1
m_nOldLineCount = 0
m_nOldLineStart = 0
End Sub

Private Sub ISource_UnGetCh(Optional ByVal nCount As Long = 1&)
If nCount > m_nPointer Then nCount = m_nPointer
If nCount > 0 Then
 m_nPointer = m_nPointer - nCount
 '///
 If nCount > m_nOldLineCount Then
  Debug.Assert False
  nCount = m_nOldLineCount
 End If
 '///
 If nCount > 0 Then
  m_nOldLineStart = (m_nOldLineStart - nCount) And (m_nOldLineMax - 1)
  m_nLine = m_nOldLine(m_nOldLineStart)
  m_nColumn = m_nOldColumn(m_nOldLineStart)
  m_nOldLineCount = m_nOldLineCount - nCount
 End If
End If
End Sub

Friend Function LoadFile(ByVal fn As String) As Boolean
On Error Resume Next
Dim m As Long
Dim b() As Byte, s1 As String
'///
Erase m_bData
m_nSize = 0
m_sFileName = vbNullString
'///
m_nPointer = 0
m_nLine = 1
m_nColumn = 1
m_nOldLineCount = 0
m_nOldLineStart = 0
'///
Err.Clear
If (GetAttr(fn) And vbDirectory) = 0 Then
 If Err.Number Then Exit Function
Else
 Exit Function
End If
'///
Open fn For Binary As #1
m = LOF(1)
If m > 0 Then
 #If StrConv Then
 ReDim b(m - 1)
 Get #1, 1, b
 #Else
 ReDim m_bData(m - 1)
 Get #1, 1, m_bData
 #End If
End If
Close
'///
If Err.Number Then Exit Function
#If StrConv Then
s1 = StrConv(b, vbUnicode)
m_nSize = Len(s1)
If m_nSize > 0 Then
 ReDim m_bData(m_nSize - 1)
 CopyMemory m_bData(0), ByVal StrPtr(s1), m_nSize * 2&
End If
#Else
m_nSize = m
#End If
m_sFileName = fn
LoadFile = True
'///
End Function
