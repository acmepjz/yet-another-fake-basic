VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DataTypeNode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'TODO:data type system

Implements IASTNode

Private m_tName() As typeToken '1-based
Private m_nNameCount As Long

Private m_objStringSize As IASTNode 'can be Nothing

Private m_objTypeNode As clsTypeNode

Private Function IASTNode_GetProperty(ByVal nProp As enumASTNodeProperty) As Long
'
End Function

Private Function IASTNode_Codegen(ByVal objContext As clsVerifyContext, ByVal nParam1 As Long, ByVal nParam2 As Long, ByVal nParam3 As Long, ByVal nParam4 As Long) As Long
'TODO:
End Function

Private Function IASTNode_GetType(nFlags As Long) As clsTypeNode
nFlags = 0
Set IASTNode_GetType = m_objTypeNode
End Function

Private Property Get IASTNode_NodeType() As enumASTNodeType
IASTNode_NodeType = node_dimtype
End Property

Friend Property Get This() As IASTNode
Set This = Me
End Property

Friend Sub AddNameToken()
m_nNameCount = m_nNameCount + 1
ReDim Preserve m_tName(1 To m_nNameCount)
m_tName(m_nNameCount) = g_tToken
End Sub

Friend Sub AddNameTokenEx(ByRef t As typeToken)
m_nNameCount = m_nNameCount + 1
ReDim Preserve m_tName(1 To m_nNameCount)
m_tName(m_nNameCount) = t
End Sub

Friend Property Get StringSize() As IASTNode
Set StringSize = m_objStringSize
End Property

Friend Property Set StringSize(ByVal obj As IASTNode)
Set m_objStringSize = obj
End Property

'--- internal function
'TODO:fixed size string
Friend Function FindDataType(ByVal objModule As clsSrcParser) As Boolean
On Error Resume Next
Dim v As clsTypeNode 'Variant
'///
If Not m_objTypeNode Is Nothing Then
 FindDataType = True
 Exit Function
End If
'///
If m_nNameCount <= 0 Then Exit Function
'TODO:if m_nNameCount>1
Debug.Assert m_nNameCount = 1
'///
If Not objModule Is Nothing Then
 Err.Clear
 Set v = objModule.SymbolTable.TypeTable.Item(m_tName(1).sValue)
 If Err.Number = 0 Then Set m_objTypeNode = v
End If
If Not m_objTypeNode Is Nothing Then
 FindDataType = True
 Exit Function
End If
'///
Err.Clear
Set v = g_objGlobalTable.TypeTable.Item(m_tName(1).sValue)
If Err.Number = 0 Then Set m_objTypeNode = v
If Not m_objTypeNode Is Nothing Then
 FindDataType = True
 Exit Function
End If
End Function

Friend Function GetDataType() As clsTypeNode
Set GetDataType = m_objTypeNode
End Function

Friend Sub ErrorNotFound()
If m_nNameCount > 0 Then
 PrintError "User defined type '" + m_tName(m_nNameCount).sValue + "' not found", m_tName(m_nNameCount).nLine, m_tName(m_nNameCount).nColumn
Else
 PrintError "User defined type not found", -1, -1
End If
End Sub

Friend Sub ErrorNoByVal()
If m_nNameCount > 0 Then
 PrintError "User defined type '" + m_tName(m_nNameCount).sValue + "' can't be ByVal argument", m_tName(m_nNameCount).nLine, m_tName(m_nNameCount).nColumn
Else
 PrintError "User defined type can't be ByVal argument", -1, -1
End If
End Sub

Friend Sub ErrorNoFixedString()
If m_nNameCount > 0 Then
 PrintError "Type '" + m_tName(m_nNameCount).sValue + "' isn't string, can't use fixed-size string syntax", m_tName(m_nNameCount).nLine, m_tName(m_nNameCount).nColumn
Else
 PrintError "Type isn't string, can't use fixed-size string syntax", -1, -1
End If
End Sub

Private Function IASTNode_Verify(ByVal objContext As clsVerifyContext) As Boolean
'TODO:
Select Case objContext.Phase
Case verify_dim
 If Not FindDataType(objContext.Module) Then
  ErrorNotFound
  Exit Function
 End If
 If Not m_objStringSize Is Nothing Then
  If Not m_objTypeNode.DataType <> vbString Then
   ErrorNoFixedString
   Exit Function
  End If
  'TODO:fixed-size string
  PrintError "Currently fixed-size string is unsupported", -1, -1
  Exit Function
 End If
End Select
'///
IASTNode_Verify = True
End Function
