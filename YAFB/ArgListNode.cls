VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ArgListNode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'<arglist>:(({byval}?<exp>)?{,})*{byval}?<exp> 'ByVal???

Implements IASTNode

Private m_objExpression() As IASTNode '1-based, can be Nothing
Private m_nFlags() As Long
'1=ByVal (exp is a pointer)
Private m_nCount As Long

Friend Property Get ArgumentCount() As Long
ArgumentCount = m_nCount
End Property

Friend Property Get ArgumentNode(ByVal nIndex As Long) As IASTNode
Set ArgumentNode = m_objExpression(nIndex)
End Property

Friend Property Get ArgumentFlags(ByVal nIndex As Long) As Long
ArgumentFlags = m_nFlags(nIndex)
End Property

Private Function IASTNode_Codegen(ByVal objContext As clsVerifyContext, ByVal nParam1 As Long, ByVal nParam2 As Long, ByVal nParam3 As Long, ByVal nParam4 As Long) As Long
'TODO:
End Function

Private Function IASTNode_GetProperty(ByVal nProp As enumASTNodeProperty) As Long
'
End Function

Private Function IASTNode_GetType(nFlags As Long) As clsTypeNode
'nothing
End Function

Private Property Get IASTNode_NodeType() As enumASTNodeType
IASTNode_NodeType = node_arglist
End Property

Friend Property Get This() As IASTNode
Set This = Me
End Property

Friend Sub AddSubNode(ByVal obj As IASTNode, ByVal nFlags As Long)
m_nCount = m_nCount + 1
ReDim Preserve m_objExpression(1 To m_nCount)
ReDim Preserve m_nFlags(1 To m_nCount)
Set m_objExpression(m_nCount) = obj
m_nFlags(m_nCount) = nFlags
End Sub

Private Function IASTNode_Verify(ByVal objContext As clsVerifyContext) As Boolean
Dim i As Long
For i = 1 To m_nCount
 If Not m_objExpression(i) Is Nothing Then
  If Not m_objExpression(i).Verify(objContext) Then Exit Function
 End If
Next i
'///
IASTNode_Verify = True
End Function


