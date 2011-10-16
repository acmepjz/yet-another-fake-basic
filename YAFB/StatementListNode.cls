VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StatementListNode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IASTNode

Private m_objStatment() As IASTNode '1-based
Private m_nCount As Long

Private Function IASTNode_GetProperty(ByVal nProp As enumASTNodeProperty) As Long
'
End Function

Private Function IASTNode_Codegen(ByVal objContext As clsVerifyContext, ByVal nParam1 As Long, ByVal nParam2 As Long, ByVal nParam3 As Long, ByVal nParam4 As Long) As Long
Dim objFunction As FunctionNode
Dim i As Long
'///
Set objFunction = objContext.CurrentFunction
'///
objFunction.ResetTempVariable objContext, 0
For i = 1 To m_nCount
 m_objStatment(i).Codegen objContext, nParam1, nParam2, nParam3, nParam4
 objFunction.ResetTempVariable objContext, 0
Next i
End Function

Private Function IASTNode_GetType(nFlags As Long) As clsTypeNode
'nothing
End Function

Private Property Get IASTNode_NodeType() As enumASTNodeType
IASTNode_NodeType = node_statlist
End Property

Friend Property Get This() As IASTNode
Set This = Me
End Property

Friend Sub AddSubNode(ByVal obj As IASTNode)
m_nCount = m_nCount + 1
ReDim Preserve m_objStatment(1 To m_nCount)
Set m_objStatment(m_nCount) = obj
End Sub

Private Function IASTNode_Verify(ByVal objContext As clsVerifyContext) As Boolean
Dim i As Long
For i = 1 To m_nCount
 If Not m_objStatment(i).Verify(objContext) Then Exit Function
Next i
IASTNode_Verify = True
End Function
