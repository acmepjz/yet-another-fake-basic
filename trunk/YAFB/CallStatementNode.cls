VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CallStatementNode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IASTNode

Private m_objFunc As VariableNode

Private Function IASTNode_GetType(nFlags As Long) As clsTypeNode
'nothing
End Function

Private Property Get IASTNode_NodeType() As enumASTNodeType
IASTNode_NodeType = node_callstat
End Property

Friend Property Get This() As IASTNode
Set This = Me
End Property

Friend Sub Create(ByVal obj As VariableNode)
Set m_objFunc = obj
End Sub

Private Function IASTNode_Verify(ByVal objContext As clsVerifyContext) As Boolean
Select Case objContext.Phase
Case verify_const
 IASTNode_Verify = True
Case Else
 IASTNode_Verify = m_objFunc.VerifyEx(objContext, , True)
End Select
End Function

Private Function IASTNode_Codegen(ByVal objContext As clsVerifyContext, ByVal nParam1 As Long, ByVal nParam2 As Long, ByVal nParam3 As Long, ByVal nParam4 As Long) As Long
IASTNode_Codegen = m_objFunc.CodegenEx(objContext, , True)
End Function

Private Function IASTNode_GetProperty(ByVal nProp As enumASTNodeProperty) As Long
'
End Function

