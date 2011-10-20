VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GoToStatement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IASTNode

Private m_t As typeToken

'================================ LLVM ================================

Private m_obj As LineNumberStatement
Private m_hBlock As Long

Private Function IASTNode_Codegen(ByVal objContext As clsVerifyContext, ByVal nParam1 As Long, ByVal nParam2 As Long, ByVal nParam3 As Long, ByVal nParam4 As Long) As Long
LLVMBuildBr g_hBuilder, m_obj.BlockHandle
'///over
LLVMPositionBuilderAtEnd g_hBuilder, _
LLVMAppendBasicBlock(objContext.CurrentFunction.FunctionHandle, StrPtr(StrConv("DeadCode", vbFromUnicode)))
End Function

Private Function IASTNode_GetProperty(ByVal nProp As enumASTNodeProperty) As Long
'nothing
End Function

Private Function IASTNode_GetType(nFlags As Long) As clsTypeNode
'nothing
End Function

Private Property Get IASTNode_NodeType() As enumASTNodeType
IASTNode_NodeType = node_gotostat
End Property

Private Function IASTNode_Verify(ByVal objContext As clsVerifyContext) As Boolean
If objContext.Phase = verify_all Then
 Set m_obj = objContext.LookUpLineNumberTable(m_t.sValue)
 If m_obj Is Nothing Then
  PrintError "Line number '" + m_t.sValue + "' not found", m_t.nLine, m_t.nColumn
  Exit Function
 End If
End If
IASTNode_Verify = True
End Function

Friend Sub SetToken()
m_t = g_tToken
End Sub

Friend Property Get This() As IASTNode
Set This = Me
End Property
