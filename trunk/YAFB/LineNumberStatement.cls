VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LineNumberStatement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IASTNode

Private m_t As typeToken

'================================ LLVM ================================

Private m_hBlock As Long

Friend Property Get BlockHandle() As Long
BlockHandle = m_hBlock
End Property

Friend Sub CodegenLineNumber(ByVal objContext As clsVerifyContext)
m_hBlock = LLVMAppendBasicBlock(objContext.CurrentFunction.FunctionHandle, StrPtr(StrConv(m_t.sValue, vbFromUnicode)))
End Sub

Private Function IASTNode_Codegen(ByVal objContext As clsVerifyContext, ByVal nParam1 As Long, ByVal nParam2 As Long, ByVal nParam3 As Long, ByVal nParam4 As Long) As Long
LLVMBuildBr g_hBuilder, m_hBlock
LLVMPositionBuilderAtEnd g_hBuilder, m_hBlock
End Function

Private Function IASTNode_GetProperty(ByVal nProp As enumASTNodeProperty) As Long
'nothing
End Function

Private Function IASTNode_GetType(nFlags As Long) As clsTypeNode
'nothing
End Function

Private Property Get IASTNode_NodeType() As enumASTNodeType
IASTNode_NodeType = node_linenumberstat
End Property

Private Function IASTNode_Verify(ByVal objContext As clsVerifyContext) As Boolean
IASTNode_Verify = True '??
End Function

Friend Sub SetToken()
m_t = g_tToken
End Sub

Friend Function Register(ByVal objTable As clsSymbolTable) As Boolean
On Error Resume Next
Err.Clear
objTable.LineNumberTable.Add Me, m_t.sValue
If Err.Number Then
 PrintError "Line number '" + m_t.sValue + "' already exists", m_t.nLine, m_t.nColumn
 Exit Function
End If
Register = True
End Function

Friend Property Get This() As IASTNode
Set This = Me
End Property

