VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ExitStatementNode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IASTNode

Private m_t As typeToken

Private m_nType As enumASTNodeType

Friend Function SetToken(Optional ByVal bGetTypeFromToken As Boolean) As Boolean
m_t = g_tToken
'///
If bGetTypeFromToken Then
 Select Case m_t.nType
 Case keyword_sub, keyword_function, keyword_property
  m_nType = node_funcstat
 Case keyword_for
  m_nType = node_forstat
 Case keyword_do
  m_nType = node_dostat
 Case Else
  PrintError "'Exit " + m_t.sValue + "' is invalid or unsupported"
  Exit Function
 End Select
End If
'///
SetToken = True
End Function

Friend Property Get ExitType() As enumASTNodeType
ExitType = m_nType
End Property

Friend Property Let ExitType(ByVal n As enumASTNodeType)
m_nType = n
End Property

'TODO:release local variables (e.g. With variable)
Private Function IASTNode_Codegen(ByVal objContext As clsVerifyContext, ByVal nParam1 As Long, ByVal nParam2 As Long, ByVal nParam3 As Long, ByVal nParam4 As Long) As Long
Dim hBlock As Long
'///
LLVMBuildBr g_hBuilder, objContext.Module.QueryExitStack(m_nType).GetProperty(prop_endblockhandle)
hBlock = LLVMAppendBasicBlock(objContext.CurrentFunction.FunctionHandle, StrPtr(StrConv("DeadCode", vbFromUnicode)))
LLVMPositionBuilderAtEnd g_hBuilder, hBlock
End Function

Private Function IASTNode_GetProperty(ByVal nProp As enumASTNodeProperty) As Long
'nothing
End Function

Private Function IASTNode_GetType(nFlags As Long) As clsTypeNode
'nothing
End Function

Private Property Get IASTNode_NodeType() As enumASTNodeType
IASTNode_NodeType = node_exitstat
End Property

Private Function IASTNode_Verify(ByVal objContext As clsVerifyContext) As Boolean
Dim bErr As Boolean
'///
If objContext.Module.QueryExitStack(m_nType) Is Nothing Then
 bErr = True
ElseIf m_nType = node_funcstat Then
 Select Case objContext.CurrentFunction.FuncType And &HF&
 Case 0 'sub
  bErr = m_t.nType <> keyword_sub
 Case 1 'function
  bErr = m_t.nType <> keyword_function
 Case 2, 3, 4 'property
  bErr = m_t.nType <> keyword_property
 End Select
End If
If bErr Then
 PrintError "'Exit " + m_t.sValue + "' without '" + m_t.sValue + "'", m_t.nLine, m_t.nColumn
 Exit Function
End If
IASTNode_Verify = True
End Function

Friend Property Get This() As IASTNode
Set This = Me
End Property
