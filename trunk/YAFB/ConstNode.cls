VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ConstNode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IASTNode

Private m_t As typeToken

Private m_objType As clsTypeNode

'================================ LLVM ================================

Private m_hType As Long
Private m_hConst As Long

Private Function IASTNode_Codegen(ByVal objContext As clsVerifyContext, ByVal nParam1 As Long, ByVal nParam2 As Long, ByVal nParam3 As Long, ByVal nParam4 As Long) As Long
IASTNode_Codegen = m_hConst '??
End Function

Private Function IASTNode_GetType(nFlags As Long) As clsTypeNode
nFlags = 0
Set IASTNode_GetType = m_objType
End Function

Private Property Get IASTNode_NodeType() As enumASTNodeType
IASTNode_NodeType = node_const
End Property

Friend Property Get This() As IASTNode
Set This = Me
End Property

Friend Sub SetToken()
m_t = g_tToken
End Sub

Private Function IASTNode_Verify(ByVal objContext As clsVerifyContext) As Boolean
On Error Resume Next
If m_hType = 0 And m_hConst = 0 Then
 Select Case m_t.nType
 Case token_decnum, token_hexnum, token_octnum
  Err.Clear
  m_hType = LLVMInt32Type '??? TODO: Byte, Integer and Long
  Set m_objType = g_objIntrinsicDataTypes(vbLong)
  m_hConst = LLVMConstInt(m_hType, CCur(Val(m_t.sValue)) * 0.0001@, 1)
  If m_hConst = 0 Or Err.Number <> 0 Then
   PrintError "Invalid integer number '" + m_t.sValue + "'", m_t.nLine, m_t.nColumn
   Exit Function
  End If
 Case token_floatnum
  Err.Clear
  If Right(m_t.sValue, 1) = "!" Then
   m_hType = LLVMFloatType
   Set m_objType = g_objIntrinsicDataTypes(vbSingle)
  Else
   m_hType = LLVMDoubleType
   Set m_objType = g_objIntrinsicDataTypes(vbDouble)
  End If
  m_hConst = LLVMConstReal(m_hType, Val(m_t.sValue))
  If m_hConst = 0 Or Err.Number <> 0 Then
   PrintError "Invalid float number '" + m_t.sValue + "'", m_t.nLine, m_t.nColumn
   Exit Function
  End If
 Case Else
  PrintError "Currently unsupported constant '" + m_t.sValue + "'", m_t.nLine, m_t.nColumn
  Exit Function
 End Select
End If
IASTNode_Verify = True '??
End Function

Friend Property Get TypeHandle() As Long
TypeHandle = m_hType
End Property

Friend Property Get ConstHandle() As Long
ConstHandle = m_hConst
End Property
