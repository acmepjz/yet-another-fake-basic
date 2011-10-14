VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTypeNode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IASTNode

Private m_nFlags As Long
'&H80000000=intrinsic

Private m_nType As VbVarType

Private m_t As typeToken

'TODO:other

'================================ LLVM ================================

Private m_hType As Long

Private Function IASTNode_Codegen(ByVal objContext As clsVerifyContext, ByVal nParam1 As Long, ByVal nParam2 As Long, ByVal nParam3 As Long, ByVal nParam4 As Long) As Long
'TODO:
End Function

Private Function IASTNode_GetType(nFlags As Long) As clsTypeNode
nFlags = 0
Set IASTNode_GetType = Me '??
End Function

Private Property Get IASTNode_NodeType() As enumASTNodeType
IASTNode_NodeType = node_typestat
End Property

Friend Property Get This() As IASTNode
Set This = Me
End Property

Friend Sub SetIntrinsic(ByVal nType As VbVarType, ByVal sName As String, ByVal hType As Long)
m_nFlags = &H80000000
m_nType = nType
m_t.nType = token_id
m_t.sValue = sName
m_hType = hType
Set g_objIntrinsicDataTypes(nType) = Me
End Sub

'TODO:other

Private Function IASTNode_Verify(ByVal objContext As clsVerifyContext) As Boolean
If m_nFlags And &H80000000 Then
 IASTNode_Verify = True
 Exit Function
End If
'TODO:
End Function

Friend Property Get Flags() As Long
Flags = m_nFlags
End Property

Friend Property Get DataType() As VbVarType
DataType = m_nType
End Property

Friend Property Get Name() As String
Name = m_t.sValue
End Property

Friend Property Get Handle() As Long
Handle = m_hType
End Property

'TODO:
Friend Sub CodegenDefaultConstructor(ByVal objContext As clsVerifyContext, ByVal hVariable As Long)
Select Case m_nType
Case vbByte, vbInteger, vbLong, vbBoolean
 LLVMBuildStore g_hBuilder, LLVMConstInt(m_hType, 0@, 1), hVariable
Case vbSingle, vbDouble
 LLVMBuildStore g_hBuilder, LLVMConstReal(m_hType, 0#), hVariable
End Select
End Sub

Friend Sub CodegenDefaultDestructor(ByVal objContext As clsVerifyContext, ByVal hVariable As Long)
'TODO:
End Sub

Friend Sub CodegenCopyConstructor(ByVal objContext As clsVerifyContext, ByVal hVariable As Long, ByVal hValue As Long)
'TODO:
End Sub

Friend Sub CodegenMoveConstructor(ByVal objContext As clsVerifyContext, ByVal hVariable As Long, ByVal hValue As Long)
'TODO:
End Sub
