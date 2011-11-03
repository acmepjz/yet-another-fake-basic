VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DoStatementNode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IASTNode

Private m_nType As enumASTNodeType

Private m_objDoCondition As IASTNode 'can be Nothing
Private m_objLoopCondition As IASTNode 'can be Nothing

Private m_nFlags As Long
'1=Do Until (0=Do While)
'2=Loop Until (0=Loop While)

Private m_objStatement As StatementListNode

'================================ LLVM ================================

Private m_hEndBlock As Long

Private Function IASTNode_Codegen(ByVal objContext As clsVerifyContext, ByVal nParam1 As Long, ByVal nParam2 As Long, ByVal nParam3 As Long, ByVal nParam4 As Long) As Long
Dim obj As IASTNode
Dim objType As clsTypeNode, nFlags As Long
Dim hValue As Long
Dim s As String, lp As Long
'///
Dim hFunction As Long
Dim hBlockDoCondition As Long
Dim hBlockDo As Long
'///
s = StrConv("ConditionTemp", vbFromUnicode)
lp = StrPtr(s)
'///
hFunction = objContext.CurrentFunction.FunctionHandle
If Not m_objDoCondition Is Nothing Then
 hBlockDoCondition = LLVMAppendBasicBlock(hFunction, StrPtr(StrConv("DoCondition", vbFromUnicode)))
End If
hBlockDo = LLVMAppendBasicBlock(hFunction, StrPtr(StrConv("DoBlock", vbFromUnicode)))
m_hEndBlock = LLVMAppendBasicBlock(hFunction, StrPtr(StrConv("DoEnd", vbFromUnicode)))
'///codegen Do
If hBlockDoCondition Then
 LLVMBuildBr g_hBuilder, hBlockDoCondition
 LLVMPositionBuilderAtEnd g_hBuilder, hBlockDoCondition
 hValue = m_objDoCondition.Codegen(objContext, nParam1, nParam2, nParam3, nParam4)
 nFlags = 0
 Set objType = m_objDoCondition.GetType(nFlags)
 '///
 hValue = objType.CodegenConvertToI1(hValue, lp)
 '///
 If m_nFlags And 1& Then 'Do Until
  LLVMBuildCondBr g_hBuilder, hValue, m_hEndBlock, hBlockDo
 Else 'Do While
  LLVMBuildCondBr g_hBuilder, hValue, hBlockDo, m_hEndBlock
 End If
Else
 LLVMBuildBr g_hBuilder, hBlockDo
End If
'///codegen statement list
If m_nType = node_dostat Then objContext.Module.AddExitStack Me
'///
LLVMPositionBuilderAtEnd g_hBuilder, hBlockDo
Set obj = m_objStatement
obj.Codegen objContext, nParam1, nParam2, nParam3, nParam4
'///
If m_nType = node_dostat Then objContext.Module.RemoveExitStack
'///codegen Loop
If hBlockDoCondition Then hBlockDo = hBlockDoCondition
If Not m_objLoopCondition Is Nothing Then
 hValue = m_objLoopCondition.Codegen(objContext, nParam1, nParam2, nParam3, nParam4)
 nFlags = 0
 Set objType = m_objLoopCondition.GetType(nFlags)
 '///
 hValue = objType.CodegenConvertToI1(hValue, lp)
 '///
 If m_nFlags And 2& Then 'Loop Until
  LLVMBuildCondBr g_hBuilder, hValue, m_hEndBlock, hBlockDo
 Else 'Loop While
  LLVMBuildCondBr g_hBuilder, hValue, hBlockDo, m_hEndBlock
 End If
Else
 LLVMBuildBr g_hBuilder, hBlockDo
End If
'///over
LLVMPositionBuilderAtEnd g_hBuilder, m_hEndBlock
End Function

Private Function IASTNode_GetProperty(ByVal nProp As enumASTNodeProperty) As Long
Select Case nProp
Case prop_endblockhandle
 IASTNode_GetProperty = m_hEndBlock
End Select
End Function

Private Function IASTNode_GetType(nFlags As Long) As clsTypeNode
'nothing
End Function

Private Property Get IASTNode_NodeType() As enumASTNodeType
IASTNode_NodeType = m_nType
End Property

Private Function IASTNode_Verify(ByVal objContext As clsVerifyContext) As Boolean
Dim obj As IASTNode
Dim objType As clsTypeNode, nFlags As Long
'///
If m_nType = node_dostat Then objContext.Module.AddExitStack Me
'///
Set obj = m_objStatement
If Not obj.Verify(objContext) Then Exit Function
'///
If m_nType = node_dostat Then objContext.Module.RemoveExitStack
'///check condition data type
If objContext.Phase = verify_all Then
 If Not m_objDoCondition Is Nothing Then
  If Not m_objDoCondition.Verify(objContext) Then Exit Function
  '///
  nFlags = 0
  Set objType = m_objDoCondition.GetType(nFlags)
  If (objType.Flags And &H10&) = 0 Or nFlags <> 0 Then
   PrintError "Data type '" + objType.NameEx(nFlags) + "' can't use for condition of 'Do' statement", -1, -1
   Exit Function
  End If
 End If
 If Not m_objLoopCondition Is Nothing Then
  If Not m_objLoopCondition.Verify(objContext) Then Exit Function
  '///
  nFlags = 0
  Set objType = m_objLoopCondition.GetType(nFlags)
  If (objType.Flags And &H10&) = 0 Or nFlags <> 0 Then
   PrintError "Data type '" + objType.NameEx(nFlags) + "' can't use for condition of 'Loop' statement", -1, -1
   Exit Function
  End If
 End If
End If
'///
IASTNode_Verify = True
End Function

Friend Property Get NodeType() As enumASTNodeType
NodeType = m_nType
End Property

Friend Property Let NodeType(ByVal n As enumASTNodeType)
m_nType = n
End Property

Friend Property Get Flags() As Long
Flags = m_nFlags
End Property

Friend Property Let Flags(ByVal n As Long)
m_nFlags = n
End Property

Friend Property Get DoCondition() As IASTNode
Set DoCondition = m_objDoCondition
End Property

Friend Property Set DoCondition(ByVal obj As IASTNode)
Set m_objDoCondition = obj
End Property

Friend Property Get LoopCondition() As IASTNode
Set LoopCondition = m_objLoopCondition
End Property

Friend Property Set LoopCondition(ByVal obj As IASTNode)
Set m_objLoopCondition = obj
End Property

Friend Property Get StatementList() As StatementListNode
Set StatementList = m_objStatement
End Property

Friend Property Set StatementList(ByVal obj As StatementListNode)
Set m_objStatement = obj
End Property
