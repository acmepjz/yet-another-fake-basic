VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SelectBlockNode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'<selectblock>:{case}<selectconditionlist><br><statlist>
'<selectconditionlist>:{else}|(<selectcondition>{,})*<selectcondition>
'<selectcondition>:<exp>|<exp>{to}<exp>|{is}({=}|{<>}|{>}|{<}|{>=}|{<=})<exp>

Implements IASTNode

Private m_objCond0() As IASTNode '1-based
Private m_objCond1() As IASTNode '1-based, can be Nothing
Private m_nCondType() As enumTokenType 'possible value: keyword_to,token_eq(default),token_ne,token_gt, etc.

Private m_nCondCount As Long 'if 0 then it's Case Else

Private m_objStatement As StatementListNode

Friend Property Get ConditionCount() As Long
ConditionCount = m_nCondCount
End Property

Friend Property Get StatementList() As StatementListNode
Set StatementList = m_objStatement
End Property

Friend Property Set StatementList(ByVal obj As StatementListNode)
Set m_objStatement = obj
End Property

Friend Sub AddCondition(ByVal nType As enumTokenType, ByVal obj0 As IASTNode, ByVal obj1 As IASTNode)
m_nCondCount = m_nCondCount + 1
ReDim Preserve m_objCond0(1 To m_nCondCount)
ReDim Preserve m_objCond1(1 To m_nCondCount)
ReDim Preserve m_nCondType(1 To m_nCondCount)
Set m_objCond0(m_nCondCount) = obj0
Set m_objCond1(m_nCondCount) = obj1
m_nCondType(m_nCondCount) = nType
End Sub

'should not call this function
Private Function IASTNode_Codegen(ByVal objContext As clsVerifyContext, ByVal nParam1 As Long, ByVal nParam2 As Long, ByVal nParam3 As Long, ByVal nParam4 As Long) As Long
PrintPanic "Unexpected IASTNode_Codegen function call on 'Select Case' statement", -1, -1
End Function

Private Function IASTNode_GetProperty(ByVal nProp As enumASTNodeProperty) As Long
'nothing
End Function

Private Function IASTNode_GetType(nFlags As Long) As clsTypeNode
'nothing
End Function

Private Property Get IASTNode_NodeType() As enumASTNodeType
IASTNode_NodeType = node_selectblock
End Property

Private Function IASTNode_Verify(ByVal objContext As clsVerifyContext) As Boolean
Dim obj As IASTNode
Dim i As Long
'///
If objContext.Phase = verify_all Then
 For i = 1 To m_nCondCount
  If Not m_objCond0(i).Verify(objContext) Then Exit Function
  Set obj = m_objCond1(i)
  If Not obj Is Nothing Then
   If Not obj.Verify(objContext) Then Exit Function
  End If
 Next i
End If
'///
Set obj = m_objStatement
If Not obj.Verify(objContext) Then Exit Function
'///
IASTNode_Verify = True
End Function

Friend Sub CodegenEx(ByVal objContext As clsVerifyContext, ByVal hVariable As Long, ByVal objDestType As clsTypeNode, ByVal hBlockEnd As Long)
Dim obj As IASTNode
Dim hFunction As Long
Dim hBlock As Long
Dim hBlockNext As Long, hBlockNext2 As Long
Dim i As Long
Dim s As String, lp As Long, lp2 As Long
Dim nType As enumTokenType
Dim hValue As Long
Dim objSrcType As clsTypeNode, nSrcFlags As Long
'///
s = StrConv("CaseCondition" + vbNullChar + "ConditionTemp", vbFromUnicode)
lp = StrPtr(s)
lp2 = lp + 14
'///
hFunction = objContext.CurrentFunction.FunctionHandle
'///
If m_nCondCount > 0 Then
 hBlock = LLVMAppendBasicBlock(hFunction, StrPtr(StrConv("CaseBlock", vbFromUnicode)))
 '///codegen conditions
 For i = 1 To m_nCondCount
  hBlockNext = LLVMAppendBasicBlock(hFunction, lp)
  nType = m_nCondType(i)
  Select Case nType
  Case token_equal, token_ne, token_gt, token_lt, token_ge, token_le
   Set objSrcType = m_objCond0(i).GetType(nSrcFlags)
   hValue = LLVMBuildLoad(g_hBuilder, hVariable, lp2)
   '///
   LLVMBuildCondBr g_hBuilder, _
   g_objTypeMgr.CodegenBinaryOperator(objContext, hValue, objDestType, _
   m_objCond0(i).Codegen(objContext, 0, 0, 0, 0), objSrcType, nType, Nothing, False), _
   hBlock, hBlockNext
  Case keyword_to
   hBlockNext2 = LLVMAppendBasicBlock(hFunction, lp)
   '///
   Set objSrcType = m_objCond0(i).GetType(nSrcFlags)
   hValue = LLVMBuildLoad(g_hBuilder, hVariable, lp2)
   '///
   LLVMBuildCondBr g_hBuilder, _
   g_objTypeMgr.CodegenBinaryOperator(objContext, hValue, objDestType, _
   m_objCond0(i).Codegen(objContext, 0, 0, 0, 0), objSrcType, token_ge, Nothing, False), _
   hBlockNext2, hBlockNext
   '///
   LLVMPositionBuilderAtEnd g_hBuilder, hBlockNext2
   LLVMBuildCondBr g_hBuilder, _
   g_objTypeMgr.CodegenBinaryOperator(objContext, hValue, objDestType, _
   m_objCond1(i).Codegen(objContext, 0, 0, 0, 0), objSrcType, token_le, Nothing, False), _
   hBlock, hBlockNext
  Case Else
   PrintPanic "Unexpected 'Case' block type: " + CStr(nType), -1, -1
  End Select
  '///
  LLVMPositionBuilderAtEnd g_hBuilder, hBlockNext
 Next i
 LLVMPositionBuilderAtEnd g_hBuilder, hBlock
End If
'///codegen statement list
Set obj = m_objStatement
obj.Codegen objContext, 0, 0, 0, 0
LLVMBuildBr g_hBuilder, hBlockEnd
'///over
If hBlockNext = 0 Then
 hBlockNext = LLVMAppendBasicBlock(hFunction, StrPtr(StrConv("DeadCode", vbFromUnicode)))
End If
LLVMPositionBuilderAtEnd g_hBuilder, hBlockNext
End Sub

Friend Function VerifyConditionDataType(ByVal objDestType As clsTypeNode, ByVal nDestFlags As Long) As Boolean
Dim i As Long
Dim objSrcType As clsTypeNode, nSrcFlags As Long
'///
For i = 1 To m_nCondCount
 nSrcFlags = 0
 Set objSrcType = m_objCond0(i).GetType(nSrcFlags)
 If g_objTypeMgr.CheckTypeConversion(objSrcType, nSrcFlags, objDestType, nDestFlags) = 0 Then
  PrintError "Type mismatch: can't convert data type '" + objSrcType.Name + "' to '" + objDestType.Name + "'", -1, -1
  Exit Function
 End If
 If Not m_objCond1(i) Is Nothing Then
  nSrcFlags = 0
  Set objSrcType = m_objCond1(i).GetType(nSrcFlags)
  If g_objTypeMgr.CheckTypeConversion(objSrcType, nSrcFlags, objDestType, nDestFlags) = 0 Then
   PrintError "Type mismatch: can't convert data type '" + objSrcType.Name + "' to '" + objDestType.Name + "'", -1, -1
   Exit Function
  End If
 End If
Next i
'///
VerifyConditionDataType = True
End Function
