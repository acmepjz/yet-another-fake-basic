VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ForStatementNode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IASTNode

Private m_objVariable As VariableNode
Private m_objStartValue As IASTNode
Private m_objEndValue As IASTNode
Private m_objStepValue As IASTNode 'can be Nothing

Private m_objStatement As StatementListNode

'================================ LLVM ================================

Private m_hForEndBlock As Long

Private Function IASTNode_GetProperty(ByVal nProp As enumASTNodeProperty) As Long
Select Case nProp
Case prop_endblockhandle
 IASTNode_GetProperty = m_hForEndBlock
End Select
End Function

Friend Property Get Variable() As VariableNode
Set Variable = m_objVariable
End Property

Friend Property Set Variable(ByVal obj As VariableNode)
Set m_objVariable = obj
End Property

Friend Property Get StartValue() As IASTNode
Set StartValue = m_objStartValue
End Property

Friend Property Set StartValue(ByVal obj As IASTNode)
Set m_objStartValue = obj
End Property

Friend Property Get EndValue() As IASTNode
Set EndValue = m_objEndValue
End Property

Friend Property Set EndValue(ByVal obj As IASTNode)
Set m_objEndValue = obj
End Property

Friend Property Get StepValue() As IASTNode
Set StepValue = m_objStepValue
End Property

Friend Property Set StepValue(ByVal obj As IASTNode)
Set m_objStepValue = obj
End Property

Friend Property Get StatementList() As StatementListNode
Set StatementList = m_objStatement
End Property

Friend Property Set StatementList(ByVal obj As StatementListNode)
Set m_objStatement = obj
End Property

Private Function IASTNode_Codegen(ByVal objContext As clsVerifyContext, ByVal nParam1 As Long, ByVal nParam2 As Long, ByVal nParam3 As Long, ByVal nParam4 As Long) As Long
Dim obj As IASTNode
Dim objSrcType As clsTypeNode, nSrcFlags As Long
Dim objDestType As clsTypeNode, nDestFlags As Long
Dim hValue As Long
Dim hVariable As Long
Dim nTag As Long
Dim hEndVariable As Long
Dim hStepVariable As Long
'///
Dim hFunction As Long
Dim hBlockForStepCheckGreaterThanZero As Long
Dim hBlockForStepCheckLesserThanZero As Long
Dim hBlockForStepGreaterThanZero As Long
Dim hBlockForStepLesserThanZero As Long
Dim hBlockFor As Long
'///
nTag = ObjPtr(Me)
'///
Set obj = m_objVariable
Set objDestType = obj.GetType(nDestFlags)
'///
nSrcFlags = 0
Set objSrcType = m_objStartValue.GetType(nSrcFlags)
'///
hValue = g_objTypeMgr.CodegenTypeConversion(objContext, m_objStartValue.Codegen(objContext, nParam1, nParam2, nParam3, nParam4), _
objSrcType, nSrcFlags, objDestType, nDestFlags, False)
'///
nSrcFlags = 0
Set objSrcType = m_objEndValue.GetType(nSrcFlags)
'///
'hEndVariable = objContext.CurrentFunction.GetNewVariable(objContext, objDestType, "ForEndVariable")
hEndVariable = objContext.CurrentFunction.GetTempVariable(objContext, objDestType, nTag)
LLVMBuildStore g_hBuilder, _
g_objTypeMgr.CodegenTypeConversion(objContext, m_objEndValue.Codegen(objContext, nParam1, nParam2, nParam3, nParam4), _
objSrcType, nSrcFlags, objDestType, nDestFlags, False), hEndVariable
'///
If m_objStepValue Is Nothing Then
 nSrcFlags = nDestFlags
 Set objSrcType = objDestType
 '///
 'hStepVariable = objContext.CurrentFunction.GetNewVariable(objContext, objSrcType, "ForStepVariable")
 hStepVariable = objContext.CurrentFunction.GetTempVariable(objContext, objSrcType, nTag)
 LLVMBuildStore g_hBuilder, objSrcType.CodegenOneValue, hStepVariable
Else
 nSrcFlags = 0
 Set objSrcType = m_objStepValue.GetType(nSrcFlags)
 '///
 'hStepVariable = objContext.CurrentFunction.GetNewVariable(objContext, objSrcType, "ForStepVariable")
 hStepVariable = objContext.CurrentFunction.GetTempVariable(objContext, objSrcType, nTag)
 LLVMBuildStore g_hBuilder, _
 g_objTypeMgr.CodegenTypeConversion(objContext, m_objStepValue.Codegen(objContext, nParam1, nParam2, nParam3, nParam4), _
 objSrcType, nSrcFlags, objDestType, nDestFlags, False), hStepVariable
End If
'///
hVariable = m_objVariable.CodegenEx(objContext, True)
LLVMBuildStore g_hBuilder, hValue, hVariable
'///
hFunction = objContext.CurrentFunction.FunctionHandle
hBlockForStepCheckGreaterThanZero = LLVMAppendBasicBlock(hFunction, StrPtr(StrConv("ForStepCheckGreaterThanZero", vbFromUnicode)))
hBlockForStepCheckLesserThanZero = LLVMAppendBasicBlock(hFunction, StrPtr(StrConv("ForStepCheckLesserThanZero", vbFromUnicode)))
hBlockForStepGreaterThanZero = LLVMAppendBasicBlock(hFunction, StrPtr(StrConv("ForStepGreaterThanZero", vbFromUnicode)))
hBlockForStepLesserThanZero = LLVMAppendBasicBlock(hFunction, StrPtr(StrConv("ForStepLesserThanZero", vbFromUnicode)))
hBlockFor = LLVMAppendBasicBlock(hFunction, StrPtr(StrConv("ForBlock", vbFromUnicode)))
m_hForEndBlock = LLVMAppendBasicBlock(hFunction, StrPtr(StrConv("ForEnd", vbFromUnicode)))
'///
LLVMBuildBr g_hBuilder, hBlockForStepCheckGreaterThanZero
'///check if step>0
LLVMPositionBuilderAtEnd g_hBuilder, hBlockForStepCheckGreaterThanZero
hValue = LLVMBuildLoad(g_hBuilder, hStepVariable, StrPtr(StrConv("StepTemp", vbFromUnicode)))
Select Case objSrcType.DataType
Case vbByte, vbUnsignedInteger, vbUnsignedLong, vbUnsignedLongLong, vbUIntPtr_t
 hValue = LLVMBuildICmp(g_hBuilder, LLVMIntUGT, hValue, LLVMConstInt(objSrcType.Handle, 0@, 1), StrPtr(StrConv("CmpTemp", vbFromUnicode)))
Case vbSignedByte, vbInteger, vbLong, vbLongLong, vbIntPtr_t, vbCurrency
 hValue = LLVMBuildICmp(g_hBuilder, LLVMIntSGT, hValue, LLVMConstInt(objSrcType.Handle, 0@, 1), StrPtr(StrConv("CmpTemp", vbFromUnicode)))
Case vbSingle, vbDouble
 hValue = LLVMBuildFCmp(g_hBuilder, LLVMRealOGT, hValue, LLVMConstReal(objSrcType.Handle, 0#), StrPtr(StrConv("CmpTemp", vbFromUnicode)))
End Select
LLVMBuildCondBr g_hBuilder, hValue, hBlockForStepGreaterThanZero, hBlockForStepCheckLesserThanZero
'///check if step<0
LLVMPositionBuilderAtEnd g_hBuilder, hBlockForStepCheckLesserThanZero
Select Case objSrcType.DataType
Case vbByte, vbUnsignedInteger, vbUnsignedLong, vbUnsignedLongLong, vbUIntPtr_t
 hValue = LLVMConstInt(LLVMInt1Type, 0@, 1)
Case vbSignedByte, vbInteger, vbLong, vbLongLong, vbIntPtr_t, vbCurrency
 hValue = LLVMBuildLoad(g_hBuilder, hStepVariable, StrPtr(StrConv("StepTemp", vbFromUnicode)))
 hValue = LLVMBuildICmp(g_hBuilder, LLVMIntSLT, hValue, LLVMConstInt(objSrcType.Handle, 0@, 1), StrPtr(StrConv("CmpTemp", vbFromUnicode)))
Case vbSingle, vbDouble
 hValue = LLVMBuildLoad(g_hBuilder, hStepVariable, StrPtr(StrConv("StepTemp", vbFromUnicode)))
 hValue = LLVMBuildFCmp(g_hBuilder, LLVMRealOLT, hValue, LLVMConstReal(objSrcType.Handle, 0#), StrPtr(StrConv("CmpTemp", vbFromUnicode)))
End Select
LLVMBuildCondBr g_hBuilder, hValue, hBlockForStepLesserThanZero, hBlockFor
'///if step>0 then if v>end then over
LLVMPositionBuilderAtEnd g_hBuilder, hBlockForStepGreaterThanZero
hValue = LLVMBuildLoad(g_hBuilder, hVariable, StrPtr(StrConv("ForVariableTemp", vbFromUnicode)))
hValue = g_objTypeMgr.CodegenBinaryOperator(objContext, hValue, objDestType, _
LLVMBuildLoad(g_hBuilder, hEndVariable, StrPtr(StrConv("ForEndTemp", vbFromUnicode))), _
objDestType, token_gt, Nothing, False)
LLVMBuildCondBr g_hBuilder, hValue, m_hForEndBlock, hBlockFor
'///if step<0 then if v<end then over
LLVMPositionBuilderAtEnd g_hBuilder, hBlockForStepLesserThanZero
hValue = LLVMBuildLoad(g_hBuilder, hVariable, StrPtr(StrConv("ForVariableTemp", vbFromUnicode)))
hValue = g_objTypeMgr.CodegenBinaryOperator(objContext, hValue, objDestType, _
LLVMBuildLoad(g_hBuilder, hEndVariable, StrPtr(StrConv("ForEndTemp", vbFromUnicode))), _
objDestType, token_lt, Nothing, False)
LLVMBuildCondBr g_hBuilder, hValue, m_hForEndBlock, hBlockFor
'///codegen for body
objContext.Module.AddExitStack Me
'///
LLVMPositionBuilderAtEnd g_hBuilder, hBlockFor
Set obj = m_objStatement
obj.Codegen objContext, nParam1, nParam2, nParam3, nParam4
'///
objContext.Module.RemoveExitStack
'///
hValue = LLVMBuildLoad(g_hBuilder, hVariable, StrPtr(StrConv("ForVariableTemp", vbFromUnicode)))
hValue = g_objTypeMgr.CodegenBinaryOperator(objContext, hValue, objDestType, _
LLVMBuildLoad(g_hBuilder, hStepVariable, StrPtr(StrConv("StepTemp", vbFromUnicode))), _
objSrcType, token_plus, objDestType, False)
LLVMBuildStore g_hBuilder, hValue, hVariable
LLVMBuildBr g_hBuilder, hBlockForStepCheckGreaterThanZero
'///over
LLVMPositionBuilderAtEnd g_hBuilder, m_hForEndBlock
objContext.CurrentFunction.ResetTempVariable objContext, nTag
End Function

Private Function IASTNode_GetType(nFlags As Long) As clsTypeNode
'nothing
End Function

Private Property Get IASTNode_NodeType() As enumASTNodeType
IASTNode_NodeType = node_forstat
End Property

Private Function IASTNode_Verify(ByVal objContext As clsVerifyContext) As Boolean
Dim obj As IASTNode
Dim objSrcType As clsTypeNode, nSrcFlags As Long
Dim objDestType As clsTypeNode, nDestFlags As Long
'///
objContext.Module.AddExitStack Me
'///
Set obj = m_objStatement
If Not obj.Verify(objContext) Then Exit Function
'///
objContext.Module.RemoveExitStack
'///
If objContext.Phase = verify_all Then
 'TODO: check it's fixed variable (e.g. For x(i)= ... is not allowed)
 If Not m_objVariable.VerifyEx(objContext, True) Then Exit Function
 Set obj = m_objVariable
 Set objDestType = obj.GetType(nDestFlags)
 '///
 If (objDestType.Flags And &H40&) = 0 Then
  PrintError "Type mismatch: can't use data type '" + objDestType.Name + "' for 'For' variable", -1, -1
  Exit Function
 End If
 '///
 If Not m_objStartValue.Verify(objContext) Then Exit Function
 nSrcFlags = 0
 Set objSrcType = m_objStartValue.GetType(nSrcFlags)
 If g_objTypeMgr.CheckTypeConversion(objSrcType, nSrcFlags, objDestType, nDestFlags) = 0 Then
  PrintError "Type mismatch: can't convert data type '" + objSrcType.Name + "' to '" + objDestType.Name + "'", -1, -1
  Exit Function
 End If
 '///
 If Not m_objEndValue.Verify(objContext) Then Exit Function
 nSrcFlags = 0
 Set objSrcType = m_objEndValue.GetType(nSrcFlags)
 If g_objTypeMgr.CheckTypeConversion(objSrcType, nSrcFlags, objDestType, nDestFlags) = 0 Then
  PrintError "Type mismatch: can't convert data type '" + objSrcType.Name + "' to '" + objDestType.Name + "'", -1, -1
  Exit Function
 End If
 '///
 If Not m_objStepValue Is Nothing Then
  If Not m_objStepValue.Verify(objContext) Then Exit Function
  nSrcFlags = 0
  Set objSrcType = m_objStepValue.GetType(nSrcFlags)
  If g_objTypeMgr.CheckTypeConversion(objSrcType, nSrcFlags, objDestType, nDestFlags) = 0 Then
   PrintError "Type mismatch: can't convert data type '" + objSrcType.Name + "' to '" + objDestType.Name + "'", -1, -1
   Exit Function
  End If
 End If
End If
'///
IASTNode_Verify = True
End Function
