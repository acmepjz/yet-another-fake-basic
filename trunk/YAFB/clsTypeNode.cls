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
'bit 0-3:
'0=other
'1=can be bit cast to signed integer (Integer,Long,Boolean,etc.) Currency isn't of this type
'2=can be bit cast to unsigned integer (Byte,etc.)
'3=can be bit cast to float number (Single,Double,Date)
'etc.
'///
'&H10&=can be Const
'&H20&=can be convert to i1
'&H40&=can be used for 'For' variable (CodegenOneValue avaliable)
'&H80000000=intrinsic

Private m_nType As VbVarType

Private m_nSize As Long

Private m_t As typeToken

'TODO:other

'================================ LLVM ================================

Private m_hType As Long

Private m_hDefaultValue As Long
Private m_hOneValue As Long

Private m_hFunctionScalarDestructor As Long
Private m_hFunctionVectorDestructor As Long

Private m_hFunctionSafeArrayDestructor As Long

Friend Function CodegenConvertToI1(ByVal hValue As Long, ByVal lpName As String, Optional ByVal bIsConstant As Boolean, Optional ByVal hBuilder As Long) As Long
'Dim hVariable As Long
'Dim hConst As Long
'///
If hBuilder = 0 Then hBuilder = g_hBuilder
'///
Select Case m_nFlags And &HF&
Case 1, 2 'integer
 If bIsConstant Then
  CodegenConvertToI1 = LLVMConstICmp(LLVMIntNE, hValue, CodegenDefaultValue)
 Else
  CodegenConvertToI1 = LLVMBuildICmp(hBuilder, LLVMIntNE, hValue, CodegenDefaultValue, lpName)
 End If
Case 3 'float point
 If bIsConstant Then
  CodegenConvertToI1 = LLVMConstFCmp(LLVMRealUNE, hValue, CodegenDefaultValue)
 Else
  CodegenConvertToI1 = LLVMBuildFCmp(hBuilder, LLVMRealUNE, hValue, CodegenDefaultValue, lpName)
 End If
Case Else
 Select Case m_nType
 Case vbCurrency
  If bIsConstant Then
   CodegenConvertToI1 = LLVMConstICmp(LLVMIntNE, hValue, CodegenDefaultValue)
  Else
   CodegenConvertToI1 = LLVMBuildICmp(hBuilder, LLVMIntNE, hValue, CodegenDefaultValue, lpName)
  End If
 Case vbDecimal
'  '///dirty workaround
'  hVariable = objContext.CurrentFunction.GetTempVariable(objContext, Me, -1)
'  LLVMBuildStore hBuilder, hValue, hVariable
'  hConst = LLVMConstInt(LLVMInt32Type, 0@, 1)
'  CodegenConvertToI1 = LLVMBuildOr(hBuilder, LLVMBuildOr(hBuilder, _
'  LLVMBuildICmp, _
'  LLVMBuildICmp, lpName), _
'  LLVMBuildICmp, lpName)
'  'blah blah blah... we don't need dirty workaround
'  objContext.CurrentFunction.RemoveTempVariable hVariable
  CodegenConvertToI1 = LLVMBuildOr(hBuilder, _
  LLVMBuildICmp(hBuilder, LLVMIntNE, LLVMBuildExtractValue(hBuilder, hValue, 3, lpName), LLVMConstInt(LLVMInt32Type, 0@, 1), lpName), _
  LLVMBuildICmp(hBuilder, LLVMIntNE, LLVMBuildExtractValue(hBuilder, hValue, 4, lpName), LLVMConstInt(LLVMInt64Type, 0@, 1), lpName), lpName)
 Case Else
  'TODO: other
 End Select
End Select
End Function

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

Friend Sub SetIntrinsic(ByVal nType As VbVarType, ByVal sName As String, ByVal hType As Long, ByVal nSize As Long, ByVal nFlags As Long)
m_nFlags = &H80000000 Or nFlags
m_nType = nType
m_nSize = nSize
m_t.nType = token_id
m_t.sValue = sName
m_hType = hType
Set g_objIntrinsicDataTypes(nType) = Me
g_objGlobalTable.TypeTable.Add Me, sName
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

Friend Function NameEx(ByVal nFlags As Long) As String
Dim s As String
s = m_t.sValue
If nFlags And &H400& Then s = s + "()"
NameEx = s
End Function

Friend Property Get Handle() As Long
Handle = m_hType
End Property

Friend Property Get Size() As Long
Size = m_nSize
End Property

'TODO: other (e.g. object etc.)
Friend Sub CodegenDefaultConstructor(ByVal hVariable As Long, Optional ByVal hBuilder As Long)
Dim h(7) As Long
Dim m As Currency
'///
If hBuilder = 0 Then hBuilder = g_hBuilder
'///
m = LLVMABISizeOfType(g_hTargetData, m_hType)
If m <= 0.0008@ Then 'LLVM generates extremely stupid code :(
 LLVMBuildStore hBuilder, CodegenDefaultValue, hVariable
Else 'so we should call 'memset' instead... (???)
 h(0) = LLVMBuildPointerCast(hBuilder, hVariable, LLVMPointerType(LLVMInt8Type, 0), ByVal StrPtr(StrConv("ByRefTemp", vbFromUnicode)))
 h(1) = LLVMConstNull(LLVMInt8Type)
 h(2) = LLVMBuildIntCast(hBuilder, LLVMSizeOf(m_hType), g_hTypeIntPtr_t, ByVal StrPtr(StrConv("SizeTemp", vbFromUnicode)))
 h(3) = LLVMConstNull(LLVMInt32Type) 'align=unknown
 h(4) = LLVMConstNull(LLVMInt1Type) 'isVolatile=False
 LLVMBuildCall hBuilder, RuntimeLibraryGetFunction(internal_llvm_memset), h(0), 5, StrPtr(vbNullChar)
End If
End Sub

Friend Sub CodegenDefaultDestructor(ByVal hVariable As Long, Optional ByVal hBuilder As Long)
If hBuilder = 0 Then hBuilder = g_hBuilder
'TODO: string, variant, type, etc.
End Sub

Friend Function GetDefaultScalarDestructorFunction() As Long
Dim hType(3) As Long
Dim p0 As Long
Dim hFunctionType As Long
Dim hFunction As Long
Dim hBuilder As Long
'///
hFunction = m_hFunctionScalarDestructor
If hFunction = 0 Then
 hBuilder = LLVMCreateBuilder
 '///
 hType(0) = LLVMPointerType(m_hType, 0)
 hFunctionType = LLVMFunctionType(LLVMVoidType, hType(0), 1, 0)
 hFunction = LLVMAddFunction(g_hModule, StrPtr(StrConv("YAFB.ScalarDestructor." + m_t.sValue, vbFromUnicode)), hFunctionType)
 p0 = LLVMGetParam(hFunction, 0)
 LLVMSetValueName p0, StrPtr(StrConv("lp", vbFromUnicode))
 LLVMAddAttribute p0, LLVMNoCaptureAttribute
 LLVMSetLinkage hFunction, LLVMPrivateLinkage
 LLVMAddFunctionAttr hFunction, LLVMInlineHintAttribute Or LLVMNoUnwindAttribute
 m_hFunctionScalarDestructor = hFunction
 '///
 LLVMPositionBuilderAtEnd hBuilder, LLVMAppendBasicBlock(hFunction, StrPtr(StrConv("FunctionEntry", vbFromUnicode)))
 CodegenDefaultDestructor p0, hBuilder
 LLVMBuildRetVoid hBuilder
 '///
 LLVMDisposeBuilder hBuilder
End If
'///
GetDefaultScalarDestructorFunction = hFunction
End Function

Friend Function GetDefaultSafeArrayDestructorFunction() As Long
'Private Type SAFEARRAY2D
'    cDims As Integer
'    fFeatures As Integer
'    cbElements As Long
'    cLocks As Long
'    pvData As Long
'    Bounds(0 To 1) As SAFEARRAYBOUND
'End Type
'Private Type SAFEARRAYBOUND
'    cElements As Long
'    lLbound As Long
'End Type
Dim hType(7) As Long
Dim p0 As Long
Dim hFunctionType As Long
Dim hFunction As Long
Dim hBuilder As Long
Dim hBlock(7) As Long
Dim hValue As Long
Dim hValue_SafeArray As Long
Dim hValue_Dimension As Long
Dim hValue_Pointer As Long
Dim hValue_SafeArrayBound As Long
Dim hVariable_i As Long
Dim hVariable_m As Long
Dim tmp As Long, lp As Long
'///
hFunction = m_hFunctionSafeArrayDestructor
If hFunction = 0 Then
 lp = VarPtr(tmp)
 hBuilder = LLVMCreateBuilder
 '///
 hType(0) = LLVMPointerType(LLVMPointerType(g_hTypeSafeArray, 0), 0)
 hFunctionType = LLVMFunctionType(LLVMVoidType, hType(0), 1, 0)
 hFunction = LLVMAddFunction(g_hModule, StrPtr(StrConv("YAFB.SafeArrayDestructor." + m_t.sValue + ".i" + CStr(g_nWordSize * 8&), vbFromUnicode)), hFunctionType)
 p0 = LLVMGetParam(hFunction, 0)
 LLVMSetValueName p0, StrPtr(StrConv("lp", vbFromUnicode))
 LLVMAddAttribute p0, LLVMNoCaptureAttribute
 LLVMSetLinkage hFunction, LLVMPrivateLinkage
 LLVMAddFunctionAttr hFunction, LLVMInlineHintAttribute Or LLVMNoUnwindAttribute
 m_hFunctionSafeArrayDestructor = hFunction
 '///
 LLVMPositionBuilderAtEnd hBuilder, LLVMAppendBasicBlock(hFunction, lp)
 hBlock(0) = LLVMAppendBasicBlock(hFunction, lp)
 hBlock(1) = LLVMAppendBasicBlock(hFunction, lp) 'loop for calc element count
 hBlock(2) = LLVMAppendBasicBlock(hFunction, lp)
 hBlock(3) = LLVMAppendBasicBlock(hFunction, lp) 'if it's fixed-size array
 hBlock(4) = LLVMAppendBasicBlock(hFunction, lp) 'if it's dynamic array
 hBlock(7) = LLVMAppendBasicBlock(hFunction, lp) 'end
 '///check if input is empty
 hVariable_m = LLVMBuildAlloca(hBuilder, g_hTypeIntPtr_t, StrPtr("m"))
 hVariable_i = LLVMBuildAlloca(hBuilder, LLVMInt32Type, StrPtr("i"))
 hValue_SafeArray = LLVMBuildLoad(hBuilder, p0, lp)
 LLVMBuildCondBr hBuilder, LLVMBuildIsNull(hBuilder, hValue_SafeArray, lp), hBlock(7), hBlock(0)
 '///calc element count
 LLVMPositionBuilderAtEnd hBuilder, hBlock(0)
 hValue_Pointer = LLVMBuildLoad(hBuilder, LLVMBuildStructGEP(hBuilder, hValue_SafeArray, 4, lp), lp)
 hValue_Dimension = LLVMBuildZExt(hBuilder, LLVMBuildLoad(hBuilder, LLVMBuildStructGEP(hBuilder, hValue_SafeArray, 0, lp), lp), LLVMInt32Type, lp)
 hValue_SafeArrayBound = LLVMBuildStructGEP(hBuilder, hValue_SafeArray, 5, lp)
 LLVMBuildStore hBuilder, LLVMConstNull(LLVMInt32Type), hVariable_i
 LLVMBuildStore hBuilder, LLVMConstInt(g_hTypeIntPtr_t, 0.0001@, 0), hVariable_m
 LLVMBuildCondBr hBuilder, _
 LLVMBuildOr(hBuilder, LLVMBuildIsNull(hBuilder, hValue_Pointer, lp), LLVMBuildIsNull(hBuilder, hValue_Dimension, lp), _
 lp), hBlock(7), hBlock(1)
 '///
 LLVMPositionBuilderAtEnd hBuilder, hBlock(1)
 hValue = LLVMBuildLoad(hBuilder, hVariable_i, lp)
 LLVMBuildStore hBuilder, LLVMBuildNUWMul(hBuilder, LLVMBuildZExt(hBuilder, _
 LLVMBuildLoad(hBuilder, LLVMBuildStructGEP(hBuilder, LLVMBuildGEP(hBuilder, hValue_SafeArrayBound, hValue, 1, lp), 0, lp), lp), g_hTypeIntPtr_t, lp), _
 LLVMBuildLoad(hBuilder, hVariable_m, lp), _
 lp), hVariable_m
 hValue = LLVMBuildAdd(hBuilder, hValue, LLVMConstInt(LLVMInt32Type, 0.0001@, 0), lp)
 LLVMBuildStore hBuilder, hValue, hVariable_i
 LLVMBuildCondBr hBuilder, LLVMBuildICmp(hBuilder, LLVMIntULT, hValue, hValue_Dimension, lp), hBlock(1), hBlock(2)
 '///done, call vector destructor
 LLVMPositionBuilderAtEnd hBuilder, hBlock(2)
 hType(0) = LLVMBuildPointerCast(hBuilder, hValue_Pointer, LLVMPointerType(m_hType, 0), lp)
 hType(1) = LLVMBuildLoad(hBuilder, hVariable_m, lp)
 LLVMBuildCall hBuilder, GetDefaultVectorDestructorFunction, hType(0), 2, lp
 'check if it's dynamic array
 LLVMBuildCondBr hBuilder, LLVMBuildIsNotNull(hBuilder, LLVMBuildAnd(hBuilder, _
 LLVMBuildLoad(hBuilder, LLVMBuildStructGEP(hBuilder, hValue_SafeArray, 1, lp), lp), _
 LLVMConstInt(LLVMInt16Type, 0.0016@, 1), lp), lp), hBlock(3), hBlock(4)
 '///it's fixed-size array
 LLVMPositionBuilderAtEnd hBuilder, hBlock(3)
 hType(0) = LLVMBuildPointerCast(hBuilder, hValue_Pointer, LLVMPointerType(LLVMInt8Type, 0), lp)
 hType(1) = LLVMConstNull(LLVMInt8Type)
 hType(2) = LLVMBuildLoad(hBuilder, hVariable_m, lp)
 hType(3) = LLVMConstNull(LLVMInt32Type)
 hType(4) = LLVMConstNull(LLVMInt1Type)
 LLVMBuildCall hBuilder, RuntimeLibraryGetFunction(internal_llvm_memset), hType(0), 5, lp
 LLVMBuildBr hBuilder, hBlock(7)
 '///it's dynamic array
 LLVMPositionBuilderAtEnd hBuilder, hBlock(4)
 LLVMBuildFree hBuilder, hValue_Pointer
 LLVMBuildFree hBuilder, hValue_SafeArray
 LLVMBuildStore hBuilder, LLVMConstPointerNull(LLVMPointerType(g_hTypeSafeArray, 0)), p0
 LLVMBuildBr hBuilder, hBlock(7)
 '///over
 LLVMPositionBuilderAtEnd hBuilder, hBlock(7)
 LLVMBuildRetVoid hBuilder
 '///
 LLVMDisposeBuilder hBuilder
End If
GetDefaultSafeArrayDestructorFunction = hFunction
End Function

Friend Function GetDefaultVectorDestructorFunction() As Long
Dim hType(3) As Long
Dim p0 As Long, p1 As Long
Dim hFunctionType As Long
Dim hFunction As Long
Dim hBuilder As Long
Dim hBlockDo As Long
Dim hBlockEnd As Long
Dim hVariable As Long
Dim tmp As Long, lp As Long
'///
hFunction = m_hFunctionVectorDestructor
If hFunction = 0 Then
 lp = VarPtr(tmp)
 hBuilder = LLVMCreateBuilder
 '///
 hType(0) = LLVMPointerType(m_hType, 0)
 hType(1) = g_hTypeIntPtr_t
 hFunctionType = LLVMFunctionType(LLVMVoidType, hType(0), 2, 0)
 hFunction = LLVMAddFunction(g_hModule, StrPtr(StrConv("YAFB.VectorDestructor." + m_t.sValue + ".i" + CStr(g_nWordSize * 8&), vbFromUnicode)), hFunctionType)
 p0 = LLVMGetParam(hFunction, 0)
 p1 = LLVMGetParam(hFunction, 1)
 LLVMSetValueName p0, StrPtr(StrConv("lp", vbFromUnicode))
 LLVMAddAttribute p0, LLVMNoCaptureAttribute
 LLVMSetValueName p1, StrPtr(StrConv("nCount", vbFromUnicode))
 LLVMSetLinkage hFunction, LLVMPrivateLinkage
 LLVMAddFunctionAttr hFunction, LLVMInlineHintAttribute Or LLVMNoUnwindAttribute
 m_hFunctionVectorDestructor = hFunction
 '///
 LLVMPositionBuilderAtEnd hBuilder, LLVMAppendBasicBlock(hFunction, StrPtr(StrConv("FunctionEntry", vbFromUnicode)))
 hBlockDo = LLVMAppendBasicBlock(hFunction, StrPtr(StrConv("DoBlock", vbFromUnicode)))
 hBlockEnd = LLVMAppendBasicBlock(hFunction, StrPtr(StrConv("FunctionEnd", vbFromUnicode)))
 '///
 hVariable = LLVMBuildAlloca(hBuilder, hType(1), StrPtr("i"))
 hType(2) = LLVMConstNull(hType(1))
 LLVMBuildStore hBuilder, hType(2), hVariable
 LLVMBuildCondBr hBuilder, LLVMBuildICmp(hBuilder, LLVMIntNE, hType(2), p1, lp), hBlockDo, hBlockEnd
 '///
 LLVMPositionBuilderAtEnd hBuilder, hBlockDo
 hType(2) = LLVMBuildLoad(hBuilder, hVariable, lp)
 CodegenDefaultDestructor LLVMBuildGEP(hBuilder, p0, hType(2), 1, lp), hBuilder
 hType(2) = LLVMBuildAdd(hBuilder, hType(2), LLVMConstInt(hType(1), 0.0001@, 0), lp)
 LLVMBuildStore hBuilder, hType(2), hVariable
 LLVMBuildCondBr hBuilder, LLVMBuildICmp(hBuilder, LLVMIntULT, hType(2), p1, lp), hBlockDo, hBlockEnd
 '///
 LLVMPositionBuilderAtEnd hBuilder, hBlockEnd
 LLVMBuildRetVoid hBuilder
 '///
 LLVMDisposeBuilder hBuilder
End If
'///
GetDefaultVectorDestructorFunction = hFunction
End Function

Friend Sub CodegenCopyConstructor(ByVal hVariable As Long, ByVal hValue As Long, Optional ByVal hBuilder As Long)
'TODO:
End Sub

Friend Sub CodegenMoveConstructor(ByVal hVariable As Long, ByVal hValue As Long, Optional ByVal hBuilder As Long)
'TODO:
End Sub

Private Function IASTNode_GetProperty(ByVal nProp As enumASTNodeProperty) As Long
'
End Function

Friend Function CodegenDefaultValue() As Long
Dim i(7) As Long
'///
If m_hDefaultValue = 0 Then
 Select Case m_nType
 Case vbDecimal
  i(0) = LLVMConstInt(LLVMInt16Type, 0.0014@, 1)
  i(1) = LLVMConstInt(LLVMInt8Type, 0@, 1)
  i(2) = i(1)
  i(3) = LLVMConstInt(LLVMInt32Type, 0@, 1)
  i(4) = LLVMConstInt(LLVMInt64Type, 0@, 1)
  m_hDefaultValue = LLVMConstStruct(i(0), 5, 0)
 Case Else 'TODO: other
  m_hDefaultValue = LLVMConstNull(m_hType)
 End Select
End If
CodegenDefaultValue = m_hDefaultValue
End Function

Friend Function CodegenOneValue() As Long
Dim i(7) As Long
'///
If m_hOneValue = 0 Then
 Select Case m_nFlags And &HF&
 Case 1, 2 'integer
  m_hOneValue = LLVMConstInt(m_hType, 0.0001@, 1)
 Case 3 'float point
  m_hOneValue = LLVMConstReal(m_hType, 1#)
 Case Else
  Select Case m_nType
  Case vbCurrency
   m_hOneValue = LLVMConstInt(m_hType, 1@, 1)
  Case vbDecimal
   i(0) = LLVMConstInt(LLVMInt16Type, 0.0014@, 1)
   i(1) = LLVMConstInt(LLVMInt8Type, 0@, 1)
   i(2) = i(1)
   i(3) = LLVMConstInt(LLVMInt32Type, 0@, 1)
   i(4) = LLVMConstInt(LLVMInt64Type, 0.0001@, 1)
   m_hOneValue = LLVMConstStruct(i(0), 5, 0)
  Case Else
   'TODO: other
  End Select
 End Select
End If
CodegenOneValue = m_hOneValue
End Function
