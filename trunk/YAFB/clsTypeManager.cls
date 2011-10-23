VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTypeManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)

'0=failed
'1=OK
'TODO:boolean, etc.
Friend Function CheckTypeConversion(ByVal objSrcType As clsTypeNode, ByVal nSrcFlags As Long, ByVal objDestType As clsTypeNode, ByVal nDestFlags As Long) As Long
Dim nSrcType As VbVarType
Dim nDestType As VbVarType
'///
nSrcType = objSrcType.DataType
nDestType = objDestType.DataType
'///
'TODO:
'///
CheckTypeConversion = 1
End Function

'TODO:boolean, string, etc.
Friend Function CodegenTypeConversion(ByVal hValue As Long, ByVal objSrcType As clsTypeNode, ByVal nSrcFlags As Long, ByVal objDestType As clsTypeNode, ByVal nDestFlags As Long, ByVal bIsConstant As Boolean) As Long
Dim nSrcType As VbVarType
Dim nDestType As VbVarType
Dim lp As Long, s As String
'///
nSrcType = objSrcType.DataType
nDestType = objDestType.DataType
'///
If nSrcType = nDestType Then
 CodegenTypeConversion = hValue
 Exit Function
End If
'///
s = StrConv("CastTemp", vbFromUnicode)
lp = StrPtr(s)
'///
Select Case nDestType
Case vbByte
 Select Case nSrcType
 Case vbInteger, vbLong, vbBoolean
  If bIsConstant Then
   CodegenTypeConversion = LLVMConstIntCast(hValue, objDestType.Handle, 1)
  Else
   CodegenTypeConversion = LLVMBuildIntCast(g_hBuilder, hValue, objDestType.Handle, lp)
  End If
 Case vbSingle, vbDouble
  If bIsConstant Then
   CodegenTypeConversion = LLVMConstFPToUI(hValue, objDestType.Handle)
  Else
   CodegenTypeConversion = LLVMBuildFPToUI(g_hBuilder, hValue, objDestType.Handle, lp)
  End If
 End Select
Case vbInteger, vbLong
 Select Case nSrcType
 Case vbByte
  If bIsConstant Then
   CodegenTypeConversion = LLVMConstZExt(hValue, objDestType.Handle)
  Else
   CodegenTypeConversion = LLVMBuildZExt(g_hBuilder, hValue, objDestType.Handle, lp)
  End If
 Case vbInteger, vbBoolean
  If bIsConstant Then
   CodegenTypeConversion = LLVMConstSExt(hValue, objDestType.Handle)
  Else
   CodegenTypeConversion = LLVMBuildSExt(g_hBuilder, hValue, objDestType.Handle, lp)
  End If
 Case vbLong
  If bIsConstant Then
   CodegenTypeConversion = LLVMConstIntCast(hValue, objDestType.Handle, 0)
  Else
   CodegenTypeConversion = LLVMBuildIntCast(g_hBuilder, hValue, objDestType.Handle, lp)
  End If
 Case vbSingle, vbDouble
  If bIsConstant Then
   CodegenTypeConversion = LLVMConstFPToSI(hValue, objDestType.Handle)
  Else
   CodegenTypeConversion = LLVMBuildFPToSI(g_hBuilder, hValue, objDestType.Handle, lp)
  End If
 End Select
Case vbSingle, vbDouble
 Select Case nSrcType
 Case vbByte
  If bIsConstant Then
   CodegenTypeConversion = LLVMConstUIToFP(hValue, objDestType.Handle)
  Else
   CodegenTypeConversion = LLVMBuildUIToFP(g_hBuilder, hValue, objDestType.Handle, lp)
  End If
 Case vbInteger, vbLong, vbBoolean
  If bIsConstant Then
   CodegenTypeConversion = LLVMConstSIToFP(hValue, objDestType.Handle)
  Else
   CodegenTypeConversion = LLVMBuildSIToFP(g_hBuilder, hValue, objDestType.Handle, lp)
  End If
 Case vbSingle, vbDouble
  If bIsConstant Then
   CodegenTypeConversion = LLVMConstFPCast(hValue, objDestType.Handle)
  Else
   CodegenTypeConversion = LLVMBuildFPCast(g_hBuilder, hValue, objDestType.Handle, lp)
  End If
 End Select
Case vbBoolean
 Select Case nSrcType
 Case vbByte, vbInteger, vbLong
  If bIsConstant Then
   hValue = LLVMConstICmp(LLVMIntNE, hValue, LLVMConstInt(objSrcType.Handle, 0@, 1))
   CodegenTypeConversion = LLVMConstSExt(hValue, objDestType.Handle)
  Else
   hValue = LLVMBuildICmp(g_hBuilder, LLVMIntNE, hValue, LLVMConstInt(objSrcType.Handle, 0@, 1), lp)
   CodegenTypeConversion = LLVMBuildSExt(g_hBuilder, hValue, objDestType.Handle, lp)
  End If
 Case vbSingle, vbDouble
  If bIsConstant Then
   hValue = LLVMConstFCmp(LLVMRealUNE, hValue, LLVMConstReal(objSrcType.Handle, 0#))
   CodegenTypeConversion = LLVMConstSExt(hValue, objDestType.Handle)
  Else
   hValue = LLVMBuildFCmp(g_hBuilder, LLVMRealUNE, hValue, LLVMConstReal(objSrcType.Handle, 0#), lp)
   CodegenTypeConversion = LLVMBuildSExt(g_hBuilder, hValue, objDestType.Handle, lp)
  End If
 End Select
End Select
End Function

'TODO:string, etc.
Friend Function CheckUnaryOperator(ByVal objType1 As clsTypeNode, ByVal nOperator As enumTokenType) As clsTypeNode
Dim nType1 As VbVarType
'///
If nOperator = token_plus Then
 Set CheckUnaryOperator = objType1
 Exit Function
End If
'///
nType1 = objType1.DataType
'///
Select Case nType1
Case vbByte, vbInteger
 Set CheckUnaryOperator = g_objIntrinsicDataTypes(vbInteger)
Case vbLong
 Set CheckUnaryOperator = g_objIntrinsicDataTypes(vbLong)
Case vbBoolean
 Select Case nOperator
 Case token_minus
  Set CheckUnaryOperator = g_objIntrinsicDataTypes(vbInteger)
 Case Else
  Set CheckUnaryOperator = objType1
 End Select
Case vbSingle, vbDouble
 Select Case nOperator
 Case keyword_not
  Set CheckUnaryOperator = g_objIntrinsicDataTypes(vbLong)
 Case Else
  Set CheckUnaryOperator = objType1
 End Select
End Select
End Function

Friend Function CodegenUnaryOperator(ByVal hValue1 As Long, ByVal objType1 As clsTypeNode, ByVal nOperator As enumTokenType, ByVal objDestType As clsTypeNode, ByVal bIsConstant As Boolean) As Long
Select Case nOperator
Case token_plus
 CodegenUnaryOperator = hValue1
Case token_minus
 hValue1 = CodegenTypeConversion(hValue1, objType1, 0, objDestType, 0, bIsConstant)
 Select Case objDestType.DataType
 Case vbSingle, vbDouble
  If bIsConstant Then
   CodegenUnaryOperator = LLVMConstFNeg(hValue1)
  Else
   CodegenUnaryOperator = LLVMBuildFNeg(g_hBuilder, hValue1, StrPtr(StrConv("NegTemp", vbFromUnicode)))
  End If
 Case Else
  If bIsConstant Then
   CodegenUnaryOperator = LLVMConstNeg(hValue1)
  Else
   CodegenUnaryOperator = LLVMBuildNeg(g_hBuilder, hValue1, StrPtr(StrConv("NegTemp", vbFromUnicode)))
  End If
 End Select
Case keyword_not
 hValue1 = CodegenTypeConversion(hValue1, objType1, 0, objDestType, 0, bIsConstant)
 If bIsConstant Then
  CodegenUnaryOperator = LLVMConstNot(hValue1)
 Else
  CodegenUnaryOperator = LLVMBuildNot(g_hBuilder, hValue1, StrPtr(StrConv("NotTemp", vbFromUnicode)))
 End If
End Select
End Function

'TODO:string, etc.
Friend Function CheckBinaryOperator(ByVal objType1 As clsTypeNode, ByVal objType2 As clsTypeNode, ByVal nOperator As enumTokenType) As clsTypeNode
Dim nType1 As VbVarType
Dim nType2 As VbVarType
'///
nType1 = objType1.DataType
nType2 = objType2.DataType
'///
Select Case nOperator
Case keyword_xor, keyword_or, keyword_and, keyword_eqv, keyword_imp
 If nType1 = vbBoolean And nType2 = vbBoolean Then
  Set CheckBinaryOperator = objType1
 Else
  Set CheckBinaryOperator = g_objIntrinsicDataTypes(vbLong) 'TODO: Byte, Integer, etc.
 End If
Case token_gt, token_lt, token_ge, token_le, token_equal, token_ne
 Set CheckBinaryOperator = g_objIntrinsicDataTypes(vbBoolean)
Case keyword_is
 'TODO:
Case token_and
 'TODO:
 Set CheckBinaryOperator = g_objIntrinsicDataTypes(vbString)
Case token_plus, token_minus, token_asterisk
 Select Case nType1
 Case vbSingle
  Select Case nType2
  Case vbByte, vbInteger, vbBoolean, vbSingle
   Set CheckBinaryOperator = g_objIntrinsicDataTypes(vbSingle)
  Case vbLong, vbDouble
   Set CheckBinaryOperator = g_objIntrinsicDataTypes(vbDouble)
  End Select
 Case vbDouble
  Set CheckBinaryOperator = g_objIntrinsicDataTypes(vbDouble)
 Case vbByte, vbInteger, vbBoolean
  Select Case nType2
  Case vbByte, vbInteger, vbBoolean
   Set CheckBinaryOperator = g_objIntrinsicDataTypes(vbInteger)
  Case vbLong
   Set CheckBinaryOperator = g_objIntrinsicDataTypes(vbLong)
  Case vbSingle
   Set CheckBinaryOperator = g_objIntrinsicDataTypes(vbSingle)
  Case vbDouble
   Set CheckBinaryOperator = g_objIntrinsicDataTypes(vbDouble)
  End Select
 Case vbLong
  Select Case nType2
  Case vbSingle, vbDouble
   Set CheckBinaryOperator = g_objIntrinsicDataTypes(vbDouble)
  Case vbByte, vbInteger, vbBoolean, vbLong
   Set CheckBinaryOperator = g_objIntrinsicDataTypes(vbLong)
  End Select
 End Select
Case token_backslash, keyword_mod
 Select Case nType1
 Case vbByte, vbInteger, vbBoolean
  Select Case nType2
  Case vbByte, vbInteger, vbBoolean
   Set CheckBinaryOperator = g_objIntrinsicDataTypes(vbInteger)
  Case Else
   Set CheckBinaryOperator = g_objIntrinsicDataTypes(vbLong)
  End Select
 Case Else
  Set CheckBinaryOperator = g_objIntrinsicDataTypes(vbLong)
 End Select
Case token_slash, token_power
 Set CheckBinaryOperator = g_objIntrinsicDataTypes(vbDouble)
End Select
End Function

Friend Function CodegenBinaryOperator(ByVal hValue1 As Long, ByVal objType1 As clsTypeNode, ByVal hValue2 As Long, ByVal objType2 As clsTypeNode, ByVal nOperator As enumTokenType, ByVal objDestType As clsTypeNode, ByVal bIsConstant As Boolean) As Long
Dim objType As clsTypeNode
Dim i As Long
Dim n(7) As Long
Dim f1 As Double, f2 As Double, f3 As Double
Dim n1 As Currency, n2 As Currency
'///
Select Case nOperator
Case token_gt, token_lt, token_ge, token_le, token_equal, token_ne
 Set objType = CheckBinaryOperator(objType1, objType2, token_plus) '??? TODO:
 hValue1 = CodegenTypeConversion(hValue1, objType1, 0, objType, 0, bIsConstant)
 hValue2 = CodegenTypeConversion(hValue2, objType2, 0, objType, 0, bIsConstant)
 '///
 Select Case objType.DataType
 Case vbSingle, vbDouble
  Select Case nOperator
  Case token_gt
   i = LLVMRealOGT
  Case token_lt
   i = LLVMRealOLT
  Case token_ge
   i = LLVMRealOGE
  Case token_le
   i = LLVMRealOLE
  Case token_equal
   i = LLVMRealOEQ
  Case token_ne
   i = LLVMRealUNE
  End Select
  '///
  If bIsConstant Then
   hValue1 = LLVMConstFCmp(i, hValue1, hValue2)
  Else
   hValue1 = LLVMBuildFCmp(g_hBuilder, i, hValue1, hValue2, StrPtr(StrConv("CmpTemp", vbFromUnicode)))
  End If
 Case Else
  Select Case nOperator
  Case token_gt
   i = LLVMIntSGT
  Case token_lt
   i = LLVMIntSLT
  Case token_ge
   i = LLVMIntSGE
  Case token_le
   i = LLVMIntSLE
  Case token_equal
   i = LLVMIntEQ
  Case token_ne
   i = LLVMIntNE
  End Select
  '///
  If bIsConstant Then
   hValue1 = LLVMConstICmp(i, hValue1, hValue2)
  Else
   hValue1 = LLVMBuildICmp(g_hBuilder, i, hValue1, hValue2, StrPtr(StrConv("CmpTemp", vbFromUnicode)))
  End If
 End Select
 '///
 If objDestType Is Nothing Then
  CodegenBinaryOperator = hValue1
 ElseIf bIsConstant Then
  CodegenBinaryOperator = LLVMConstSExt(hValue1, objDestType.Handle)
 Else
  CodegenBinaryOperator = LLVMBuildSExt(g_hBuilder, hValue1, objDestType.Handle, StrPtr(StrConv("CmpTemp", vbFromUnicode)))
 End If
Case keyword_is
 'TODO:
Case Else
 hValue1 = CodegenTypeConversion(hValue1, objType1, 0, objDestType, 0, bIsConstant)
 hValue2 = CodegenTypeConversion(hValue2, objType2, 0, objDestType, 0, bIsConstant)
 '///
 Select Case nOperator
 Case keyword_xor
  If bIsConstant Then
   CodegenBinaryOperator = LLVMConstXor(hValue1, hValue2)
  Else
   CodegenBinaryOperator = LLVMBuildXor(g_hBuilder, hValue1, hValue2, StrPtr(StrConv("XorTemp", vbFromUnicode)))
  End If
 Case keyword_eqv
  If bIsConstant Then
   CodegenBinaryOperator = LLVMConstNot(LLVMConstXor(hValue1, hValue2))
  Else
   CodegenBinaryOperator = LLVMBuildNot(g_hBuilder, _
   LLVMBuildXor(g_hBuilder, hValue1, hValue2, StrPtr(StrConv("XorTemp", vbFromUnicode))), _
   StrPtr(StrConv("EqvTemp", vbFromUnicode)))
  End If
 Case keyword_or
  If bIsConstant Then
   CodegenBinaryOperator = LLVMConstOr(hValue1, hValue2)
  Else
   CodegenBinaryOperator = LLVMBuildOr(g_hBuilder, hValue1, hValue2, StrPtr(StrConv("OrTemp", vbFromUnicode)))
  End If
 Case keyword_imp
  If bIsConstant Then
   CodegenBinaryOperator = LLVMConstOr(LLVMConstNot(hValue1), hValue2)
  Else
   CodegenBinaryOperator = LLVMBuildOr(g_hBuilder, _
   LLVMBuildNot(g_hBuilder, hValue1, StrPtr(StrConv("NotTemp", vbFromUnicode))) _
   , hValue2, StrPtr(StrConv("ImpTemp", vbFromUnicode)))
  End If
 Case keyword_and
  If bIsConstant Then
   CodegenBinaryOperator = LLVMConstAnd(hValue1, hValue2)
  Else
   CodegenBinaryOperator = LLVMBuildAnd(g_hBuilder, hValue1, hValue2, StrPtr(StrConv("AndTemp", vbFromUnicode)))
  End If
 Case token_plus
  Select Case objDestType.DataType
  Case vbSingle, vbDouble
   If bIsConstant Then
    CodegenBinaryOperator = LLVMConstFAdd(hValue1, hValue2)
   Else
    CodegenBinaryOperator = LLVMBuildFAdd(g_hBuilder, hValue1, hValue2, StrPtr(StrConv("AddTemp", vbFromUnicode)))
   End If
  Case Else
   If bIsConstant Then
    CodegenBinaryOperator = LLVMConstAdd(hValue1, hValue2)
   Else
    CodegenBinaryOperator = LLVMBuildAdd(g_hBuilder, hValue1, hValue2, StrPtr(StrConv("AddTemp", vbFromUnicode)))
   End If
  End Select
 Case token_minus
  Select Case objDestType.DataType
  Case vbSingle, vbDouble
   If bIsConstant Then
    CodegenBinaryOperator = LLVMConstFSub(hValue1, hValue2)
   Else
    CodegenBinaryOperator = LLVMBuildFSub(g_hBuilder, hValue1, hValue2, StrPtr(StrConv("SubTemp", vbFromUnicode)))
   End If
  Case Else
   If bIsConstant Then
    CodegenBinaryOperator = LLVMConstSub(hValue1, hValue2)
   Else
    CodegenBinaryOperator = LLVMBuildSub(g_hBuilder, hValue1, hValue2, StrPtr(StrConv("SubTemp", vbFromUnicode)))
   End If
  End Select
 Case token_asterisk
  Select Case objDestType.DataType
  Case vbSingle, vbDouble
   If bIsConstant Then
    CodegenBinaryOperator = LLVMConstFMul(hValue1, hValue2)
   Else
    CodegenBinaryOperator = LLVMBuildFMul(g_hBuilder, hValue1, hValue2, StrPtr(StrConv("MulTemp", vbFromUnicode)))
   End If
  Case Else
   If bIsConstant Then
    CodegenBinaryOperator = LLVMConstMul(hValue1, hValue2)
   Else
    CodegenBinaryOperator = LLVMBuildMul(g_hBuilder, hValue1, hValue2, StrPtr(StrConv("MulTemp", vbFromUnicode)))
   End If
  End Select
 Case token_backslash
  If bIsConstant Then
   CodegenBinaryOperator = LLVMConstSDiv(hValue1, hValue2)
  Else
   CodegenBinaryOperator = LLVMBuildSDiv(g_hBuilder, hValue1, hValue2, StrPtr(StrConv("SubTemp", vbFromUnicode)))
  End If
 Case keyword_mod
  If bIsConstant Then
   CodegenBinaryOperator = LLVMConstSRem(hValue1, hValue2)
  Else
   CodegenBinaryOperator = LLVMBuildSRem(g_hBuilder, hValue1, hValue2, StrPtr(StrConv("SubTemp", vbFromUnicode)))
  End If
 Case token_slash
  If bIsConstant Then
   CodegenBinaryOperator = LLVMConstFDiv(hValue1, hValue2)
  Else
   CodegenBinaryOperator = LLVMBuildFDiv(g_hBuilder, hValue1, hValue2, StrPtr(StrConv("SubTemp", vbFromUnicode)))
  End If
 Case token_power
  '///dirty workaround
  If bIsConstant Then
   hValue1 = LLVMConstBitCast(hValue1, LLVMInt64Type)
   hValue2 = LLVMConstBitCast(hValue2, LLVMInt64Type)
   n1 = LLVMConstIntGetSExtValue(hValue1)
   n2 = LLVMConstIntGetSExtValue(hValue2)
   CopyMemory f1, n1, 8&
   CopyMemory f2, n2, 8&
   On Error Resume Next
   Err.Clear
   f3 = f1 ^ f2
   If Err.Number Then
    PrintError "Error when calculating " + CStr(f1) + "^" + CStr(f2), -1, -1
    Exit Function
   Else
    CodegenBinaryOperator = LLVMConstReal(LLVMDoubleType, f3)
   End If
  Else
   n(0) = hValue1
   n(1) = hValue2
   CodegenBinaryOperator = LLVMBuildCall(g_hBuilder, g_hFunctionPow, n(0), 2, StrPtr(StrConv("PowTemp", vbFromUnicode)))
  End If
 End Select
End Select
End Function

