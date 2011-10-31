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
'TODO:currency, etc.
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

'TODO:currency, string, etc.
Friend Function CodegenTypeConversion(ByVal objContext As clsVerifyContext, ByVal hValue As Long, ByVal objSrcType As clsTypeNode, ByVal nSrcFlags As Long, ByVal objDestType As clsTypeNode, ByVal nDestFlags As Long, ByVal bIsConstant As Boolean) As Long
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
Case vbBoolean
 hValue = objSrcType.CodegenConvertToI1(objContext, hValue, lp, bIsConstant)
 If bIsConstant Then
  CodegenTypeConversion = LLVMConstSExt(hValue, objDestType.Handle)
 Else
  CodegenTypeConversion = LLVMBuildSExt(g_hBuilder, hValue, objDestType.Handle, lp)
 End If
Case Else
 Select Case objDestType.Flags And &HF&
 Case 1, 2 'signed and unsigned integer
  Select Case objSrcType.Flags And &HF&
  Case 1, 2 'signed and unsigned integer
   If objDestType.Size <= objSrcType.Size Then
    If bIsConstant Then
     CodegenTypeConversion = LLVMConstIntCast(hValue, objDestType.Handle, (objDestType.Flags And &HF&) = 1)
    Else
     CodegenTypeConversion = LLVMBuildIntCast(g_hBuilder, hValue, objDestType.Handle, lp)
    End If
   ElseIf (objSrcType.Flags And &HF&) = 1 Then
    If bIsConstant Then
     CodegenTypeConversion = LLVMConstSExt(hValue, objDestType.Handle)
    Else
     CodegenTypeConversion = LLVMBuildSExt(g_hBuilder, hValue, objDestType.Handle, lp)
    End If
   Else
    If bIsConstant Then
     CodegenTypeConversion = LLVMConstZExt(hValue, objDestType.Handle)
    Else
     CodegenTypeConversion = LLVMBuildZExt(g_hBuilder, hValue, objDestType.Handle, lp)
    End If
   End If
  Case 3 'flost number
   If (objDestType.Flags And &HF&) = 1 Then
    If bIsConstant Then
     CodegenTypeConversion = LLVMConstFPToSI(hValue, objDestType.Handle)
    Else
     CodegenTypeConversion = LLVMBuildFPToSI(g_hBuilder, hValue, objDestType.Handle, lp)
    End If
   Else
    If bIsConstant Then
     CodegenTypeConversion = LLVMConstFPToUI(hValue, objDestType.Handle)
    Else
     CodegenTypeConversion = LLVMBuildFPToUI(g_hBuilder, hValue, objDestType.Handle, lp)
    End If
   End If
  End Select
 Case 3 'float number
  Select Case objSrcType.Flags And &HF&
  Case 1 'signed integer
   If bIsConstant Then
    CodegenTypeConversion = LLVMConstSIToFP(hValue, objDestType.Handle)
   Else
    CodegenTypeConversion = LLVMBuildSIToFP(g_hBuilder, hValue, objDestType.Handle, lp)
   End If
  Case 2 'unsigned integer
   If bIsConstant Then
    CodegenTypeConversion = LLVMConstUIToFP(hValue, objDestType.Handle)
   Else
    CodegenTypeConversion = LLVMBuildUIToFP(g_hBuilder, hValue, objDestType.Handle, lp)
   End If
  Case 3 'flost number
   If bIsConstant Then
    CodegenTypeConversion = LLVMConstFPCast(hValue, objDestType.Handle)
   Else
    CodegenTypeConversion = LLVMBuildFPCast(g_hBuilder, hValue, objDestType.Handle, lp)
   End If
  End Select
 End Select
End Select
'///
End Function

'TODO:currency, string, etc.
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
Select Case nOperator
Case keyword_not
 Select Case nType1
 Case Else
  Select Case objType1.Flags And &HF&
  Case 1, 2 'integers
   Set CheckUnaryOperator = objType1
  Case 3 'float number
   Set CheckUnaryOperator = g_objIntrinsicDataTypes(vbLong)
  End Select
 End Select
Case token_plus
 Set CheckUnaryOperator = objType1
Case token_minus
 Select Case nType1
 Case vbByte
  Set CheckUnaryOperator = g_objIntrinsicDataTypes(vbInteger)
 Case vbUnsignedInteger, vbBoolean
  Set CheckUnaryOperator = g_objIntrinsicDataTypes(vbLong)
 Case vbUnsignedLong
  Set CheckUnaryOperator = g_objIntrinsicDataTypes(vbLong)
  'Set CheckUnaryOperator = g_objIntrinsicDataTypes(vbLongLong)
 Case vbUnsignedLongLong
  Set CheckUnaryOperator = g_objIntrinsicDataTypes(vbLongLong)
 Case vbUIntPtr_t
  If g_nWordSize > 4 Then
   Set CheckUnaryOperator = g_objIntrinsicDataTypes(vbLongLong)
  Else
   Set CheckUnaryOperator = g_objIntrinsicDataTypes(vbLong)
   'Set CheckUnaryOperator = g_objIntrinsicDataTypes(vbLongLong)
  End If
 Case vbCurrency, vbDecimal
  Set CheckUnaryOperator = objType1
 Case Else
  Select Case objType1.Flags And &HF&
  Case 1, 3 'signed integer and float numbers
   Set CheckUnaryOperator = objType1
  End Select
 End Select
End Select
End Function

'TODO:currency, string, etc.
'TODO:overflow, etc.
Friend Function CodegenUnaryOperator(ByVal objContext As clsVerifyContext, ByVal hValue1 As Long, ByVal objType1 As clsTypeNode, ByVal nOperator As enumTokenType, ByVal objDestType As clsTypeNode, ByVal bIsConstant As Boolean) As Long
Dim s As String, lp As Long
'///
Select Case nOperator
Case token_plus
 CodegenUnaryOperator = hValue1
Case token_minus
 s = StrConv("NegTemp", vbFromUnicode)
 lp = StrPtr(s)
 '///
 hValue1 = CodegenTypeConversion(objContext, hValue1, objType1, 0, objDestType, 0, bIsConstant)
 Select Case objDestType.DataType
 Case vbDecimal
  If bIsConstant Then
   hValue1 = LLVMConstInsertValue(hValue1, LLVMConstXor(LLVMConstExtractValue(hValue1, 2&, 1), LLVMConstInt(LLVMInt8Type, 0.0128@, 0)), 2&, 1)
   CodegenUnaryOperator = hValue1
  Else
   hValue1 = LLVMBuildInsertValue(g_hBuilder, hValue1, _
   LLVMBuildXor(g_hBuilder, LLVMBuildExtractValue(g_hBuilder, hValue1, 2&, lp), LLVMConstInt(LLVMInt8Type, 0.0128@, 0), lp), 2&, lp)
   CodegenUnaryOperator = hValue1
  End If
 Case Else
  Select Case objDestType.Flags And &HF&
  Case 3
   If bIsConstant Then
    CodegenUnaryOperator = LLVMConstFNeg(hValue1)
   Else
    CodegenUnaryOperator = LLVMBuildFNeg(g_hBuilder, hValue1, lp)
   End If
  Case Else
   If bIsConstant Then
    CodegenUnaryOperator = LLVMConstNeg(hValue1)
   Else
    CodegenUnaryOperator = LLVMBuildNeg(g_hBuilder, hValue1, lp)
   End If
  End Select
 End Select
Case keyword_not
 hValue1 = CodegenTypeConversion(objContext, hValue1, objType1, 0, objDestType, 0, bIsConstant)
 If bIsConstant Then
  CodegenUnaryOperator = LLVMConstNot(hValue1)
 Else
  CodegenUnaryOperator = LLVMBuildNot(g_hBuilder, hValue1, StrPtr(StrConv("NotTemp", vbFromUnicode)))
 End If
End Select
End Function

'TODO:currency, string, etc.
Friend Function CheckBinaryOperator(ByVal objType1 As clsTypeNode, ByVal objType2 As clsTypeNode, ByVal nOperator As enumTokenType) As clsTypeNode
Dim nType1 As VbVarType
Dim nType2 As VbVarType
Dim nFlags1 As Long
Dim nFlags2 As Long
Dim nMaxSize As Long
'///
nType1 = objType1.DataType
nType2 = objType2.DataType
nFlags1 = objType1.Flags
nFlags2 = objType2.Flags
'///
Select Case nOperator
Case keyword_xor, keyword_or, keyword_and, keyword_eqv, keyword_imp
 If nType1 = vbBoolean And nType2 = vbBoolean Then
  Set CheckBinaryOperator = objType1
 ElseIf (nFlags1 And &HF&) = 3 Or (nFlags2 And &HF&) = 3 Then
  If ((nFlags1 And &HF&) = 1 Or (nFlags1 And &HF&) = 2) And objType1.Size >= 4 Then
   Set CheckBinaryOperator = objType1
  ElseIf ((nFlags2 And &HF&) = 1 Or (nFlags2 And &HF&) = 2) And objType2.Size >= 4 Then
   Set CheckBinaryOperator = objType2
  Else
   Set CheckBinaryOperator = g_objIntrinsicDataTypes(vbLong)
  End If
 ElseIf ((nFlags1 And &HF&) = 1 Or (nFlags1 And &HF&) = 2) And ((nFlags2 And &HF&) = 1 Or (nFlags2 And &HF&) = 2) Then
  If objType1.Size >= objType2.Size Then
   If nType1 = vbBoolean Then
    Set CheckBinaryOperator = g_objIntrinsicDataTypes(vbInteger)
   Else
    Set CheckBinaryOperator = objType1
   End If
  Else
   If nType2 = vbBoolean Then
    Set CheckBinaryOperator = g_objIntrinsicDataTypes(vbInteger)
   Else
    Set CheckBinaryOperator = objType2
   End If
  End If
 End If
Case token_gt, token_lt, token_ge, token_le, token_equal, token_ne
 Set CheckBinaryOperator = g_objIntrinsicDataTypes(vbBoolean)
Case keyword_is
 'TODO:
Case token_and
 'TODO:
 Set CheckBinaryOperator = g_objIntrinsicDataTypes(vbString)
Case token_plus, token_minus, token_asterisk
 If (nFlags1 And &HF&) = 3 Or (nFlags2 And &HF&) = 3 Then
  If ((nFlags1 And &HF&) = 1 Or (nFlags1 And &HF&) = 2) And objType1.Size <= 2 Then
   Set CheckBinaryOperator = objType2
  ElseIf ((nFlags2 And &HF&) = 1 Or (nFlags2 And &HF&) = 2) And objType2.Size <= 2 Then
   Set CheckBinaryOperator = objType1
  ElseIf (nFlags1 And &HF&) = 3 And (nFlags2 And &HF&) = 3 Then
   If objType1.Size >= objType2.Size Then
    Set CheckBinaryOperator = objType1
   Else
    Set CheckBinaryOperator = objType2
   End If
  Else
   Set CheckBinaryOperator = g_objIntrinsicDataTypes(vbDouble)
  End If
 ElseIf ((nFlags1 And &HF&) = 1 Or (nFlags1 And &HF&) = 2) And ((nFlags2 And &HF&) = 1 Or (nFlags2 And &HF&) = 2) Then
  nMaxSize = objType1.Size
  If nMaxSize < objType2.Size Then nMaxSize = objType2.Size
  'XXX TODO:signed constant --> unsigned constant
  If (nFlags1 And &HF&) = 1 Or (nFlags2 And &HF&) = 1 Or nOperator = token_minus Then
   Select Case nMaxSize
   Case 1
    Set CheckBinaryOperator = g_objIntrinsicDataTypes(vbSignedByte)
   Case 2
    Set CheckBinaryOperator = g_objIntrinsicDataTypes(vbInteger)
   Case 4
    Set CheckBinaryOperator = g_objIntrinsicDataTypes(vbLong)
   Case 8
    Set CheckBinaryOperator = g_objIntrinsicDataTypes(vbLongLong)
   End Select
  Else
   Select Case nMaxSize
   Case 1
    Set CheckBinaryOperator = g_objIntrinsicDataTypes(vbByte)
   Case 2
    Set CheckBinaryOperator = g_objIntrinsicDataTypes(vbUnsignedInteger)
   Case 4
    Set CheckBinaryOperator = g_objIntrinsicDataTypes(vbUnsignedLong)
   Case 8
    Set CheckBinaryOperator = g_objIntrinsicDataTypes(vbUnsignedLongLong)
   End Select
  End If
 End If
Case token_backslash, keyword_mod
 If (nFlags1 And &HF&) = 3 Or (nFlags2 And &HF&) = 3 Then
  If ((nFlags1 And &HF&) = 1 Or (nFlags1 And &HF&) = 2) Then
   If objType1.Size <= 4 Then
    Set CheckBinaryOperator = g_objIntrinsicDataTypes(vbLong)
   Else
    Set CheckBinaryOperator = g_objIntrinsicDataTypes(vbLongLong)
   End If
  ElseIf ((nFlags2 And &HF&) = 1 Or (nFlags2 And &HF&) = 2) Then
   If objType2.Size <= 4 Then
    Set CheckBinaryOperator = g_objIntrinsicDataTypes(vbLong)
   Else
    Set CheckBinaryOperator = g_objIntrinsicDataTypes(vbLongLong)
   End If
  Else
   Set CheckBinaryOperator = g_objIntrinsicDataTypes(vbLong)
  End If
 ElseIf ((nFlags1 And &HF&) = 1 Or (nFlags1 And &HF&) = 2) And ((nFlags2 And &HF&) = 1 Or (nFlags2 And &HF&) = 2) Then
  nMaxSize = objType1.Size
  If nMaxSize < objType2.Size Then nMaxSize = objType2.Size
  'XXX TODO:signed constant --> unsigned constant
  'If (nFlags1 And &HF&) = 1 Or (nFlags2 And &HF&) = 1 Then
  If (nFlags1 And &HF&) = 1 Then 'dirty workaround
   Select Case nMaxSize
   Case 1
    Set CheckBinaryOperator = g_objIntrinsicDataTypes(vbSignedByte)
   Case 2
    Set CheckBinaryOperator = g_objIntrinsicDataTypes(vbInteger)
   Case 4
    Set CheckBinaryOperator = g_objIntrinsicDataTypes(vbLong)
   Case 8
    Set CheckBinaryOperator = g_objIntrinsicDataTypes(vbLongLong)
   End Select
  Else
   Select Case nMaxSize
   Case 1
    Set CheckBinaryOperator = g_objIntrinsicDataTypes(vbByte)
   Case 2
    Set CheckBinaryOperator = g_objIntrinsicDataTypes(vbUnsignedInteger)
   Case 4
    Set CheckBinaryOperator = g_objIntrinsicDataTypes(vbUnsignedLong)
   Case 8
    Set CheckBinaryOperator = g_objIntrinsicDataTypes(vbUnsignedLongLong)
   End Select
  End If
 End If
Case token_slash, token_power
 Set CheckBinaryOperator = g_objIntrinsicDataTypes(vbDouble)
End Select
End Function

'TODO:currency, string, etc.
'TODO:overflow, etc. (using LLVM intrinsic functions)
Friend Function CodegenBinaryOperator(ByVal objContext As clsVerifyContext, ByVal hValue1 As Long, ByVal objType1 As clsTypeNode, ByVal hValue2 As Long, ByVal objType2 As clsTypeNode, ByVal nOperator As enumTokenType, ByVal objDestType As clsTypeNode, ByVal bIsConstant As Boolean) As Long
Dim objType As clsTypeNode
Dim i As Long
Dim n(7) As Long
Dim f1 As Double, f2 As Double, f3 As Double
Dim n1 As Currency, n2 As Currency
Dim nFlags1 As Long
Dim nFlags2 As Long
'///
nFlags1 = objType1.Flags
nFlags2 = objType2.Flags
'///
Select Case nOperator
Case token_gt, token_lt, token_ge, token_le, token_equal, token_ne
 Set objType = CheckBinaryOperator(objType1, objType2, token_plus) '??? TODO:
 hValue1 = CodegenTypeConversion(objContext, hValue1, objType1, 0, objType, 0, bIsConstant)
 hValue2 = CodegenTypeConversion(objContext, hValue2, objType2, 0, objType, 0, bIsConstant)
 '///
 Select Case objType.Flags And &HF&
 Case 3 'float number
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
 Case 1, 2 'signed and unsigned integer
  If (nFlags1 And &HF&) = 2 And (nFlags1 And &HF&) = 2 Then
   Select Case nOperator
   Case token_gt
    i = LLVMIntUGT
   Case token_lt
    i = LLVMIntULT
   Case token_ge
    i = LLVMIntUGE
   Case token_le
    i = LLVMIntULE
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
  ElseIf ((nFlags1 And &HF&) = 1 And (objType1.Size > objType2.Size Or (nFlags2 And &HF&) = 1)) Or _
  ((nFlags2 And &HF&) = 1 And objType1.Size < objType2.Size) Then
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
  Else 'XXX: signed/unsigned mismatch
   '///
   PrintWarning "signed/unsigned mismatch: '" + objType1.Name + "' and '" + objType2.Name + "'", -1, -1
   '///
   If (nFlags1 And &HF&) = 2 Then
    Select Case nOperator
    Case token_gt
     i = LLVMIntUGT
    Case token_lt
     i = LLVMIntULT
    Case token_ge
     i = LLVMIntUGE
    Case token_le
     i = LLVMIntULE
    Case token_equal
     i = LLVMIntEQ
    Case token_ne
     i = LLVMIntNE
    End Select
   Else
    i = hValue1
    hValue1 = hValue2
    hValue2 = i
    Select Case nOperator
    Case token_gt
     i = LLVMIntULT
    Case token_lt
     i = LLVMIntUGT
    Case token_ge
     i = LLVMIntULE
    Case token_le
     i = LLVMIntUGE
    Case token_equal
     i = LLVMIntEQ
    Case token_ne
     i = LLVMIntNE
    End Select
   End If
   '///now assume op1 is unsigned, op2 is signed
   If bIsConstant Then
    n(0) = LLVMConstICmp(i, hValue1, hValue2)
    Select Case i
    Case LLVMIntULT, LLVMIntULE, LLVMIntEQ
     n(1) = LLVMConstICmp(LLVMIntSGE, hValue2, LLVMConstInt(objType.Handle, 0@, 1))
     hValue1 = LLVMConstAnd(n(0), n(1))
    Case LLVMIntUGT, LLVMIntUGE, LLVMIntNE
     n(1) = LLVMConstICmp(LLVMIntSLT, hValue2, LLVMConstInt(objType.Handle, 0@, 1))
     hValue1 = LLVMConstOr(n(0), n(1))
    End Select
   Else
    n(0) = LLVMBuildICmp(g_hBuilder, i, hValue1, hValue2, StrPtr(StrConv("CmpTemp", vbFromUnicode)))
    Select Case i
    Case LLVMIntULT, LLVMIntULE, LLVMIntEQ
     n(1) = LLVMBuildICmp(g_hBuilder, LLVMIntSGE, hValue2, LLVMConstInt(objType.Handle, 0@, 1), StrPtr(StrConv("CmpTemp", vbFromUnicode)))
     hValue1 = LLVMBuildAnd(g_hBuilder, n(0), n(1), StrPtr(StrConv("CmpTemp", vbFromUnicode)))
    Case LLVMIntUGT, LLVMIntUGE, LLVMIntNE
     n(1) = LLVMBuildICmp(g_hBuilder, LLVMIntSLT, hValue2, LLVMConstInt(objType.Handle, 0@, 1), StrPtr(StrConv("CmpTemp", vbFromUnicode)))
     hValue1 = LLVMBuildOr(g_hBuilder, n(0), n(1), StrPtr(StrConv("CmpTemp", vbFromUnicode)))
    End Select
   End If
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
 hValue1 = CodegenTypeConversion(objContext, hValue1, objType1, 0, objDestType, 0, bIsConstant)
 hValue2 = CodegenTypeConversion(objContext, hValue2, objType2, 0, objDestType, 0, bIsConstant)
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
   LLVMBuildXor(g_hBuilder, hValue1, hValue2, StrPtr(StrConv("EqvTemp", vbFromUnicode))), _
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
   LLVMBuildNot(g_hBuilder, hValue1, StrPtr(StrConv("ImpTemp", vbFromUnicode))) _
   , hValue2, StrPtr(StrConv("ImpTemp", vbFromUnicode)))
  End If
 Case keyword_and
  If bIsConstant Then
   CodegenBinaryOperator = LLVMConstAnd(hValue1, hValue2)
  Else
   CodegenBinaryOperator = LLVMBuildAnd(g_hBuilder, hValue1, hValue2, StrPtr(StrConv("AndTemp", vbFromUnicode)))
  End If
 Case token_plus
  Select Case objDestType.Flags And &HF&
  Case 3
   If bIsConstant Then
    CodegenBinaryOperator = LLVMConstFAdd(hValue1, hValue2)
   Else
    CodegenBinaryOperator = LLVMBuildFAdd(g_hBuilder, hValue1, hValue2, StrPtr(StrConv("AddTemp", vbFromUnicode)))
   End If
  Case 1, 2
   If bIsConstant Then
    CodegenBinaryOperator = LLVMConstAdd(hValue1, hValue2)
   Else
    CodegenBinaryOperator = LLVMBuildAdd(g_hBuilder, hValue1, hValue2, StrPtr(StrConv("AddTemp", vbFromUnicode)))
   End If
  End Select
 Case token_minus
  Select Case objDestType.Flags And &HF&
  Case 3
   If bIsConstant Then
    CodegenBinaryOperator = LLVMConstFSub(hValue1, hValue2)
   Else
    CodegenBinaryOperator = LLVMBuildFSub(g_hBuilder, hValue1, hValue2, StrPtr(StrConv("SubTemp", vbFromUnicode)))
   End If
  Case 1, 2
   If bIsConstant Then
    CodegenBinaryOperator = LLVMConstSub(hValue1, hValue2)
   Else
    CodegenBinaryOperator = LLVMBuildSub(g_hBuilder, hValue1, hValue2, StrPtr(StrConv("SubTemp", vbFromUnicode)))
   End If
  End Select
 Case token_asterisk
  Select Case objDestType.Flags And &HF&
  Case 3
   If bIsConstant Then
    CodegenBinaryOperator = LLVMConstFMul(hValue1, hValue2)
   Else
    CodegenBinaryOperator = LLVMBuildFMul(g_hBuilder, hValue1, hValue2, StrPtr(StrConv("MulTemp", vbFromUnicode)))
   End If
  Case 1, 2 'signed and unsigned (??)
   If bIsConstant Then
    CodegenBinaryOperator = LLVMConstMul(hValue1, hValue2)
   Else
    CodegenBinaryOperator = LLVMBuildMul(g_hBuilder, hValue1, hValue2, StrPtr(StrConv("MulTemp", vbFromUnicode)))
   End If
  End Select
 Case token_backslash
  Select Case objDestType.Flags And &HF&
  Case 1 'signed
   If bIsConstant Then
    CodegenBinaryOperator = LLVMConstSDiv(hValue1, hValue2)
   Else
    CodegenBinaryOperator = LLVMBuildSDiv(g_hBuilder, hValue1, hValue2, StrPtr(StrConv("IDivTemp", vbFromUnicode)))
   End If
  Case 2 'unsigned
   If bIsConstant Then
    CodegenBinaryOperator = LLVMConstUDiv(hValue1, hValue2)
   Else
    CodegenBinaryOperator = LLVMBuildUDiv(g_hBuilder, hValue1, hValue2, StrPtr(StrConv("IDivTemp", vbFromUnicode)))
   End If
  End Select
 Case keyword_mod
  Select Case objDestType.Flags And &HF&
  Case 1 'signed
   If bIsConstant Then
    CodegenBinaryOperator = LLVMConstSRem(hValue1, hValue2)
   Else
    CodegenBinaryOperator = LLVMBuildSRem(g_hBuilder, hValue1, hValue2, StrPtr(StrConv("ModTemp", vbFromUnicode)))
   End If
  Case 2 'unsigned
   If bIsConstant Then
    CodegenBinaryOperator = LLVMConstURem(hValue1, hValue2)
   Else
    CodegenBinaryOperator = LLVMBuildURem(g_hBuilder, hValue1, hValue2, StrPtr(StrConv("ModTemp", vbFromUnicode)))
   End If
  End Select
 Case token_slash
  If bIsConstant Then
   CodegenBinaryOperator = LLVMConstFDiv(hValue1, hValue2)
  Else
   CodegenBinaryOperator = LLVMBuildFDiv(g_hBuilder, hValue1, hValue2, StrPtr(StrConv("DivTemp", vbFromUnicode)))
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
   CodegenBinaryOperator = LLVMBuildCall(g_hBuilder, RuntimeLibraryGetFunction(internal_pow), n(0), 2, StrPtr(StrConv("PowTemp", vbFromUnicode)))
  End If
 End Select
End Select
End Function

