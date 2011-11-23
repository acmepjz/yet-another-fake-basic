Attribute VB_Name = "mdlPreprocessor"
Option Explicit

Public Type typePreprocessorConstant
 nType As VbVarType 'currently only vbDouble is avaliable
 fValue As Double
End Type

Public Type typeNamedPreprocessorConstant
 sName As String
 tValue As typePreprocessorConstant
End Type

Public Type typePreprocessorConstCollection
 tConst() As typeNamedPreprocessorConstant '1-based
 nCount As Long
 nMax As Long
 obj As New Collection
End Type

Public g_tGlobalPreprocessorConst As typePreprocessorConstCollection
Public g_tLocalPreprocessorConst As typePreprocessorConstCollection

'////////

Private Type typePreprocessorStack
 nType As enumASTNodeType
 'possible values:
 'node_ifstat
 '///
 bEnabled As Boolean '?
 bParentEnabled As Boolean '?
End Type

Private m_tStack() As typePreprocessorStack '1-based
Private m_nStackPointer As Long, m_nStackMax As Long

Public Sub ClearPreprocessorConst()
Dim t As typePreprocessorConstCollection
g_tGlobalPreprocessorConst = t
g_tLocalPreprocessorConst = t
End Sub

Public Function FindPreprocessorConst(ByVal sName As String) As Long
On Error Resume Next
Dim i As Long
Err.Clear
i = g_tLocalPreprocessorConst.obj.Item(sName)
If Err.Number = 0 Then
 FindPreprocessorConst = i
 Exit Function
End If
Err.Clear
i = g_tGlobalPreprocessorConst.obj.Item(sName)
If Err.Number = 0 Then
 FindPreprocessorConst = i Or &H80000000
 Exit Function
End If
End Function

Public Function QueryPreprocessorConst(ByVal sName As String) As Double
On Error Resume Next
Dim i As Long
Err.Clear
i = g_tLocalPreprocessorConst.obj.Item(sName)
If Err.Number = 0 Then
 QueryPreprocessorConst = g_tLocalPreprocessorConst.tConst(i).tValue.fValue
 Exit Function
End If
Err.Clear
i = g_tGlobalPreprocessorConst.obj.Item(sName)
If Err.Number = 0 Then
 QueryPreprocessorConst = g_tGlobalPreprocessorConst.tConst(i).tValue.fValue
 Exit Function
End If
End Function

Public Sub SetPreprocessorConst(ByRef t As typePreprocessorConstCollection, ByVal sName As String, ByVal fValue As Double)
On Error Resume Next
Dim i As Long
Err.Clear
i = t.obj.Item(sName)
If Err.Number <> 0 Then
 i = t.nCount + 1
 t.nCount = i
 If i > t.nMax Then
  t.nMax = t.nMax + 256&
  ReDim Preserve t.tConst(1 To t.nMax)
 End If
 t.obj.Add i, sName
 t.tConst(i).sName = sName
 t.tConst(i).tValue.nType = vbDouble
End If
t.tConst(i).tValue.fValue = fValue
End Sub

'with preprocessor
Public Function GetNextToken(ByVal objFile As ISource, ByRef t As typeToken) As Boolean
Dim bSkip As Boolean
Dim i As Long
Dim s As String, f As Double
'///
Do
 If Not GetNextToken_Internal(objFile, t) Then Exit Function
 '///
 Select Case t.nType
 Case preprocessor_const
  If bSkip Then
   If Not PreprocessorParseSkip(objFile, t) Then Exit Function
  Else
   If Not GetNextToken_Internal(objFile, t) Then Exit Function
   If t.nType <> token_id Then
    PrintError "Identifier expected", t.nLine, t.nColumn
    Exit Function
   End If
   s = t.sValue
   i = FindPreprocessorConst(s)
   If i > 0 Then
    PrintError "Ambiguous name detected: '" + s + "'", t.nLine, t.nColumn
    Exit Function
   ElseIf i < 0 Then
    PrintWarning "Preprocessor const '" + s + "' is already defined in global const table", t.nLine, t.nColumn
   End If
   '///
   If Not GetNextToken_Internal(objFile, t) Then Exit Function
   If t.nType <> token_equal Then
    PrintError "'=' expected", t.nLine, t.nColumn
    Exit Function
   End If
   '///
   If Not GetNextToken_Internal(objFile, t) Then Exit Function
   If Not PreprocessorParseExpression(objFile, t, f) Then Exit Function
   Select Case t.nType
   Case token_eof, token_crlf, token_colon
   Case Else
    PrintError "end of line expected", t.nLine, t.nColumn
    Exit Function
   End Select
   '///
   SetPreprocessorConst g_tLocalPreprocessorConst, s, f
  End If
 Case preprocessor_if
  'TODO:
 Case preprocessor_elseif
  'TODO:
 Case preprocessor_else
  'TODO:
 Case preprocessor_end
  If Not GetNextToken_Internal(objFile, t) Then Exit Function
  If t.nType <> keyword_if Then
   PrintError "'#End If' expected", t.nLine, t.nColumn
   Exit Function
  End If
  '///pop stack
  'TODO:
  '///
  If Not GetNextToken_Internal(objFile, t) Then Exit Function
  Select Case t.nType
  Case token_eof, token_crlf, token_colon
  Case Else
   PrintError "end of line expected", t.nLine, t.nColumn
   Exit Function
  End Select
 End Select
Loop While bSkip
'///
GetNextToken = True
End Function

Public Sub InitializePreprocessor()
Dim t As typePreprocessorConstCollection
g_tLocalPreprocessorConst = t
ReDim m_tStack(1 To 32)
m_nStackPointer = 0
m_nStackMax = 32
End Sub

Public Function FinalizePreprocessor() As Boolean
'TODO:
FinalizePreprocessor = True
End Function

Private Function PreprocessorParseSkip(ByVal objFile As ISource, ByRef t As typeToken) As Boolean
Do
 If t.nType = 0 Then Exit Do
 '///
 Select Case t.nType
 Case token_eof, token_crlf, token_colon
  PreprocessorParseSkip = True
  Exit Function
 Case Else
  If Not GetNextToken_Internal(objFile, t) Then Exit Function
 End Select
Loop
End Function

Private Function PreprocessorParseExpression(ByVal objFile As ISource, ByRef t As typeToken, ByRef LHS As Double, Optional ByVal ExprPrec As Long, Optional ByVal LHSEnabled As Boolean) As Boolean
Dim TokPrec As Long, NextPrec As Long
Dim i As Long
Dim RHS As Double
'//
If Not LHSEnabled Then
 If Not PreprocessorParseUnaryOp(objFile, t, LHS) Then Exit Function
End If
'// If this is a binop, find its precedence.
Do
 TokPrec = GetBinaryTokPrecedence(t.nType)
 '// If this is a binop that binds at least as tightly as the current binop,
 '// consume it, otherwise we are done.
 If TokPrec < ExprPrec Then
  PreprocessorParseExpression = True
  Exit Function
 End If
 '// Okay, we know this is a binop.
 i = t.nType
 If Not GetNextToken_Internal(objFile, t) Then Exit Function '// eat binop
 '// Parse the unary expression after the binary operator.
 If Not PreprocessorParseUnaryOp(objFile, t, RHS) Then Exit Function
 '// If BinOp binds less tightly with RHS than the operator after RHS, let
 '// the pending operator take RHS as its LHS.
 NextPrec = GetBinaryTokPrecedence(t.nType)
 If TokPrec < NextPrec Then
  If Not PreprocessorParseExpression(objFile, t, RHS, TokPrec + 1, True) Then Exit Function
 End If
 '// Merge LHS/RHS.
 Select Case i
 Case keyword_imp
  LHS = CLng(LHS) Imp CLng(RHS)
 Case keyword_xor
  LHS = CLng(LHS) Xor CLng(RHS)
 Case keyword_eqv
  LHS = CLng(LHS) Eqv CLng(RHS)
 Case keyword_or
  LHS = CLng(LHS) Or CLng(RHS)
 Case keyword_and
  LHS = CLng(LHS) And CLng(RHS)
 Case token_gt
  LHS = LHS > RHS
 Case token_lt
  LHS = LHS < RHS
 Case token_ge
  LHS = LHS >= RHS
 Case token_le
  LHS = LHS <= RHS
 Case token_equal
  LHS = LHS = RHS
 Case token_ne
  LHS = LHS <> RHS
 Case keyword_is
  'TODO:
  PrintError "Object required", t.nLine, t.nColumn
  Exit Function
 Case token_shl, token_shr, token_rol, token_ror
  'TODO:
  PrintError "Currently unsupported '" + GetOperatorName(i) + "'", t.nLine, t.nColumn
  Exit Function
 Case token_and
  'TODO:
  PrintError "Type mismatch", t.nLine, t.nColumn
  Exit Function
 Case token_plus
  LHS = LHS + RHS
 Case token_minus
  LHS = LHS - RHS
 Case keyword_mod
  LHS = CLng(LHS) Mod CLng(RHS)
 Case token_backslash
  LHS = CLng(LHS) \ CLng(RHS)
 Case token_asterisk
  LHS = LHS * RHS
 Case token_slash
  LHS = LHS / RHS
 Case token_power
  LHS = LHS ^ RHS
 Case Else
  Debug.Assert False
  PrintError "Unknown operator '" + GetOperatorName(i) + "'", t.nLine, t.nColumn
  Exit Function
 End Select
Loop  '// loop around to the top of the while loop.
End Function

Private Function PreprocessorParseUnaryOp(ByVal objFile As ISource, ByRef t As typeToken, ByRef LHS As Double) As Boolean
Dim i As Long, TokPrec As Long
TokPrec = GetUnaryTokPrecedence(t.nType)
If TokPrec > 0 Then
 '// If this is a unary operator, read it.
 i = t.nType
 If Not GetNextToken_Internal(objFile, t) Then Exit Function
 If Not PreprocessorParseExpression(objFile, t, LHS, TokPrec) Then Exit Function
 '// evaluate
 Select Case i
 Case keyword_not
  LHS = Not CLng(LHS)
 Case token_plus
 Case token_minus
  LHS = -LHS
 Case Else
  Debug.Assert False
  PrintError "Unknown operator '" + GetOperatorName(i) + "'", t.nLine, t.nColumn
  Exit Function
 End Select
 '// over
 PreprocessorParseUnaryOp = True
Else
 '// If the current token is not an operator, it must be a primary expr.
 PreprocessorParseUnaryOp = PreprocessorParseExpressionTerm(objFile, t, LHS)
End If
End Function

Private Function PreprocessorParseExpressionTerm(ByVal objFile As ISource, ByRef t As typeToken, ByRef LHS As Double) As Boolean
'///
Select Case t.nType
Case token_lbracket '"("<expression>")"
 If Not GetNextToken_Internal(objFile, t) Then Exit Function
 '///
 If Not PreprocessorParseExpression(objFile, t, LHS) Then Exit Function
 '///
 If t.nType <> token_rbracket Then
  PrintError "')' expected", t.nLine, t.nColumn
  Exit Function
 End If
 '///
 If Not GetNextToken_Internal(objFile, t) Then Exit Function
 PreprocessorParseExpressionTerm = True
Case token_id
 LHS = QueryPreprocessorConst(t.sValue)
 If Not GetNextToken_Internal(objFile, t) Then Exit Function
 PreprocessorParseExpressionTerm = True
Case token_decnum, token_hexnum, token_octnum, token_floatnum
 LHS = Val(t.sValue)
 If Not GetNextToken_Internal(objFile, t) Then Exit Function
 PreprocessorParseExpressionTerm = True
Case keyword_true
 LHS = -1
 If Not GetNextToken_Internal(objFile, t) Then Exit Function
 PreprocessorParseExpressionTerm = True
Case keyword_false
 LHS = 0
 If Not GetNextToken_Internal(objFile, t) Then Exit Function
 PreprocessorParseExpressionTerm = True
Case token_string, token_currencynum, token_datenum
 'TODO:
 PrintError "Currently unsupported '" + t.sValue + "'", t.nLine, t.nColumn
Case Else
 PrintError "Identifier or constant expected", t.nLine, t.nColumn
End Select
End Function
