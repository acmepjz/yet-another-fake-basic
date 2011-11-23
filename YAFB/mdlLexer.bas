Attribute VB_Name = "mdlLexer"
Option Explicit

Public sKeywords() As String, nKeywordCount As Long

'////////
'without preprocessor
'TODO:frx process

Public Function GetNextToken_Internal(ByVal objFile As ISource, ByRef ret As typeToken) As Boolean
Dim m As Long
Dim s1 As String, s2 As String
Dim c As Long
Dim c1 As Long
Dim i As Long
'///
Dim nFlags As Long
Dim CanBeLineNumber As Boolean
'///
If nKeywordCount = 0 Then pInitKeyword
'///
ret.nFlags = 0
ret.nFlags2 = 0
ret.sValue = vbNullString
GetNextToken_Internal = True
'///
CanBeLineNumber = objFile.Column = 1
'///
label_start:
c1 = objFile.GetCh
s1 = ChrW(c1)
Select Case c1
Case -1
 ret.nType = token_eof
 GoTo label_over
Case [" "], ["\t"]
 If objFile.GetCh = ["_"] Then
  Select Case objFile.GetCh
  Case ["\r"]
   If objFile.GetCh <> ["\n"] Then objFile.UnGetCh
  Case ["\n"]
   If objFile.GetCh <> ["\r"] Then objFile.UnGetCh
  Case Else
   objFile.UnGetCh 2
  End Select
 Else
  objFile.UnGetCh
 End If
 nFlags = nFlags Or 1&
 GoTo label_start
Case ["0"] To ["9"]
 GoTo label_num
Case ["&"]
 c = objFile.GetCh
 Select Case c
 Case ["h"], ["hh"]
  s1 = s1 + ChrW(c)
  c = objFile.GetCh
  Select Case c
  Case ["0"] To ["9"], ["a"] To ["f"], ["aa"] To ["ff"]
   s1 = s1 + ChrW(c)
   GoTo label_hexnum
  Case Else
   objFile.UnGetCh 2
  End Select
 Case ["o"], ["oo"]
  s1 = s1 + ChrW(c)
  c = objFile.GetCh
  Select Case c
  Case ["0"] To ["7"]
   s1 = s1 + ChrW(c)
   GoTo label_octnum
  Case Else
   objFile.UnGetCh 2
  End Select
 Case ["0"] To ["7"]
  s1 = s1 + ChrW(c)
  GoTo label_octnum
 Case Else
  objFile.UnGetCh
 End Select
 ret.nType = token_and
 GoTo label_over
Case ["""]
 s1 = vbNullString
 Do
  c = objFile.GetCh
  Select Case c
  Case -1, ["\r"], ["\n"]
   s1 = "Missing end-of-string quote"
   GoTo label_error
  Case ["""]
   If objFile.GetCh = ["""] Then
    s1 = s1 + ChrW(c)
   Else
    objFile.UnGetCh
    Exit Do
   End If
  Case Else
   s1 = s1 + ChrW(c)
  End Select
 Loop
 ret.nType = token_string
 ret.sValue = s1
 GoTo label_over
Case ["\r"], ["\n"]
 Do
  c = objFile.GetCh
  Select Case c
  Case ["\r"], ["\n"]
  Case Else
   Exit Do
  End Select
 Loop
 objFile.UnGetCh
 ret.nType = token_crlf
 GoTo label_over
Case [":"]
 Do While objFile.GetCh = [":"]
 Loop
 objFile.UnGetCh
 ret.nType = token_colon
 GoTo label_over
Case ["."]
 c = objFile.GetCh
 Select Case c
 Case ["0"] To ["9"]
  s1 = s1 + ChrW(c)
  GoTo label_floatnum
 Case Else
  objFile.UnGetCh
 End Select
 ret.nType = token_dot
 GoTo label_over
Case [","]
 ret.nType = token_comma
 GoTo label_over
Case [";"]
 ret.nType = token_semicolon
 GoTo label_over
Case ["#"]
 GoTo label_poundsign
Case ["("]
 ret.nType = token_lbracket
 GoTo label_over
Case [")"]
 ret.nType = token_rbracket
 GoTo label_over
Case ["+"]
 ret.nType = token_plus
 GoTo label_over
Case ["-"]
 ret.nType = token_minus
 GoTo label_over
Case ["*"]
 ret.nType = token_asterisk
 GoTo label_over
Case ["/"]
 ret.nType = token_slash
 GoTo label_over
Case ["\"]
 ret.nType = token_backslash
 GoTo label_over
Case ["="]
 Select Case objFile.GetCh
 Case ["<"]
  ret.nType = token_le
 Case [">"]
  ret.nType = token_ge
 Case Else
  objFile.UnGetCh
  ret.nType = token_equal
 End Select
 GoTo label_over
Case ["^"]
 ret.nType = token_power
 Exit Function
Case [">"]
 Select Case objFile.GetCh
 Case ["<"]
  ret.nType = token_ne
 Case ["="]
  ret.nType = token_ge
 Case [">"]
  Select Case objFile.GetCh
  Case [">"]
   ret.nType = token_ror
  Case Else
   objFile.UnGetCh
   ret.nType = token_shr
  End Select
 Case Else
  objFile.UnGetCh
  ret.nType = token_gt
 End Select
 GoTo label_over
Case ["<"]
 Select Case objFile.GetCh
 Case [">"]
  ret.nType = token_ne
 Case ["="]
  ret.nType = token_le
 Case ["<"]
  Select Case objFile.GetCh
  Case ["<"]
   ret.nType = token_rol
  Case Else
   objFile.UnGetCh
   ret.nType = token_shl
  End Select
 Case Else
  objFile.UnGetCh
  ret.nType = token_lt
 End Select
 GoTo label_over
Case ["'"]
 GoTo label_comment
Case Else
 i = c1
 If (i >= ["a"] And i <= ["z"]) Or (i >= ["aa"] And i <= ["zz"]) Or i > 127 Then
  Do
   c = objFile.GetCh
   i = c
   If (i >= ["0"] And i <= ["9"]) Or (i >= ["a"] And i <= ["z"]) Or i = ["_"] Or (i >= ["aa"] And i <= ["zz"]) Or i > 127 Then
    s1 = s1 + ChrW(c)
   ElseIf i = ["!"] Or i = ["#"] Or i = ["$"] Or i = ["%"] Or i = ["&"] Or i = ["@"] Then
    s1 = s1 + ChrW(c)
    Exit Do
   ElseIf StrComp(s1, "Rem", vbTextCompare) = 0 Then
    objFile.UnGetCh
    GoTo label_comment
   Else
    Exit Do
   End If
  Loop
  '///
  ret.nType = 1000 + pIsKeyword(s1)
  If ret.nType = 1000 And i = [":"] And CanBeLineNumber Then
   ret.nType = token_linenumber
  Else
   objFile.UnGetCh
  End If
  ret.nLine = objFile.Line
  ret.nColumn = objFile.Column
  ret.sValue = s1
  Exit Function
 ElseIf i = ["lll"] Then
  Do
   c = objFile.GetCh
   i = c
   If i = ["\r"] Or i = ["\n"] Then
    s1 = "Unexpected end of line"
    GoTo label_error
   ElseIf i = ["rrr"] Then
    s1 = s1 + ChrW(c)
    Exit Do
   Else
    s1 = s1 + ChrW(c)
   End If
  Loop
  ret.nType = 1000 + pIsKeyword(s1)
  ret.nLine = objFile.Line
  ret.nColumn = objFile.Column
  ret.sValue = s1
  Exit Function
 Else
  s1 = "Invalid character"
  GoTo label_error
 End If
End Select
'////////////////////////////////////////
label_comment:
Do
 c = objFile.GetCh
 Select Case c
 Case [" "], ["\t"]
  's1 = s1 + ChrW(c)
  If objFile.GetCh = ["_"] Then
   Select Case objFile.GetCh
   Case ["\r"]
    If objFile.GetCh <> ["\n"] Then objFile.UnGetCh
   Case ["\n"]
    If objFile.GetCh <> ["\r"] Then objFile.UnGetCh
   Case Else
    objFile.UnGetCh 2
   End Select
  Else
   objFile.UnGetCh
  End If
 Case -1, ["\r"], ["\n"]
  objFile.UnGetCh
  Exit Do
 Case Else
  's1 = s1 + ChrW(c)
 End Select
Loop
nFlags = nFlags Or 1&
GoTo label_start
'////////////////////////////////////////
label_num:
Do
 c = objFile.GetCh
 Select Case c
 Case ["0"] To ["9"]
  s1 = s1 + ChrW(c)
 Case ["."]
  s1 = s1 + ChrW(c)
  GoTo label_floatnum
 Case ["e"], ["d"], ["ee"], ["dd"]
  GoTo label_floatnum_e
 Case ["%"], ["&"]
  s1 = s1 + ChrW(c)
  ret.nType = token_decnum
  ret.nLine = objFile.Line
  ret.nColumn = objFile.Column
  ret.sValue = s1
  Exit Function
 Case ["!"], ["#"]
  s1 = s1 + ChrW(c)
  ret.nType = token_floatnum
  ret.nLine = objFile.Line
  ret.nColumn = objFile.Column
  ret.sValue = s1
  Exit Function
 Case ["@"]
  s1 = s1 + ChrW(c)
  ret.nType = token_currencynum
  ret.nLine = objFile.Line
  ret.nColumn = objFile.Column
  ret.sValue = s1
  Exit Function
 Case Else
  objFile.UnGetCh
  If CanBeLineNumber Then ret.nType = token_linenumber _
  Else ret.nType = token_decnum
  ret.nLine = objFile.Line
  ret.nColumn = objFile.Column
  ret.sValue = s1
  Exit Function
 End Select
Loop
'////////////////////////////////////////
label_hexnum:
Do
 c = objFile.GetCh
 Select Case c
 Case ["0"] To ["9"], ["a"] To ["f"], ["aa"] To ["ff"]
  s1 = s1 + ChrW(c)
 Case ["%"], ["&"]
  s1 = s1 + ChrW(c)
  ret.nType = token_hexnum
  ret.nLine = objFile.Line
  ret.nColumn = objFile.Column
  ret.sValue = s1
  Exit Function
 Case Else
  objFile.UnGetCh
  ret.nType = token_hexnum
  ret.nLine = objFile.Line
  ret.nColumn = objFile.Column
  ret.sValue = s1
  Exit Function
 End Select
Loop
'////////////////////////////////////////
label_octnum:
Do
 c = objFile.GetCh
 Select Case c
 Case ["0"] To ["7"]
  s1 = s1 + ChrW(c)
 Case ["%"], ["&"]
  s1 = s1 + ChrW(c)
  ret.nType = token_octnum
  ret.nLine = objFile.Line
  ret.nColumn = objFile.Column
  ret.sValue = s1
  Exit Function
 Case Else
  objFile.UnGetCh
  ret.nType = token_octnum
  ret.nLine = objFile.Line
  ret.nColumn = objFile.Column
  ret.sValue = s1
  Exit Function
 End Select
Loop
'////////////////////////////////////////
label_floatnum:
Do
 c = objFile.GetCh
 Select Case c
 Case ["0"] To ["9"]
  s1 = s1 + ChrW(c)
 Case ["e"], ["d"], ["ee"], ["dd"]
  GoTo label_floatnum_e
 Case ["!"], ["#"]
  s1 = s1 + ChrW(c)
  ret.nType = token_floatnum
  ret.nLine = objFile.Line
  ret.nColumn = objFile.Column
  ret.sValue = s1
  Exit Function
 Case ["@"]
  s1 = s1 + ChrW(c)
  ret.nType = token_currencynum
  ret.nLine = objFile.Line
  ret.nColumn = objFile.Column
  ret.sValue = s1
  Exit Function
 Case Else
  objFile.UnGetCh
  ret.nType = token_floatnum
  ret.nLine = objFile.Line
  ret.nColumn = objFile.Column
  ret.sValue = s1
  Exit Function
 End Select
Loop
'////////////////////////////////////////
label_floatnum_e:
s2 = "E"
i = 2
c = objFile.GetCh
Select Case c
Case ["+"], ["-"]
 s2 = s2 + ChrW(c)
 i = i + 1
 c = objFile.GetCh
End Select
Select Case c
Case ["0"] To ["9"]
 i = token_floatnum
 s2 = s2 + ChrW(c)
 Do
  c = objFile.GetCh
  Select Case c
  Case ["0"] To ["9"]
   s2 = s2 + ChrW(c)
  Case ["!"], ["#"]
   s2 = s2 + ChrW(c)
   Exit Do
  Case ["@"]
   i = token_currencynum
   s2 = s2 + ChrW(c)
   Exit Do
  Case Else
   objFile.UnGetCh
   Exit Do
  End Select
 Loop
 ret.nType = i
 ret.nLine = objFile.Line
 ret.nColumn = objFile.Column
 ret.sValue = s1 + s2
 Exit Function
Case Else
 objFile.UnGetCh i
 ret.nType = token_floatnum
 ret.nLine = objFile.Line
 ret.nColumn = objFile.Column
 ret.sValue = s1
 Exit Function
End Select
'////////////////////////////////////////
label_poundsign:
If CanBeLineNumber Then
 '///must be preprocessor instructions
 Do
  c = objFile.GetCh
  Select Case c
  Case ["a"] To ["z"], ["aa"] To ["zz"]
   s1 = s1 + ChrW(c)
  Case Else
   objFile.UnGetCh
   Exit Do
  End Select
 Loop
 '///
 Select Case LCase(s1)
 Case "#const"
  ret.nType = preprocessor_const
 Case "#else"
  ret.nType = preprocessor_else
 Case "#elseif"
  ret.nType = preprocessor_elseif
 Case "#end"
  ret.nType = preprocessor_end
 Case "#error"
  ret.nType = preprocessor_error
 Case "#if"
  ret.nType = preprocessor_if
 Case Else
  s1 = "'#Const' or '#Else' or '#ElseIf' or '#End' or '#If' expected"
  GoTo label_error
 End Select
 ret.sValue = s1
Else
 i = 0
 s1 = vbNullString
 Do While i < 30
  i = i + 1
  c = objFile.GetCh
  Select Case c
  Case ["#"]
   If i < 4 Then Exit Do
   ret.nType = token_datenum
   ret.nLine = objFile.Line
   ret.nColumn = objFile.Column
   ret.sValue = s1
   Exit Function
  Case ["0"] To ["9"], [" "], ["\t"], ["/"], ["-"], [":"], ["a"] To ["z"], ["aa"] To ["zz"]
   s1 = s1 + ChrW(c)
  Case Else
   Exit Do
  End Select
 Loop
 objFile.UnGetCh i
 ret.nType = token_poundsign
End If
GoTo label_over
'////////////////////////////////////////
label_over:
ret.nLine = objFile.Line
ret.nColumn = objFile.Column
ret.nFlags = nFlags
Exit Function
'////////////////////////////////////////
label_error:
GetNextToken_Internal = False
'///
ret.nType = token_err
ret.nLine = objFile.Line
ret.nColumn = objFile.Column
'///
PrintError s1, ret.nLine, ret.nColumn
End Function

Private Function pIsKeyword(ByRef s As String) As Long
Dim i As Long, j As Long, k As Long
Dim n As Long
i = 1
j = nKeywordCount
Do Until i > j
 k = (i + j) \ 2
 n = StrComp(s, sKeywords(k), vbTextCompare)
 If n = 0 Then
  pIsKeyword = k
  Exit Function
 ElseIf n > 0 Then
  i = k + 1
 Else
  j = k - 1
 End If
Loop
pIsKeyword = 0
End Function

Private Sub pInitKeyword()
'### BEGIN INIT KEYWORD
nKeywordCount = 76
ReDim sKeywords(1 To 76)
sKeywords(1) = "alias"
sKeywords(2) = "and"
sKeywords(3) = "as"
sKeywords(4) = "attribute"
sKeywords(5) = "byref"
sKeywords(6) = "byval"
sKeywords(7) = "call"
sKeywords(8) = "case"
sKeywords(9) = "cdecl"
sKeywords(10) = "close"
sKeywords(11) = "const"
sKeywords(12) = "declare"
sKeywords(13) = "dim"
sKeywords(14) = "do"
sKeywords(15) = "each"
sKeywords(16) = "else"
sKeywords(17) = "elseif"
sKeywords(18) = "end"
sKeywords(19) = "enum"
sKeywords(20) = "eqv"
sKeywords(21) = "erase"
sKeywords(22) = "exit"
sKeywords(23) = "false"
sKeywords(24) = "fastcall"
sKeywords(25) = "for"
sKeywords(26) = "friend"
sKeywords(27) = "function"
sKeywords(28) = "get"
sKeywords(29) = "global"
sKeywords(30) = "goto"
sKeywords(31) = "if"
sKeywords(32) = "imp"
sKeywords(33) = "in"
sKeywords(34) = "input"
sKeywords(35) = "is"
sKeywords(36) = "let"
sKeywords(37) = "lib"
sKeywords(38) = "line"
sKeywords(39) = "loop"
sKeywords(40) = "lset"
sKeywords(41) = "mod"
sKeywords(42) = "new"
sKeywords(43) = "next"
sKeywords(44) = "not"
sKeywords(45) = "on"
sKeywords(46) = "open"
sKeywords(47) = "option"
sKeywords(48) = "optional"
sKeywords(49) = "or"
sKeywords(50) = "paramarray"
sKeywords(51) = "preserve"
sKeywords(52) = "print"
sKeywords(53) = "private"
sKeywords(54) = "property"
sKeywords(55) = "public"
sKeywords(56) = "put"
sKeywords(57) = "raiseevent"
sKeywords(58) = "redim"
sKeywords(59) = "rset"
sKeywords(60) = "select"
sKeywords(61) = "set"
sKeywords(62) = "static"
sKeywords(63) = "stdcall"
sKeywords(64) = "step"
sKeywords(65) = "sub"
sKeywords(66) = "then"
sKeywords(67) = "to"
sKeywords(68) = "true"
sKeywords(69) = "type"
sKeywords(70) = "until"
sKeywords(71) = "wend"
sKeywords(72) = "while"
sKeywords(73) = "with"
sKeywords(74) = "withevents"
sKeywords(75) = "write"
sKeywords(76) = "xor"
 '### END INIT KEYWORD
End Sub

Public Function GetBinaryTokPrecedence(ByVal nType As enumTokenType) As Long
Select Case nType
Case keyword_imp '??
 GetBinaryTokPrecedence = 7
Case keyword_xor, keyword_eqv '??
 GetBinaryTokPrecedence = 10
Case keyword_or
 GetBinaryTokPrecedence = 20
Case keyword_and
 GetBinaryTokPrecedence = 30
Case token_gt, token_lt, token_ge, token_le, token_equal, token_ne, keyword_is
 GetBinaryTokPrecedence = 40
Case token_shl, token_shr, token_rol, token_ror
 GetBinaryTokPrecedence = 47
Case token_and
 GetBinaryTokPrecedence = 50
Case token_plus, token_minus
 GetBinaryTokPrecedence = 60
Case keyword_mod
 GetBinaryTokPrecedence = 70
Case token_backslash
 GetBinaryTokPrecedence = 80
Case token_asterisk, token_slash
 GetBinaryTokPrecedence = 90
Case token_power
 GetBinaryTokPrecedence = 100
Case Else
 GetBinaryTokPrecedence = -1
End Select
End Function

Public Function GetUnaryTokPrecedence(ByVal nType As enumTokenType) As Long
Select Case nType
Case keyword_not
 GetUnaryTokPrecedence = 35
Case token_plus, token_minus
 GetUnaryTokPrecedence = 95
Case Else
 GetUnaryTokPrecedence = -1
End Select
End Function
