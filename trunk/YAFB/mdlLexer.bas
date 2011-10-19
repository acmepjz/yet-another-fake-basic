Attribute VB_Name = "mdlLexer"
Option Explicit

Public sKeywords() As String, nKeywordCount As Long

'Private Function pAddStringContent(ByRef t As typeSourceFile, ByRef s As String) As Long
'With t.tContents
' .nValueCount = .nValueCount + 1
' If .nValueCount > .nValueMax Then
'  .nValueMax = .nValueMax + .nValueMax \ 2& + 256&
'  ReDim Preserve .sValues(1 To .nValueMax)
' End If
' .sValues(.nValueCount) = s
' pAddStringContent = .nValueCount
'End With
'End Function
'
'Private Function pAddToken(ByRef t As typeSourceFile, ByVal nType As enumTokenType, Optional ByVal nValueIndex As Long = 0) As Long
'With t.tContents
' .nTokenCount = .nTokenCount + 1
' If .nTokenCount > .nTokenMax Then
'  .nTokenMax = .nTokenMax + .nTokenMax \ 2& + 1024&
'  ReDim Preserve .tTokens(1 To .nTokenMax)
' End If
' With .tTokens(.nTokenCount)
'  .nType = nType
'  .nValueIndex = nValueIndex
' End With
' pAddToken = .nTokenCount
'End With
'End Function
'
'Private Function pAddTokenWithString(ByRef t As typeSourceFile, ByVal nType As enumTokenType, ByRef s As String) As Long
'Dim i As Long
'With t.tContents
' i = .nValueCount + 1
' .nValueCount = i
' If i > .nValueMax Then
'  .nValueMax = .nValueMax + .nValueMax \ 2& + 256&
'  ReDim Preserve .sValues(1 To .nValueMax)
' End If
' .sValues(i) = s
' '///
' .nTokenCount = .nTokenCount + 1
' If .nTokenCount > .nTokenMax Then
'  .nTokenMax = .nTokenMax + .nTokenMax \ 2& + 1024&
'  ReDim Preserve .tTokens(1 To .nTokenMax)
' End If
' With .tTokens(.nTokenCount)
'  .nType = nType
'  .nValueIndex = i
' End With
' pAddTokenWithString = .nTokenCount
'End With
'End Function

'////////
'TODO:preprocessor
'TODO:frx process

Public Function GetNextToken(ByVal objFile As ISource, ByRef ret As typeToken) As Boolean
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
GetNextToken = True
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
  s1 = s1 + ChrW(c)
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
  s1 = s1 + ChrW(c)
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
GoTo label_over
'////////////////////////////////////////
label_over:
ret.nLine = objFile.Line
ret.nColumn = objFile.Column
ret.nFlags = nFlags
Exit Function
'////////////////////////////////////////
label_error:
GetNextToken = False
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
nKeywordCount = 72
ReDim sKeywords(1 To 72)
sKeywords(1) = "alias"
sKeywords(2) = "and"
sKeywords(3) = "as"
sKeywords(4) = "attribute"
sKeywords(5) = "byref"
sKeywords(6) = "byval"
sKeywords(7) = "call"
sKeywords(8) = "case"
sKeywords(9) = "close"
sKeywords(10) = "const"
sKeywords(11) = "declare"
sKeywords(12) = "dim"
sKeywords(13) = "do"
sKeywords(14) = "each"
sKeywords(15) = "else"
sKeywords(16) = "elseif"
sKeywords(17) = "end"
sKeywords(18) = "enum"
sKeywords(19) = "eqv"
sKeywords(20) = "exit"
sKeywords(21) = "false"
sKeywords(22) = "for"
sKeywords(23) = "friend"
sKeywords(24) = "function"
sKeywords(25) = "get"
sKeywords(26) = "global"
sKeywords(27) = "goto"
sKeywords(28) = "if"
sKeywords(29) = "imp"
sKeywords(30) = "in"
sKeywords(31) = "input"
sKeywords(32) = "is"
sKeywords(33) = "let"
sKeywords(34) = "lib"
sKeywords(35) = "line"
sKeywords(36) = "loop"
sKeywords(37) = "lset"
sKeywords(38) = "mod"
sKeywords(39) = "new"
sKeywords(40) = "next"
sKeywords(41) = "not"
sKeywords(42) = "on"
sKeywords(43) = "open"
sKeywords(44) = "option"
sKeywords(45) = "optional"
sKeywords(46) = "or"
sKeywords(47) = "paramarray"
sKeywords(48) = "preserve"
sKeywords(49) = "print"
sKeywords(50) = "private"
sKeywords(51) = "property"
sKeywords(52) = "public"
sKeywords(53) = "put"
sKeywords(54) = "raiseevent"
sKeywords(55) = "redim"
sKeywords(56) = "rset"
sKeywords(57) = "select"
sKeywords(58) = "set"
sKeywords(59) = "static"
sKeywords(60) = "step"
sKeywords(61) = "sub"
sKeywords(62) = "then"
sKeywords(63) = "to"
sKeywords(64) = "true"
sKeywords(65) = "type"
sKeywords(66) = "until"
sKeywords(67) = "wend"
sKeywords(68) = "while"
sKeywords(69) = "with"
sKeywords(70) = "withevents"
sKeywords(71) = "write"
sKeywords(72) = "xor"
 '### END INIT KEYWORD
End Sub
