VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSrcParser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_nFileType As Long
'0-bas
'1-cls
'2-frm
'3-ctl

Private m_objNodes As New Collection

Private m_objTable As clsSymbolTable

Private m_objForStack() As ForStatementNode '1-based
Private m_nForStackSize As Long

Private m_objExitStack() As IASTNode '1-based
Private m_nExitStackSize As Long

Friend Sub AddExitStack(ByVal obj As IASTNode)
m_nExitStackSize = m_nExitStackSize + 1
ReDim Preserve m_objExitStack(1 To m_nExitStackSize)
Set m_objExitStack(m_nExitStackSize) = obj
End Sub

Friend Sub RemoveExitStack()
If m_nExitStackSize > 0 Then
 Set m_objExitStack(m_nExitStackSize) = Nothing
 m_nExitStackSize = m_nExitStackSize - 1
End If
End Sub

Friend Sub ResetExitStack()
Erase m_objExitStack
m_nExitStackSize = 0
End Sub

Friend Function QueryExitStack(ByVal nType As enumASTNodeType) As IASTNode
Dim i As Long
For i = m_nExitStackSize To 1 Step -1
 If m_objExitStack(i).NodeType = nType Then
  Set QueryExitStack = m_objExitStack(i)
  Exit Function
 End If
Next i
End Function

Friend Property Get SymbolTable() As clsSymbolTable
Set SymbolTable = m_objTable
End Property

Friend Property Get FileType() As Long
FileType = m_nFileType
End Property

Friend Function ParseFile(ByVal objFile As ISource) As Boolean
Dim nFlags As Long
Dim obj As IASTNode
'///
Set m_objNodes = New Collection
'///get file type
Select Case LCase(Right(objFile.FileName, 4))
Case ".cls"
 m_nFileType = 1
Case ".frm"
 m_nFileType = 2
Case ".ctl"
 m_nFileType = 3
Case Else
 m_nFileType = 0
End Select
'///get first token
If Not GetNextToken(objFile, g_tToken) Then Exit Function
'////////
Do
 If g_tToken.nType <= 0 Then Exit Do
 '///
 nFlags = 0
 '///
 Select Case g_tToken.nType
 Case token_crlf, token_colon
 Case keyword_attribute
  If Not GetNextToken(objFile, g_tToken) Then Exit Function
  'TODO:
  If Not ParseSkip(objFile) Then Exit Function
 Case keyword_option
  If Not GetNextToken(objFile, g_tToken) Then Exit Function
  'TODO:
  If Not ParseSkip(objFile) Then Exit Function
 Case keyword_private, keyword_public, keyword_friend, _
 keyword_static, keyword_type, keyword_enum, _
 keyword_declare, keyword_sub, keyword_function, keyword_property
  '////////
  Select Case g_tToken.nType
  Case keyword_public
   If Not GetNextToken(objFile, g_tToken) Then Exit Function
  Case keyword_private
   nFlags = &H10&
   If Not GetNextToken(objFile, g_tToken) Then Exit Function
  Case keyword_friend
   nFlags = &H20&
   If Not GetNextToken(objFile, g_tToken) Then Exit Function
  End Select
  '///
  If g_tToken.nType = keyword_static Then
   nFlags = nFlags Or &H200&
   If Not GetNextToken(objFile, g_tToken) Then Exit Function
  End If
  '///
  If g_tToken.nType = keyword_declare Then
   If nFlags And &H220& Then
    PrintError "'Sub' or 'Function' or 'Property' expected"
    Exit Function
   End If
   nFlags = nFlags Or &H100&
   If Not GetNextToken(objFile, g_tToken) Then Exit Function
  End If
  '///
  Select Case g_tToken.nType
  Case keyword_sub
   If Not GetNextToken(objFile, g_tToken) Then Exit Function
   Set obj = ParseFunction(objFile, nFlags)
   If obj Is Nothing Then Exit Function
   m_objNodes.Add obj
  Case keyword_function
   If Not GetNextToken(objFile, g_tToken) Then Exit Function
   Set obj = ParseFunction(objFile, nFlags Or 1&)
   If obj Is Nothing Then Exit Function
   m_objNodes.Add obj
  Case keyword_property
   If nFlags And &H100& Then
    PrintError "'Sub' or 'Function' expected"
    Exit Function
   End If
   If Not GetNextToken(objFile, g_tToken) Then Exit Function
   '///
   Select Case g_tToken.nType
   Case keyword_get
    nFlags = nFlags Or 2&
   Case keyword_let
    nFlags = nFlags Or 3&
   Case keyword_set
    nFlags = nFlags Or 4&
   Case Else
    PrintError "'Get' or 'Let' or 'Set' expected"
    Exit Function
   End Select
   '///
   If Not GetNextToken(objFile, g_tToken) Then Exit Function
   Set obj = ParseFunction(objFile, nFlags)
   If obj Is Nothing Then Exit Function
   m_objNodes.Add obj
   '///
  Case keyword_type
   If nFlags And &H320& Then
    PrintError "'Sub' or 'Function' or 'Property' expected"
    Exit Function
   End If
   If Not GetNextToken(objFile, g_tToken) Then Exit Function
   Set obj = ParseType(objFile, nFlags)
   If obj Is Nothing Then Exit Function
   m_objNodes.Add obj
  Case keyword_enum
   If nFlags And &H320& Then
    PrintError "'Sub' or 'Function' or 'Property' expected"
    Exit Function
   End If
   If Not GetNextToken(objFile, g_tToken) Then Exit Function
   Set obj = ParseEnum(objFile, nFlags)
   If obj Is Nothing Then Exit Function
   m_objNodes.Add obj
  Case keyword_const
   If nFlags And &H320& Then
    PrintError "'Sub' or 'Function' or 'Property' expected"
    Exit Function
   End If
   If Not GetNextToken(objFile, g_tToken) Then Exit Function
   Set obj = ParseConst(objFile, nFlags)
   If obj Is Nothing Then Exit Function
   m_objNodes.Add obj
  Case Else
   If nFlags And &H320& Then
    PrintError "'Sub' or 'Function' or 'Property' expected"
    Exit Function
   End If
   Set obj = ParseDim(objFile, nFlags)
   If obj Is Nothing Then Exit Function
   m_objNodes.Add obj
  End Select
 Case keyword_dim
  If Not GetNextToken(objFile, g_tToken) Then Exit Function
  Set obj = ParseDim(objFile, &H10&)
  If obj Is Nothing Then Exit Function
  m_objNodes.Add obj
 Case keyword_const
  If Not GetNextToken(objFile, g_tToken) Then Exit Function
  Set obj = ParseConst(objFile, &H10&)
  If obj Is Nothing Then Exit Function
  m_objNodes.Add obj
 Case Else
  PrintError "Currently unsupported"
  Exit Function
 End Select
 '///
 If Not ParseEndOfLine(objFile) Then Exit Function
Loop
'///
ParseFile = g_tToken.nType < 0
End Function

'///nEndType:
'0=sub
'1=function
'2=property get
'3=property let
'4=property set
'///and node_XXX etc.
'node_ifstat       -- xx:xx:xx {else}|{vbCrLf}
'node_ifblock      -- end with {else}|{elseif}|{end}{if} <-- nFlagsReturn=1 means {end}{if}
'node_elseblock    -- end with {end}{if}
'node_forstat      -- end with {next} (or {,})
'node_dostat       -- end with {loop}
'node_whilestat    -- end with {wend}
'node_selectblock  -- end with {case}|{end} <-- XXX
Friend Function ParseStatementList(ByVal objFile As ISource, ByVal nEndType As Long, Optional ByVal nReserved As Long, Optional ByVal objFunction As FunctionNode, Optional ByRef nFlagsReturn As Long) As StatementListNode
Dim objList As New StatementListNode
Dim obj As IASTNode
Dim objVariable As VariableNode
Dim bParsedLineNumber As Boolean
'///
nFlagsReturn = 0
'///
Do
 bParsedLineNumber = False
 '///
 Select Case g_tToken.nType
 Case token_crlf, token_colon
 Case keyword_dim
  If Not GetNextToken(objFile, g_tToken) Then Exit Function
  Set obj = ParseDim(objFile, 0, objFunction)
  If obj Is Nothing Then Exit Function
  objList.AddSubNode obj
 Case keyword_const
  If Not GetNextToken(objFile, g_tToken) Then Exit Function
  Set obj = ParseConst(objFile, 0, objFunction)
  If obj Is Nothing Then Exit Function
  objList.AddSubNode obj
 Case keyword_static
  If Not GetNextToken(objFile, g_tToken) Then Exit Function
  Set obj = ParseDim(objFile, &H200&, objFunction)
  If obj Is Nothing Then Exit Function
  objList.AddSubNode obj
 Case keyword_const
  If Not GetNextToken(objFile, g_tToken) Then Exit Function
  Set obj = ParseConst(objFile, 0, objFunction)
  If obj Is Nothing Then Exit Function
  objList.AddSubNode obj
 Case keyword_end
  If nEndType = node_selectblock Then
   'don't eat current token
   Exit Do
  End If
  If Not GetNextToken(objFile, g_tToken) Then Exit Function
  Select Case g_tToken.nType
  Case keyword_sub
   If nEndType <> 0 Then
    PrintError "'End Sub' without 'Sub'"
    Exit Function
   End If
   If Not GetNextToken(objFile, g_tToken) Then Exit Function
   Exit Do
  Case keyword_function
   If nEndType <> 1 Then
    PrintError "'End Function' without 'Function'"
    Exit Function
   End If
   If Not GetNextToken(objFile, g_tToken) Then Exit Function
   Exit Do
  Case keyword_property
   If nEndType <> 2 And nEndType <> 3 And nEndType <> 4 Then
    PrintError "'End Property' without 'Property'"
    Exit Function
   End If
   If Not GetNextToken(objFile, g_tToken) Then Exit Function
   Exit Do
  Case keyword_if
   If nEndType <> node_ifblock And nEndType <> node_elseblock Then
    PrintError "'End If' without 'If'"
    Exit Function
   End If
   If Not GetNextToken(objFile, g_tToken) Then Exit Function
   nFlagsReturn = 1
   Exit Do
  Case Else
   PrintError "Currently unsupported 'End'"
   Exit Function
  End Select
 Case keyword_exit
  If Not GetNextToken(objFile, g_tToken) Then Exit Function
  '///
  If g_tToken.nType > 1000 And g_tToken.nType < 2000 Then
   With New ExitStatementNode
    If Not .SetToken(True) Then Exit Function
    objList.AddSubNode .This
   End With
   If Not GetNextToken(objFile, g_tToken) Then Exit Function
  Else
   PrintError "Keyword expected"
   Exit Function
  End If
 Case keyword_if
  If Not GetNextToken(objFile, g_tToken) Then Exit Function
  Set obj = ParseIf(objFile, objFunction)
  If obj Is Nothing Then Exit Function
  objList.AddSubNode obj
 Case keyword_else
  If nEndType <> node_ifstat And nEndType <> node_ifblock Then
   PrintError "'Else' without 'If'"
   Exit Function
  End If
  Exit Do
 Case keyword_elseif
  If nEndType <> node_ifblock Then
   PrintError "'ElseIf' without 'If'"
   Exit Function
  End If
  Exit Do
 Case keyword_for
  If Not GetNextToken(objFile, g_tToken) Then Exit Function
  Set obj = ParseFor(objFile, objFunction)
  If obj Is Nothing Then Exit Function
  objList.AddSubNode obj
  '///
  If g_tToken.nType = token_comma Then
   If nEndType <> node_forstat Then
    PrintError "End of line expected"
    Exit Function
   End If
   If Not GetNextToken(objFile, g_tToken) Then Exit Function
   Exit Do
  End If
 Case keyword_do, keyword_while
  'don't eat current token
  Set obj = ParseDo(objFile, objFunction)
  If obj Is Nothing Then Exit Function
  objList.AddSubNode obj
 Case keyword_loop
  If nEndType <> node_dostat Then
   PrintError "'Loop' without 'Do'"
   Exit Function
  End If
  If Not GetNextToken(objFile, g_tToken) Then Exit Function
  Exit Do
 Case keyword_wend
  If nEndType <> node_whilestat Then
   PrintError "'Wend' without 'While'"
   Exit Function
  End If
  If Not GetNextToken(objFile, g_tToken) Then Exit Function
  Exit Do
 Case keyword_next
  If nEndType <> node_forstat Then
   PrintError "'Next' without 'For'"
   Exit Function
  End If
  If Not GetNextToken(objFile, g_tToken) Then Exit Function
  Exit Do
 Case keyword_select
  If Not GetNextToken(objFile, g_tToken) Then Exit Function
  Set obj = ParseSelectCase(objFile, Nothing)
  If obj Is Nothing Then Exit Function
  objList.AddSubNode obj
 Case keyword_case
  If nEndType <> node_selectblock Then
   PrintError "'Case' without 'Select Case'"
   Exit Function
  End If
  'don't eat current token
  Exit Do
 Case keyword_let
  If Not GetNextToken(objFile, g_tToken) Then Exit Function
  Set obj = ParseMake(objFile, 0, Nothing)
  If obj Is Nothing Then Exit Function
  objList.AddSubNode obj
 Case keyword_set
  If Not GetNextToken(objFile, g_tToken) Then Exit Function
  Set obj = ParseMake(objFile, 1, Nothing)
  If obj Is Nothing Then Exit Function
  objList.AddSubNode obj
 Case keyword_lset
  If Not GetNextToken(objFile, g_tToken) Then Exit Function
  Set obj = ParseMake(objFile, 2, Nothing)
  If obj Is Nothing Then Exit Function
  objList.AddSubNode obj
 Case keyword_rset
  If Not GetNextToken(objFile, g_tToken) Then Exit Function
  Set obj = ParseMake(objFile, 3, Nothing)
  If obj Is Nothing Then Exit Function
  objList.AddSubNode obj
 Case token_id, token_dot
  Set objVariable = ParseVariable(objFile)
  If objVariable Is Nothing Then Exit Function
  '///
  If g_tToken.nType = token_equal Then
   Set obj = ParseMake(objFile, 0, objVariable)
   If obj Is Nothing Then Exit Function
   objList.AddSubNode obj
  Else
   Set obj = ParseCall(objFile, 0, objVariable)
   If obj Is Nothing Then Exit Function
   objList.AddSubNode obj
  End If
 Case keyword_call
  If Not GetNextToken(objFile, g_tToken) Then Exit Function
  Set obj = ParseCall(objFile, 1, Nothing)
  If obj Is Nothing Then Exit Function
  objList.AddSubNode obj
 Case token_linenumber
  With New LineNumberStatement
   .SetToken
   If Not .Register(objFunction.SymbolTable) Then Exit Function
   objList.AddSubNode .This
  End With
  If Not GetNextToken(objFile, g_tToken) Then Exit Function
  bParsedLineNumber = True
 Case keyword_goto
  If Not GetNextToken(objFile, g_tToken) Then Exit Function
  Select Case g_tToken.nType
  Case token_id, token_decnum
   With New GoToStatement
    .SetToken
    objList.AddSubNode .This
   End With
   If Not GetNextToken(objFile, g_tToken) Then Exit Function
  Case Else
   PrintError "Identifier or decimal number expected"
   Exit Function
  End Select
 Case Else
  PrintError "Syntax error or currently unsupported"
  Exit Function
 End Select
 '///
 If nEndType = node_ifstat Then
  'if statement, etc.
  Select Case g_tToken.nType
  Case token_colon
   If Not GetNextToken(objFile, g_tToken) Then Exit Function
  Case token_crlf, token_eof, keyword_else
   Exit Do
  Case Else
   PrintError "':' or 'Else' or end of line expected"
   Exit Function
  End Select
 ElseIf Not bParsedLineNumber Then
  If Not ParseEndOfLine(objFile) Then Exit Function
 End If
Loop
'///
Set ParseStatementList = objList
End Function

Friend Function ParseMake(ByVal objFile As ISource, ByVal nFlags As Long, ByVal LHS As VariableNode) As IASTNode
Dim obj As New MakeStatementNode
Dim RHS As IASTNode
'///
If LHS Is Nothing Then
 Set LHS = ParseVariable(objFile)
 If LHS Is Nothing Then Exit Function
End If
'///
If g_tToken.nType <> token_equal Then
 PrintError "'=' expected"
 Exit Function
End If
If Not GetNextToken(objFile, g_tToken) Then Exit Function
'///
Set RHS = ParseExpression(objFile)
If RHS Is Nothing Then Exit Function
'///
obj.Create nFlags, LHS, RHS
Set ParseMake = obj
End Function

'<ifstat>:{if}<exp>{then}<ifstatlist>@1@|{if}<exp>{then}<ifstatlist>{else}<ifstatlist>@1@|{if}<exp>{then}{else}<ifstatlist>@1@
'<ifblock>:{if}<exp>{then}<br><statlist><elseifblock>*<elseblock>?{end}{if}
'<elseifblock>:({elseif}|{else}{if})<exp>{then}<br><statlist>
'<elseblock>:|{else}<br><statlist>
Friend Function ParseIf(ByVal objFile As ISource, ByVal objFunction As FunctionNode) As IASTNode
Dim obj As New IfStatementNode
Dim objCondition As IASTNode
Dim objThen As StatementListNode
Dim objElse As StatementListNode
Dim nFlags As Long
'///
Set objCondition = ParseExpression(objFile)
If objCondition Is Nothing Then Exit Function
'///
If g_tToken.nType <> keyword_then Then
 PrintError "'Then' expected"
 Exit Function
End If
If Not GetNextToken(objFile, g_tToken) Then Exit Function
'///now check it's single-line statement or blocked statement
Select Case g_tToken.nType
Case token_eof
 PrintError "Unexpected end of file"
 Exit Function
Case token_colon, token_crlf
 '///blocked statement
 If Not GetNextToken(objFile, g_tToken) Then Exit Function
 '///
 Do
  Set objThen = ParseStatementList(objFile, node_ifblock, , objFunction, nFlags)
  If objThen Is Nothing Then Exit Function
  '///
  obj.AddThenBlock objCondition, objThen
  '///check if it's End If
  If nFlags = 1 Then Exit Do
  '///
  Select Case g_tToken.nType
  Case keyword_else
   If Not GetNextToken(objFile, g_tToken) Then Exit Function
   If Not ParseEndOfLine(objFile) Then Exit Function
   Set objElse = ParseStatementList(objFile, node_elseblock, , objFunction, nFlags)
   If objElse Is Nothing Then Exit Function
   If nFlags <> 1 Then
    PrintError "'End If' expected"
    Exit Function
   End If
   Set obj.ElseBlock = objElse
   Exit Do
  Case keyword_elseif
   If Not GetNextToken(objFile, g_tToken) Then Exit Function
   Set objCondition = ParseExpression(objFile)
   If objCondition Is Nothing Then Exit Function
   '///
   If g_tToken.nType <> keyword_then Then
    PrintError "'Then' expected"
    Exit Function
   End If
   If Not GetNextToken(objFile, g_tToken) Then Exit Function
   '///
   If Not ParseEndOfLine(objFile) Then Exit Function
  Case Else
   PrintError "'Else' or 'ElseIf' expected"
   Exit Function
  End Select
 Loop
Case Else
 '///single-line statement
 Set objThen = ParseStatementList(objFile, node_ifstat, , objFunction)
 If objThen Is Nothing Then Exit Function
 '///
 Select Case g_tToken.nType
 Case token_crlf, token_eof
 Case keyword_else
  If Not GetNextToken(objFile, g_tToken) Then Exit Function
  Set objElse = ParseStatementList(objFile, node_ifstat, , objFunction)
  If objElse Is Nothing Then Exit Function
 Case Else
  PrintError "':' or 'Else' or end of line expected"
  Exit Function
 End Select
 '///over
 obj.AddThenBlock objCondition, objThen
 Set obj.ElseBlock = objElse
End Select
'///
Set ParseIf = obj
End Function

'{for}(<var>{=}<exp>{to}<exp>|{each}<var>{in}<exp>)<br><statlist>{next}<var>
Friend Function ParseFor(ByVal objFile As ISource, ByVal objFunction As FunctionNode) As IASTNode
Dim obj As New ForStatementNode
Dim objVariable As VariableNode
Dim objVar2 As VariableNode
Dim objStart As IASTNode, objEnd As IASTNode, objStep As IASTNode
Dim objStatement As StatementListNode
Dim i As Long
'///
If g_tToken.nType = keyword_each Then
 'TODO:
 PrintError "Currently 'For Each' is unsupported"
 Exit Function
Else
 Set objVariable = ParseVariable(objFile)
 If objVariable Is Nothing Then Exit Function
 '///check if current variable is conflict with previous for variable
 For i = 1 To m_nForStackSize
  If objVariable.IsSameVariable(m_objForStack(i).Variable) Then
   PrintError "For variable is already used in previous 'For' statement"
   Exit Function
  End If
 Next i
 '///
 If g_tToken.nType <> token_equal Then
  PrintError "'=' expected"
  Exit Function
 End If
 If Not GetNextToken(objFile, g_tToken) Then Exit Function
 '///
 Set objStart = ParseExpression(objFile)
 If objStart Is Nothing Then Exit Function
 '///
 If g_tToken.nType <> keyword_to Then
  PrintError "'To' expected"
  Exit Function
 End If
 If Not GetNextToken(objFile, g_tToken) Then Exit Function
 '///
 Set objEnd = ParseExpression(objFile)
 If objEnd Is Nothing Then Exit Function
 '///
 If g_tToken.nType = keyword_step Then
  If Not GetNextToken(objFile, g_tToken) Then Exit Function
  Set objStep = ParseExpression(objFile)
  If objStep Is Nothing Then Exit Function
 End If
 '///
 If Not ParseEndOfLine(objFile) Then Exit Function
 '///add current statement to the stack
 Set obj.Variable = objVariable
 Set obj.StartValue = objStart
 Set obj.EndValue = objEnd
 Set obj.StepValue = objStep
 m_nForStackSize = m_nForStackSize + 1
 ReDim Preserve m_objForStack(1 To m_nForStackSize)
 Set m_objForStack(m_nForStackSize) = obj
 '///
 Set objStatement = ParseStatementList(objFile, node_forstat)
 '///remove current statement from the stack
 Set m_objForStack(m_nForStackSize) = Nothing
 m_nForStackSize = m_nForStackSize - 1
 '///
 If objStatement Is Nothing Then Exit Function
 '///
 Select Case g_tToken.nType
 Case token_crlf, token_colon, token_eof, token_comma 'token_comma ??
 Case Else
  Set objVar2 = ParseVariable(objFile)
  If objVar2 Is Nothing Then Exit Function
  'check objVariable and objVar2 is the same
  If Not objVariable.IsSameVariable(objVar2) Then
   PrintError "For variable mismatch"
   Exit Function
  End If
 End Select
 '///over
 Set obj.StatementList = objStatement
End If
'///
Set ParseFor = obj
End Function

'{do}(({while}|{until})<exp>)?<br><statlist>{loop}(({while}|{until})<exp>)?
'{while}<exp><br><statlist>{wend}
Friend Function ParseDo(ByVal objFile As ISource, ByVal objFunction As FunctionNode) As IASTNode
Dim obj As New DoStatementNode
Dim obj1 As IASTNode
Dim objStatement As StatementListNode
Dim nFlags As Long
'///
Select Case g_tToken.nType
Case keyword_do
 obj.NodeType = node_dostat
 If Not GetNextToken(objFile, g_tToken) Then Exit Function
 '///
 Select Case g_tToken.nType
 Case keyword_while
  If Not GetNextToken(objFile, g_tToken) Then Exit Function
  Set obj1 = ParseExpression(objFile)
  If obj1 Is Nothing Then Exit Function
  Set obj.DoCondition = obj1
 Case keyword_until
  nFlags = nFlags Or 1&
  If Not GetNextToken(objFile, g_tToken) Then Exit Function
  Set obj1 = ParseExpression(objFile)
  If obj1 Is Nothing Then Exit Function
  Set obj.DoCondition = obj1
 Case token_colon, token_crlf, token_eof
 Case Else
  PrintError "'While' or 'Until' or end of line expected"
  Exit Function
 End Select
 '///
 If Not ParseEndOfLine(objFile) Then Exit Function
 '///
 Set objStatement = ParseStatementList(objFile, obj.NodeType, , objFunction)
 If objStatement Is Nothing Then Exit Function
 '///
 Select Case g_tToken.nType
 Case keyword_while
  If Not GetNextToken(objFile, g_tToken) Then Exit Function
  Set obj1 = ParseExpression(objFile)
  If obj1 Is Nothing Then Exit Function
  Set obj.LoopCondition = obj1
 Case keyword_until
  nFlags = nFlags Or 2&
  If Not GetNextToken(objFile, g_tToken) Then Exit Function
  Set obj1 = ParseExpression(objFile)
  If obj1 Is Nothing Then Exit Function
  Set obj.LoopCondition = obj1
 Case token_colon, token_crlf, token_eof
 Case Else
  PrintError "'While' or 'Until' or end of line expected"
  Exit Function
 End Select
 '///
Case keyword_while
 obj.NodeType = node_whilestat
 If Not GetNextToken(objFile, g_tToken) Then Exit Function
 '///
 Set obj1 = ParseExpression(objFile)
 If obj1 Is Nothing Then Exit Function
 Set obj.DoCondition = obj1
 '///
 If Not ParseEndOfLine(objFile) Then Exit Function
 '///
 Set objStatement = ParseStatementList(objFile, obj.NodeType, , objFunction)
 If objStatement Is Nothing Then Exit Function
 '///
Case Else
 PrintError "'Do' or 'While' expected"
 Exit Function
End Select
'///over
obj.Flags = nFlags
Set obj.StatementList = objStatement
Set ParseDo = obj
End Function

'<selectstat>:{select}{case}<exp><br><selectblock>*{end}{select}
Friend Function ParseSelectCase(ByVal objFile As ISource, ByVal objFunction As FunctionNode) As IASTNode
Dim obj As New SelectStatementNode
Dim objCondition As IASTNode
Dim objBlock As SelectBlockNode
'///
If g_tToken.nType <> keyword_case Then
 PrintError "'Case' expected"
 Exit Function
End If
If Not GetNextToken(objFile, g_tToken) Then Exit Function
'///
Set objCondition = ParseExpression(objFile)
If objCondition Is Nothing Then Exit Function
Set obj.Condition = objCondition
'///eat whitespaces
If Not ParseEndOfLine(objFile) Then Exit Function
Do
 Select Case g_tToken.nType
 Case token_crlf, token_colon
  If Not GetNextToken(objFile, g_tToken) Then Exit Function
 Case token_eof
  PrintError "Unexpected end of file"
  Exit Function
 Case Else
  Exit Do
 End Select
Loop
'///
Do Until g_tToken.nType = keyword_end
 Set objBlock = ParseSelectBlock(objFile, objFunction)
 If objBlock Is Nothing Then Exit Function
 obj.AddBlock objBlock
 '///
 If objBlock.ConditionCount = 0 Then Exit Do
Loop
'///
If g_tToken.nType <> keyword_end Then
 PrintError "'End Select' expected"
 Exit Function
End If
If Not GetNextToken(objFile, g_tToken) Then Exit Function
If g_tToken.nType <> keyword_select Then
 PrintError "'End Select' expected"
 Exit Function
End If
If Not GetNextToken(objFile, g_tToken) Then Exit Function
'///
Set ParseSelectCase = obj
End Function

'<selectblock>:{case}<selectconditionlist><br><statlist>
'<selectconditionlist>:{else}|(<selectcondition>{,})*<selectcondition>
'<selectcondition>:<exp>|<exp>{to}<exp>|{is}({=}|{<>}|{>}|{<}|{>=}|{<=})<exp>
Friend Function ParseSelectBlock(ByVal objFile As ISource, ByVal objFunction As FunctionNode) As SelectBlockNode
Dim obj As New SelectBlockNode
Dim obj0 As IASTNode
Dim obj1 As IASTNode
Dim nType As enumTokenType
Dim objStatement As StatementListNode
'///
If g_tToken.nType <> keyword_case Then
 PrintError "'Case' expected"
 Exit Function
End If
If Not GetNextToken(objFile, g_tToken) Then Exit Function
'///
If g_tToken.nType = keyword_else Then
 If Not GetNextToken(objFile, g_tToken) Then Exit Function
Else
 Do
  Set obj1 = Nothing
  nType = token_equal
  If g_tToken.nType = keyword_is Then
   If Not GetNextToken(objFile, g_tToken) Then Exit Function
   Select Case g_tToken.nType
   Case token_equal, token_ne, token_gt, token_lt, token_ge, token_le
    nType = g_tToken.nType
    If Not GetNextToken(objFile, g_tToken) Then Exit Function
    Set obj0 = ParseExpression(objFile)
    If obj0 Is Nothing Then Exit Function
   Case Else
    PrintError "'=' or '<>' or '>' or '<' or '>=' or '<=' expected"
    Exit Function
   End Select
  Else
   Set obj0 = ParseExpression(objFile)
   If obj0 Is Nothing Then Exit Function
   If g_tToken.nType = keyword_to Then
    nType = keyword_to
    If Not GetNextToken(objFile, g_tToken) Then Exit Function
    Set obj1 = ParseExpression(objFile)
    If obj1 Is Nothing Then Exit Function
   End If
  End If
  '///
  obj.AddCondition nType, obj0, obj1
  '///
  Select Case g_tToken.nType
  Case token_comma
   If Not GetNextToken(objFile, g_tToken) Then Exit Function
  Case token_crlf, token_colon, token_eof
   Exit Do
  Case Else
   PrintError "',' or end of line expected"
   Exit Function
  End Select
 Loop
End If
'///
If Not ParseEndOfLine(objFile) Then Exit Function
'///
Set objStatement = ParseStatementList(objFile, node_selectblock, , objFunction)
If objStatement Is Nothing Then Exit Function
Set obj.StatementList = objStatement
'///
Set ParseSelectBlock = obj
End Function

Friend Function ParseCall(ByVal objFile As ISource, ByVal nFlags As Long, ByVal LHS As VariableNode) As IASTNode
Dim obj As New CallStatementNode
Dim objArgList As ArgListNode
'///
If LHS Is Nothing Then
 Set LHS = ParseVariable(objFile)
 If LHS Is Nothing Then Exit Function
End If
'///
If nFlags = 0 Then
 Select Case g_tToken.nType
 Case token_eof, token_colon, token_crlf, keyword_else
 Case Else
  Set objArgList = ParseArgListNode(objFile)
  If objArgList Is Nothing Then Exit Function
  LHS.AddArgListToLastSubNode objArgList
 End Select
End If
'///
obj.Create LHS
Set ParseCall = obj
End Function

Friend Function ParseConst(ByVal objFile As ISource, ByVal nFlags As Long, Optional ByVal objFunction As FunctionNode) As IASTNode
Dim obj As New DimListNode
Dim objNode As New DimNode
Dim objDataType As DataTypeNode
Dim objValue As IASTNode
'///
nFlags = nFlags Or 4&
'///
Do
 Set objNode = Nothing
 '///
 If g_tToken.nType <> token_id Then
  PrintError "Identifier expected"
  Exit Function
 End If
 objNode.SetToken
 If Not GetNextToken(objFile, g_tToken) Then Exit Function
 '///
 If g_tToken.nType = keyword_as Then
  If Not GetNextToken(objFile, g_tToken) Then Exit Function
  Set objDataType = ParseDataType(objFile)
  If objDataType Is Nothing Then Exit Function
  Set objNode.DataType = objDataType
 Else
  Set objDataType = Nothing
 End If
 '///
 If g_tToken.nType <> token_equal Then
  PrintError "'=' expected"
  Exit Function
 End If
 If Not GetNextToken(objFile, g_tToken) Then Exit Function
 '///
 Set objValue = ParseExpression(objFile)
 If objValue Is Nothing Then Exit Function
 Set objNode.InitValue = objValue
 '///
 objNode.DimType = nFlags
 obj.AddSubNode objNode
 '///
 If g_tToken.nType <> token_comma Then Exit Do
 If Not GetNextToken(objFile, g_tToken) Then Exit Function
Loop
'///
If objFunction Is Nothing Then
 If Not obj.Register(m_objTable) Then Exit Function
Else
 If Not obj.Register(objFunction.SymbolTable) Then Exit Function
End If
'///
Set ParseConst = obj
End Function

Friend Function ParseDim(ByVal objFile As ISource, ByVal nFlags As Long, Optional ByVal objFunction As FunctionNode) As IASTNode
Dim obj As New DimListNode
Dim objNode As New DimNode
'///
Dim obj0 As IASTNode, obj1 As IASTNode
Dim objDataType As DataTypeNode
Dim nType As Long
'///
Do
 Set objNode = Nothing
 nType = nFlags And &H2F0&
 '///
 If g_tToken.nType = keyword_withevents Then
  nType = nType Or &H100&
  If Not GetNextToken(objFile, g_tToken) Then Exit Function
 End If
 '///
 If g_tToken.nType <> token_id Then
  PrintError "Identifier expected"
  Exit Function
 End If
 objNode.SetToken
 If Not GetNextToken(objFile, g_tToken) Then Exit Function
 '///
 If g_tToken.nType = token_lbracket Then
  nType = nType Or &H400&
  If Not GetNextToken(objFile, g_tToken) Then Exit Function
  If g_tToken.nType = token_rbracket Then
   If Not GetNextToken(objFile, g_tToken) Then Exit Function
  Else
   '///get array dimensions
   Do
    Set obj0 = ParseExpression(objFile)
    If obj0 Is Nothing Then Exit Function
    If g_tToken.nType = keyword_to Then
     If Not GetNextToken(objFile, g_tToken) Then Exit Function
     Set obj1 = ParseExpression(objFile)
     If obj1 Is Nothing Then Exit Function
     objNode.AddDimension obj0, obj1
    Else
     objNode.AddDimension Nothing, obj0
    End If
    '///
    Select Case g_tToken.nType
    Case token_comma
     If Not GetNextToken(objFile, g_tToken) Then Exit Function
    Case token_rbracket
     If Not GetNextToken(objFile, g_tToken) Then Exit Function
     Exit Do
    Case Else
     PrintError "',' or ')' expected"
     Exit Function
    End Select
   Loop
  End If
 End If
 '///
 If g_tToken.nType = keyword_as Then
  If Not GetNextToken(objFile, g_tToken) Then Exit Function
  '///
  If g_tToken.nType = keyword_new Then
   nType = nType Or &H800&
   If Not GetNextToken(objFile, g_tToken) Then Exit Function
  End If
  '///
  Set objDataType = ParseDataType(objFile)
  If objDataType Is Nothing Then Exit Function
  Set objNode.DataType = objDataType
 Else
  PrintWarning "Data type unspecified; use default type 'Variant'"
  objNode.SetDataTypeFromString "Variant"
 End If
 '///
 objNode.DimType = nType
 obj.AddSubNode objNode
 '///
 If g_tToken.nType <> token_comma Then Exit Do
 If Not GetNextToken(objFile, g_tToken) Then Exit Function
Loop
'///
If objFunction Is Nothing Then
 If Not obj.Register(m_objTable) Then Exit Function
Else
 If Not obj.Register(objFunction.SymbolTable) Then Exit Function
End If
'///
Set ParseDim = obj
End Function

Friend Function ParseFunction(ByVal objFile As ISource, ByVal nFlags As Long) As IASTNode
Dim obj As New FunctionNode
Dim objArgList As ArgDeclareListNode
Dim objReturnType As DataTypeNode
Dim nReturnFlags As Long
Dim objStatement As StatementListNode
'///
obj.FuncType = nFlags
'///
If g_tToken.nType <> token_id Then
 PrintError "Identifier expected"
 Exit Function
End If
obj.SetName g_tToken
If Not GetNextToken(objFile, g_tToken) Then Exit Function
'///
If nFlags And &H100& Then
 If g_tToken.nType <> keyword_lib Then
  PrintError "'Lib' expected"
  Exit Function
 End If
 If Not GetNextToken(objFile, g_tToken) Then Exit Function
 If g_tToken.nType <> token_string Then
  PrintError "String constant expected"
  Exit Function
 End If
 obj.SetLib g_tToken
 If Not GetNextToken(objFile, g_tToken) Then Exit Function
 If g_tToken.nType = keyword_alias Then
  If Not GetNextToken(objFile, g_tToken) Then Exit Function
  If g_tToken.nType <> token_string Then
   PrintError "String constant expected"
   Exit Function
  End If
  obj.SetAlias g_tToken
  If Not GetNextToken(objFile, g_tToken) Then Exit Function
 End If
End If
'///
If g_tToken.nType <> token_lbracket Then
 PrintError "'(' expected"
 Exit Function
End If
If Not GetNextToken(objFile, g_tToken) Then Exit Function
If g_tToken.nType <> token_rbracket Then
 Set objArgList = ParseArgDeclareList(objFile)
 If objArgList Is Nothing Then Exit Function
 If g_tToken.nType <> token_rbracket Then
  PrintError "')' expected"
  Exit Function
 End If
 obj.SetArgList objArgList
End If
If Not GetNextToken(objFile, g_tToken) Then Exit Function
'///
If g_tToken.nType = keyword_as Then
 Select Case nFlags And &HF&
 Case 0, 3, 4
  PrintError "'Sub' or 'Property Let' or 'Property Set' doesn't have return value"
  Exit Function
 End Select
 If Not GetNextToken(objFile, g_tToken) Then Exit Function
 Set objReturnType = ParseDataType(objFile)
 If objReturnType Is Nothing Then Exit Function
 If g_tToken.nType = token_lbracket Then
  If Not GetNextToken(objFile, g_tToken) Then Exit Function
  If g_tToken.nType <> token_rbracket Then
   PrintError "')' expected"
   Exit Function
  End If
  If Not GetNextToken(objFile, g_tToken) Then Exit Function
  nReturnFlags = &H400&
 End If
 '///
 obj.SetReturnType objReturnType, nReturnFlags
Else
 Select Case nFlags And &HF&
 Case 1, 2
  PrintWarning "Data type unspecified; use default type 'Variant'"
  obj.SetReturnType Nothing, 0
 End Select
End If
'///
If Not obj.Register(m_objTable) Then Exit Function
'///
If (nFlags And &H100&) = 0 Then
 If Not ParseEndOfLine(objFile) Then Exit Function
 '///
 Erase m_objForStack
 m_nForStackSize = 0
 '///
 Set objStatement = ParseStatementList(objFile, nFlags And &HF&, , obj)
 '///
 Erase m_objForStack
 m_nForStackSize = 0
 '///
 If objStatement Is Nothing Then Exit Function
 obj.SetStatement objStatement
End If
'///
Set ParseFunction = obj
End Function

Friend Function ParseEndOfLine(ByVal objFile As ISource) As Boolean
Select Case g_tToken.nType
Case token_eof
Case token_crlf, token_colon
 If Not GetNextToken(objFile, g_tToken) Then Exit Function
Case Else
 PrintError "End of line expected"
 Exit Function
End Select
ParseEndOfLine = True
End Function

Friend Function ParseType(ByVal objFile As ISource, ByVal nFlags As Long) As IASTNode
'TODO:
PrintError "Currently unsupported 'Type'"
End Function

Friend Function ParseEnum(ByVal objFile As ISource, ByVal nFlags As Long) As IASTNode
'TODO:
PrintError "Currently unsupported 'Enum'"
End Function

Friend Function ParseDataType(ByVal objFile As ISource) As DataTypeNode
Dim obj As New DataTypeNode
Dim obj1 As IASTNode
'///
Do
 If g_tToken.nType <> token_id Then
  PrintError "Identifier expected"
  Exit Function
 End If
 obj.AddNameToken
 If Not GetNextToken(objFile, g_tToken) Then Exit Function
 If g_tToken.nType <> token_dot Then Exit Do
 If Not GetNextToken(objFile, g_tToken) Then Exit Function
Loop
'///
If g_tToken.nType = token_asterisk Then
 If Not GetNextToken(objFile, g_tToken) Then Exit Function
 Set obj1 = ParseExpression(objFile)
 If obj1 Is Nothing Then Exit Function
 Set obj.StringSize = obj1
End If
'///
Set ParseDataType = obj
End Function

Friend Function ParseArgDeclareList(ByVal objFile As ISource) As ArgDeclareListNode
Dim obj As New ArgDeclareListNode
Dim nFlags As Long
Dim tName As typeToken
Dim objDataType As DataTypeNode
Dim objDefault As IASTNode
'///
Do
 '///
 nFlags = 0
 Set objDataType = Nothing
 Set objDefault = Nothing
 '///
 Select Case g_tToken.nType
 Case keyword_optional
  nFlags = nFlags Or 4&
  If Not GetNextToken(objFile, g_tToken) Then Exit Function
 Case keyword_paramarray
  nFlags = nFlags Or 8&
  If Not GetNextToken(objFile, g_tToken) Then Exit Function
 End Select
 '///
 Select Case g_tToken.nType
 Case keyword_byval
  nFlags = nFlags Or 1&
  If Not GetNextToken(objFile, g_tToken) Then Exit Function
 Case keyword_byref
  If Not GetNextToken(objFile, g_tToken) Then Exit Function
 End Select
 '///
 If g_tToken.nType <> token_id Then
  PrintError "Identifier expected"
  Exit Function
 End If
 tName = g_tToken
 If Not GetNextToken(objFile, g_tToken) Then Exit Function
 '///
 If g_tToken.nType = token_lbracket Then
  If nFlags And 1& Then
   PrintError "ByVal argument can't be an array"
   Exit Function
  End If
  If nFlags And 4& Then
   PrintError "Optional argument can't be an array"
   Exit Function
  End If
  If Not GetNextToken(objFile, g_tToken) Then Exit Function
  If g_tToken.nType <> token_rbracket Then
   PrintError "')' expected"
   Exit Function
  End If
  If Not GetNextToken(objFile, g_tToken) Then Exit Function
  nFlags = nFlags Or 2&
 ElseIf nFlags And 8& Then
  PrintError "ParamArray argument must be an array"
  Exit Function
 End If
 '///
 If g_tToken.nType = keyword_as Then
  If Not GetNextToken(objFile, g_tToken) Then Exit Function
  Set objDataType = ParseDataType(objFile)
  If objDataType Is Nothing Then Exit Function
 Else
  PrintWarning "Data type unspecified; use default type 'Variant'"
 End If
 '///
 If g_tToken.nType = token_equal Then
  If Not GetNextToken(objFile, g_tToken) Then Exit Function
  If nFlags And &H4& Then
   Set objDefault = ParseExpression(objFile)
   If objDefault Is Nothing Then Exit Function
  Else
   PrintError "Can't set default argument value on non-optional argument"
  End If
 End If
 '///
 obj.AddSubNode tName, nFlags, objDataType, objDefault
 '///
 If g_tToken.nType <> token_comma Then Exit Do
 If Not GetNextToken(objFile, g_tToken) Then Exit Function
Loop
'///
Set ParseArgDeclareList = obj
End Function

Friend Function ParseSkip(ByVal objFile As ISource) As Boolean
Do
 If g_tToken.nType = 0 Then Exit Do
 '///
 Select Case g_tToken.nType
 Case token_eof, token_crlf, token_colon
  ParseSkip = True
  Exit Function
 Case Else
  If Not GetNextToken(objFile, g_tToken) Then Exit Function
 End Select
Loop
End Function

Friend Function ParseExpression(ByVal objFile As ISource, Optional ByVal ExprPrec As Long, Optional ByVal LHS As IASTNode) As IASTNode
Dim TokPrec As Long, NextPrec As Long
Dim i As Long
Dim RHS As IASTNode
'//
If LHS Is Nothing Then
 Set LHS = ParseUnaryOp(objFile)
 If LHS Is Nothing Then Exit Function
End If
'// If this is a binop, find its precedence.
Do
 TokPrec = GetBinaryTokPrecedence
 '// If this is a binop that binds at least as tightly as the current binop,
 '// consume it, otherwise we are done.
 If TokPrec < ExprPrec Then
  Set ParseExpression = LHS
  Exit Function
 End If
 '// Okay, we know this is a binop.
 i = g_tToken.nType
 If Not GetNextToken(objFile, g_tToken) Then Exit Function '// eat binop
 '// Parse the unary expression after the binary operator.
 Set RHS = ParseUnaryOp(objFile)
 If RHS Is Nothing Then Exit Function
 '// If BinOp binds less tightly with RHS than the operator after RHS, let
 '// the pending operator take RHS as its LHS.
 NextPrec = GetBinaryTokPrecedence
 If TokPrec < NextPrec Then
  Set RHS = ParseExpression(objFile, TokPrec + 1, RHS)
  If RHS Is Nothing Then Exit Function
 End If
 '// Merge LHS/RHS.
 With New ExpressionNode
  .SetTypeAndSubNode i, 2, LHS, RHS
  Set LHS = .This
 End With
Loop  '// loop around to the top of the while loop.
End Function

Friend Function ParseUnaryOp(ByVal objFile As ISource) As IASTNode
Dim i As Long, TokPrec As Long
Dim obj As IASTNode
TokPrec = GetUnaryTokPrecedence
If TokPrec > 0 Then
 '// If this is a unary operator, read it.
 i = g_tToken.nType
 If Not GetNextToken(objFile, g_tToken) Then Exit Function
 Set obj = ParseExpression(objFile, TokPrec)
 If obj Is Nothing Then Exit Function
 With New ExpressionNode
  .SetTypeAndSubNode i, 1, obj
  Set ParseUnaryOp = .This
 End With
Else
 '// If the current token is not an operator, it must be a primary expr.
 Set ParseUnaryOp = ParseExpressionTerm(objFile)
End If
End Function

Friend Function GetBinaryTokPrecedence() As Long
Select Case g_tToken.nType
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

Friend Function GetUnaryTokPrecedence() As Long
Select Case g_tToken.nType
Case keyword_not
 GetUnaryTokPrecedence = 35
Case token_plus, token_minus
 GetUnaryTokPrecedence = 95
Case Else
 GetUnaryTokPrecedence = -1
End Select
End Function

Friend Function ParseExpressionTerm(ByVal objFile As ISource) As IASTNode
Dim obj As IASTNode
'///
Select Case g_tToken.nType
Case token_lbracket '"("<expression>")"
 If Not GetNextToken(objFile, g_tToken) Then Exit Function
 '///
 Set obj = ParseExpression(objFile)
 If obj Is Nothing Then Exit Function
 '///
 If g_tToken.nType <> token_rbracket Then
  PrintError "')' expected"
  Exit Function
 End If
 '///
 If Not GetNextToken(objFile, g_tToken) Then Exit Function
 '///
 'force ByVal (create a copy of variable)
 With New ExpressionNode
  .SetTypeAndSubNode token_plus, 1, obj
  Set ParseExpressionTerm = .This
 End With
Case token_id, token_dot
 Set obj = ParseVariable(objFile)
 If obj Is Nothing Then Exit Function
 '///
 Set ParseExpressionTerm = obj
Case token_decnum, token_hexnum, token_octnum, token_floatnum, token_string, token_currencynum, token_datenum, keyword_true, keyword_false
 With New ConstNode
  .SetToken
  Set obj = .This
 End With
 '///
 If Not GetNextToken(objFile, g_tToken) Then Exit Function
 Set ParseExpressionTerm = obj
Case Else
 PrintError "Identifier or constant expected"
End Select
End Function

'<var>:(<array_or_func>|<membervar>)<membervar>*
Friend Function ParseVariable(ByVal objFile As ISource) As VariableNode
Dim obj As New VariableNode
Dim obj1 As ArrayOrFuncNode
Dim nFlags2 As Long
'///
Select Case g_tToken.nType
Case token_dot
 nFlags2 = 1
 If Not GetNextToken(objFile, g_tToken) Then Exit Function
End Select
'///
Do
 Set obj1 = ParseArrayOrFuncNode(objFile, nFlags2)
 If obj1 Is Nothing Then Exit Function
 obj.AddSubNode obj1
 '///
 If g_tToken.nType <> token_dot Then Exit Do
 If g_tToken.nFlags And 1& Then Exit Do 'dirty workaround of parsing "a.b" and "a .b"
 '///
 nFlags2 = 1
 If Not GetNextToken(objFile, g_tToken) Then Exit Function
Loop
'///
Set ParseVariable = obj
End Function

'<array_or_func>:<id>({(}<arglist>{)})+
Friend Function ParseArrayOrFuncNode(ByVal objFile As ISource, ByVal nFlags2 As Long) As ArrayOrFuncNode
Dim obj As New ArrayOrFuncNode
Dim obj1 As ArgListNode
'///
Select Case g_tToken.nType
Case token_id
 obj.SetToken nFlags2
 If Not GetNextToken(objFile, g_tToken) Then Exit Function
Case Else
 PrintError "Identifier expected"
 Exit Function
End Select
'///
Do While g_tToken.nType = token_lbracket And (g_tToken.nFlags And 1&) = 0 'dirty workaround of "a (b),(c)"
 If Not GetNextToken(objFile, g_tToken) Then Exit Function
 If g_tToken.nType = token_rbracket Then
  obj.AddSubNode Nothing
 Else
  Set obj1 = ParseArgListNode(objFile)
  If obj1 Is Nothing Then Exit Function
  If g_tToken.nType <> token_rbracket Then
   PrintError "')' expected"
   Exit Function
  End If
  obj.AddSubNode obj1
 End If
 If Not GetNextToken(objFile, g_tToken) Then Exit Function
Loop
'///
Set ParseArrayOrFuncNode = obj
End Function

'<arglist>:(({byval}?<exp>)?{,})*{byval}?<exp> 'ByVal???
Friend Function ParseArgListNode(ByVal objFile As ISource) As ArgListNode
Dim obj As New ArgListNode
Dim obj1 As IASTNode
Dim nFlags As Long
'///
Do
 If g_tToken.nType = token_comma Then
  obj.AddSubNode Nothing, 0
  If Not GetNextToken(objFile, g_tToken) Then Exit Function
 Else
  nFlags = 0
  If g_tToken.nType = keyword_byval Then
   nFlags = nFlags Or 1&
   If Not GetNextToken(objFile, g_tToken) Then Exit Function
  End If
  Set obj1 = ParseExpression(objFile)
  If obj1 Is Nothing Then Exit Function
  obj.AddSubNode obj1, nFlags
  If g_tToken.nType <> token_comma Then Exit Do
  If Not GetNextToken(objFile, g_tToken) Then Exit Function
 End If
Loop
'///
Set ParseArgListNode = obj
End Function

Friend Function Verify(ByVal nStep As enumASTNodeVerifyStep) As Boolean
Dim v As Variant
Dim obj As IASTNode
Dim objContext As New clsVerifyContext
'///
objContext.Phase = nStep
Set objContext.Module = Me
'///
For Each v In m_objNodes
 Set obj = v
 If Not obj.Verify(objContext) Then Exit Function
Next v
'///
Verify = True
End Function

Friend Sub Codegen()
Dim v As Variant
Dim obj As IASTNode
Dim objContext As New clsVerifyContext
'///
Set objContext.Module = Me
'///
m_objTable.Codegen objContext
'///
For Each v In m_objNodes
 Set obj = v
 obj.Codegen objContext, 0, 0, 0, 0
Next v
End Sub

Friend Property Get This() As clsSrcParser
Set This = Me
End Property

Private Sub Class_Initialize()
Set m_objTable = New clsSymbolTable
End Sub
