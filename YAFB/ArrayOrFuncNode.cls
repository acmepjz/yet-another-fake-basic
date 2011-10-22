VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ArrayOrFuncNode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'<array_or_func>:<id>({(}<arglist>{)})+

Implements IASTNode

Private m_t As typeToken
'nFlags2:
'1=id begin with "."

Private m_objArgList() As ArgListNode '1-based, can be Nothing
Private m_nCount As Long

'================================ LLVM ================================

Private m_objVariable As DimNode
Private m_objFunction As FunctionNode

Private m_hConst As Long

Friend Function IsSameVariable(ByVal obj As ArrayOrFuncNode) As Boolean
Dim t As typeToken
'///
obj.GetToken t
If m_t.sValue <> t.sValue Then Exit Function
If m_t.nFlags2 <> t.nFlags2 Then Exit Function
'///
If m_nCount <> obj.ArgListCount Then Exit Function
'///
'TODO:
If m_nCount > 0 Then Exit Function
'///
IsSameVariable = True
End Function

Friend Property Get ArgListCount() As Long
ArgListCount = m_nCount
End Property

Private Function IASTNode_GetProperty(ByVal nProp As enumASTNodeProperty) As Long
Dim obj As IASTNode
'///
Select Case nProp
Case action_const_codegen
 If m_hConst = 0 Then
  Set obj = m_objVariable
  m_hConst = obj.GetProperty(action_const_codegen)
 End If
 IASTNode_GetProperty = m_hConst
End Select
End Function

'TODO:
Private Function IASTNode_GetType(nFlags As Long) As clsTypeNode
Dim obj As IASTNode
If Not m_objVariable Is Nothing Then
 Set obj = m_objVariable
 Set IASTNode_GetType = obj.GetType(nFlags)
 Exit Function
End If
If Not m_objFunction Is Nothing Then
 Set obj = m_objFunction
 Set IASTNode_GetType = obj.GetType(nFlags)
 Exit Function
End If
End Function

Private Property Get IASTNode_NodeType() As enumASTNodeType
IASTNode_NodeType = node_array_or_func
End Property

Friend Property Get This() As IASTNode
Set This = Me
End Property

Friend Sub GetToken(ByRef t As typeToken)
t = m_t
End Sub

'1=id begin with "."
Friend Sub SetToken(ByVal nFlags2 As Long)
m_t = g_tToken
m_t.nFlags2 = nFlags2
End Sub

Friend Sub AddSubNode(ByVal obj As ArgListNode)
m_nCount = m_nCount + 1
ReDim Preserve m_objArgList(1 To m_nCount)
Set m_objArgList(m_nCount) = obj
End Sub

'TODO:
Friend Function VerifyEx(ByVal objContext As clsVerifyContext, Optional ByVal bIsLValue As Boolean, Optional ByVal bIsCall As Boolean) As Boolean
Dim obj As IASTNode, obj1 As IASTNode
Dim objVar1 As VariableNode
'///
Dim objVariable As DimNode
Dim objFunction As FunctionNode
Dim objArgDeclareList As ArgDeclareListNode
Dim nArgCount As Long
Dim objSrcType As clsTypeNode, nSrcFlags As Long
Dim objDestType As clsTypeNode, nDestFlags As Long
Dim i As Long, m As Long
'///
Select Case objContext.Phase
Case verify_const
 If m_t.nFlags2 And 1& Then
  PrintError "Member variable in constant is unsupported", -1, -1
  Exit Function
 End If
 If m_nCount > 0 Then
  PrintError "Array and function in constant is unsupported", -1, -1
  Exit Function
 End If
 '///look up the symbol table
 Set objVariable = objContext.LookUpVariableTable(m_t.sValue)
 If objVariable Is Nothing Then
  PrintError "Variable '" + m_t.sValue + "' not found", m_t.nLine, m_t.nColumn
  Exit Function
 ElseIf (objVariable.DimType And &HF&) <> 4 Then
  PrintError "Variable '" + m_t.sValue + "' is not a constant", m_t.nLine, m_t.nColumn
  Exit Function
 Else
  '///register dependency
  g_objConstDAG.AddEdge objVariable, Me
 End If
 '///over
 Set m_objVariable = objVariable
Case verify_all
 If m_t.nFlags2 And 1& Then
  PrintError "Currently member variable and function is unsupported", -1, -1
  Exit Function
 End If
 If m_nCount > 1 Then
  PrintError "Currently Array() etc. is unsupported", -1, -1
  Exit Function
 End If
 '///
 For i = 1 To m_nCount
  Set obj = m_objArgList(i)
  If Not obj.Verify(objContext) Then Exit Function
 Next i
 '///look up the symbol table
 '///check array (??) TEST ONLY!!! TODO:
 If Not bIsCall And m_nCount = 0 Then
  Set objVariable = objContext.LookUpVariableTable(m_t.sValue)
 End If
 If objVariable Is Nothing Then
  If bIsLValue Then
   'TODO:property let/set
   PrintError "Variable '" + m_t.sValue + "' not found", m_t.nLine, m_t.nColumn
   Exit Function
  End If
  Set objFunction = objContext.LookUpFunctionTable(m_t.sValue)
  If objFunction Is Nothing Then
   If bIsCall Then
    PrintError "Function '" + m_t.sValue + "' not found", m_t.nLine, m_t.nColumn
   Else
    PrintError "Variable or function '" + m_t.sValue + "' not found", m_t.nLine, m_t.nColumn
   End If
   Exit Function
  End If
 End If
 '///
 Set m_objVariable = objVariable
 Set m_objFunction = objFunction
 'TODO:array,etc.
 '///if it's function call then verify argument
 If Not objFunction Is Nothing Then
  If Not bIsCall Then
   Select Case objFunction.FuncType And &HF&
   Case 1, 2
   Case Else
    PrintError "'" + m_t.sValue + "' is not Function or Property Get, it doesn't have return value", m_t.nLine, m_t.nColumn
    Exit Function
   End Select
  End If
  '///
  Set objArgDeclareList = objFunction.ArgumentDeclareList
  '///
  If m_nCount > 0 Then
   If Not m_objArgList(1) Is Nothing Then
    nArgCount = m_objArgList(1).ArgumentCount
   End If
  End If
  '///
  If objArgDeclareList Is Nothing Then
   If nArgCount > 0 Then
    PrintError "'" + m_t.sValue + "' doesn't accept arguments", m_t.nLine, m_t.nColumn
    Exit Function
   End If
  Else
   m = objArgDeclareList.ArgumentCount
   For i = 1 To nArgCount
    If i > m Then
     PrintError "'" + m_t.sValue + "' doesn't accept more than " + CStr(m) + " arguments", m_t.nLine, m_t.nColumn
     Exit Function
    End If
    '///verify
    Set obj = m_objArgList(1).ArgumentNode(i)
    If obj Is Nothing Then
     If (objArgDeclareList.ArgumentFlags(i) And 4&) = 0 Then
      PrintError "Argument " + CStr(i) + " of function '" + m_t.sValue + "' is not optional", m_t.nLine, m_t.nColumn
      Exit Function
     End If
    Else
     Set objSrcType = obj.GetType(nSrcFlags)
     Set obj1 = objArgDeclareList.ArgumentDeclaration(i).DataType
     Set objDestType = obj1.GetType(nDestFlags)
     If g_objTypeMgr.CheckTypeConversion(objSrcType, nSrcFlags, objDestType, nDestFlags) = 0 Then
      PrintError "Argument " + CStr(i) + " of function '" + m_t.sValue + "' type mismatch: trying convert '" + _
      objSrcType.Name + "' to '" + objDestType.Name + "'", m_t.nLine, m_t.nColumn
      Exit Function
     End If
     '///
     If (objArgDeclareList.ArgumentFlags(i) And 1&) = 0 Then 'declaration is ByRef xxx
      If m_objArgList(1).ArgumentFlags(i) And 1& Then 'ByVal xxx
       If g_objTypeMgr.CheckTypeConversion(objSrcType, nSrcFlags, g_objIntrinsicDataTypes(vbLong), 0) = 0 Then
        PrintError "Argument " + CStr(i) + " of function '" + m_t.sValue + "' type mismatch: trying convert '" + _
        objSrcType.Name + "' to 'Long'", m_t.nLine, m_t.nColumn
        Exit Function
       End If
      ElseIf objDestType.DataType <> vbEmpty Then 'Any
       On Error Resume Next
       Set objVar1 = Nothing
       Set objVar1 = obj
       On Error GoTo 0
       If Not objVar1 Is Nothing Then
        If objVar1.IsLValue Then
         If Not objSrcType Is objDestType Then
          PrintError "ByRef argument " + CStr(i) + " of function '" + m_t.sValue + "' type mismatch: '" + _
          objSrcType.Name + "' and '" + objDestType.Name + "'", m_t.nLine, m_t.nColumn
          Exit Function
         End If
        End If
       End If
      End If
     End If
     '///
    End If
   Next i
   '///
  End If
 End If
 '///TODO:etc.
End Select
'///
VerifyEx = True
End Function

'TODO:
Friend Function IsLValue() As Boolean
If Not m_objVariable Is Nothing Then
 If (m_objVariable.DimType And &HF&) = 0 Then
  IsLValue = True
 End If
End If
End Function

'TODO:
Friend Function CodegenEx(ByVal objContext As clsVerifyContext, Optional ByVal bIsLValue As Boolean, Optional ByVal bIsCall As Boolean) As Long
Dim hArguments() As Long
'///
Dim obj As IASTNode, obj1 As IASTNode
Dim objVar1 As VariableNode
'///
Dim objArgDeclareList As ArgDeclareListNode
Dim nArgCount As Long
Dim objSrcType As clsTypeNode, nSrcFlags As Long
Dim objDestType As clsTypeNode, nDestFlags As Long
Dim h As Long, i As Long, j As Long, m As Long
Dim nTag As Long
'///
nTag = node_array_or_func
'///variable?
If Not m_objVariable Is Nothing Then
 Select Case m_objVariable.DimType And &HF&
 Case 0 'dim
  If bIsLValue Then
   CodegenEx = m_objVariable.VariableHandle
  Else
   CodegenEx = LLVMBuildLoad(g_hBuilder, m_objVariable.VariableHandle, StrPtr(StrConv("LoadTemp", vbFromUnicode)))
  End If
 Case 4 'const
  If bIsLValue Then
   Panic
  Else
   CodegenEx = m_objVariable.ConstHandle
  End If
 End Select
 '///
 Exit Function
End If
'///function?
If Not m_objFunction Is Nothing Then
 Set objArgDeclareList = m_objFunction.ArgumentDeclareList
 '///
 If m_nCount > 0 Then
  If Not m_objArgList(1) Is Nothing Then
   nArgCount = m_objArgList(1).ArgumentCount
  End If
 End If
 '///
 If objArgDeclareList Is Nothing Then
  ReDim hArguments(1 To 1)
 Else
  m = objArgDeclareList.ArgumentCount
  ReDim hArguments(1 To m + 1)
  For i = 1 To m
   '///get argument node
   If i <= nArgCount Then
    Set obj = m_objArgList(1).ArgumentNode(i)
   Else
    Set obj = Nothing
   End If
   '///default argument
   If obj Is Nothing Then
    Set obj = objArgDeclareList.ArgumentDefaultValue(i)
    'TODO: if obj is nothing then ...
   End If
   '///
   Set objSrcType = obj.GetType(nSrcFlags)
   Set obj1 = objArgDeclareList.ArgumentDeclaration(i).DataType
   Set objDestType = obj1.GetType(nDestFlags)
   '///
   If (objArgDeclareList.ArgumentFlags(i) And 1&) = 0 Then 'declaration is ByRef xxx
    If m_objArgList(1).ArgumentFlags(i) And 1& Then 'ByVal xxx
     h = obj.Codegen(objContext, 0, 0, 0, 0)
     h = g_objTypeMgr.CodegenTypeConversion(h, objSrcType, nSrcFlags, g_objIntrinsicDataTypes(vbLong), 0, False)
     hArguments(i) = LLVMBuildIntToPtr(g_hBuilder, h, LLVMPointerType(objDestType.Handle, 0), StrPtr(StrConv("ByValTemp", vbFromUnicode)))
    Else
     On Error Resume Next
     Set objVar1 = Nothing
     Set objVar1 = obj
     On Error GoTo 0
     If Not objVar1 Is Nothing Then
      If Not objVar1.IsLValue Then Set objVar1 = Nothing
     End If
     '///
     If objVar1 Is Nothing Then
      '///create a temp variable TODO: should reuse temp variable because LLVM can't optimize it
      h = obj.Codegen(objContext, 0, 0, 0, 0)
      If objDestType.DataType = vbEmpty Then 'Any
       j = objContext.CurrentFunction.GetTempVariable(objContext, objSrcType, nTag)
       LLVMBuildStore g_hBuilder, h, j
       hArguments(i) = LLVMBuildPointerCast(g_hBuilder, j, LLVMPointerType(objDestType.Handle, 0), StrPtr(StrConv("ByRefTemp", vbFromUnicode)))
      Else
       h = g_objTypeMgr.CodegenTypeConversion(h, objSrcType, nSrcFlags, objDestType, nDestFlags, False)
       j = objContext.CurrentFunction.GetTempVariable(objContext, objDestType, nTag)
       LLVMBuildStore g_hBuilder, h, j
       hArguments(i) = j
      End If
     Else
      '///get address of variable
      h = objVar1.CodegenEx(objContext, True)
      If objSrcType Is objDestType Then
       hArguments(i) = h
      Else
       hArguments(i) = LLVMBuildPointerCast(g_hBuilder, h, LLVMPointerType(objDestType.Handle, 0), StrPtr(StrConv("ByRefTemp", vbFromUnicode)))
      End If
     End If
    End If
   Else
    h = obj.Codegen(objContext, 0, 0, 0, 0)
    hArguments(i) = g_objTypeMgr.CodegenTypeConversion(h, objSrcType, nSrcFlags, objDestType, nDestFlags, False)
   End If
   '///
  Next i
 End If
 '///over, call the function
 h = m_objFunction.FunctionHandle
 Select Case m_objFunction.FuncType And &HF&
 Case 1, 2
  i = LLVMBuildCall(g_hBuilder, h, hArguments(1), m, StrPtr(StrConv("CallTemp", vbFromUnicode)))
 Case Else
  'weird LLVM behavior: if it's Sub then it can't have a name
  i = LLVMBuildCall(g_hBuilder, h, hArguments(1), m, StrPtr(vbNullChar))
 End Select
 LLVMSetInstructionCallConv i, LLVMGetFunctionCallConv(h)
 'objContext.CurrentFunction.ResetTempVariable objContext, nTag 'wrong code, e.g. CopyMemory xxx,ByVal StrPtr(s),xxx
 CodegenEx = i
End If
End Function

Private Function IASTNode_Verify(ByVal objContext As clsVerifyContext) As Boolean
IASTNode_Verify = VerifyEx(objContext, False)
End Function

Private Function IASTNode_Codegen(ByVal objContext As clsVerifyContext, ByVal nParam1 As Long, ByVal nParam2 As Long, ByVal nParam3 As Long, ByVal nParam4 As Long) As Long
IASTNode_Codegen = CodegenEx(objContext, False)
End Function

