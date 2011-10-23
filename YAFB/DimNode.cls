VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DimNode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IASTNode

Private m_nType As Long
'0=dim
'1=redim
'2=redim preserve
'3=erase
'4=const (??)
'///
'0=public
'&H10&=private
'&H30&=protected (currently unsupported)
'&HF0&=function argument
'///
'&H100&=withevents
'&H200&=static
'&H400&=array
'&H800&=new
'&H1000&=byref (??)

Private m_t As typeToken

Private m_objDimension() As IASTNode '0 to 1,1 to m_nDimemsionCount, (0,x) can be Nothing, (1,x) can't be Nothing
Private m_nDimemsionCount As Long

Private m_objDataType As DataTypeNode 'must not be Nothing (??)

Private m_objInitValue As IASTNode 'optional (??) only when 'const' ??

'================================ LLVM ================================

Private m_hType As Long 'not exactly
Private m_hVariable As Long

Private m_hConst As Long

Friend Property Get InitValue() As IASTNode
Set InitValue = m_objInitValue
End Property

Friend Property Set InitValue(ByVal obj As IASTNode)
Set m_objInitValue = obj
End Property

Private Function IASTNode_GetProperty(ByVal nProp As enumASTNodeProperty) As Long
Dim obj As IASTNode
Dim objSrcType As clsTypeNode, nSrcFlags As Long
Dim objDestType As clsTypeNode, nDestFlags As Long
'///
Select Case m_nType And &HF&
Case 4 'const
 Select Case nProp
 Case action_const_codegen
  If m_hConst = 0 Then
   m_hConst = m_objInitValue.GetProperty(action_const_codegen)
   If m_hConst = 0 Then Exit Function
   '///verify data type
   Set objSrcType = m_objInitValue.GetType(nSrcFlags)
   If m_objDataType Is Nothing Then
    '///automatic data detect
    Select Case objSrcType.DataType
    Case vbByte, vbInteger, vbLong, vbBoolean, vbSingle, vbDouble, vbCurrency, vbString, vbDate
     SetDataTypeFromString objSrcType.Name
    Case Else
     PrintError "Const variable can't have data type '" + objSrcType.Name + "'", m_t.nLine, m_t.nColumn
     m_hConst = 0
     Exit Function
    End Select
   End If
   m_objDataType.FindDataType Nothing '???
   '///check type conversion
   Set obj = m_objDataType
   Set objDestType = obj.GetType(nDestFlags)
   Select Case g_objTypeMgr.CheckTypeConversion(objSrcType, nSrcFlags, objDestType, nDestFlags)
   Case 0
    PrintError "Type mismatch: can't convert data type '" + objSrcType.Name + "' to '" + objDestType.Name + "'", -1, -1
    m_hConst = 0
    Exit Function
   End Select
   '///codegen
   m_hConst = g_objTypeMgr.CodegenTypeConversion(m_hConst, objSrcType, nSrcFlags, objDestType, nDestFlags, True)
   '///over
  End If
  IASTNode_GetProperty = m_hConst
 End Select
End Select
End Function

Friend Sub CodegenDefaultDestructor(ByVal objContext As clsVerifyContext)
Select Case m_nType And &HF&
Case 0 'dim
 If (m_nType And &HF0&) = &HF0& Then
  'function argument. TODO: ByVal String, reference count, etc.
 Else
  'TODO:array, etc.
  m_objDataType.GetDataType.CodegenDefaultDestructor objContext, m_hVariable
 End If
Case Else
 'TODO: things other than 'Dim', etc.
End Select
End Sub

Friend Function CodegenType() As Long
Dim i As Long
Select Case m_nType And &HF&
Case 0 'dim
 If m_hType Then
  CodegenType = m_hType
  Exit Function
 End If
 '///
 If m_nType And &H400& Then
  'TODO:array
  Panic
 Else
  i = m_objDataType.GetDataType.Handle
  If m_nType And &H1000& Then
   m_hType = LLVMPointerType(i, 0)
  Else
   m_hType = i
  End If
  CodegenType = m_hType
 End If
Case Else
 'TODO: things other than 'Dim', etc.
End Select
End Function

'TODO:array, byref, etc.
'TODO:default value
Friend Function CodegenVariable(ByVal objContext As clsVerifyContext) As Long
Select Case m_nType And &HF&
Case 0 'dim
 If m_hVariable Then
  CodegenVariable = m_hVariable
  Exit Function
 End If
 If Not objContext.CurrentFunction Is Nothing Then
  '///local variable
  m_hVariable = LLVMBuildAlloca(g_hBuilder, m_hType, ByVal StrPtr(StrConv(m_t.sValue, vbFromUnicode)))
  m_objDataType.GetDataType.CodegenDefaultConstructor objContext, m_hVariable
 ElseIf objContext.Module.FileType = 0 Then
  '///global variable
  m_hVariable = LLVMAddGlobal(g_hModule, m_hType, ByVal StrPtr(StrConv(m_t.sValue, vbFromUnicode)))
  '/// ???
  Select Case m_objDataType.GetDataType.DataType
  Case vbByte, vbInteger, vbLong, vbBoolean, vbCurrency
   LLVMSetInitializer m_hVariable, LLVMConstInt(m_hType, 0@, 1)
  Case vbSingle, vbDouble, vbDate
   LLVMSetInitializer m_hVariable, LLVMConstReal(m_hType, 0#)
  End Select
 Else
  'TODO: class member variable
 End If
 '///
 CodegenVariable = m_hVariable
Case Else
 'TODO: things other than 'Dim', etc.
End Select
End Function

Friend Function CodegenArgument(ByVal hFunction As Long, ByVal nIndex As Long) As Long
Dim i As Long
Dim s As String, lp As Long
If m_hVariable Then
 CodegenArgument = m_hVariable
 Exit Function
End If
'///
s = StrConv(m_t.sValue, vbFromUnicode)
lp = StrPtr(s)
'///
i = LLVMGetParam(hFunction, nIndex)
LLVMSetValueName i, lp
If m_nType And &H1000& Then
 m_hVariable = i
Else
 m_hVariable = LLVMBuildAlloca(g_hBuilder, m_hType, lp)
 LLVMBuildStore g_hBuilder, i, m_hVariable
End If
CodegenArgument = m_hVariable
End Function

Friend Property Get TypeHandle() As Long
'not exactly
TypeHandle = m_hType
End Property

Friend Property Get VariableHandle() As Long
VariableHandle = m_hVariable
End Property

Private Function IASTNode_Codegen(ByVal objContext As clsVerifyContext, ByVal nParam1 As Long, ByVal nParam2 As Long, ByVal nParam3 As Long, ByVal nParam4 As Long) As Long
'TODO: things other than 'Dim', etc.
End Function

'TODO:flags
Private Function IASTNode_GetType(nFlags As Long) As clsTypeNode
Dim obj As IASTNode
'///
Select Case m_nType And &HF&
Case 0, 4 'dim, const (??)
 Set obj = m_objDataType
 Set IASTNode_GetType = obj.GetType(nFlags)
Case Else
 'TODO: things other than 'Dim', etc.
End Select
End Function

Private Property Get IASTNode_NodeType() As enumASTNodeType
IASTNode_NodeType = node_dimitem
End Property

Friend Property Get This() As IASTNode
Set This = Me
End Property

Friend Property Get DimType() As Long
DimType = m_nType
End Property

Friend Property Let DimType(ByVal n As Long)
m_nType = n
End Property

Friend Sub SetToken()
m_t = g_tToken
End Sub

Friend Sub SetTokenEx(ByRef t As typeToken)
m_t = t
End Sub

Friend Sub AddDimension(ByVal obj0 As IASTNode, ByVal obj1 As IASTNode)
m_nDimemsionCount = m_nDimemsionCount + 1
ReDim Preserve m_objDimension(1, 1 To m_nDimemsionCount)
Set m_objDimension(0, m_nDimemsionCount) = obj0
Set m_objDimension(1, m_nDimemsionCount) = obj1
End Sub

Friend Property Get DataType() As DataTypeNode
Set DataType = m_objDataType
End Property

Friend Property Set DataType(ByVal obj As DataTypeNode)
Set m_objDataType = obj
End Property

Friend Sub SetDataTypeFromString(ByVal s As String)
Dim t As typeToken
Set m_objDataType = New DataTypeNode
t.nType = token_id
t.sValue = s
t.nLine = -1
t.nColumn = -1
m_objDataType.AddNameTokenEx t
End Sub

Friend Function Register(ByVal objTable As clsSymbolTable) As Boolean
On Error Resume Next
'///
Select Case m_nType And &HF&
Case 0, 4 'dim and const
 Err.Clear
 objTable.VariableTable.Add Me, m_t.sValue
 If Err.Number Then
  PrintError "Ambiguous name detected: '" + m_t.sValue + "'", m_t.nLine, m_t.nColumn
  Exit Function
 End If
End Select
'///
Register = True
End Function

Friend Sub ErrorMustBeOptional()
PrintError "Argument '" + m_t.sValue + "' must be optional because previous argument is optional", m_t.nLine, m_t.nColumn
End Sub

Private Function IASTNode_Verify(ByVal objContext As clsVerifyContext) As Boolean
Dim i As Long
Dim obj As IASTNode
Dim nType As VbVarType
'///
Select Case m_nType And &HF&
Case 0 'dim
 Select Case objContext.Phase
 Case verify_const
  If m_nType And &H400& Then
   For i = 1 To m_nDimemsionCount
    Set obj = m_objDimension(0, i)
    If Not obj Is Nothing Then
     If Not obj.Verify(objContext) Then Exit Function
    End If
    Set obj = m_objDimension(1, i)
    If Not obj Is Nothing Then
     If Not obj.Verify(objContext) Then Exit Function
    End If
   Next i
  End If
 Case verify_dim
  'TODO:array dimension
  If m_nType And &H400& Then
   PrintError "Currently array is unsupported", -1, -1
   Exit Function
 '  For i = 1 To m_nDimemsionCount
 '  Next i
  End If
  '///verify data type
  If m_objDataType Is Nothing Then
   nType = vbVariant
  Else
   Set obj = m_objDataType
   If Not obj.Verify(objContext) Then Exit Function
   nType = m_objDataType.GetDataType.DataType
  End If
  'check "Object" (??)
  If nType <> vbObject Then
   If m_nType And &H100& Then
    PrintError "Non-object variable can't be WithEvents", m_t.nLine, m_t.nColumn
    Exit Function
   End If
   If m_nType And &H800& Then
    PrintError "Non-object variable can't be New", m_t.nLine, m_t.nColumn
    Exit Function
   End If
  End If
  'check "Any"
  If nType = vbEmpty Then
   If (m_nType And &H1DFF&) <> &H10F0& Then
    PrintError "Invalid use of 'Any' data type", m_t.nLine, m_t.nColumn
    Exit Function
   End If
  End If
  '///
  CodegenType
 End Select
Case 4 'const
 Select Case objContext.Phase
 Case verify_const
  '///verify data type
  If Not m_objDataType Is Nothing Then
   Set obj = m_objDataType
   With New clsVerifyContext
    .Clone objContext
    .Phase = verify_dim
    If Not obj.Verify(.This) Then Exit Function
   End With
   Select Case m_objDataType.GetDataType.DataType
   Case vbByte, vbInteger, vbLong, vbBoolean, vbSingle, vbDouble, vbCurrency, vbString, vbDate
   Case Else
    PrintError "Const variable can't have data type '" + m_objDataType.GetDataType.Name + "'", m_t.nLine, m_t.nColumn
    Exit Function
   End Select
  End If
  '///verify expression
  If Not m_objInitValue.Verify(objContext) Then Exit Function
  '///register dependency
  g_objConstDAG.AddEdge m_objInitValue, Me
 End Select
Case Else
 'TODO: things other than 'Dim', etc.
End Select
'///
IASTNode_Verify = True
End Function

Friend Property Get ConstHandle() As Long
ConstHandle = m_hConst
End Property
