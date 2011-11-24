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

Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)

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
Private m_nArrayDimSize() As Currency '1 to m_nDimemsionCount
Private m_nDimemsionCount As Long

Private m_objDataType As DataTypeNode 'must not be Nothing (??)

Private m_objInitValue As IASTNode 'optional (??) only when 'const' ??

Private m_objVariable As VariableNode 'optional (??) only when 'redim', 'redim preserve' and 'erase' ??

'================================ LLVM ================================

Private m_hType As Long 'not exactly
Private m_hVariable As Long

Private m_hConst As Long

Private m_nArrayElementCount As Currency
Private m_hConst_ArrayElementCount As Long

Private m_bArrayInHeap As Boolean

Friend Property Get VariableObject() As VariableNode
Set VariableObject = m_objVariable
End Property

Friend Property Set VariableObject(ByVal obj As VariableNode)
Set m_objVariable = obj
End Property

Friend Property Get ArrayDimensionSize(ByVal nIndex As Long) As Currency
ArrayDimensionSize = m_nArrayDimSize(nIndex)
End Property

Friend Property Get DimensionCount() As Long
DimensionCount = m_nDimemsionCount
End Property

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
   If Not CheckConstant(m_hConst) Then Exit Function
   '///verify data type
   Set objSrcType = m_objInitValue.GetType(nSrcFlags)
   If m_objDataType Is Nothing Then
    '///automatic data detect
    If (objSrcType.Flags And &H10&) <> 0 And nSrcFlags = 0 Then
     SetDataTypeFromString objSrcType.Name
    Else
     PrintError "Const variable can't have data type '" + objSrcType.NameEx(nSrcFlags) + "'", m_t.nLine, m_t.nColumn
     m_hConst = 0
     Exit Function
    End If
   End If
   m_objDataType.FindDataType Nothing '???
   '///check type conversion
   Set obj = m_objDataType
   Set objDestType = obj.GetType(nDestFlags)
   Select Case g_objTypeMgr.CheckTypeConversion(objSrcType, nSrcFlags, objDestType, nDestFlags)
   Case 0
    PrintError "Type mismatch: can't convert data type '" + objSrcType.NameEx(nSrcFlags) + "' to '" + objDestType.NameEx(nDestFlags) + "'", -1, -1
    m_hConst = 0
    Exit Function
   End Select
   '///codegen
   m_hConst = g_objTypeMgr.CodegenTypeConversion(Nothing, m_hConst, objSrcType, nSrcFlags, objDestType, nDestFlags, True)
   If Not CheckConstant(m_hConst) Then Exit Function
   '///over
  End If
  IASTNode_GetProperty = m_hConst
 End Select
End Select
End Function

Friend Sub CodegenDefaultDestructor()
Dim h(3) As Long
'///
Select Case m_nType And &HF&
Case 0 'dim
 If (m_nType And &HF0&) = &HF0& Then
  'function argument. TODO: ByVal String, reference count, etc.
 Else
  If m_nType And &H400& Then 'array
   If m_nDimemsionCount > 0 Then 'fixed-size array
    '///generate destructor of elements in array
    h(0) = LLVMBuildPointerCast(g_hBuilder, m_hVariable, LLVMPointerType(m_objDataType.GetDataType.Handle, 0), StrPtr(StrConv("ByRefTemp", vbFromUnicode)))
    h(1) = m_hConst_ArrayElementCount
    LLVMBuildCall g_hBuilder, m_objDataType.GetDataType.GetDefaultVectorDestructorFunction, h(0), 2, StrPtr(vbNullChar)
    '///
    If m_bArrayInHeap Then
     LLVMBuildFree g_hBuilder, m_hVariable
    End If
   Else 'dynamic array
    h(0) = m_hVariable
    LLVMBuildCall g_hBuilder, m_objDataType.GetDataType.GetDefaultSafeArrayDestructorFunction, h(0), 1, StrPtr(vbNullChar)
   End If
  Else
   'TODO:etc.
   m_objDataType.GetDataType.CodegenDefaultDestructor m_hVariable
  End If
 End If
Case Else
 'TODO: things other than 'Dim', etc.
End Select
End Sub

Friend Function CodegenType() As Long
Dim i As Long, j As Long
Select Case m_nType And &HF&
Case 0 'dim
 If m_hType Then
  CodegenType = m_hType
  Exit Function
 End If
 '///
 If m_nType And &H400& Then
  '///array type (test only) TODO:
  If m_nDimemsionCount > 0 Then
   i = m_objDataType.GetDataType.Handle
   '@#^$$#%%$&#$????? TODO:
   CopyMemory j, m_nArrayElementCount, 4&
   i = LLVMArrayType(i, j)
  Else
   i = LLVMPointerType(g_hTypeSafeArray, 0)
  End If
 Else
  i = m_objDataType.GetDataType.Handle
 End If
 '///ByRef ?
 If m_nType And &H1000& Then
  m_hType = LLVMPointerType(i, 0)
 Else
  m_hType = i
 End If
 CodegenType = m_hType
Case Else
 'TODO: things other than 'Dim', etc.
End Select
End Function

'TODO:array, byref, etc.
'TODO:default value
Friend Function CodegenVariable(ByVal objContext As clsVerifyContext) As Long
Dim h(7) As Long
Dim m As Currency
'///
Select Case m_nType And &HF&
Case 0 'dim
 If m_hVariable Then
  CodegenVariable = m_hVariable
  Exit Function
 End If
 If Not objContext.CurrentFunction Is Nothing Then
  '///local variable
  If m_nType And &H400& Then 'array
   If m_nDimemsionCount > 0 Then 'fixed-size array
    '///check if array is too big to be in stack
    m = LLVMABISizeOfType(g_hTargetData, m_hType)
    If m > 6.5535@ Then
     PrintWarning "Size of local array '" + m_t.sValue + "' is bigger than 64KB, allocated in heap instead", m_t.nLine, m_t.nColumn
     m_bArrayInHeap = True
    End If
    '///
    If m_bArrayInHeap Then
     m_hVariable = LLVMBuildMalloc(g_hBuilder, m_hType, ByVal StrPtr(StrConv(m_t.sValue, vbFromUnicode)))
    Else
     m_hVariable = LLVMBuildAlloca(g_hBuilder, m_hType, ByVal StrPtr(StrConv(m_t.sValue, vbFromUnicode)))
    End If
    '///
    'If m <= 0.0064@ Then 'LLVM generates extremely stupid code :(
    ' LLVMBuildStore g_hBuilder, LLVMConstNull(m_hType), m_hVariable
    'Else 'so we should call 'memset' instead...
     h(0) = LLVMBuildPointerCast(g_hBuilder, m_hVariable, LLVMPointerType(LLVMInt8Type, 0), ByVal StrPtr(StrConv("ByRefTemp", vbFromUnicode)))
     h(1) = LLVMConstNull(LLVMInt8Type)
     h(2) = LLVMBuildIntCast(g_hBuilder, LLVMSizeOf(m_hType), g_hTypeIntPtr_t, ByVal StrPtr(StrConv("SizeTemp", vbFromUnicode)))
     h(3) = LLVMConstNull(LLVMInt32Type) 'align=unknown
     h(4) = LLVMConstNull(LLVMInt1Type) 'isVolatile=False
     LLVMBuildCall g_hBuilder, RuntimeLibraryGetFunction(internal_llvm_memset), h(0), 5, StrPtr(vbNullChar)
    'End If
   Else 'dynamic array
    m_hVariable = LLVMBuildAlloca(g_hBuilder, m_hType, ByVal StrPtr(StrConv(m_t.sValue, vbFromUnicode)))
    LLVMBuildStore g_hBuilder, LLVMConstPointerNull(m_hType), m_hVariable
   End If
  Else
   m_hVariable = LLVMBuildAlloca(g_hBuilder, m_hType, ByVal StrPtr(StrConv(m_t.sValue, vbFromUnicode)))
   m_objDataType.GetDataType.CodegenDefaultConstructor m_hVariable
  End If
 ElseIf objContext.Module.FileType = 0 Then
  '///global variable
  'TODO: LLVM generates extremely stupid code. Should use BSS section.
  If m_nType And &H400& Then 'array
   If m_nDimemsionCount > 0 Then
    m_hVariable = LLVMAddGlobal(g_hModule, m_hType, ByVal StrPtr(StrConv(m_t.sValue, vbFromUnicode)))
    LLVMSetInitializer m_hVariable, LLVMConstNull(m_hType)
   Else
    m_hVariable = LLVMAddGlobal(g_hBuilder, m_hType, ByVal StrPtr(StrConv(m_t.sValue, vbFromUnicode)))
    LLVMSetInitializer m_hVariable, LLVMConstPointerNull(m_hType)
   End If
  Else
   m_hVariable = LLVMAddGlobal(g_hModule, m_hType, ByVal StrPtr(StrConv(m_t.sValue, vbFromUnicode)))
   LLVMSetInitializer m_hVariable, m_objDataType.GetDataType.CodegenDefaultValue
  End If
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

Friend Sub CodegenErase(ByVal objContext As clsVerifyContext)
Dim h(7) As Long
'///
If m_nDimemsionCount > 0 Then 'fixed-size array
 '///generate destructor of elements in array
 h(0) = LLVMBuildPointerCast(g_hBuilder, m_hVariable, LLVMPointerType(m_objDataType.GetDataType.Handle, 0), StrPtr(StrConv("ByRefTemp", vbFromUnicode)))
 h(1) = m_hConst_ArrayElementCount
 LLVMBuildCall g_hBuilder, m_objDataType.GetDataType.GetDefaultVectorDestructorFunction, h(0), 2, StrPtr(vbNullChar)
 '///then memset
 h(0) = LLVMBuildPointerCast(g_hBuilder, m_hVariable, LLVMPointerType(LLVMInt8Type, 0), ByVal StrPtr(StrConv("ByRefTemp", vbFromUnicode)))
 h(1) = LLVMConstNull(LLVMInt8Type)
 h(2) = LLVMBuildIntCast(g_hBuilder, LLVMSizeOf(m_hType), g_hTypeIntPtr_t, ByVal StrPtr(StrConv("SizeTemp", vbFromUnicode)))
 h(3) = LLVMConstNull(LLVMInt32Type) 'align=unknown
 h(4) = LLVMConstNull(LLVMInt1Type) 'isVolatile=False
 LLVMBuildCall g_hBuilder, RuntimeLibraryGetFunction(internal_llvm_memset), h(0), 5, StrPtr(vbNullChar)
Else 'dynamic array
 h(0) = m_hVariable
 LLVMBuildCall g_hBuilder, m_objDataType.GetDataType.GetDefaultSafeArrayDestructorFunction, h(0), 1, StrPtr(vbNullChar)
End If
End Sub

Private Function IASTNode_Codegen(ByVal objContext As clsVerifyContext, ByVal nParam1 As Long, ByVal nParam2 As Long, ByVal nParam3 As Long, ByVal nParam4 As Long) As Long
Dim obj As IASTNode
Dim objType As clsTypeNode, nFlags As Long
Dim objDim As DimNode
'///
Select Case m_nType And &HF&
Case 3 'erase
 '///check type
 Set obj = m_objVariable
 Set objType = obj.GetType(nFlags)
 If nFlags And &H400& Then
  'OK it's an array
  Set objDim = m_objVariable.GetDimNode
  If objDim Is Nothing Then
   'TODO:
   PrintPanic "Can't find DimNode!!! Something in unimplemented...", -1, -1
  Else
   objDim.CodegenErase objContext
  End If
 Else 'err, TODO:
  PrintPanic "Passed a non-array data type '" + objType.NameEx(nFlags) + "' to 'Erase' statement", -1, -1
 End If
End Select
End Function

'TODO:flags
Private Function IASTNode_GetType(nFlags As Long) As clsTypeNode
Dim obj As IASTNode
'///
nFlags = 0
'///
Select Case m_nType And &HF&
Case 0, 4 'dim, const (??)
 Set obj = m_objDataType
 Set IASTNode_GetType = obj.GetType(0) '?
 nFlags = m_nType And &H400&
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

Friend Property Get ArrayBound(ByVal nLBoundOrUBound As Long, ByVal nIndex As Long) As IASTNode
Set ArrayBound = m_objDimension(nLBoundOrUBound, nIndex)
End Property

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
Dim obj As IASTNode, obj1 As IASTNode
Dim objType As clsTypeNode, nFlags As Long
Dim nType As VbVarType
Dim hValue As Long
Dim n1 As Currency, n2 As Currency
Dim nArrayElementCount As Currency
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
  '///array dimension
  If m_nType And &H400& Then
   If m_nDimemsionCount = 0 Then
'    PrintError "Currently dynamic array is unsupported", -1, -1
'    Exit Function
   ElseIf m_nDimemsionCount >= 65536 Then
    PrintError "Array dimension must be smaller than 65536", -1, -1
    Exit Function
   End If
   '///check array
   If m_nDimemsionCount > 0 Then
    ReDim m_nArrayDimSize(1 To m_nDimemsionCount)
    nArrayElementCount = 0.0001@
    For i = 1 To m_nDimemsionCount
     '///check LBound
     Set obj = m_objDimension(0, i)
     If obj Is Nothing Then
      Set obj = objContext.Module.DefaultBase
     Else
      hValue = obj.GetProperty(action_const_codegen)
      If Not CheckConstant(hValue) Then Exit Function
      '///check type
      nFlags = 0
      Set objType = obj.GetType(nFlags)
      If g_objTypeMgr.CheckTypeConversion(objType, nFlags, g_objIntrinsicDataTypes(vbLong), 0) = 0 Then
       PrintError "Array lower bound " + CStr(i) + " of variable '" + m_t.sValue + "' type mismatch: trying convert '" + _
       objType.NameEx(nFlags) + "' to 'Long'", m_t.nLine, m_t.nColumn
       Exit Function
      End If
      hValue = g_objTypeMgr.CodegenTypeConversion(objContext, hValue, objType, nFlags, g_objIntrinsicDataTypes(vbLong), 0, True)
      If Not CheckConstant(hValue) Then Exit Function
      '///
      With New ConstNode
       .CreateIndirect g_objIntrinsicDataTypes(vbLong), hValue
       Set obj = .This
      End With
     End If
     '///check UBound
     Set obj1 = m_objDimension(1, i)
     If obj1 Is Nothing Then
      PrintError "Must specify array upper bound " + CStr(i) + " of variable '" + m_t.sValue + "'", m_t.nLine, m_t.nColumn
      Exit Function
     Else
      hValue = obj1.GetProperty(action_const_codegen)
      If Not CheckConstant(hValue) Then Exit Function
      '///check type
      nFlags = 0
      Set objType = obj1.GetType(nFlags)
      If g_objTypeMgr.CheckTypeConversion(objType, nFlags, g_objIntrinsicDataTypes(vbLong), 0) = 0 Then
       PrintError "Array upper bound " + CStr(i) + " of variable '" + m_t.sValue + "' type mismatch: trying convert '" + _
       objType.NameEx(nFlags) + "' to 'Long'", m_t.nLine, m_t.nColumn
       Exit Function
      End If
      hValue = g_objTypeMgr.CodegenTypeConversion(objContext, hValue, objType, nFlags, g_objIntrinsicDataTypes(vbLong), 0, True)
      If Not CheckConstant(hValue) Then Exit Function
      '///
      With New ConstNode
       .CreateIndirect g_objIntrinsicDataTypes(vbLong), hValue
       Set obj1 = .This
      End With
     End If
     '///check LBound<=UBound
     n1 = LLVMConstIntGetSExtValue(obj.GetProperty(action_const_codegen))
     n2 = LLVMConstIntGetSExtValue(obj1.GetProperty(action_const_codegen))
     If n1 > n2 Then
      PrintError "Lower bound '" + CStr(n1 * 10000@) + "' is greater than upper bound '" + CStr(n2 * 10000@) + _
      "' for array subscript " + CStr(i) + " of variable '" + m_t.sValue + "'", m_t.nLine, m_t.nColumn
      Exit Function
     End If
     '///calc array size
     m_nArrayDimSize(i) = n2 - n1 + 0.0001@
     If m_nArrayDimSize(i) >= 429496.7296@ Then
      PrintError "Array size of one dimension must be smaller than 4294967296", m_t.nLine, m_t.nColumn
      Exit Function
     End If
     '///
     nArrayElementCount = UnsignedLongLongMultiply(nArrayElementCount, m_nArrayDimSize(i), nFlags)
     If ((nFlags And 1&) <> 0 And g_nWordSize = 4) Or ((nFlags And 2&) <> 0 And g_nWordSize = 8) Then
      PrintError "Array '" + m_t.sValue + "' too big", m_t.nLine, m_t.nColumn
      Exit Function
     End If
     '///over
     Set m_objDimension(0, i) = obj
     Set m_objDimension(1, i) = obj1
    Next i
    '///XXX TODO:
    If nArrayElementCount >= 429496.7296@ Then
     PrintError "Currently more than 4294967296 elements in fixed-size array is unsupported", m_t.nLine, m_t.nColumn
     Exit Function
    End If
    '///save array size
    m_nArrayElementCount = nArrayElementCount
    m_hConst_ArrayElementCount = LLVMConstInt(g_hTypeIntPtr_t, nArrayElementCount, 0)
   End If
   '///
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
   If (m_objDataType.GetDataType.Flags And &H10&) = 0 Then
    PrintError "Const variable can't have data type '" + m_objDataType.GetDataType.Name + "'", m_t.nLine, m_t.nColumn
    Exit Function
   End If
  End If
  '///verify expression
  If Not m_objInitValue.Verify(objContext) Then Exit Function
  '///register dependency
  g_objConstDAG.AddEdge m_objInitValue, Me
 End Select
Case 3 'erase
 Select Case objContext.Phase
 Case verify_all
  Set obj = m_objVariable
  If Not obj.Verify(objContext) Then Exit Function
  '///check type
  Set objType = obj.GetType(nFlags)
  If nFlags And &H400& Then
   'OK it's an array
  ElseIf objType.DataType = vbVariant Then 'TODO:
   PrintError "Currently Variant as array is unsupport", -1, -1
   Exit Function
  Else 'err
   PrintError "Passed a non-array data type '" + objType.NameEx(nFlags) + "' to 'Erase' statement", -1, -1
   Exit Function
  End If
 End Select
Case 1, 2 'redim and redim preserve
 'TODO:
 Exit Function
Case Else
 'XXX
 Debug.Assert False
 Exit Function
End Select
'///
IASTNode_Verify = True
End Function

Friend Property Get ConstHandle() As Long
ConstHandle = m_hConst
End Property
