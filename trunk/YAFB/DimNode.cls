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

Private m_objDataType As DataTypeNode 'must not be Nothing

'================================ LLVM ================================

Private m_hType As Long 'not exactly
Private m_hVariable As Long

Friend Function CodegenType() As Long
Dim i As Long
'///
If m_hType Then
 CodegenType = m_hType
 Exit Function
End If
'///
If m_nType And &H400& Then
 'TODO:array
 abort
Else
 i = m_objDataType.GetDataType.Handle
 If m_nType And &H1000& Then
  m_hType = LLVMPointerType(i, 0)
 Else
  m_hType = i
 End If
 CodegenType = m_hType
End If
End Function

'TODO:array, byref, etc.
'TODO:default value
Friend Function CodegenVariable(ByVal objContext As clsVerifyContext) As Long
If m_hVariable Then
 CodegenVariable = m_hVariable
 Exit Function
End If
If Not objContext.CurrentFunction Is Nothing Then
 '///local variable
 m_hVariable = LLVMBuildAlloca(g_hBuilder, m_hType, ByVal StrPtr(StrConv(m_t.sValue, vbFromUnicode)))
ElseIf objContext.Module.FileType = 0 Then
 '///global variable
 m_hVariable = LLVMAddGlobal(g_hModule, m_hType, ByVal StrPtr(StrConv(m_t.sValue, vbFromUnicode)))
Else
 'TODO: class member variable
End If
'///
CodegenVariable = m_hVariable
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
'TODO:
End Function

'TODO:flags
Private Function IASTNode_GetType(nFlags As Long) As clsTypeNode
Dim obj As IASTNode
Set obj = m_objDataType
Set IASTNode_GetType = obj.GetType(nFlags)
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

Friend Sub SetVariantType()
Dim t As typeToken
Set m_objDataType = New DataTypeNode
t.nType = token_id
t.sValue = "Variant"
t.nLine = -1
t.nColumn = -1
m_objDataType.AddNameTokenEx t
End Sub

Friend Function Register(ByVal objTable As clsSymbolTable) As Boolean
On Error Resume Next
'///
Err.Clear
objTable.VariableTable.Add Me, m_t.sValue
If Err.Number Then
 PrintError "Ambiguous name detected: '" + m_t.sValue + "'", m_t.nLine, m_t.nColumn
 Exit Function
End If
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
'TODO:
Select Case objContext.Phase
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
'///
IASTNode_Verify = True
End Function
