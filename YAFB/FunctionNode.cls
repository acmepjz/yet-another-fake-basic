VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FunctionNode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IASTNode

Private m_nType As Long
'0=sub
'1=function
'2=property get
'3=property let
'4=property set
'///
'0=public
'&H10&=private
'&H20&=friend
'&H30&=protected (currently unsupported)
'///
'&H100&=declare
'&H200&=static

Private m_tName As typeToken
Private m_tLib As typeToken
Private m_tAlias As typeToken

Private m_objStatement As StatementListNode 'can be Nothing

Private m_objArgList As ArgDeclareListNode 'can be Nothing
Private m_objReturnVariable As DimNode 'can be Nothing

Private m_objTable As New clsSymbolTable

'================================ LLVM ================================

Private m_hFunctionType As Long
Private m_hFunction As Long
Private m_hFunctionEndBlock As Long

Friend Property Get EndBlockHandle() As Long
EndBlockHandle = m_hFunctionEndBlock
End Property

Friend Sub CodegenFunctionDeclaration(ByVal objContext As clsVerifyContext)
Dim s As String
'///
s = m_tName.sValue
If m_nType And &H100& Then
 'TODO:Lib
 'TODO:Alias
 If m_tAlias.nType Then
  s = m_tAlias.sValue
 End If
End If
m_hFunction = LLVMAddFunction(g_hModule, StrPtr(StrConv(s, vbFromUnicode)), m_hFunctionType)
LLVMSetLinkage m_hFunction, LLVMExternalLinkage '??
LLVMSetFunctionCallConv m_hFunction, LLVMX86StdcallCallConv '?? TODO:cdecl support
End Sub

Friend Property Get ArgumentDeclareList() As ArgDeclareListNode
Set ArgumentDeclareList = m_objArgList
End Property

Private Function IASTNode_Codegen(ByVal objContext As clsVerifyContext, ByVal nParam1 As Long, ByVal nParam2 As Long, ByVal nParam3 As Long, ByVal nParam4 As Long) As Long
Dim obj As IASTNode
Dim objNewContext As New clsVerifyContext
Dim hBlock As Long
'///function body
If (m_nType And &H100&) = 0 Then
 objNewContext.Clone objContext
 Set objNewContext.CurrentFunction = Me
 '///
 hBlock = LLVMAppendBasicBlock(m_hFunction, StrPtr(StrConv("entry", vbFromUnicode)))
 m_hFunctionEndBlock = LLVMAppendBasicBlock(m_hFunction, StrPtr(StrConv("end", vbFromUnicode)))
 LLVMPositionBuilderAtEnd g_hBuilder, hBlock
 '///
 If Not m_objArgList Is Nothing Then
  m_objArgList.CodegenArgument m_hFunction
 End If
 '///
 m_objTable.Codegen objNewContext
 Set obj = m_objStatement
 obj.Codegen objNewContext, 0, 0, 0, 0
 '///
 LLVMBuildBr g_hBuilder, m_hFunctionEndBlock
 LLVMPositionBuilderAtEnd g_hBuilder, m_hFunctionEndBlock
 '///
 'TODO: local variable destructor, etc.
 '///
 Select Case m_nType And &HF&
 Case 1, 2
  LLVMBuildRet g_hBuilder, LLVMBuildLoad(g_hBuilder, m_objReturnVariable.VariableHandle, StrPtr(StrConv("ReturnValue", vbFromUnicode)))
 Case Else
  LLVMBuildRetVoid g_hBuilder
 End Select
End If
End Function

'TODO:property get/let/set
Friend Function Register(ByVal objTable As clsSymbolTable) As Boolean
On Error Resume Next
'///
Err.Clear
objTable.FunctionTable.Add Me, m_tName.sValue
If Err.Number Then
 PrintError "Ambiguous name detected: '" + m_tName.sValue + "'", m_tName.nLine, m_tName.nColumn
 Exit Function
End If
'///
If (m_nType And &H100&) = 0 Then
 If Not m_objReturnVariable Is Nothing Then
  If Not m_objReturnVariable.Register(m_objTable) Then Exit Function
 End If
 If Not m_objArgList Is Nothing Then
  If Not m_objArgList.Register(m_objTable) Then Exit Function
 End If
End If
'///
Register = True
End Function

Friend Property Get SymbolTable() As clsSymbolTable
Set SymbolTable = m_objTable
End Property

Private Function IASTNode_GetType(nFlags As Long) As clsTypeNode
Dim obj As IASTNode
If Not m_objReturnVariable Is Nothing Then
 Set obj = m_objReturnVariable
 Set IASTNode_GetType = obj.GetType(nFlags)
End If
End Function

Private Property Get IASTNode_NodeType() As enumASTNodeType
IASTNode_NodeType = node_funcstat
End Property

Friend Property Get This() As IASTNode
Set This = Me
End Property

Friend Property Get FuncTypeHandle() As Long
FuncTypeHandle = m_hFunctionType
End Property

Friend Property Get FuncType() As Long
FuncType = m_nType
End Property

Friend Property Let FuncType(ByVal n As Long)
m_nType = n
End Property

Friend Sub SetName(ByRef t As typeToken)
m_tName = t
End Sub

Friend Sub SetLib(ByRef t As typeToken)
m_tLib = t
End Sub

Friend Sub SetAlias(ByRef t As typeToken)
m_tAlias = t
End Sub

Friend Sub SetStatement(ByVal obj As StatementListNode)
Set m_objStatement = obj
End Sub

Friend Sub SetArgList(ByVal obj As ArgDeclareListNode)
Set m_objArgList = obj
End Sub

'nFlags: same as DimNode
Friend Sub SetReturnType(ByVal obj As DataTypeNode, ByVal nFlags As Long)
Set m_objReturnVariable = New DimNode
m_objReturnVariable.DimType = nFlags
m_objReturnVariable.SetTokenEx m_tName
If obj Is Nothing Then
 m_objReturnVariable.SetVariantType
Else
 Set m_objReturnVariable.DataType = obj
End If
End Sub

Friend Property Get FunctionHandle() As Long
FunctionHandle = m_hFunction
End Property

Private Function IASTNode_Verify(ByVal objContext As clsVerifyContext) As Boolean
Dim obj As IASTNode
Dim j As Long
'///
If Not m_objArgList Is Nothing Then
 Set obj = m_objArgList
 If Not obj.Verify(objContext) Then Exit Function
End If
'///
If Not m_objReturnVariable Is Nothing Then
 Set obj = m_objReturnVariable
 If Not obj.Verify(objContext) Then Exit Function
End If
'///
If Not m_objStatement Is Nothing Then
 Set obj = m_objStatement
 With New clsVerifyContext
  .Clone objContext
  Set .CurrentFunction = Me
  If Not obj.Verify(.This) Then Exit Function
 End With
End If
'///???
If objContext.Phase = verify_all Then
 Select Case m_nType And &HF&
 Case 1, 2
  j = m_objReturnVariable.DataType.GetDataType.Handle
 Case Else
  j = LLVMVoidType
 End Select
 If m_objArgList Is Nothing Then
  m_hFunctionType = LLVMFunctionType(j, 0&, 0&, 0)
 Else
  m_hFunctionType = m_objArgList.CodegenFunctionType(j)
 End If
End If
'///
IASTNode_Verify = True
End Function