VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsVerifyContext"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_nPhase As enumASTNodeVerifyStep
Private m_objModule As clsSrcParser
Private m_objFunction As FunctionNode

Private m_objScope() As IASTNode '1-based
Private m_nScopeCount As Long

Friend Sub AddScope(ByVal obj As IASTNode)
m_nScopeCount = m_nScopeCount + 1
ReDim Preserve m_objScope(1 To m_nScopeCount)
Set m_objScope(m_nScopeCount) = obj
End Sub

Friend Sub RemoveScope()
If m_nScopeCount > 0 Then
 Set m_objScope(m_nScopeCount) = Nothing
 m_nScopeCount = m_nScopeCount - 1
End If
End Sub

Friend Sub ClearScope()
Erase m_objScope
m_nScopeCount = 0
End Sub

Friend Property Get Phase() As enumASTNodeVerifyStep
Phase = m_nPhase
End Property

Friend Property Let Phase(ByVal n As enumASTNodeVerifyStep)
m_nPhase = n
End Property

Friend Property Get Module() As clsSrcParser
Set Module = m_objModule
End Property

Friend Property Set Module(ByVal obj As clsSrcParser)
Set m_objModule = obj
End Property

Friend Property Get CurrentFunction() As FunctionNode
Set CurrentFunction = m_objFunction
End Property

Friend Property Set CurrentFunction(ByVal obj As FunctionNode)
Set m_objFunction = obj
End Property

'doesn't copy m_objScope
Friend Sub Clone(ByVal objSrc As clsVerifyContext)
m_nPhase = objSrc.Phase
Set m_objModule = objSrc.Module
Set m_objFunction = objSrc.CurrentFunction
End Sub

Friend Property Get This() As clsVerifyContext
Set This = Me
End Property

Friend Function LookUpVariableTable(ByVal sName As String) As DimNode
On Error Resume Next
Dim obj As DimNode
'///
If Not m_objFunction Is Nothing Then
 Err.Clear
 Set obj = m_objFunction.SymbolTable.VariableTable.Item(sName)
 If Err.Number = 0 Then
  Set LookUpVariableTable = obj
  Exit Function
 End If
End If
'///
If Not m_objModule Is Nothing Then
 Err.Clear
 Set obj = m_objModule.SymbolTable.VariableTable.Item(sName)
 If Err.Number = 0 Then
  Set LookUpVariableTable = obj
  Exit Function
 End If
End If
'///
Err.Clear
Set obj = g_objGlobalTable.VariableTable.Item(sName)
If Err.Number = 0 Then
 Set LookUpVariableTable = obj
 Exit Function
End If
End Function

Friend Function LookUpFunctionTable(ByVal sName As String) As FunctionNode
On Error Resume Next
Dim obj As FunctionNode
'///
If Not m_objModule Is Nothing Then
 Err.Clear
 Set obj = m_objModule.SymbolTable.FunctionTable.Item(sName)
 If Err.Number = 0 Then
  Set LookUpFunctionTable = obj
  Exit Function
 End If
End If
'///
Err.Clear
Set obj = g_objGlobalTable.FunctionTable.Item(sName)
If Err.Number = 0 Then
 Set LookUpFunctionTable = obj
 Exit Function
End If
End Function

Friend Function LookUpTypeTable(ByVal sName As String) As clsTypeNode
On Error Resume Next
Dim obj As clsTypeNode
'///
If Not m_objModule Is Nothing Then
 Err.Clear
 Set obj = m_objModule.SymbolTable.TypeTable.Item(sName)
 If Err.Number = 0 Then
  Set LookUpTypeTable = obj
  Exit Function
 End If
End If
'///
Err.Clear
Set obj = g_objGlobalTable.TypeTable.Item(sName)
If Err.Number = 0 Then
 Set LookUpTypeTable = obj
 Exit Function
End If
End Function
