VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "VariableNode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'<var>:(<array_or_func>|<membervar>)<membervar>*

Implements IASTNode

Private m_objSubNode() As ArrayOrFuncNode '1-based
Private m_nCount As Long

Private Function IASTNode_GetProperty(ByVal nProp As enumASTNodeProperty) As Long
'
End Function

Friend Property Get SubNodeCount() As Long
SubNodeCount = m_nCount
End Property

Friend Property Get SubNode(ByVal nIndex As Long) As ArrayOrFuncNode
Set SubNode = m_objSubNode(nIndex)
End Property

Friend Function IsSameVariable(ByVal obj As VariableNode) As Boolean
Dim i As Long
'///
If m_nCount <> obj.SubNodeCount Then Exit Function
'///
For i = 1 To m_nCount
 If Not m_objSubNode(i).IsSameVariable(obj.SubNode(i)) Then Exit Function
Next i
'///
IsSameVariable = True
End Function

'TODO:
Private Function IASTNode_GetType(nFlags As Long) As clsTypeNode
Dim obj As IASTNode
Set obj = m_objSubNode(1)
Set IASTNode_GetType = obj.GetType(nFlags)
End Function

Private Property Get IASTNode_NodeType() As enumASTNodeType
IASTNode_NodeType = node_var
End Property

Friend Property Get This() As IASTNode
Set This = Me
End Property

Friend Sub AddSubNode(ByVal obj As ArrayOrFuncNode)
m_nCount = m_nCount + 1
ReDim Preserve m_objSubNode(1 To m_nCount)
Set m_objSubNode(m_nCount) = obj
End Sub

Friend Sub AddArgListToLastSubNode(ByVal obj As ArgListNode)
If m_nCount > 0 Then m_objSubNode(m_nCount).AddSubNode obj
End Sub

'TODO:
Friend Function VerifyEx(ByVal objContext As clsVerifyContext, Optional ByVal bIsLValue As Boolean, Optional ByVal bIsCall As Boolean) As Boolean
If objContext.Phase = verify_all Then
 If m_nCount <> 1 Then
  PrintError "Currently member variable and function is unsupported", -1, -1
  Exit Function
 End If
 If Not m_objSubNode(1).VerifyEx(objContext, bIsLValue, bIsCall) Then Exit Function
End If
'///
VerifyEx = True
End Function

'TODO:
Friend Function IsLValue() As Boolean
If m_nCount = 1 Then
 IsLValue = m_objSubNode(1).IsLValue
End If
End Function

'TODO:
Friend Function CodegenEx(ByVal objContext As clsVerifyContext, Optional ByVal bIsLValue As Boolean, Optional ByVal bIsCall As Boolean) As Long
CodegenEx = m_objSubNode(1).CodegenEx(objContext, bIsLValue, bIsCall)
End Function

Private Function IASTNode_Verify(ByVal objContext As clsVerifyContext) As Boolean
IASTNode_Verify = VerifyEx(objContext)
End Function

Private Function IASTNode_Codegen(ByVal objContext As clsVerifyContext, ByVal nParam1 As Long, ByVal nParam2 As Long, ByVal nParam3 As Long, ByVal nParam4 As Long) As Long
IASTNode_Codegen = CodegenEx(objContext)
End Function

