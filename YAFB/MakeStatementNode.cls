VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MakeStatementNode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IASTNode

Private m_nFlags As Long
'0=let
'1=set
'2=lset
'3=rset

Private m_objLHS As VariableNode
Private m_objRHS As IASTNode

Private Function IASTNode_GetProperty(ByVal nProp As enumASTNodeProperty) As Long
'
End Function

'??
Private Function IASTNode_GetType(nFlags As Long) As clsTypeNode
Dim obj As IASTNode
Set obj = m_objLHS
Set IASTNode_GetType = obj.GetType(nFlags)
End Function

Private Property Get IASTNode_NodeType() As enumASTNodeType
IASTNode_NodeType = node_makestat
End Property

Friend Property Get This() As IASTNode
Set This = Me
End Property

Friend Sub Create(ByVal nFlags As Long, ByVal LHS As VariableNode, ByVal RHS As IASTNode)
m_nFlags = nFlags
Set m_objLHS = LHS
Set m_objRHS = RHS
End Sub

Private Function IASTNode_Verify(ByVal objContext As clsVerifyContext) As Boolean
Dim obj As IASTNode
Dim objSrcType As clsTypeNode, nSrcFlags As Long
Dim objDestType As clsTypeNode, nDestFlags As Long
'///
If objContext.Phase = verify_all Then
 If Not m_objLHS.VerifyEx(objContext, True) Then Exit Function
 If Not m_objRHS.Verify(objContext) Then Exit Function
 '///
 Set objSrcType = m_objRHS.GetType(nSrcFlags)
 Set obj = m_objLHS
 Set objDestType = obj.GetType(nDestFlags)
 '///
 Select Case g_objTypeMgr.CheckTypeConversion(objSrcType, nSrcFlags, objDestType, nDestFlags)
 Case 0
  PrintError "Type mismatch: can't convert data type '" + objSrcType.Name + "' to '" + objDestType.Name + "'", -1, -1
  Exit Function
 End Select
End If
'///
IASTNode_Verify = True
End Function

Private Function IASTNode_Codegen(ByVal objContext As clsVerifyContext, ByVal nParam1 As Long, ByVal nParam2 As Long, ByVal nParam3 As Long, ByVal nParam4 As Long) As Long
Dim hVariable As Long
Dim hValue As Long
Dim obj As IASTNode
Dim objSrcType As clsTypeNode, nSrcFlags As Long
Dim objDestType As clsTypeNode, nDestFlags As Long
'///
hValue = m_objRHS.Codegen(objContext, nParam1, nParam2, nParam3, nParam4)
hVariable = m_objLHS.CodegenEx(objContext, True)
'///
Set objSrcType = m_objRHS.GetType(nSrcFlags)
Set obj = m_objLHS
Set objDestType = obj.GetType(nDestFlags)
'///
hValue = g_objTypeMgr.CodegenTypeConversion(hValue, objSrcType, nSrcFlags, objDestType, nDestFlags, False)
IASTNode_Codegen = LLVMBuildStore(g_hBuilder, hValue, hVariable)
End Function


