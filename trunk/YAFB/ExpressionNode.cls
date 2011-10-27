VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ExpressionNode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IASTNode

Private m_nOperatorType As enumTokenType '0 means trasnparent force-byval unary operator
Private m_nOperandCount As Long

Private m_objSubNode(1) As IASTNode

Private m_objType As clsTypeNode

'================================ LLVM ================================

Private m_hConst As Long

Private Function IASTNode_GetProperty(ByVal nProp As enumASTNodeProperty) As Long
Dim hValue1 As Long, hValue2 As Long
Dim objType1 As clsTypeNode, nFlags1 As Long
Dim objType2 As clsTypeNode, nFlags2 As Long
'///
Select Case nProp
Case action_const_codegen
 If m_hConst = 0 Then
  Select Case m_nOperandCount
  Case 1
   hValue1 = m_objSubNode(0).GetProperty(action_const_codegen)
   If hValue1 = 0 Then Exit Function
   '///
   Set objType1 = m_objSubNode(0).GetType(nFlags1)
   Set m_objType = g_objTypeMgr.CheckUnaryOperator(objType1, m_nOperatorType)
   If m_objType Is Nothing Then
    PrintError "Can't apply operator <TODO> on data type '" + objType1.Name + "'", -1, -1
    Exit Function
   End If
   '///
   m_hConst = g_objTypeMgr.CodegenUnaryOperator(Nothing, hValue1, objType1, m_nOperatorType, m_objType, True)
  Case 2
   hValue1 = m_objSubNode(0).GetProperty(action_const_codegen)
   If hValue1 = 0 Then Exit Function
   hValue2 = m_objSubNode(1).GetProperty(action_const_codegen)
   If hValue2 = 0 Then Exit Function
   '///
   Set objType1 = m_objSubNode(0).GetType(nFlags1)
   Set objType2 = m_objSubNode(1).GetType(nFlags2)
   Set m_objType = g_objTypeMgr.CheckBinaryOperator(objType1, objType2, m_nOperatorType)
   If m_objType Is Nothing Then
    PrintError "Can't apply operator <TODO> on data type '" + objType1.Name + "' and '" + objType2.Name + "'", -1, -1
    Exit Function
   End If
   '///
   m_hConst = g_objTypeMgr.CodegenBinaryOperator(Nothing, hValue1, objType1, hValue2, objType2, m_nOperatorType, m_objType, True)
  End Select
 End If
 IASTNode_GetProperty = m_hConst
End Select
End Function

Private Function IASTNode_GetType(nFlags As Long) As clsTypeNode
nFlags = 0 '???
Set IASTNode_GetType = m_objType
End Function

Private Property Get IASTNode_NodeType() As enumASTNodeType
IASTNode_NodeType = node_exp
End Property

Friend Sub SetTypeAndSubNode(ByVal nType As enumTokenType, ByVal nCount As Long, ByVal obj0 As IASTNode, Optional ByVal obj1 As IASTNode)
m_nOperatorType = nType
m_nOperandCount = nCount
Set m_objSubNode(0) = obj0
Set m_objSubNode(1) = obj1
End Sub

Friend Property Get This() As IASTNode
Set This = Me
End Property

Private Function IASTNode_Verify(ByVal objContext As clsVerifyContext) As Boolean
Dim objType1 As clsTypeNode, nFlags1 As Long
Dim objType2 As clsTypeNode, nFlags2 As Long
Dim i As Long
'///
For i = 0 To m_nOperandCount - 1
 If Not m_objSubNode(i).Verify(objContext) Then Exit Function
Next i
'///
Select Case objContext.Phase
Case verify_const
 '///register dependency
 For i = 0 To m_nOperandCount - 1
  g_objConstDAG.AddEdge m_objSubNode(i), Me
 Next i
Case verify_all
 Select Case m_nOperandCount
 Case 1
  Set objType1 = m_objSubNode(0).GetType(nFlags1)
  Set m_objType = g_objTypeMgr.CheckUnaryOperator(objType1, m_nOperatorType)
  If m_objType Is Nothing Then
   PrintError "Can't apply operator <TODO> on data type '" + objType1.Name + "'", -1, -1
   Exit Function
  End If
 Case 2
  Set objType1 = m_objSubNode(0).GetType(nFlags1)
  Set objType2 = m_objSubNode(1).GetType(nFlags2)
  Set m_objType = g_objTypeMgr.CheckBinaryOperator(objType1, objType2, m_nOperatorType)
  If m_objType Is Nothing Then
   PrintError "Can't apply operator <TODO> on data type '" + objType1.Name + "' and '" + objType2.Name + "'", -1, -1
   Exit Function
  End If
 End Select
End Select
'///
IASTNode_Verify = True
End Function

Private Function IASTNode_Codegen(ByVal objContext As clsVerifyContext, ByVal nParam1 As Long, ByVal nParam2 As Long, ByVal nParam3 As Long, ByVal nParam4 As Long) As Long
Dim objType1 As clsTypeNode, nFlags1 As Long
Dim objType2 As clsTypeNode, nFlags2 As Long
Dim hValue1 As Long
Dim hValue2 As Long
'///
Select Case m_nOperandCount
Case 1
 hValue1 = m_objSubNode(0).Codegen(objContext, 0, 0, 0, 0)
 Set objType1 = m_objSubNode(0).GetType(nFlags1)
 IASTNode_Codegen = g_objTypeMgr.CodegenUnaryOperator(objContext, hValue1, objType1, m_nOperatorType, m_objType, False)
Case 2
 hValue1 = m_objSubNode(0).Codegen(objContext, 0, 0, 0, 0)
 hValue2 = m_objSubNode(1).Codegen(objContext, 0, 0, 0, 0)
 Set objType1 = m_objSubNode(0).GetType(nFlags1)
 Set objType2 = m_objSubNode(1).GetType(nFlags1)
 IASTNode_Codegen = g_objTypeMgr.CodegenBinaryOperator(objContext, hValue1, objType1, hValue2, objType2, m_nOperatorType, m_objType, False)
End Select
End Function


