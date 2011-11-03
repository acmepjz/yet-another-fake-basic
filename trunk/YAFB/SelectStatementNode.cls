VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SelectStatementNode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'<selectstat>:{select}{case}<exp><br><selectblock>*{end}{select}

Implements IASTNode

Private m_objCondition As IASTNode

Private m_objBlock() As SelectBlockNode '1-based
Private m_nBlockCount As Long

Friend Property Get Condition() As IASTNode
Set Condition = m_objCondition
End Property

Friend Property Set Condition(ByVal obj As IASTNode)
Set m_objCondition = obj
End Property

Friend Sub AddBlock(ByVal obj As SelectBlockNode)
m_nBlockCount = m_nBlockCount + 1
ReDim Preserve m_objBlock(1 To m_nBlockCount)
Set m_objBlock(m_nBlockCount) = obj
End Sub

Private Function IASTNode_Codegen(ByVal objContext As clsVerifyContext, ByVal nParam1 As Long, ByVal nParam2 As Long, ByVal nParam3 As Long, ByVal nParam4 As Long) As Long
Dim objDestType As clsTypeNode, nDestFlags As Long
Dim nTag As Long
Dim hVariable As Long
Dim hBlockEnd As Long
Dim i As Long
'///
hBlockEnd = LLVMAppendBasicBlock(objContext.CurrentFunction.FunctionHandle, StrPtr(StrConv("SelectCaseEnd", vbFromUnicode)))
'///create temp variable
nTag = ObjPtr(Me)
Set objDestType = m_objCondition.GetType(nDestFlags)
hVariable = objContext.CurrentFunction.GetTempVariable(objContext, objDestType, nTag)
'///store it
LLVMBuildStore g_hBuilder, m_objCondition.Codegen(objContext, nParam1, nParam2, nParam3, nParam4), hVariable
'///
For i = 1 To m_nBlockCount
 m_objBlock(i).CodegenEx objContext, hVariable, objDestType, nDestFlags, hBlockEnd
Next i
LLVMBuildBr g_hBuilder, hBlockEnd
'///over
LLVMPositionBuilderAtEnd g_hBuilder, hBlockEnd
objContext.CurrentFunction.ResetTempVariable objContext, nTag
End Function

Private Function IASTNode_GetProperty(ByVal nProp As enumASTNodeProperty) As Long
'nothing
End Function

Private Function IASTNode_GetType(nFlags As Long) As clsTypeNode
'nothing
End Function

Private Property Get IASTNode_NodeType() As enumASTNodeType
IASTNode_NodeType = node_selectstat
End Property

Private Function IASTNode_Verify(ByVal objContext As clsVerifyContext) As Boolean
Dim obj As IASTNode
Dim objDestType As clsTypeNode, nDestFlags As Long
Dim i As Long
'///
For i = 1 To m_nBlockCount
 Set obj = m_objBlock(i)
 If Not obj.Verify(objContext) Then Exit Function
Next i
'///check data type
If objContext.Phase = verify_all Then
 If Not m_objCondition.Verify(objContext) Then Exit Function
 '///
 Set objDestType = m_objCondition.GetType(nDestFlags)
 For i = 1 To m_nBlockCount
  If Not m_objBlock(i).VerifyConditionDataType(objDestType, nDestFlags) Then Exit Function
 Next i
End If
'///
IASTNode_Verify = True
End Function
