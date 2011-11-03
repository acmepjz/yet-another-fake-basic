VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IfStatementNode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IASTNode

Private m_nCount As Long

'<if> cond1 <then> then1
'<elseif> cond2 <then> then2
'......
'<elseif> condN <then> thenN
'<else> else

Private m_objCondition() As IASTNode '1-based, must not be Nothing
Private m_objThen() As StatementListNode 'can be Nothing

Private m_objElse As StatementListNode 'can be Nothing

Private Function IASTNode_GetProperty(ByVal nProp As enumASTNodeProperty) As Long
'
End Function

Friend Property Get ThenBlockCount() As Long
ThenBlockCount = m_nCount
End Property

Friend Sub AddThenBlock(ByVal objCondition As IASTNode, ByVal objThen As StatementListNode)
m_nCount = m_nCount + 1
ReDim Preserve m_objCondition(1 To m_nCount)
ReDim Preserve m_objThen(1 To m_nCount)
Set m_objCondition(m_nCount) = objCondition
Set m_objThen(m_nCount) = objThen
End Sub

Friend Property Get Condition(ByVal nIndex As Long) As IASTNode
Set Condition = m_objCondition(nIndex)
End Property

Friend Property Set Condition(ByVal nIndex As Long, ByVal obj As IASTNode)
Set m_objCondition(nIndex) = obj
End Property

Friend Property Get ThenBlock(ByVal nIndex As Long) As StatementListNode
Set ThenBlock = m_objThen(nIndex)
End Property

Friend Property Set ThenBlock(ByVal nIndex As Long, ByVal obj As StatementListNode)
Set m_objThen(nIndex) = obj
End Property

Friend Property Get ElseBlock() As StatementListNode
Set ElseBlock = m_objElse
End Property

Friend Property Set ElseBlock(ByVal obj As StatementListNode)
Set m_objElse = obj
End Property

Private Function IASTNode_Codegen(ByVal objContext As clsVerifyContext, ByVal nParam1 As Long, ByVal nParam2 As Long, ByVal nParam3 As Long, ByVal nParam4 As Long) As Long
Dim obj As IASTNode
Dim objType As clsTypeNode, nFlags As Long
Dim hFunction As Long
Dim hBlock() As Long, hTmpBlock As Long
Dim hValue As Long
Dim i As Long
Dim s As String, lp As Long, lp2 As Long
'///
hFunction = objContext.CurrentFunction.FunctionHandle
'///
ReDim hBlock(1 To m_nCount + 2)
s = StrConv("IfBlock", vbFromUnicode)
lp = StrPtr(s)
For i = 1 To m_nCount + 2
 hBlock(i) = LLVMAppendBasicBlock(hFunction, lp)
Next i
'///if temp block
s = StrConv("ConditionTemp" + vbNullChar + "IfTempBlock", vbFromUnicode)
lp = StrPtr(s)
lp2 = lp + 14
For i = 1 To m_nCount
 hValue = m_objCondition(i).Codegen(objContext, nParam1, nParam2, nParam3, nParam4)
 Set objType = m_objCondition(i).GetType(nFlags)
 '///
 hValue = objType.CodegenConvertToI1(hValue, lp)
 '///
 hTmpBlock = LLVMAppendBasicBlock(hFunction, lp2)
 LLVMBuildCondBr g_hBuilder, hValue, hBlock(i), hTmpBlock
 LLVMPositionBuilderAtEnd g_hBuilder, hTmpBlock
Next i
LLVMBuildBr g_hBuilder, hBlock(m_nCount + 1)
'///then block
For i = 1 To m_nCount
 LLVMPositionBuilderAtEnd g_hBuilder, hBlock(i)
 Set obj = m_objThen(i)
 If Not obj Is Nothing Then obj.Codegen objContext, nParam1, nParam2, nParam3, nParam4
 LLVMBuildBr g_hBuilder, hBlock(m_nCount + 2)
Next i
'///else block
LLVMPositionBuilderAtEnd g_hBuilder, hBlock(m_nCount + 1)
Set obj = m_objElse
If Not obj Is Nothing Then obj.Codegen objContext, nParam1, nParam2, nParam3, nParam4
LLVMBuildBr g_hBuilder, hBlock(m_nCount + 2)
'///
LLVMPositionBuilderAtEnd g_hBuilder, hBlock(m_nCount + 2)
End Function

Private Function IASTNode_GetType(nFlags As Long) As clsTypeNode
'nothing
End Function

Private Property Get IASTNode_NodeType() As enumASTNodeType
IASTNode_NodeType = node_ifstat
End Property

Private Function IASTNode_Verify(ByVal objContext As clsVerifyContext) As Boolean
Dim i As Long
Dim obj As IASTNode
Dim objType As clsTypeNode, nFlags As Long
Dim nPhase As enumASTNodeVerifyStep
'///
nPhase = objContext.Phase
'///
For i = 1 To m_nCount
 If nPhase = verify_all Then
  If Not m_objCondition(i).Verify(objContext) Then Exit Function
 End If
 '///
 Set obj = m_objThen(i)
 If Not obj Is Nothing Then
  If Not obj.Verify(objContext) Then Exit Function
 End If
Next i
'///
Set obj = m_objElse
If Not obj Is Nothing Then
 If Not obj.Verify(objContext) Then Exit Function
End If
'///check condition data type
If nPhase = verify_all Then
 For i = 1 To m_nCount
  Set objType = m_objCondition(i).GetType(nFlags)
  If (objType.Flags And &H20&) = 0 Or nFlags <> 0 Then
   PrintError "Data type '" + objType.NameEx(nFlags) + "' can't use for condition of 'If' statement", -1, -1
   Exit Function
  End If
 Next i
End If
'///
IASTNode_Verify = True
End Function
