VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ArgDeclareListNode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IASTNode

Private m_nFlags() As Long '1-based
'1=byval (0=byref)
'2=array
'4=optional
'8=paramarray
Private m_obj() As DimNode
Private m_objDefault() As IASTNode

Private m_nCount As Long

Friend Property Get ArgumentCount() As Long
ArgumentCount = m_nCount
End Property

Friend Property Get ArgumentDeclaration(ByVal nIndex As Long) As DimNode
Set ArgumentDeclaration = m_obj(nIndex)
End Property

Friend Property Get ArgumentDefaultValue(ByVal nIndex As Long) As IASTNode
Set ArgumentDefaultValue = m_objDefault(nIndex)
End Property

Friend Property Get ArgumentFlags(ByVal nIndex As Long) As Long
ArgumentFlags = m_nFlags(nIndex)
End Property

Private Function IASTNode_Codegen(ByVal objContext As clsVerifyContext, ByVal nParam1 As Long, ByVal nParam2 As Long, ByVal nParam3 As Long, ByVal nParam4 As Long) As Long
'nothing
End Function

Friend Function CodegenFunctionType(ByVal hReturnType As Long) As Long
Dim hType() As Long
Dim i As Long, j As Long
'///
ReDim hType(1 To m_nCount + 2)
'///
For i = 1 To m_nCount
 hType(i) = m_obj(i).CodegenType
Next i
'///
CodegenFunctionType = LLVMFunctionType(hReturnType, hType(1), m_nCount, 0)
End Function

Friend Sub CodegenArgument(ByVal hFunction As Long)
Dim i As Long
For i = 1 To m_nCount
 m_obj(i).CodegenArgument hFunction, i - 1
Next i
End Sub

Private Function IASTNode_GetProperty(ByVal nProp As enumASTNodeProperty) As Long
'
End Function

Private Function IASTNode_GetType(nFlags As Long) As clsTypeNode
'nothing
End Function

Private Property Get IASTNode_NodeType() As enumASTNodeType
IASTNode_NodeType = node_argumentlist
End Property

Friend Property Get This() As IASTNode
Set This = Me
End Property

Friend Sub AddSubNode(ByRef tName As typeToken, ByVal nFlags As Long, ByVal objDataType As DataTypeNode, ByVal objDefault As IASTNode)
Dim obj As New DimNode
'///
obj.DimType = &HF0& Or (((nFlags And &H1&) = 0) And &H1000&) Or (((nFlags And &H2&) <> 0) And &H400&)
obj.SetTokenEx tName
If objDataType Is Nothing Then
 obj.SetDataTypeFromString "Variant"
Else
 Set obj.DataType = objDataType
End If
'///
m_nCount = m_nCount + 1
ReDim Preserve m_nFlags(1 To m_nCount)
ReDim Preserve m_obj(1 To m_nCount)
ReDim Preserve m_objDefault(1 To m_nCount)
m_nFlags(m_nCount) = nFlags
Set m_obj(m_nCount) = obj
Set m_objDefault(m_nCount) = objDefault
End Sub

Friend Function Register(ByVal objTable As clsSymbolTable) As Boolean
Dim i As Long
For i = 1 To m_nCount
 If Not m_obj(i).Register(objTable) Then Exit Function
Next i
Register = True
End Function

Private Function IASTNode_Verify(ByVal objContext As clsVerifyContext) As Boolean
Dim i As Long
Dim hValue As Long
Dim nLastFlags As Long
Dim obj As IASTNode
Dim objSrcType As clsTypeNode, nSrcFlags As Long
Dim objDestType As clsTypeNode, nDestFlags As Long
'///
Select Case objContext.Phase
Case verify_const
 For i = 1 To m_nCount
  If Not m_objDefault(i) Is Nothing Then
   If Not m_objDefault(i).Verify(objContext) Then Exit Function
  End If
 Next i
Case verify_dim
 For i = 1 To m_nCount
  Set obj = m_obj(i)
  If Not obj.Verify(objContext) Then Exit Function
  '///
  If Not m_obj(i).DataType Is Nothing Then
   If m_obj(i).DataType.GetDataType.DataType = vbUserDefinedType Then
    m_obj(i).DataType.ErrorNoByVal
    Exit Function
   End If
  End If
  '///
  If nLastFlags And 4& Then
   If (m_nFlags(i) And 4&) = 0 Then
    m_obj(i).ErrorMustBeOptional
    Exit Function
   End If
  End If
  nLastFlags = m_nFlags(i)
 Next i
Case verify_all
 For i = 1 To m_nCount
  If m_nFlags(i) And 4& Then
   If m_objDefault(i) Is Nothing Then
    '///get default value of data type
    Set obj = m_obj(i)
    Set objDestType = obj.GetType(nDestFlags)
    If nDestFlags = 0 Then
     hValue = objDestType.CodegenDefaultValue
    Else
     hValue = 0
    End If
    '///TODO:
    If hValue Then
     With New ConstNode
      .CreateIndirect objDestType, hValue
      Set m_objDefault(i) = .This
     End With
    Else
     PrintError "Current default value of data type '" + objDestType.NameEx(nDestFlags) + "' is unsupported", -1, -1
     Exit Function
    End If
   Else
    '///convert data type
    Set objSrcType = m_objDefault(i).GetType(nSrcFlags)
    Set obj = m_obj(i)
    Set objDestType = obj.GetType(nDestFlags)
    Select Case g_objTypeMgr.CheckTypeConversion(objSrcType, nSrcFlags, objDestType, nDestFlags)
    Case 0
     PrintError "Type mismatch: can't convert data type '" + objSrcType.NameEx(nSrcFlags) + "' to '" + objDestType.NameEx(nDestFlags) + "'", -1, -1
     Exit Function
    End Select
    '///
    hValue = m_objDefault(i).GetProperty(action_const_codegen)
    If Not CheckConstant(hValue) Then Exit Function
    hValue = g_objTypeMgr.CodegenTypeConversion(objContext, hValue, objSrcType, nSrcFlags, objDestType, nDestFlags, True)
    If Not CheckConstant(hValue) Then Exit Function
    With New ConstNode
     .CreateIndirect objDestType, hValue
     Set m_objDefault(i) = .This
    End With
   End If
  End If
 Next i
End Select
'///
IASTNode_Verify = True
End Function
