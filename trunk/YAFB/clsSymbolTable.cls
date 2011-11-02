VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSymbolTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_objFunc As Collection
Private m_objVariable As Collection 'not all items are variables, some are constants
Private m_objType As Collection
Private m_objLineNumber As Collection

Friend Property Get FunctionTable() As Collection
Set FunctionTable = m_objFunc
End Property

Friend Property Get VariableTable() As Collection
Set VariableTable = m_objVariable
End Property

Friend Property Get TypeTable() As Collection
Set TypeTable = m_objType
End Property

Friend Property Get LineNumberTable() As Collection
Set LineNumberTable = m_objLineNumber
End Property

Private Sub Class_Initialize()
Set m_objFunc = New Collection
Set m_objVariable = New Collection
Set m_objType = New Collection
Set m_objLineNumber = New Collection
End Sub

Friend Sub Codegen(ByVal objContext As clsVerifyContext)
Dim v As Variant
Dim objFunction As FunctionNode
Dim objVariable As DimNode
'///
For Each v In m_objFunc
 Set objFunction = v
 objFunction.CodegenFunctionDeclaration objContext
Next v
'///
For Each v In m_objVariable
 Set objVariable = v
 objVariable.CodegenVariable objContext
Next v
End Sub

Friend Sub CodegenDefaultDestructor()
Dim v As Variant
Dim objVariable As DimNode
'///
For Each v In m_objVariable
 Set objVariable = v
 objVariable.CodegenDefaultDestructor
Next v
End Sub

Friend Sub CodegenLineNumber(ByVal objContext As clsVerifyContext)
Dim v As Variant
Dim obj As LineNumberStatement
'///
For Each v In m_objLineNumber
 Set obj = v
 obj.CodegenLineNumber objContext
Next v
End Sub
