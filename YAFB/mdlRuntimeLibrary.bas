Attribute VB_Name = "mdlRuntimeLibrary"
Option Explicit

Public g_objIntrinsicDataTypes(255) As clsTypeNode

Public g_hFunctionPow As Long

Public Sub SetupRuntimeLibrary()
Dim i(3) As Long
'///setup default types
With New clsTypeNode
 .SetIntrinsic vbByte, "Byte", LLVMInt8Type
 g_objGlobalTable.TypeTable.Add .This, "Byte"
End With
With New clsTypeNode
 .SetIntrinsic vbInteger, "Integer", LLVMInt16Type
 g_objGlobalTable.TypeTable.Add .This, "Integer"
End With
With New clsTypeNode
 .SetIntrinsic vbLong, "Long", LLVMInt32Type
 g_objGlobalTable.TypeTable.Add .This, "Long"
End With
With New clsTypeNode
 .SetIntrinsic vbBoolean, "Boolean", LLVMInt16Type
 g_objGlobalTable.TypeTable.Add .This, "Boolean"
End With
With New clsTypeNode
 .SetIntrinsic vbSingle, "Single", LLVMFloatType
 g_objGlobalTable.TypeTable.Add .This, "Single"
End With
With New clsTypeNode
 .SetIntrinsic vbDouble, "Double", LLVMDoubleType
 g_objGlobalTable.TypeTable.Add .This, "Double"
End With
'///??
With New clsTypeNode
 'LLVM doesn't support void* so void* becomes char*
 .SetIntrinsic vbEmpty, "Any", LLVMInt8Type
 g_objGlobalTable.TypeTable.Add .This, "Any"
End With
'///??
i(0) = LLVMInt32Type
i(1) = i(0)
i(2) = i(0)
i(3) = i(0)
With New clsTypeNode
 .SetIntrinsic vbVariant, "Variant", LLVMStructType(i(0), 4, 1)
 g_objGlobalTable.TypeTable.Add .This, "Variant"
End With
'///
'TODO:etc.
End Sub

Public Sub SetupRuntimeLibraryFunctions()
Dim hType(7) As Long
Dim hFunctionType As Long
'///
hType(0) = LLVMDoubleType
hType(1) = hType(0)
hFunctionType = LLVMFunctionType(hType(0), hType(0), 2, 0)
g_hFunctionPow = LLVMAddFunction(g_hModule, StrPtr(StrConv("pow", vbFromUnicode)), hFunctionType)
LLVMSetLinkage g_hFunctionPow, LLVMExternalLinkage
'///
'TODO:other
End Sub
