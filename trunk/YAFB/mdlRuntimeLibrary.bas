Attribute VB_Name = "mdlRuntimeLibrary"
Option Explicit

Public g_objIntrinsicDataTypes(255) As clsTypeNode

Public g_hFunctionPow As Long

Public g_hTypeVariant As Long
Public g_hTypeSafeArray As Long
Public g_hTypeSafeArrayBound As Long

Public Sub SetupRuntimeLibrary()
Dim i(7) As Long
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
'///variant (??)
i(0) = LLVMInt16Type
i(1) = i(0)
i(2) = i(0)
i(3) = i(0)
i(4) = LLVMInt64Type
With New clsTypeNode
 g_hTypeVariant = LLVMStructType(i(0), 5, 0)
 .SetIntrinsic vbVariant, "Variant", g_hTypeVariant
 g_objGlobalTable.TypeTable.Add .This, "Variant"
End With
'///safearraybound (??)
'Private Type SAFEARRAYBOUND
'    cElements As Long
'    lLbound As Long
'End Type
i(0) = LLVMInt32Type
i(1) = i(0)
g_hTypeSafeArrayBound = LLVMStructType(i(0), 2, 0)
'///safearray (??)
'Private Type SAFEARRAY2D
'    cDims As Integer
'    fFeatures As Integer
'    cbElements As Long
'    cLocks As Long
'    pvData As Long
'    Bounds(0 To 1) As SAFEARRAYBOUND
'End Type
i(0) = LLVMInt16Type
i(1) = i(0)
i(2) = LLVMInt32Type
i(3) = i(2)
i(4) = LLVMPointerType(LLVMInt8Type, 0)
i(5) = g_hTypeSafeArrayBound
g_hTypeSafeArray = LLVMStructType(i(0), 6, 0)
'///
'TODO:etc.
End Sub

Public Sub SetupRuntimeLibraryFunctions()
Dim hType(7) As Long
Dim hFunctionType As Long
'///
LLVMAddTypeName g_hModule, StrPtr(StrConv("VARIANTARG", vbFromUnicode)), g_hTypeVariant
LLVMAddTypeName g_hModule, StrPtr(StrConv("SAFEARRAYBOUND", vbFromUnicode)), g_hTypeSafeArrayBound
LLVMAddTypeName g_hModule, StrPtr(StrConv("SAFEARRAY", vbFromUnicode)), g_hTypeSafeArray
'///
hType(0) = LLVMDoubleType
hType(1) = hType(0)
hFunctionType = LLVMFunctionType(hType(0), hType(0), 2, 0)
g_hFunctionPow = LLVMAddFunction(g_hModule, StrPtr(StrConv("pow", vbFromUnicode)), hFunctionType)
LLVMSetLinkage g_hFunctionPow, LLVMExternalLinkage
'///
'TODO:other
End Sub
