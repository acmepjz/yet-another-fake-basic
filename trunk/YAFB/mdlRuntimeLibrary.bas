Attribute VB_Name = "mdlRuntimeLibrary"
Option Explicit

Public g_objIntrinsicDataTypes(255) As clsTypeNode

Private g_hFunction(1023) As Long

Public Enum enumIntrinsicFunctions
 internal_pow = 0
 internal_memcpy
 internal_memmove
 internal_memset
 internal_llvm_memcpy
 internal_llvm_memmove
 internal_llvm_memset
End Enum

Public g_hTypeVariant As Long
Public g_hTypeSafeArray As Long
Public g_hTypeSafeArrayBound As Long
Public g_hTypeDecimal As Long

'================================ YAFB Extensions ================================

'/*
' * VARENUM usage key,
' *
' * * [V] - may appear in a VARIANT
' * * [T] - may appear in a TYPEDESC
' * * [P] - may appear in an OLE property set
' * * [S] - may appear in a Safe Array
' *
' *
' *  VT_EMPTY            [V]   [P]     nothing
' *  VT_NULL             [V]   [P]     SQL style Null
' *  VT_I2               [V][T][P][S]  2 byte signed int
' *  VT_I4               [V][T][P][S]  4 byte signed int
' *  VT_R4               [V][T][P][S]  4 byte real
' *  VT_R8               [V][T][P][S]  8 byte real
' *  VT_CY               [V][T][P][S]  currency
' *  VT_DATE             [V][T][P][S]  date
' *  VT_BSTR             [V][T][P][S]  OLE Automation string
' *  VT_DISPATCH         [V][T]   [S]  IDispatch *
' *  VT_ERROR            [V][T][P][S]  SCODE
' *  VT_BOOL             [V][T][P][S]  True=-1, False=0
' *  VT_VARIANT          [V][T][P][S]  VARIANT *
' *  VT_UNKNOWN          [V][T]   [S]  IUnknown *
' *  VT_DECIMAL          [V][T]   [S]  16 byte fixed point
' *  VT_RECORD           [V]   [P][S]  user defined type
' *  VT_I1               [V][T][P][s]  signed char
' *  VT_UI1              [V][T][P][S]  unsigned char
' *  VT_UI2              [V][T][P][S]  unsigned short
' *  VT_UI4              [V][T][P][S]  unsigned long
' *  VT_I8                  [T][P]     signed 64-bit int
' *  VT_UI8                 [T][P]     unsigned 64-bit int
' *  VT_INT              [V][T][P][S]  signed machine int
' *  VT_UINT             [V][T]   [S]  unsigned machine int
' *  VT_INT_PTR             [T]        signed machine register size width
' *  VT_UINT_PTR            [T]        unsigned machine register size width
' *  VT_VOID                [T]        C style void
' *  VT_HRESULT             [T]        Standard return type
' *  VT_PTR                 [T]        pointer type
' *  VT_SAFEARRAY           [T]        (use VT_ARRAY in VARIANT)
' *  VT_CARRAY              [T]        C style array
' *  VT_USERDEFINED         [T]        user defined type
' *  VT_LPSTR               [T][P]     null terminated string
' *  VT_LPWSTR              [T][P]     wide null terminated string
' *  VT_FILETIME               [P]     FILETIME
' *  VT_BLOB                   [P]     Length prefixed bytes
' *  VT_STREAM                 [P]     Name of the stream follows
' *  VT_STORAGE                [P]     Name of the storage follows
' *  VT_STREAMED_OBJECT        [P]     Stream contains an object
' *  VT_STORED_OBJECT          [P]     Storage contains an object
' *  VT_VERSIONED_STREAM       [P]     Stream with a GUID version
' *  VT_BLOB_OBJECT            [P]     Blob contains an object
' *  VT_CF                     [P]     Clipboard format
' *  VT_CLSID                  [P]     A Class ID
' *  VT_VECTOR                 [P]     simple counted array
' *  VT_ARRAY            [V]           SAFEARRAY*
' *  VT_BYREF            [V]           void* for local use
' *  VT_BSTR_BLOB                      Reserved for system use
' */
'
'enum VARENUM
'    {  VT_EMPTY    = 0,
'   VT_NULL = 1,
'   VT_I2   = 2,
'   VT_I4   = 3,
'   VT_R4   = 4,
'   VT_R8   = 5,
'   VT_CY   = 6,
'   VT_DATE = 7,
'   VT_BSTR = 8,
'   VT_DISPATCH = 9,
'   VT_ERROR    = 10,
'   VT_BOOL = 11,
'   VT_VARIANT  = 12,
'   VT_UNKNOWN  = 13,
'   VT_DECIMAL  = 14,
'   VT_I1   = 16,
'   VT_UI1  = 17,
'   VT_UI2  = 18,
'   VT_UI4  = 19,
'   VT_I8   = 20,
'   VT_UI8  = 21,
'   VT_INT  = 22,
'   VT_UINT = 23,
'   VT_VOID = 24,
'   VT_HRESULT  = 25,
'   VT_PTR  = 26,
'   VT_SAFEARRAY    = 27,
'   VT_CARRAY   = 28,
'   VT_USERDEFINED  = 29,
'   VT_LPSTR    = 30,
'   VT_LPWSTR   = 31,
'   VT_RECORD   = 36,
'   VT_INT_PTR  = 37,
'   VT_UINT_PTR = 38,
'   VT_FILETIME = 64,
'   VT_BLOB = 65,
'   VT_STREAM   = 66,
'   VT_STORAGE  = 67,
'   VT_STREAMED_OBJECT  = 68,
'   VT_STORED_OBJECT    = 69,
'   VT_BLOB_OBJECT  = 70,
'   VT_CF   = 71,
'   VT_CLSID    = 72,
'   VT_VERSIONED_STREAM = 73,
'   VT_BSTR_BLOB    = 0xfff,
'   VT_VECTOR   = 0x1000,
'   VT_ARRAY    = 0x2000,
'   VT_BYREF    = 0x4000,
'   VT_RESERVED = 0x8000,
'   VT_ILLEGAL  = 0xffff,
'   VT_ILLEGALMASKED    = 0xfff,
'   VT_TYPEMASK = 0xfff
'    } ;

Public Const vbSignedByte As Long = 16
Public Const vbUnsignedInteger As Long = 18
Public Const vbUnsignedLong As Long = 19
Public Const vbLongLong As Long = 20
Public Const vbUnsignedLongLong As Long = 21
Public Const vbIntPtr_t As Long = 37
Public Const vbUIntPtr_t As Long = 38

Public Const vbLongLongLong As Long = &HC0&
Public Const vbUnsignedLongLongLong As Long = &HC1&

Public Sub SetupRuntimeLibrary()
Dim i(7) As Long
'////////setup default (and extension) types
'///int8
With New clsTypeNode
 .SetIntrinsic vbSignedByte, "SignedByte", LLVMInt8Type, 1, &H71
End With
With New clsTypeNode
 .SetIntrinsic vbByte, "Byte", LLVMInt8Type, 1, &H72
End With
'///int16
With New clsTypeNode
 .SetIntrinsic vbInteger, "Integer", LLVMInt16Type, 2, &H71
End With
With New clsTypeNode
 .SetIntrinsic vbUnsignedInteger, "UnsignedInteger", LLVMInt16Type, 2, &H72
End With
'///int32
With New clsTypeNode
 .SetIntrinsic vbLong, "Long", LLVMInt32Type, 4, &H71
End With
With New clsTypeNode
 .SetIntrinsic vbUnsignedLong, "UnsignedLong", LLVMInt32Type, 4, &H72
End With
'///int64
With New clsTypeNode
 .SetIntrinsic vbLongLong, "LongLong", LLVMInt64Type, 8, &H71
End With
With New clsTypeNode
 .SetIntrinsic vbUnsignedLongLong, "UnsignedLongLong", LLVMInt64Type, 8, &H72
End With
'///int128 (experimental)
i(0) = LLVMIntType(128)
With New clsTypeNode
 .SetIntrinsic vbLongLongLong, "LongLongLong", i(0), 16, &H71
End With
With New clsTypeNode
 .SetIntrinsic vbUnsignedLongLongLong, "UnsignedLongLongLong", i(0), 16, &H72
End With
'///intptr_t
Select Case g_nWordSize
Case 4
 With New clsTypeNode
  .SetIntrinsic vbIntPtr_t, "IntPtr_t", LLVMInt32Type, 4, &H71
 End With
 With New clsTypeNode
  .SetIntrinsic vbUIntPtr_t, "UIntPtr_t", LLVMInt32Type, 4, &H72
 End With
Case 8
 With New clsTypeNode
  .SetIntrinsic vbIntPtr_t, "IntPtr_t", LLVMInt64Type, 8, &H71
 End With
 With New clsTypeNode
  .SetIntrinsic vbUIntPtr_t, "UIntPtr_t", LLVMInt64Type, 8, &H72
 End With
Case Else
 PrintPanic "Unknown word size: " + CStr(g_nWordSize), -1, -1
End Select
'///currency
With New clsTypeNode
 .SetIntrinsic vbCurrency, "Currency", LLVMInt64Type, 8, &H70
End With
'///boolean
With New clsTypeNode
 .SetIntrinsic vbBoolean, "Boolean", LLVMInt16Type, 2, &H31
End With
'///float
With New clsTypeNode
 .SetIntrinsic vbSingle, "Single", LLVMFloatType, 4, &H73
End With
With New clsTypeNode
 .SetIntrinsic vbDouble, "Double", LLVMDoubleType, 8, &H73
End With
'///date
With New clsTypeNode
 .SetIntrinsic vbDate, "Date", LLVMDoubleType, 8, &H33
End With
'///any (??)
With New clsTypeNode
 'LLVM doesn't support void* so void* becomes char*
 .SetIntrinsic vbEmpty, "Any", LLVMInt8Type, 1, 0
End With
'///variant (??)
i(0) = LLVMInt16Type
i(1) = i(0)
i(2) = i(0)
i(3) = i(0)
i(4) = LLVMInt64Type
With New clsTypeNode
 g_hTypeVariant = LLVMStructType(i(0), 5, 0)
 .SetIntrinsic vbVariant, "Variant", g_hTypeVariant, 16, 0
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
'///decimal (??)
'///union with tagVariant
'typedef struct tagDEC
'    {
'    USHORT wReserved; 'should be vbDecimal
'    BYTE scale;
'    BYTE sign;
'    ULONG Hi32;
'    ULONGLONG Lo64;
'    }   DECIMAL;
i(0) = LLVMInt16Type
i(1) = LLVMInt8Type
i(2) = i(1)
i(3) = LLVMInt32Type
i(4) = LLVMInt64Type
With New clsTypeNode
 g_hTypeDecimal = LLVMStructType(i(0), 5, 0)
 .SetIntrinsic vbDecimal, "Decimal", g_hTypeDecimal, 16, &H70
End With
'///
'TODO:etc.
End Sub

Public Sub SetupRuntimeLibraryFunctions()
'///
LLVMAddTypeName g_hModule, StrPtr(StrConv("VARIANT", vbFromUnicode)), g_hTypeVariant
LLVMAddTypeName g_hModule, StrPtr(StrConv("SAFEARRAYBOUND", vbFromUnicode)), g_hTypeSafeArrayBound
LLVMAddTypeName g_hModule, StrPtr(StrConv("SAFEARRAY", vbFromUnicode)), g_hTypeSafeArray
LLVMAddTypeName g_hModule, StrPtr(StrConv("DECIMAL", vbFromUnicode)), g_hTypeDecimal
'///
'TODO:other
End Sub

Public Function RuntimeLibraryGetFunction(ByVal nIndex As enumIntrinsicFunctions) As Long
Dim hType(7) As Long
Dim hFunctionType As Long
Dim hFunction As Long
'///
hFunction = g_hFunction(nIndex)
If hFunction = 0 Then
 Select Case nIndex
 Case internal_pow
  hType(0) = LLVMDoubleType
  hType(1) = hType(0)
  hFunctionType = LLVMFunctionType(hType(0), hType(0), 2, 0)
  'hFunction = LLVMAddFunction(g_hModule, StrPtr(StrConv("pow", vbFromUnicode)), hFunctionType)
  hFunction = LLVMAddFunction(g_hModule, StrPtr(StrConv("llvm.pow.f64", vbFromUnicode)), hFunctionType)
  LLVMAddFunctionAttr hFunction, LLVMNoUnwindAttribute
  LLVMSetLinkage hFunction, LLVMExternalLinkage
 Case internal_memcpy
  hType(0) = LLVMPointerType(LLVMInt8Type, 0)
  hType(1) = hType(0)
  hType(2) = g_objIntrinsicDataTypes(vbIntPtr_t).Handle
  hFunctionType = LLVMFunctionType(LLVMVoidType, hType(0), 3, 0)
  hFunction = LLVMAddFunction(g_hModule, StrPtr(StrConv("memcpy", vbFromUnicode)), hFunctionType)
  LLVMAddAttribute LLVMGetParam(hFunction, 0), LLVMNoCaptureAttribute
  LLVMAddAttribute LLVMGetParam(hFunction, 1), LLVMNoCaptureAttribute
  LLVMAddFunctionAttr hFunction, LLVMNoUnwindAttribute
  LLVMSetLinkage hFunction, LLVMExternalLinkage
 Case internal_memmove
  hType(0) = LLVMPointerType(LLVMInt8Type, 0)
  hType(1) = hType(0)
  hType(2) = g_objIntrinsicDataTypes(vbIntPtr_t).Handle
  hFunctionType = LLVMFunctionType(LLVMVoidType, hType(0), 3, 0)
  hFunction = LLVMAddFunction(g_hModule, StrPtr(StrConv("memmove", vbFromUnicode)), hFunctionType)
  LLVMAddAttribute LLVMGetParam(hFunction, 0), LLVMNoCaptureAttribute
  LLVMAddAttribute LLVMGetParam(hFunction, 1), LLVMNoCaptureAttribute
  LLVMAddFunctionAttr hFunction, LLVMNoUnwindAttribute
  LLVMSetLinkage hFunction, LLVMExternalLinkage
 Case internal_memset
  hType(0) = LLVMPointerType(LLVMInt8Type, 0)
  hType(1) = LLVMInt32Type
  hType(2) = g_objIntrinsicDataTypes(vbIntPtr_t).Handle
  hFunctionType = LLVMFunctionType(LLVMVoidType, hType(0), 3, 0)
  hFunction = LLVMAddFunction(g_hModule, StrPtr(StrConv("memset", vbFromUnicode)), hFunctionType)
  LLVMAddAttribute LLVMGetParam(hFunction, 0), LLVMNoCaptureAttribute
  LLVMAddFunctionAttr hFunction, LLVMNoUnwindAttribute
  LLVMSetLinkage hFunction, LLVMExternalLinkage
 Case internal_llvm_memcpy
  hType(0) = LLVMPointerType(LLVMInt8Type, 0)
  hType(1) = hType(0)
  hType(2) = g_objIntrinsicDataTypes(vbIntPtr_t).Handle
  hType(3) = LLVMInt32Type
  hType(4) = LLVMInt1Type
  hFunctionType = LLVMFunctionType(LLVMVoidType, hType(0), 5, 0)
  hFunction = LLVMAddFunction(g_hModule, StrPtr(StrConv("llvm.memcpy.p0i8.p0i8.i" + CStr(g_nWordSize * 8&), vbFromUnicode)), hFunctionType)
  LLVMAddAttribute LLVMGetParam(hFunction, 0), LLVMNoCaptureAttribute
  LLVMAddAttribute LLVMGetParam(hFunction, 1), LLVMNoCaptureAttribute
  LLVMAddFunctionAttr hFunction, LLVMNoUnwindAttribute
  LLVMSetLinkage hFunction, LLVMExternalLinkage
 Case internal_llvm_memmove
  hType(0) = LLVMPointerType(LLVMInt8Type, 0)
  hType(1) = hType(0)
  hType(2) = g_objIntrinsicDataTypes(vbIntPtr_t).Handle
  hType(3) = LLVMInt32Type
  hType(4) = LLVMInt1Type
  hFunctionType = LLVMFunctionType(LLVMVoidType, hType(0), 5, 0)
  hFunction = LLVMAddFunction(g_hModule, StrPtr(StrConv("llvm.memmove.p0i8.p0i8.i" + CStr(g_nWordSize * 8&), vbFromUnicode)), hFunctionType)
  LLVMAddAttribute LLVMGetParam(hFunction, 0), LLVMNoCaptureAttribute
  LLVMAddAttribute LLVMGetParam(hFunction, 1), LLVMNoCaptureAttribute
  LLVMAddFunctionAttr hFunction, LLVMNoUnwindAttribute
  LLVMSetLinkage hFunction, LLVMExternalLinkage
 Case internal_llvm_memset
  hType(0) = LLVMPointerType(LLVMInt8Type, 0)
  hType(1) = LLVMInt8Type
  hType(2) = g_objIntrinsicDataTypes(vbIntPtr_t).Handle
  hType(3) = LLVMInt32Type
  hType(4) = LLVMInt1Type
  hFunctionType = LLVMFunctionType(LLVMVoidType, hType(0), 5, 0)
  hFunction = LLVMAddFunction(g_hModule, StrPtr(StrConv("llvm.memset.p0i8.i" + CStr(g_nWordSize * 8&), vbFromUnicode)), hFunctionType)
  LLVMAddAttribute LLVMGetParam(hFunction, 0), LLVMNoCaptureAttribute
  LLVMAddFunctionAttr hFunction, LLVMNoUnwindAttribute
  LLVMSetLinkage hFunction, LLVMExternalLinkage
 End Select
 g_hFunction(nIndex) = hFunction
End If
'///
RuntimeLibraryGetFunction = hFunction
End Function
