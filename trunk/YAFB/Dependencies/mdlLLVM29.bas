Attribute VB_Name = "mdlLLVM29"
Option Explicit

'--- .\Lib\swig.swg ---
'--- .\Lib\swigwarnings.swg ---
'--- .\Lib\swigwarn.swg ---
'--- .\Lib\swigwarnings.swg ---
'--- .\Lib\swig.swg ---
'--- F:\Projects\llvm-2.9\include\llvm-c\test.i ---
'--- F:\Projects\llvm-2.9\llvm2.9\Util_CdeclCallbackWrapper.h ---
'Public Declare Function Util_CreateCdeclCallbackWrapper Lib "llvm2.9.dll" (ByVal func_ As Long, ByVal ArgCount_ As Long) As Long 'Void*
'Public Declare Sub Util_DestroyCdeclCallbackWrapper Lib "llvm2.9.dll" (ByRef lp_ As Any)
'--- F:\Projects\llvm-2.9\include\llvm-c\test.i ---
'--- F:\Projects\llvm-2.9\include\llvm-c\test.h ---
'--- F:\Projects\llvm-2.9\include\llvm-c\Core.h ---
'--- F:\Projects\llvm-2.9\vs2008\include\llvm\Support\DataTypes.h ---
'--- F:\Projects\llvm-2.9\include\llvm-c\Core.h ---
'Public Declare Sub LLVMDisposeMessage Lib "llvm2.9.dll" (ByVal Message_ As String)
'Public Declare Function LLVMContextCreate Lib "llvm2.9.dll" () As Long
'Public Declare Function LLVMGetGlobalContext Lib "llvm2.9.dll" () As Long
'Public Declare Sub LLVMContextDispose Lib "llvm2.9.dll" (ByVal C_ As Long)
'Public Declare Function LLVMGetMDKindIDInContext Lib "llvm2.9.dll" (ByVal C_ As Long, ByVal Name_ As String, ByVal SLen_ As Long) As Long
'Public Declare Function LLVMGetMDKindID Lib "llvm2.9.dll" (ByVal Name_ As String, ByVal SLen_ As Long) As Long
'Public Declare Function LLVMModuleCreateWithName Lib "llvm2.9.dll" (ByVal ModuleID_ As String) As Long
'Public Declare Function LLVMModuleCreateWithNameInContext Lib "llvm2.9.dll" (ByVal ModuleID_ As String, ByVal C_ As Long) As Long
'Public Declare Sub LLVMDisposeModule Lib "llvm2.9.dll" (ByVal M_ As Long)
'Public Declare Function LLVMGetDataLayout Lib "llvm2.9.dll" (ByVal M_ As Long) As Long 'Byte*
'Public Declare Sub LLVMSetDataLayout Lib "llvm2.9.dll" (ByVal M_ As Long, ByVal Triple_ As String)
'Public Declare Function LLVMGetTarget Lib "llvm2.9.dll" (ByVal M_ As Long) As Long 'Byte*
'Public Declare Sub LLVMSetTarget Lib "llvm2.9.dll" (ByVal M_ As Long, ByVal Triple_ As String)
'Public Declare Function LLVMAddTypeName Lib "llvm2.9.dll" (ByVal M_ As Long, ByVal Name_ As String, ByVal Ty_ As Long) As Long
'Public Declare Sub LLVMDeleteTypeName Lib "llvm2.9.dll" (ByVal M_ As Long, ByVal Name_ As String)
'Public Declare Function LLVMGetTypeByName Lib "llvm2.9.dll" (ByVal M_ As Long, ByVal Name_ As String) As Long
'Public Declare Function LLVMGetTypeName Lib "llvm2.9.dll" (ByVal M_ As Long, ByVal Ty_ As Long) As Long 'Byte*
'Public Declare Sub LLVMDumpModule Lib "llvm2.9.dll" (ByVal M_ As Long)
'Public Declare Sub LLVMSetModuleInlineAsm Lib "llvm2.9.dll" (ByVal M_ As Long, ByVal Asm_ As String)
'Public Declare Function LLVMGetModuleContext Lib "llvm2.9.dll" (ByVal M_ As Long) As Long
'Public Declare Function LLVMGetTypeKind Lib "llvm2.9.dll" (ByVal Ty_ As Long) As LLVMTypeKind
'Public Declare Function LLVMGetTypeContext Lib "llvm2.9.dll" (ByVal Ty_ As Long) As Long
'Public Declare Function LLVMInt1TypeInContext Lib "llvm2.9.dll" (ByVal C_ As Long) As Long
'Public Declare Function LLVMInt8TypeInContext Lib "llvm2.9.dll" (ByVal C_ As Long) As Long
'Public Declare Function LLVMInt16TypeInContext Lib "llvm2.9.dll" (ByVal C_ As Long) As Long
'Public Declare Function LLVMInt32TypeInContext Lib "llvm2.9.dll" (ByVal C_ As Long) As Long
'Public Declare Function LLVMInt64TypeInContext Lib "llvm2.9.dll" (ByVal C_ As Long) As Long
'Public Declare Function LLVMIntTypeInContext Lib "llvm2.9.dll" (ByVal C_ As Long, ByVal NumBits_ As Long) As Long
'Public Declare Function LLVMInt1Type Lib "llvm2.9.dll" () As Long
'Public Declare Function LLVMInt8Type Lib "llvm2.9.dll" () As Long
'Public Declare Function LLVMInt16Type Lib "llvm2.9.dll" () As Long
'Public Declare Function LLVMInt32Type Lib "llvm2.9.dll" () As Long
'Public Declare Function LLVMInt64Type Lib "llvm2.9.dll" () As Long
'Public Declare Function LLVMIntType Lib "llvm2.9.dll" (ByVal NumBits_ As Long) As Long
'Public Declare Function LLVMGetIntTypeWidth Lib "llvm2.9.dll" (ByVal IntegerTy_ As Long) As Long
'Public Declare Function LLVMFloatTypeInContext Lib "llvm2.9.dll" (ByVal C_ As Long) As Long
'Public Declare Function LLVMDoubleTypeInContext Lib "llvm2.9.dll" (ByVal C_ As Long) As Long
'Public Declare Function LLVMX86FP80TypeInContext Lib "llvm2.9.dll" (ByVal C_ As Long) As Long
'Public Declare Function LLVMFP128TypeInContext Lib "llvm2.9.dll" (ByVal C_ As Long) As Long
'Public Declare Function LLVMPPCFP128TypeInContext Lib "llvm2.9.dll" (ByVal C_ As Long) As Long
'Public Declare Function LLVMFloatType Lib "llvm2.9.dll" () As Long
'Public Declare Function LLVMDoubleType Lib "llvm2.9.dll" () As Long
'Public Declare Function LLVMX86FP80Type Lib "llvm2.9.dll" () As Long
'Public Declare Function LLVMFP128Type Lib "llvm2.9.dll" () As Long
'Public Declare Function LLVMPPCFP128Type Lib "llvm2.9.dll" () As Long
'Public Declare Function LLVMFunctionType Lib "llvm2.9.dll" (ByVal ReturnType_ As Long, ByRef ParamTypes_ As Long, ByVal ParamCount_ As Long, ByVal IsVarArg_ As Long) As Long
'Public Declare Function LLVMIsFunctionVarArg Lib "llvm2.9.dll" (ByVal FunctionTy_ As Long) As Long
'Public Declare Function LLVMGetReturnType Lib "llvm2.9.dll" (ByVal FunctionTy_ As Long) As Long
'Public Declare Function LLVMCountParamTypes Lib "llvm2.9.dll" (ByVal FunctionTy_ As Long) As Long
'Public Declare Sub LLVMGetParamTypes Lib "llvm2.9.dll" (ByVal FunctionTy_ As Long, ByRef Dest_ As Long)
'Public Declare Function LLVMStructTypeInContext Lib "llvm2.9.dll" (ByVal C_ As Long, ByRef ElementTypes_ As Long, ByVal ElementCount_ As Long, ByVal Packed_ As Long) As Long
'Public Declare Function LLVMStructType Lib "llvm2.9.dll" (ByRef ElementTypes_ As Long, ByVal ElementCount_ As Long, ByVal Packed_ As Long) As Long
'Public Declare Function LLVMCountStructElementTypes Lib "llvm2.9.dll" (ByVal StructTy_ As Long) As Long
'Public Declare Sub LLVMGetStructElementTypes Lib "llvm2.9.dll" (ByVal StructTy_ As Long, ByRef Dest_ As Long)
'Public Declare Function LLVMIsPackedStruct Lib "llvm2.9.dll" (ByVal StructTy_ As Long) As Long
'Public Declare Function LLVMArrayType Lib "llvm2.9.dll" (ByVal ElementType_ As Long, ByVal ElementCount_ As Long) As Long
'Public Declare Function LLVMPointerType Lib "llvm2.9.dll" (ByVal ElementType_ As Long, ByVal AddressSpace_ As Long) As Long
'Public Declare Function LLVMVectorType Lib "llvm2.9.dll" (ByVal ElementType_ As Long, ByVal ElementCount_ As Long) As Long
'Public Declare Function LLVMGetElementType Lib "llvm2.9.dll" (ByVal Ty_ As Long) As Long
'Public Declare Function LLVMGetArrayLength Lib "llvm2.9.dll" (ByVal ArrayTy_ As Long) As Long
'Public Declare Function LLVMGetPointerAddressSpace Lib "llvm2.9.dll" (ByVal PointerTy_ As Long) As Long
'Public Declare Function LLVMGetVectorSize Lib "llvm2.9.dll" (ByVal VectorTy_ As Long) As Long
'Public Declare Function LLVMVoidTypeInContext Lib "llvm2.9.dll" (ByVal C_ As Long) As Long
'Public Declare Function LLVMLabelTypeInContext Lib "llvm2.9.dll" (ByVal C_ As Long) As Long
'Public Declare Function LLVMOpaqueTypeInContext Lib "llvm2.9.dll" (ByVal C_ As Long) As Long
'Public Declare Function LLVMX86MMXTypeInContext Lib "llvm2.9.dll" (ByVal C_ As Long) As Long
'Public Declare Function LLVMVoidType Lib "llvm2.9.dll" () As Long
'Public Declare Function LLVMLabelType Lib "llvm2.9.dll" () As Long
'Public Declare Function LLVMOpaqueType Lib "llvm2.9.dll" () As Long
'Public Declare Function LLVMX86MMXType Lib "llvm2.9.dll" () As Long
'Public Declare Function LLVMCreateTypeHandle Lib "llvm2.9.dll" (ByVal PotentiallyAbstractTy_ As Long) As Long
'Public Declare Sub LLVMRefineType Lib "llvm2.9.dll" (ByVal AbstractTy_ As Long, ByVal ConcreteTy_ As Long)
'Public Declare Function LLVMResolveTypeHandle Lib "llvm2.9.dll" (ByVal TypeHandle_ As Long) As Long
'Public Declare Sub LLVMDisposeTypeHandle Lib "llvm2.9.dll" (ByVal TypeHandle_ As Long)
'Public Declare Function LLVMTypeOf Lib "llvm2.9.dll" (ByVal Val_ As Long) As Long
'Public Declare Function LLVMGetValueName Lib "llvm2.9.dll" (ByVal Val_ As Long) As Long 'Byte*
'Public Declare Sub LLVMSetValueName Lib "llvm2.9.dll" (ByVal Val_ As Long, ByVal Name_ As String)
'Public Declare Sub LLVMDumpValue Lib "llvm2.9.dll" (ByVal Val_ As Long)
'Public Declare Sub LLVMReplaceAllUsesWith Lib "llvm2.9.dll" (ByVal OldVal_ As Long, ByVal NewVal_ As Long)
'Public Declare Function LLVMHasMetadata Lib "llvm2.9.dll" (ByVal Val_ As Long) As Long
'Public Declare Function LLVMGetMetadata Lib "llvm2.9.dll" (ByVal Val_ As Long, ByVal KindID_ As Long) As Long
'Public Declare Sub LLVMSetMetadata Lib "llvm2.9.dll" (ByVal Val_ As Long, ByVal KindID_ As Long, ByVal Node_ As Long)
'Public Declare Function LLVMIsAArgument Lib "llvm2.9.dll" (ByVal Val_ As Long) As Long
'Public Declare Function LLVMIsABasicBlock Lib "llvm2.9.dll" (ByVal Val_ As Long) As Long
'Public Declare Function LLVMIsAInlineAsm Lib "llvm2.9.dll" (ByVal Val_ As Long) As Long
'Public Declare Function LLVMIsAUser Lib "llvm2.9.dll" (ByVal Val_ As Long) As Long
'Public Declare Function LLVMIsAConstant Lib "llvm2.9.dll" (ByVal Val_ As Long) As Long
'Public Declare Function LLVMIsAConstantAggregateZero Lib "llvm2.9.dll" (ByVal Val_ As Long) As Long
'Public Declare Function LLVMIsAConstantArray Lib "llvm2.9.dll" (ByVal Val_ As Long) As Long
'Public Declare Function LLVMIsAConstantExpr Lib "llvm2.9.dll" (ByVal Val_ As Long) As Long
'Public Declare Function LLVMIsAConstantFP Lib "llvm2.9.dll" (ByVal Val_ As Long) As Long
'Public Declare Function LLVMIsAConstantInt Lib "llvm2.9.dll" (ByVal Val_ As Long) As Long
'Public Declare Function LLVMIsAConstantPointerNull Lib "llvm2.9.dll" (ByVal Val_ As Long) As Long
'Public Declare Function LLVMIsAConstantStruct Lib "llvm2.9.dll" (ByVal Val_ As Long) As Long
'Public Declare Function LLVMIsAConstantVector Lib "llvm2.9.dll" (ByVal Val_ As Long) As Long
'Public Declare Function LLVMIsAGlobalValue Lib "llvm2.9.dll" (ByVal Val_ As Long) As Long
'Public Declare Function LLVMIsAFunction Lib "llvm2.9.dll" (ByVal Val_ As Long) As Long
'Public Declare Function LLVMIsAGlobalAlias Lib "llvm2.9.dll" (ByVal Val_ As Long) As Long
'Public Declare Function LLVMIsAGlobalVariable Lib "llvm2.9.dll" (ByVal Val_ As Long) As Long
'Public Declare Function LLVMIsAUndefValue Lib "llvm2.9.dll" (ByVal Val_ As Long) As Long
'Public Declare Function LLVMIsAInstruction Lib "llvm2.9.dll" (ByVal Val_ As Long) As Long
'Public Declare Function LLVMIsABinaryOperator Lib "llvm2.9.dll" (ByVal Val_ As Long) As Long
'Public Declare Function LLVMIsACallInst Lib "llvm2.9.dll" (ByVal Val_ As Long) As Long
'Public Declare Function LLVMIsAIntrinsicInst Lib "llvm2.9.dll" (ByVal Val_ As Long) As Long
'Public Declare Function LLVMIsADbgInfoIntrinsic Lib "llvm2.9.dll" (ByVal Val_ As Long) As Long
'Public Declare Function LLVMIsADbgDeclareInst Lib "llvm2.9.dll" (ByVal Val_ As Long) As Long
'Public Declare Function LLVMIsAEHSelectorInst Lib "llvm2.9.dll" (ByVal Val_ As Long) As Long
'Public Declare Function LLVMIsAMemIntrinsic Lib "llvm2.9.dll" (ByVal Val_ As Long) As Long
'Public Declare Function LLVMIsAMemCpyInst Lib "llvm2.9.dll" (ByVal Val_ As Long) As Long
'Public Declare Function LLVMIsAMemMoveInst Lib "llvm2.9.dll" (ByVal Val_ As Long) As Long
'Public Declare Function LLVMIsAMemSetInst Lib "llvm2.9.dll" (ByVal Val_ As Long) As Long
'Public Declare Function LLVMIsACmpInst Lib "llvm2.9.dll" (ByVal Val_ As Long) As Long
'Public Declare Function LLVMIsAFCmpInst Lib "llvm2.9.dll" (ByVal Val_ As Long) As Long
'Public Declare Function LLVMIsAICmpInst Lib "llvm2.9.dll" (ByVal Val_ As Long) As Long
'Public Declare Function LLVMIsAExtractElementInst Lib "llvm2.9.dll" (ByVal Val_ As Long) As Long
'Public Declare Function LLVMIsAGetElementPtrInst Lib "llvm2.9.dll" (ByVal Val_ As Long) As Long
'Public Declare Function LLVMIsAInsertElementInst Lib "llvm2.9.dll" (ByVal Val_ As Long) As Long
'Public Declare Function LLVMIsAInsertValueInst Lib "llvm2.9.dll" (ByVal Val_ As Long) As Long
'Public Declare Function LLVMIsAPHINode Lib "llvm2.9.dll" (ByVal Val_ As Long) As Long
'Public Declare Function LLVMIsASelectInst Lib "llvm2.9.dll" (ByVal Val_ As Long) As Long
'Public Declare Function LLVMIsAShuffleVectorInst Lib "llvm2.9.dll" (ByVal Val_ As Long) As Long
'Public Declare Function LLVMIsAStoreInst Lib "llvm2.9.dll" (ByVal Val_ As Long) As Long
'Public Declare Function LLVMIsATerminatorInst Lib "llvm2.9.dll" (ByVal Val_ As Long) As Long
'Public Declare Function LLVMIsABranchInst Lib "llvm2.9.dll" (ByVal Val_ As Long) As Long
'Public Declare Function LLVMIsAInvokeInst Lib "llvm2.9.dll" (ByVal Val_ As Long) As Long
'Public Declare Function LLVMIsAReturnInst Lib "llvm2.9.dll" (ByVal Val_ As Long) As Long
'Public Declare Function LLVMIsASwitchInst Lib "llvm2.9.dll" (ByVal Val_ As Long) As Long
'Public Declare Function LLVMIsAUnreachableInst Lib "llvm2.9.dll" (ByVal Val_ As Long) As Long
'Public Declare Function LLVMIsAUnwindInst Lib "llvm2.9.dll" (ByVal Val_ As Long) As Long
'Public Declare Function LLVMIsAUnaryInstruction Lib "llvm2.9.dll" (ByVal Val_ As Long) As Long
'Public Declare Function LLVMIsAAllocaInst Lib "llvm2.9.dll" (ByVal Val_ As Long) As Long
'Public Declare Function LLVMIsACastInst Lib "llvm2.9.dll" (ByVal Val_ As Long) As Long
'Public Declare Function LLVMIsABitCastInst Lib "llvm2.9.dll" (ByVal Val_ As Long) As Long
'Public Declare Function LLVMIsAFPExtInst Lib "llvm2.9.dll" (ByVal Val_ As Long) As Long
'Public Declare Function LLVMIsAFPToSIInst Lib "llvm2.9.dll" (ByVal Val_ As Long) As Long
'Public Declare Function LLVMIsAFPToUIInst Lib "llvm2.9.dll" (ByVal Val_ As Long) As Long
'Public Declare Function LLVMIsAFPTruncInst Lib "llvm2.9.dll" (ByVal Val_ As Long) As Long
'Public Declare Function LLVMIsAIntToPtrInst Lib "llvm2.9.dll" (ByVal Val_ As Long) As Long
'Public Declare Function LLVMIsAPtrToIntInst Lib "llvm2.9.dll" (ByVal Val_ As Long) As Long
'Public Declare Function LLVMIsASExtInst Lib "llvm2.9.dll" (ByVal Val_ As Long) As Long
'Public Declare Function LLVMIsASIToFPInst Lib "llvm2.9.dll" (ByVal Val_ As Long) As Long
'Public Declare Function LLVMIsATruncInst Lib "llvm2.9.dll" (ByVal Val_ As Long) As Long
'Public Declare Function LLVMIsAUIToFPInst Lib "llvm2.9.dll" (ByVal Val_ As Long) As Long
'Public Declare Function LLVMIsAZExtInst Lib "llvm2.9.dll" (ByVal Val_ As Long) As Long
'Public Declare Function LLVMIsAExtractValueInst Lib "llvm2.9.dll" (ByVal Val_ As Long) As Long
'Public Declare Function LLVMIsALoadInst Lib "llvm2.9.dll" (ByVal Val_ As Long) As Long
'Public Declare Function LLVMIsAVAArgInst Lib "llvm2.9.dll" (ByVal Val_ As Long) As Long
'Public Declare Function LLVMGetFirstUse Lib "llvm2.9.dll" (ByVal Val_ As Long) As Long
'Public Declare Function LLVMGetNextUse Lib "llvm2.9.dll" (ByVal U_ As Long) As Long
'Public Declare Function LLVMGetUser Lib "llvm2.9.dll" (ByVal U_ As Long) As Long
'Public Declare Function LLVMGetUsedValue Lib "llvm2.9.dll" (ByVal U_ As Long) As Long
'Public Declare Function LLVMGetOperand Lib "llvm2.9.dll" (ByVal Val_ As Long, ByVal Index_ As Long) As Long
'Public Declare Sub LLVMSetOperand Lib "llvm2.9.dll" (ByVal User_ As Long, ByVal Index_ As Long, ByVal Val_ As Long)
'Public Declare Function LLVMGetNumOperands Lib "llvm2.9.dll" (ByVal Val_ As Long) As Long
'Public Declare Function LLVMConstNull Lib "llvm2.9.dll" (ByVal Ty_ As Long) As Long
'Public Declare Function LLVMConstAllOnes Lib "llvm2.9.dll" (ByVal Ty_ As Long) As Long
'Public Declare Function LLVMGetUndef Lib "llvm2.9.dll" (ByVal Ty_ As Long) As Long
'Public Declare Function LLVMIsConstant Lib "llvm2.9.dll" (ByVal Val_ As Long) As Long
'Public Declare Function LLVMIsNull Lib "llvm2.9.dll" (ByVal Val_ As Long) As Long
'Public Declare Function LLVMIsUndef Lib "llvm2.9.dll" (ByVal Val_ As Long) As Long
'Public Declare Function LLVMConstPointerNull Lib "llvm2.9.dll" (ByVal Ty_ As Long) As Long
'Public Declare Function LLVMMDStringInContext Lib "llvm2.9.dll" (ByVal C_ As Long, ByVal Str_ As String, ByVal SLen_ As Long) As Long
'Public Declare Function LLVMMDString Lib "llvm2.9.dll" (ByVal Str_ As String, ByVal SLen_ As Long) As Long
'Public Declare Function LLVMMDNodeInContext Lib "llvm2.9.dll" (ByVal C_ As Long, ByRef Vals_ As Long, ByVal Count_ As Long) As Long
'Public Declare Function LLVMMDNode Lib "llvm2.9.dll" (ByRef Vals_ As Long, ByVal Count_ As Long) As Long
'Public Declare Function LLVMConstInt Lib "llvm2.9.dll" (ByVal IntTy_ As Long, ByVal N_ As Currency, ByVal SignExtend_ As Long) As Long
'Public Declare Function LLVMConstIntOfArbitraryPrecision Lib "llvm2.9.dll" (ByVal IntTy_ As Long, ByVal NumWords_ As Long, ByRef Words_ As Currency) As Long
'Public Declare Function LLVMConstIntOfString Lib "llvm2.9.dll" (ByVal IntTy_ As Long, ByVal Text_ As String, ByVal Radix_ As Byte) As Long
'Public Declare Function LLVMConstIntOfStringAndSize Lib "llvm2.9.dll" (ByVal IntTy_ As Long, ByVal Text_ As String, ByVal SLen_ As Long, ByVal Radix_ As Byte) As Long
'Public Declare Function LLVMConstReal Lib "llvm2.9.dll" (ByVal RealTy_ As Long, ByVal N_ As Double) As Long
'Public Declare Function LLVMConstRealOfString Lib "llvm2.9.dll" (ByVal RealTy_ As Long, ByVal Text_ As String) As Long
'Public Declare Function LLVMConstRealOfStringAndSize Lib "llvm2.9.dll" (ByVal RealTy_ As Long, ByVal Text_ As String, ByVal SLen_ As Long) As Long
'Public Declare Function LLVMConstIntGetZExtValue Lib "llvm2.9.dll" (ByVal ConstantVal_ As Long) As Currency
'Public Declare Function LLVMConstIntGetSExtValue Lib "llvm2.9.dll" (ByVal ConstantVal_ As Long) As Currency
'Public Declare Function LLVMConstStringInContext Lib "llvm2.9.dll" (ByVal C_ As Long, ByVal Str_ As String, ByVal Length_ As Long, ByVal DontNullTerminate_ As Long) As Long
'Public Declare Function LLVMConstStructInContext Lib "llvm2.9.dll" (ByVal C_ As Long, ByRef ConstantVals_ As Long, ByVal Count_ As Long, ByVal Packed_ As Long) As Long
'Public Declare Function LLVMConstString Lib "llvm2.9.dll" (ByVal Str_ As String, ByVal Length_ As Long, ByVal DontNullTerminate_ As Long) As Long
'Public Declare Function LLVMConstArray Lib "llvm2.9.dll" (ByVal ElementTy_ As Long, ByRef ConstantVals_ As Long, ByVal Length_ As Long) As Long
'Public Declare Function LLVMConstStruct Lib "llvm2.9.dll" (ByRef ConstantVals_ As Long, ByVal Count_ As Long, ByVal Packed_ As Long) As Long
'Public Declare Function LLVMConstVector Lib "llvm2.9.dll" (ByRef ScalarConstantVals_ As Long, ByVal Size_ As Long) As Long
'Public Declare Function LLVMGetConstOpcode Lib "llvm2.9.dll" (ByVal ConstantVal_ As Long) As LLVMOpcode
'Public Declare Function LLVMAlignOf Lib "llvm2.9.dll" (ByVal Ty_ As Long) As Long
'Public Declare Function LLVMSizeOf Lib "llvm2.9.dll" (ByVal Ty_ As Long) As Long
'Public Declare Function LLVMConstNeg Lib "llvm2.9.dll" (ByVal ConstantVal_ As Long) As Long
'Public Declare Function LLVMConstNSWNeg Lib "llvm2.9.dll" (ByVal ConstantVal_ As Long) As Long
'Public Declare Function LLVMConstNUWNeg Lib "llvm2.9.dll" (ByVal ConstantVal_ As Long) As Long
'Public Declare Function LLVMConstFNeg Lib "llvm2.9.dll" (ByVal ConstantVal_ As Long) As Long
'Public Declare Function LLVMConstNot Lib "llvm2.9.dll" (ByVal ConstantVal_ As Long) As Long
'Public Declare Function LLVMConstAdd Lib "llvm2.9.dll" (ByVal LHSConstant_ As Long, ByVal RHSConstant_ As Long) As Long
'Public Declare Function LLVMConstNSWAdd Lib "llvm2.9.dll" (ByVal LHSConstant_ As Long, ByVal RHSConstant_ As Long) As Long
'Public Declare Function LLVMConstNUWAdd Lib "llvm2.9.dll" (ByVal LHSConstant_ As Long, ByVal RHSConstant_ As Long) As Long
'Public Declare Function LLVMConstFAdd Lib "llvm2.9.dll" (ByVal LHSConstant_ As Long, ByVal RHSConstant_ As Long) As Long
'Public Declare Function LLVMConstSub Lib "llvm2.9.dll" (ByVal LHSConstant_ As Long, ByVal RHSConstant_ As Long) As Long
'Public Declare Function LLVMConstNSWSub Lib "llvm2.9.dll" (ByVal LHSConstant_ As Long, ByVal RHSConstant_ As Long) As Long
'Public Declare Function LLVMConstNUWSub Lib "llvm2.9.dll" (ByVal LHSConstant_ As Long, ByVal RHSConstant_ As Long) As Long
'Public Declare Function LLVMConstFSub Lib "llvm2.9.dll" (ByVal LHSConstant_ As Long, ByVal RHSConstant_ As Long) As Long
'Public Declare Function LLVMConstMul Lib "llvm2.9.dll" (ByVal LHSConstant_ As Long, ByVal RHSConstant_ As Long) As Long
'Public Declare Function LLVMConstNSWMul Lib "llvm2.9.dll" (ByVal LHSConstant_ As Long, ByVal RHSConstant_ As Long) As Long
'Public Declare Function LLVMConstNUWMul Lib "llvm2.9.dll" (ByVal LHSConstant_ As Long, ByVal RHSConstant_ As Long) As Long
'Public Declare Function LLVMConstFMul Lib "llvm2.9.dll" (ByVal LHSConstant_ As Long, ByVal RHSConstant_ As Long) As Long
'Public Declare Function LLVMConstUDiv Lib "llvm2.9.dll" (ByVal LHSConstant_ As Long, ByVal RHSConstant_ As Long) As Long
'Public Declare Function LLVMConstSDiv Lib "llvm2.9.dll" (ByVal LHSConstant_ As Long, ByVal RHSConstant_ As Long) As Long
'Public Declare Function LLVMConstExactSDiv Lib "llvm2.9.dll" (ByVal LHSConstant_ As Long, ByVal RHSConstant_ As Long) As Long
'Public Declare Function LLVMConstFDiv Lib "llvm2.9.dll" (ByVal LHSConstant_ As Long, ByVal RHSConstant_ As Long) As Long
'Public Declare Function LLVMConstURem Lib "llvm2.9.dll" (ByVal LHSConstant_ As Long, ByVal RHSConstant_ As Long) As Long
'Public Declare Function LLVMConstSRem Lib "llvm2.9.dll" (ByVal LHSConstant_ As Long, ByVal RHSConstant_ As Long) As Long
'Public Declare Function LLVMConstFRem Lib "llvm2.9.dll" (ByVal LHSConstant_ As Long, ByVal RHSConstant_ As Long) As Long
'Public Declare Function LLVMConstAnd Lib "llvm2.9.dll" (ByVal LHSConstant_ As Long, ByVal RHSConstant_ As Long) As Long
'Public Declare Function LLVMConstOr Lib "llvm2.9.dll" (ByVal LHSConstant_ As Long, ByVal RHSConstant_ As Long) As Long
'Public Declare Function LLVMConstXor Lib "llvm2.9.dll" (ByVal LHSConstant_ As Long, ByVal RHSConstant_ As Long) As Long
'Public Declare Function LLVMConstICmp Lib "llvm2.9.dll" (ByVal Predicate_ As LLVMIntPredicate, ByVal LHSConstant_ As Long, ByVal RHSConstant_ As Long) As Long
'Public Declare Function LLVMConstFCmp Lib "llvm2.9.dll" (ByVal Predicate_ As LLVMRealPredicate, ByVal LHSConstant_ As Long, ByVal RHSConstant_ As Long) As Long
'Public Declare Function LLVMConstShl Lib "llvm2.9.dll" (ByVal LHSConstant_ As Long, ByVal RHSConstant_ As Long) As Long
'Public Declare Function LLVMConstLShr Lib "llvm2.9.dll" (ByVal LHSConstant_ As Long, ByVal RHSConstant_ As Long) As Long
'Public Declare Function LLVMConstAShr Lib "llvm2.9.dll" (ByVal LHSConstant_ As Long, ByVal RHSConstant_ As Long) As Long
'Public Declare Function LLVMConstGEP Lib "llvm2.9.dll" (ByVal ConstantVal_ As Long, ByRef ConstantIndices_ As Long, ByVal NumIndices_ As Long) As Long
'Public Declare Function LLVMConstInBoundsGEP Lib "llvm2.9.dll" (ByVal ConstantVal_ As Long, ByRef ConstantIndices_ As Long, ByVal NumIndices_ As Long) As Long
'Public Declare Function LLVMConstTrunc Lib "llvm2.9.dll" (ByVal ConstantVal_ As Long, ByVal ToType_ As Long) As Long
'Public Declare Function LLVMConstSExt Lib "llvm2.9.dll" (ByVal ConstantVal_ As Long, ByVal ToType_ As Long) As Long
'Public Declare Function LLVMConstZExt Lib "llvm2.9.dll" (ByVal ConstantVal_ As Long, ByVal ToType_ As Long) As Long
'Public Declare Function LLVMConstFPTrunc Lib "llvm2.9.dll" (ByVal ConstantVal_ As Long, ByVal ToType_ As Long) As Long
'Public Declare Function LLVMConstFPExt Lib "llvm2.9.dll" (ByVal ConstantVal_ As Long, ByVal ToType_ As Long) As Long
'Public Declare Function LLVMConstUIToFP Lib "llvm2.9.dll" (ByVal ConstantVal_ As Long, ByVal ToType_ As Long) As Long
'Public Declare Function LLVMConstSIToFP Lib "llvm2.9.dll" (ByVal ConstantVal_ As Long, ByVal ToType_ As Long) As Long
'Public Declare Function LLVMConstFPToUI Lib "llvm2.9.dll" (ByVal ConstantVal_ As Long, ByVal ToType_ As Long) As Long
'Public Declare Function LLVMConstFPToSI Lib "llvm2.9.dll" (ByVal ConstantVal_ As Long, ByVal ToType_ As Long) As Long
'Public Declare Function LLVMConstPtrToInt Lib "llvm2.9.dll" (ByVal ConstantVal_ As Long, ByVal ToType_ As Long) As Long
'Public Declare Function LLVMConstIntToPtr Lib "llvm2.9.dll" (ByVal ConstantVal_ As Long, ByVal ToType_ As Long) As Long
'Public Declare Function LLVMConstBitCast Lib "llvm2.9.dll" (ByVal ConstantVal_ As Long, ByVal ToType_ As Long) As Long
'Public Declare Function LLVMConstZExtOrBitCast Lib "llvm2.9.dll" (ByVal ConstantVal_ As Long, ByVal ToType_ As Long) As Long
'Public Declare Function LLVMConstSExtOrBitCast Lib "llvm2.9.dll" (ByVal ConstantVal_ As Long, ByVal ToType_ As Long) As Long
'Public Declare Function LLVMConstTruncOrBitCast Lib "llvm2.9.dll" (ByVal ConstantVal_ As Long, ByVal ToType_ As Long) As Long
'Public Declare Function LLVMConstPointerCast Lib "llvm2.9.dll" (ByVal ConstantVal_ As Long, ByVal ToType_ As Long) As Long
'Public Declare Function LLVMConstIntCast Lib "llvm2.9.dll" (ByVal ConstantVal_ As Long, ByVal ToType_ As Long, ByVal isSigned_ As Long) As Long
'Public Declare Function LLVMConstFPCast Lib "llvm2.9.dll" (ByVal ConstantVal_ As Long, ByVal ToType_ As Long) As Long
'Public Declare Function LLVMConstSelect Lib "llvm2.9.dll" (ByVal ConstantCondition_ As Long, ByVal ConstantIfTrue_ As Long, ByVal ConstantIfFalse_ As Long) As Long
'Public Declare Function LLVMConstExtractElement Lib "llvm2.9.dll" (ByVal VectorConstant_ As Long, ByVal IndexConstant_ As Long) As Long
'Public Declare Function LLVMConstInsertElement Lib "llvm2.9.dll" (ByVal VectorConstant_ As Long, ByVal ElementValueConstant_ As Long, ByVal IndexConstant_ As Long) As Long
'Public Declare Function LLVMConstShuffleVector Lib "llvm2.9.dll" (ByVal VectorAConstant_ As Long, ByVal VectorBConstant_ As Long, ByVal MaskConstant_ As Long) As Long
'Public Declare Function LLVMConstExtractValue Lib "llvm2.9.dll" (ByVal AggConstant_ As Long, ByRef IdxList_ As Long, ByVal NumIdx_ As Long) As Long
'Public Declare Function LLVMConstInsertValue Lib "llvm2.9.dll" (ByVal AggConstant_ As Long, ByVal ElementValueConstant_ As Long, ByRef IdxList_ As Long, ByVal NumIdx_ As Long) As Long
'Public Declare Function LLVMConstInlineAsm Lib "llvm2.9.dll" (ByVal Ty_ As Long, ByVal AsmString_ As String, ByVal Constraints_ As String, ByVal HasSideEffects_ As Long, ByVal IsAlignStack_ As Long) As Long
'Public Declare Function LLVMBlockAddress Lib "llvm2.9.dll" (ByVal F_ As Long, ByVal BB_ As Long) As Long
'Public Declare Function LLVMGetGlobalParent Lib "llvm2.9.dll" (ByVal Global_ As Long) As Long
'Public Declare Function LLVMIsDeclaration Lib "llvm2.9.dll" (ByVal Global_ As Long) As Long
'Public Declare Function LLVMGetLinkage Lib "llvm2.9.dll" (ByVal Global_ As Long) As LLVMLinkage
'Public Declare Sub LLVMSetLinkage Lib "llvm2.9.dll" (ByVal Global_ As Long, ByVal Linkage_ As LLVMLinkage)
'Public Declare Function LLVMGetSection Lib "llvm2.9.dll" (ByVal Global_ As Long) As Long 'Byte*
'Public Declare Sub LLVMSetSection Lib "llvm2.9.dll" (ByVal Global_ As Long, ByVal Section_ As String)
'Public Declare Function LLVMGetVisibility Lib "llvm2.9.dll" (ByVal Global_ As Long) As LLVMVisibility
'Public Declare Sub LLVMSetVisibility Lib "llvm2.9.dll" (ByVal Global_ As Long, ByVal Viz_ As LLVMVisibility)
'Public Declare Function LLVMGetAlignment Lib "llvm2.9.dll" (ByVal Global_ As Long) As Long
'Public Declare Sub LLVMSetAlignment Lib "llvm2.9.dll" (ByVal Global_ As Long, ByVal Bytes_ As Long)
'Public Declare Function LLVMAddGlobal Lib "llvm2.9.dll" (ByVal M_ As Long, ByVal Ty_ As Long, ByVal Name_ As String) As Long
'Public Declare Function LLVMAddGlobalInAddressSpace Lib "llvm2.9.dll" (ByVal M_ As Long, ByVal Ty_ As Long, ByVal Name_ As String, ByVal AddressSpace_ As Long) As Long
'Public Declare Function LLVMGetNamedGlobal Lib "llvm2.9.dll" (ByVal M_ As Long, ByVal Name_ As String) As Long
'Public Declare Function LLVMGetFirstGlobal Lib "llvm2.9.dll" (ByVal M_ As Long) As Long
'Public Declare Function LLVMGetLastGlobal Lib "llvm2.9.dll" (ByVal M_ As Long) As Long
'Public Declare Function LLVMGetNextGlobal Lib "llvm2.9.dll" (ByVal GlobalVar_ As Long) As Long
'Public Declare Function LLVMGetPreviousGlobal Lib "llvm2.9.dll" (ByVal GlobalVar_ As Long) As Long
'Public Declare Sub LLVMDeleteGlobal Lib "llvm2.9.dll" (ByVal GlobalVar_ As Long)
'Public Declare Function LLVMGetInitializer Lib "llvm2.9.dll" (ByVal GlobalVar_ As Long) As Long
'Public Declare Sub LLVMSetInitializer Lib "llvm2.9.dll" (ByVal GlobalVar_ As Long, ByVal ConstantVal_ As Long)
'Public Declare Function LLVMIsThreadLocal Lib "llvm2.9.dll" (ByVal GlobalVar_ As Long) As Long
'Public Declare Sub LLVMSetThreadLocal Lib "llvm2.9.dll" (ByVal GlobalVar_ As Long, ByVal IsThreadLocal_ As Long)
'Public Declare Function LLVMIsGlobalConstant Lib "llvm2.9.dll" (ByVal GlobalVar_ As Long) As Long
'Public Declare Sub LLVMSetGlobalConstant Lib "llvm2.9.dll" (ByVal GlobalVar_ As Long, ByVal IsConstant_ As Long)
'Public Declare Function LLVMAddAlias Lib "llvm2.9.dll" (ByVal M_ As Long, ByVal Ty_ As Long, ByVal Aliasee_ As Long, ByVal Name_ As String) As Long
'Public Declare Function LLVMAddFunction Lib "llvm2.9.dll" (ByVal M_ As Long, ByVal Name_ As String, ByVal FunctionTy_ As Long) As Long
'Public Declare Function LLVMGetNamedFunction Lib "llvm2.9.dll" (ByVal M_ As Long, ByVal Name_ As String) As Long
'Public Declare Function LLVMGetFirstFunction Lib "llvm2.9.dll" (ByVal M_ As Long) As Long
'Public Declare Function LLVMGetLastFunction Lib "llvm2.9.dll" (ByVal M_ As Long) As Long
'Public Declare Function LLVMGetNextFunction Lib "llvm2.9.dll" (ByVal Fn_ As Long) As Long
'Public Declare Function LLVMGetPreviousFunction Lib "llvm2.9.dll" (ByVal Fn_ As Long) As Long
'Public Declare Sub LLVMDeleteFunction Lib "llvm2.9.dll" (ByVal Fn_ As Long)
'Public Declare Function LLVMGetIntrinsicID Lib "llvm2.9.dll" (ByVal Fn_ As Long) As Long
'Public Declare Function LLVMGetFunctionCallConv Lib "llvm2.9.dll" (ByVal Fn_ As Long) As Long
'Public Declare Sub LLVMSetFunctionCallConv Lib "llvm2.9.dll" (ByVal Fn_ As Long, ByVal CC_ As Long)
'Public Declare Function LLVMGetGC Lib "llvm2.9.dll" (ByVal Fn_ As Long) As Long 'Byte*
'Public Declare Sub LLVMSetGC Lib "llvm2.9.dll" (ByVal Fn_ As Long, ByVal Name_ As String)
'Public Declare Sub LLVMAddFunctionAttr Lib "llvm2.9.dll" (ByVal Fn_ As Long, ByVal PA_ As LLVMAttribute)
'Public Declare Function LLVMGetFunctionAttr Lib "llvm2.9.dll" (ByVal Fn_ As Long) As LLVMAttribute
'Public Declare Sub LLVMRemoveFunctionAttr Lib "llvm2.9.dll" (ByVal Fn_ As Long, ByVal PA_ As LLVMAttribute)
'Public Declare Function LLVMCountParams Lib "llvm2.9.dll" (ByVal Fn_ As Long) As Long
'Public Declare Sub LLVMGetParams Lib "llvm2.9.dll" (ByVal Fn_ As Long, ByRef Params_ As Long)
'Public Declare Function LLVMGetParam Lib "llvm2.9.dll" (ByVal Fn_ As Long, ByVal Index_ As Long) As Long
'Public Declare Function LLVMGetParamParent Lib "llvm2.9.dll" (ByVal Inst_ As Long) As Long
'Public Declare Function LLVMGetFirstParam Lib "llvm2.9.dll" (ByVal Fn_ As Long) As Long
'Public Declare Function LLVMGetLastParam Lib "llvm2.9.dll" (ByVal Fn_ As Long) As Long
'Public Declare Function LLVMGetNextParam Lib "llvm2.9.dll" (ByVal Arg_ As Long) As Long
'Public Declare Function LLVMGetPreviousParam Lib "llvm2.9.dll" (ByVal Arg_ As Long) As Long
'Public Declare Sub LLVMAddAttribute Lib "llvm2.9.dll" (ByVal Arg_ As Long, ByVal PA_ As LLVMAttribute)
'Public Declare Sub LLVMRemoveAttribute Lib "llvm2.9.dll" (ByVal Arg_ As Long, ByVal PA_ As LLVMAttribute)
'Public Declare Function LLVMGetAttribute Lib "llvm2.9.dll" (ByVal Arg_ As Long) As LLVMAttribute
'Public Declare Sub LLVMSetParamAlignment Lib "llvm2.9.dll" (ByVal Arg_ As Long, ByVal align_ As Long)
'Public Declare Function LLVMBasicBlockAsValue Lib "llvm2.9.dll" (ByVal BB_ As Long) As Long
'Public Declare Function LLVMValueIsBasicBlock Lib "llvm2.9.dll" (ByVal Val_ As Long) As Long
'Public Declare Function LLVMValueAsBasicBlock Lib "llvm2.9.dll" (ByVal Val_ As Long) As Long
'Public Declare Function LLVMGetBasicBlockParent Lib "llvm2.9.dll" (ByVal BB_ As Long) As Long
'Public Declare Function LLVMCountBasicBlocks Lib "llvm2.9.dll" (ByVal Fn_ As Long) As Long
'Public Declare Sub LLVMGetBasicBlocks Lib "llvm2.9.dll" (ByVal Fn_ As Long, ByRef BasicBlocks_ As Long)
'Public Declare Function LLVMGetFirstBasicBlock Lib "llvm2.9.dll" (ByVal Fn_ As Long) As Long
'Public Declare Function LLVMGetLastBasicBlock Lib "llvm2.9.dll" (ByVal Fn_ As Long) As Long
'Public Declare Function LLVMGetNextBasicBlock Lib "llvm2.9.dll" (ByVal BB_ As Long) As Long
'Public Declare Function LLVMGetPreviousBasicBlock Lib "llvm2.9.dll" (ByVal BB_ As Long) As Long
'Public Declare Function LLVMGetEntryBasicBlock Lib "llvm2.9.dll" (ByVal Fn_ As Long) As Long
'Public Declare Function LLVMAppendBasicBlockInContext Lib "llvm2.9.dll" (ByVal C_ As Long, ByVal Fn_ As Long, ByVal Name_ As String) As Long
'Public Declare Function LLVMInsertBasicBlockInContext Lib "llvm2.9.dll" (ByVal C_ As Long, ByVal BB_ As Long, ByVal Name_ As String) As Long
'Public Declare Function LLVMAppendBasicBlock Lib "llvm2.9.dll" (ByVal Fn_ As Long, ByVal Name_ As String) As Long
'Public Declare Function LLVMInsertBasicBlock Lib "llvm2.9.dll" (ByVal InsertBeforeBB_ As Long, ByVal Name_ As String) As Long
'Public Declare Sub LLVMDeleteBasicBlock Lib "llvm2.9.dll" (ByVal BB_ As Long)
'Public Declare Sub LLVMMoveBasicBlockBefore Lib "llvm2.9.dll" (ByVal BB_ As Long, ByVal MovePos_ As Long)
'Public Declare Sub LLVMMoveBasicBlockAfter Lib "llvm2.9.dll" (ByVal BB_ As Long, ByVal MovePos_ As Long)
'Public Declare Function LLVMGetInstructionParent Lib "llvm2.9.dll" (ByVal Inst_ As Long) As Long
'Public Declare Function LLVMGetFirstInstruction Lib "llvm2.9.dll" (ByVal BB_ As Long) As Long
'Public Declare Function LLVMGetLastInstruction Lib "llvm2.9.dll" (ByVal BB_ As Long) As Long
'Public Declare Function LLVMGetNextInstruction Lib "llvm2.9.dll" (ByVal Inst_ As Long) As Long
'Public Declare Function LLVMGetPreviousInstruction Lib "llvm2.9.dll" (ByVal Inst_ As Long) As Long
'Public Declare Sub LLVMSetInstructionCallConv Lib "llvm2.9.dll" (ByVal Instr_ As Long, ByVal CC_ As Long)
'Public Declare Function LLVMGetInstructionCallConv Lib "llvm2.9.dll" (ByVal Instr_ As Long) As Long
'Public Declare Sub LLVMAddInstrAttribute Lib "llvm2.9.dll" (ByVal Instr_ As Long, ByVal index_ As Long, ByVal a3_ As LLVMAttribute)
'Public Declare Sub LLVMRemoveInstrAttribute Lib "llvm2.9.dll" (ByVal Instr_ As Long, ByVal index_ As Long, ByVal a3_ As LLVMAttribute)
'Public Declare Sub LLVMSetInstrParamAlignment Lib "llvm2.9.dll" (ByVal Instr_ As Long, ByVal index_ As Long, ByVal align_ As Long)
'Public Declare Function LLVMIsTailCall Lib "llvm2.9.dll" (ByVal CallInst_ As Long) As Long
'Public Declare Sub LLVMSetTailCall Lib "llvm2.9.dll" (ByVal CallInst_ As Long, ByVal IsTailCall_ As Long)
'Public Declare Sub LLVMAddIncoming Lib "llvm2.9.dll" (ByVal PhiNode_ As Long, ByRef IncomingValues_ As Long, ByRef IncomingBlocks_ As Long, ByVal Count_ As Long)
'Public Declare Function LLVMCountIncoming Lib "llvm2.9.dll" (ByVal PhiNode_ As Long) As Long
'Public Declare Function LLVMGetIncomingValue Lib "llvm2.9.dll" (ByVal PhiNode_ As Long, ByVal Index_ As Long) As Long
'Public Declare Function LLVMGetIncomingBlock Lib "llvm2.9.dll" (ByVal PhiNode_ As Long, ByVal Index_ As Long) As Long
'Public Declare Function LLVMCreateBuilderInContext Lib "llvm2.9.dll" (ByVal C_ As Long) As Long
'Public Declare Function LLVMCreateBuilder Lib "llvm2.9.dll" () As Long
'Public Declare Sub LLVMPositionBuilder Lib "llvm2.9.dll" (ByVal Builder_ As Long, ByVal Block_ As Long, ByVal Instr_ As Long)
'Public Declare Sub LLVMPositionBuilderBefore Lib "llvm2.9.dll" (ByVal Builder_ As Long, ByVal Instr_ As Long)
'Public Declare Sub LLVMPositionBuilderAtEnd Lib "llvm2.9.dll" (ByVal Builder_ As Long, ByVal Block_ As Long)
'Public Declare Function LLVMGetInsertBlock Lib "llvm2.9.dll" (ByVal Builder_ As Long) As Long
'Public Declare Sub LLVMClearInsertionPosition Lib "llvm2.9.dll" (ByVal Builder_ As Long)
'Public Declare Sub LLVMInsertIntoBuilder Lib "llvm2.9.dll" (ByVal Builder_ As Long, ByVal Instr_ As Long)
'Public Declare Sub LLVMInsertIntoBuilderWithName Lib "llvm2.9.dll" (ByVal Builder_ As Long, ByVal Instr_ As Long, ByVal Name_ As String)
'Public Declare Sub LLVMDisposeBuilder Lib "llvm2.9.dll" (ByVal Builder_ As Long)
'Public Declare Sub LLVMSetCurrentDebugLocation Lib "llvm2.9.dll" (ByVal Builder_ As Long, ByVal L_ As Long)
'Public Declare Function LLVMGetCurrentDebugLocation Lib "llvm2.9.dll" (ByVal Builder_ As Long) As Long
'Public Declare Sub LLVMSetInstDebugLocation Lib "llvm2.9.dll" (ByVal Builder_ As Long, ByVal Inst_ As Long)
'Public Declare Function LLVMBuildRetVoid Lib "llvm2.9.dll" (ByVal a1_ As Long) As Long
'Public Declare Function LLVMBuildRet Lib "llvm2.9.dll" (ByVal a1_ As Long, ByVal V_ As Long) As Long
'Public Declare Function LLVMBuildAggregateRet Lib "llvm2.9.dll" (ByVal a1_ As Long, ByRef RetVals_ As Long, ByVal N_ As Long) As Long
'Public Declare Function LLVMBuildBr Lib "llvm2.9.dll" (ByVal a1_ As Long, ByVal Dest_ As Long) As Long
'Public Declare Function LLVMBuildCondBr Lib "llvm2.9.dll" (ByVal a1_ As Long, ByVal If_ As Long, ByVal Then_ As Long, ByVal Else_ As Long) As Long
'Public Declare Function LLVMBuildSwitch Lib "llvm2.9.dll" (ByVal a1_ As Long, ByVal V_ As Long, ByVal Else_ As Long, ByVal NumCases_ As Long) As Long
'Public Declare Function LLVMBuildIndirectBr Lib "llvm2.9.dll" (ByVal B_ As Long, ByVal Addr_ As Long, ByVal NumDests_ As Long) As Long
'Public Declare Function LLVMBuildInvoke Lib "llvm2.9.dll" (ByVal a1_ As Long, ByVal Fn_ As Long, ByRef Args_ As Long, ByVal NumArgs_ As Long, ByVal Then_ As Long, ByVal Catch_ As Long, ByVal Name_ As String) As Long
'Public Declare Function LLVMBuildUnwind Lib "llvm2.9.dll" (ByVal a1_ As Long) As Long
'Public Declare Function LLVMBuildUnreachable Lib "llvm2.9.dll" (ByVal a1_ As Long) As Long
'Public Declare Sub LLVMAddCase Lib "llvm2.9.dll" (ByVal Switch_ As Long, ByVal OnVal_ As Long, ByVal Dest_ As Long)
'Public Declare Sub LLVMAddDestination Lib "llvm2.9.dll" (ByVal IndirectBr_ As Long, ByVal Dest_ As Long)
'Public Declare Function LLVMBuildAdd Lib "llvm2.9.dll" (ByVal a1_ As Long, ByVal LHS_ As Long, ByVal RHS_ As Long, ByVal Name_ As String) As Long
'Public Declare Function LLVMBuildNSWAdd Lib "llvm2.9.dll" (ByVal a1_ As Long, ByVal LHS_ As Long, ByVal RHS_ As Long, ByVal Name_ As String) As Long
'Public Declare Function LLVMBuildNUWAdd Lib "llvm2.9.dll" (ByVal a1_ As Long, ByVal LHS_ As Long, ByVal RHS_ As Long, ByVal Name_ As String) As Long
'Public Declare Function LLVMBuildFAdd Lib "llvm2.9.dll" (ByVal a1_ As Long, ByVal LHS_ As Long, ByVal RHS_ As Long, ByVal Name_ As String) As Long
'Public Declare Function LLVMBuildSub Lib "llvm2.9.dll" (ByVal a1_ As Long, ByVal LHS_ As Long, ByVal RHS_ As Long, ByVal Name_ As String) As Long
'Public Declare Function LLVMBuildNSWSub Lib "llvm2.9.dll" (ByVal a1_ As Long, ByVal LHS_ As Long, ByVal RHS_ As Long, ByVal Name_ As String) As Long
'Public Declare Function LLVMBuildNUWSub Lib "llvm2.9.dll" (ByVal a1_ As Long, ByVal LHS_ As Long, ByVal RHS_ As Long, ByVal Name_ As String) As Long
'Public Declare Function LLVMBuildFSub Lib "llvm2.9.dll" (ByVal a1_ As Long, ByVal LHS_ As Long, ByVal RHS_ As Long, ByVal Name_ As String) As Long
'Public Declare Function LLVMBuildMul Lib "llvm2.9.dll" (ByVal a1_ As Long, ByVal LHS_ As Long, ByVal RHS_ As Long, ByVal Name_ As String) As Long
'Public Declare Function LLVMBuildNSWMul Lib "llvm2.9.dll" (ByVal a1_ As Long, ByVal LHS_ As Long, ByVal RHS_ As Long, ByVal Name_ As String) As Long
'Public Declare Function LLVMBuildNUWMul Lib "llvm2.9.dll" (ByVal a1_ As Long, ByVal LHS_ As Long, ByVal RHS_ As Long, ByVal Name_ As String) As Long
'Public Declare Function LLVMBuildFMul Lib "llvm2.9.dll" (ByVal a1_ As Long, ByVal LHS_ As Long, ByVal RHS_ As Long, ByVal Name_ As String) As Long
'Public Declare Function LLVMBuildUDiv Lib "llvm2.9.dll" (ByVal a1_ As Long, ByVal LHS_ As Long, ByVal RHS_ As Long, ByVal Name_ As String) As Long
'Public Declare Function LLVMBuildSDiv Lib "llvm2.9.dll" (ByVal a1_ As Long, ByVal LHS_ As Long, ByVal RHS_ As Long, ByVal Name_ As String) As Long
'Public Declare Function LLVMBuildExactSDiv Lib "llvm2.9.dll" (ByVal a1_ As Long, ByVal LHS_ As Long, ByVal RHS_ As Long, ByVal Name_ As String) As Long
'Public Declare Function LLVMBuildFDiv Lib "llvm2.9.dll" (ByVal a1_ As Long, ByVal LHS_ As Long, ByVal RHS_ As Long, ByVal Name_ As String) As Long
'Public Declare Function LLVMBuildURem Lib "llvm2.9.dll" (ByVal a1_ As Long, ByVal LHS_ As Long, ByVal RHS_ As Long, ByVal Name_ As String) As Long
'Public Declare Function LLVMBuildSRem Lib "llvm2.9.dll" (ByVal a1_ As Long, ByVal LHS_ As Long, ByVal RHS_ As Long, ByVal Name_ As String) As Long
'Public Declare Function LLVMBuildFRem Lib "llvm2.9.dll" (ByVal a1_ As Long, ByVal LHS_ As Long, ByVal RHS_ As Long, ByVal Name_ As String) As Long
'Public Declare Function LLVMBuildShl Lib "llvm2.9.dll" (ByVal a1_ As Long, ByVal LHS_ As Long, ByVal RHS_ As Long, ByVal Name_ As String) As Long
'Public Declare Function LLVMBuildLShr Lib "llvm2.9.dll" (ByVal a1_ As Long, ByVal LHS_ As Long, ByVal RHS_ As Long, ByVal Name_ As String) As Long
'Public Declare Function LLVMBuildAShr Lib "llvm2.9.dll" (ByVal a1_ As Long, ByVal LHS_ As Long, ByVal RHS_ As Long, ByVal Name_ As String) As Long
'Public Declare Function LLVMBuildAnd Lib "llvm2.9.dll" (ByVal a1_ As Long, ByVal LHS_ As Long, ByVal RHS_ As Long, ByVal Name_ As String) As Long
'Public Declare Function LLVMBuildOr Lib "llvm2.9.dll" (ByVal a1_ As Long, ByVal LHS_ As Long, ByVal RHS_ As Long, ByVal Name_ As String) As Long
'Public Declare Function LLVMBuildXor Lib "llvm2.9.dll" (ByVal a1_ As Long, ByVal LHS_ As Long, ByVal RHS_ As Long, ByVal Name_ As String) As Long
'Public Declare Function LLVMBuildBinOp Lib "llvm2.9.dll" (ByVal B_ As Long, ByVal Op_ As LLVMOpcode, ByVal LHS_ As Long, ByVal RHS_ As Long, ByVal Name_ As String) As Long
'Public Declare Function LLVMBuildNeg Lib "llvm2.9.dll" (ByVal a1_ As Long, ByVal V_ As Long, ByVal Name_ As String) As Long
'Public Declare Function LLVMBuildNSWNeg Lib "llvm2.9.dll" (ByVal B_ As Long, ByVal V_ As Long, ByVal Name_ As String) As Long
'Public Declare Function LLVMBuildNUWNeg Lib "llvm2.9.dll" (ByVal B_ As Long, ByVal V_ As Long, ByVal Name_ As String) As Long
'Public Declare Function LLVMBuildFNeg Lib "llvm2.9.dll" (ByVal a1_ As Long, ByVal V_ As Long, ByVal Name_ As String) As Long
'Public Declare Function LLVMBuildNot Lib "llvm2.9.dll" (ByVal a1_ As Long, ByVal V_ As Long, ByVal Name_ As String) As Long
'Public Declare Function LLVMBuildMalloc Lib "llvm2.9.dll" (ByVal a1_ As Long, ByVal Ty_ As Long, ByVal Name_ As String) As Long
'Public Declare Function LLVMBuildArrayMalloc Lib "llvm2.9.dll" (ByVal a1_ As Long, ByVal Ty_ As Long, ByVal Val_ As Long, ByVal Name_ As String) As Long
'Public Declare Function LLVMBuildAlloca Lib "llvm2.9.dll" (ByVal a1_ As Long, ByVal Ty_ As Long, ByVal Name_ As String) As Long
'Public Declare Function LLVMBuildArrayAlloca Lib "llvm2.9.dll" (ByVal a1_ As Long, ByVal Ty_ As Long, ByVal Val_ As Long, ByVal Name_ As String) As Long
'Public Declare Function LLVMBuildFree Lib "llvm2.9.dll" (ByVal a1_ As Long, ByVal PointerVal_ As Long) As Long
'Public Declare Function LLVMBuildLoad Lib "llvm2.9.dll" (ByVal a1_ As Long, ByVal PointerVal_ As Long, ByVal Name_ As String) As Long
'Public Declare Function LLVMBuildStore Lib "llvm2.9.dll" (ByVal a1_ As Long, ByVal Val_ As Long, ByVal Ptr_ As Long) As Long
'Public Declare Function LLVMBuildGEP Lib "llvm2.9.dll" (ByVal B_ As Long, ByVal Pointer_ As Long, ByRef Indices_ As Long, ByVal NumIndices_ As Long, ByVal Name_ As String) As Long
'Public Declare Function LLVMBuildInBoundsGEP Lib "llvm2.9.dll" (ByVal B_ As Long, ByVal Pointer_ As Long, ByRef Indices_ As Long, ByVal NumIndices_ As Long, ByVal Name_ As String) As Long
'Public Declare Function LLVMBuildStructGEP Lib "llvm2.9.dll" (ByVal B_ As Long, ByVal Pointer_ As Long, ByVal Idx_ As Long, ByVal Name_ As String) As Long
'Public Declare Function LLVMBuildGlobalString Lib "llvm2.9.dll" (ByVal B_ As Long, ByVal Str_ As String, ByVal Name_ As String) As Long
'Public Declare Function LLVMBuildGlobalStringPtr Lib "llvm2.9.dll" (ByVal B_ As Long, ByVal Str_ As String, ByVal Name_ As String) As Long
'Public Declare Function LLVMBuildTrunc Lib "llvm2.9.dll" (ByVal a1_ As Long, ByVal Val_ As Long, ByVal DestTy_ As Long, ByVal Name_ As String) As Long
'Public Declare Function LLVMBuildZExt Lib "llvm2.9.dll" (ByVal a1_ As Long, ByVal Val_ As Long, ByVal DestTy_ As Long, ByVal Name_ As String) As Long
'Public Declare Function LLVMBuildSExt Lib "llvm2.9.dll" (ByVal a1_ As Long, ByVal Val_ As Long, ByVal DestTy_ As Long, ByVal Name_ As String) As Long
'Public Declare Function LLVMBuildFPToUI Lib "llvm2.9.dll" (ByVal a1_ As Long, ByVal Val_ As Long, ByVal DestTy_ As Long, ByVal Name_ As String) As Long
'Public Declare Function LLVMBuildFPToSI Lib "llvm2.9.dll" (ByVal a1_ As Long, ByVal Val_ As Long, ByVal DestTy_ As Long, ByVal Name_ As String) As Long
'Public Declare Function LLVMBuildUIToFP Lib "llvm2.9.dll" (ByVal a1_ As Long, ByVal Val_ As Long, ByVal DestTy_ As Long, ByVal Name_ As String) As Long
'Public Declare Function LLVMBuildSIToFP Lib "llvm2.9.dll" (ByVal a1_ As Long, ByVal Val_ As Long, ByVal DestTy_ As Long, ByVal Name_ As String) As Long
'Public Declare Function LLVMBuildFPTrunc Lib "llvm2.9.dll" (ByVal a1_ As Long, ByVal Val_ As Long, ByVal DestTy_ As Long, ByVal Name_ As String) As Long
'Public Declare Function LLVMBuildFPExt Lib "llvm2.9.dll" (ByVal a1_ As Long, ByVal Val_ As Long, ByVal DestTy_ As Long, ByVal Name_ As String) As Long
'Public Declare Function LLVMBuildPtrToInt Lib "llvm2.9.dll" (ByVal a1_ As Long, ByVal Val_ As Long, ByVal DestTy_ As Long, ByVal Name_ As String) As Long
'Public Declare Function LLVMBuildIntToPtr Lib "llvm2.9.dll" (ByVal a1_ As Long, ByVal Val_ As Long, ByVal DestTy_ As Long, ByVal Name_ As String) As Long
'Public Declare Function LLVMBuildBitCast Lib "llvm2.9.dll" (ByVal a1_ As Long, ByVal Val_ As Long, ByVal DestTy_ As Long, ByVal Name_ As String) As Long
'Public Declare Function LLVMBuildZExtOrBitCast Lib "llvm2.9.dll" (ByVal a1_ As Long, ByVal Val_ As Long, ByVal DestTy_ As Long, ByVal Name_ As String) As Long
'Public Declare Function LLVMBuildSExtOrBitCast Lib "llvm2.9.dll" (ByVal a1_ As Long, ByVal Val_ As Long, ByVal DestTy_ As Long, ByVal Name_ As String) As Long
'Public Declare Function LLVMBuildTruncOrBitCast Lib "llvm2.9.dll" (ByVal a1_ As Long, ByVal Val_ As Long, ByVal DestTy_ As Long, ByVal Name_ As String) As Long
'Public Declare Function LLVMBuildCast Lib "llvm2.9.dll" (ByVal B_ As Long, ByVal Op_ As LLVMOpcode, ByVal Val_ As Long, ByVal DestTy_ As Long, ByVal Name_ As String) As Long
'Public Declare Function LLVMBuildPointerCast Lib "llvm2.9.dll" (ByVal a1_ As Long, ByVal Val_ As Long, ByVal DestTy_ As Long, ByVal Name_ As String) As Long
'Public Declare Function LLVMBuildIntCast Lib "llvm2.9.dll" (ByVal a1_ As Long, ByVal Val_ As Long, ByVal DestTy_ As Long, ByVal Name_ As String) As Long
'Public Declare Function LLVMBuildFPCast Lib "llvm2.9.dll" (ByVal a1_ As Long, ByVal Val_ As Long, ByVal DestTy_ As Long, ByVal Name_ As String) As Long
'Public Declare Function LLVMBuildICmp Lib "llvm2.9.dll" (ByVal a1_ As Long, ByVal Op_ As LLVMIntPredicate, ByVal LHS_ As Long, ByVal RHS_ As Long, ByVal Name_ As String) As Long
'Public Declare Function LLVMBuildFCmp Lib "llvm2.9.dll" (ByVal a1_ As Long, ByVal Op_ As LLVMRealPredicate, ByVal LHS_ As Long, ByVal RHS_ As Long, ByVal Name_ As String) As Long
'Public Declare Function LLVMBuildPhi Lib "llvm2.9.dll" (ByVal a1_ As Long, ByVal Ty_ As Long, ByVal Name_ As String) As Long
'Public Declare Function LLVMBuildCall Lib "llvm2.9.dll" (ByVal a1_ As Long, ByVal Fn_ As Long, ByRef Args_ As Long, ByVal NumArgs_ As Long, ByVal Name_ As String) As Long
'Public Declare Function LLVMBuildSelect Lib "llvm2.9.dll" (ByVal a1_ As Long, ByVal If_ As Long, ByVal Then_ As Long, ByVal Else_ As Long, ByVal Name_ As String) As Long
'Public Declare Function LLVMBuildVAArg Lib "llvm2.9.dll" (ByVal a1_ As Long, ByVal List_ As Long, ByVal Ty_ As Long, ByVal Name_ As String) As Long
'Public Declare Function LLVMBuildExtractElement Lib "llvm2.9.dll" (ByVal a1_ As Long, ByVal VecVal_ As Long, ByVal Index_ As Long, ByVal Name_ As String) As Long
'Public Declare Function LLVMBuildInsertElement Lib "llvm2.9.dll" (ByVal a1_ As Long, ByVal VecVal_ As Long, ByVal EltVal_ As Long, ByVal Index_ As Long, ByVal Name_ As String) As Long
'Public Declare Function LLVMBuildShuffleVector Lib "llvm2.9.dll" (ByVal a1_ As Long, ByVal V1_ As Long, ByVal V2_ As Long, ByVal Mask_ As Long, ByVal Name_ As String) As Long
'Public Declare Function LLVMBuildExtractValue Lib "llvm2.9.dll" (ByVal a1_ As Long, ByVal AggVal_ As Long, ByVal Index_ As Long, ByVal Name_ As String) As Long
'Public Declare Function LLVMBuildInsertValue Lib "llvm2.9.dll" (ByVal a1_ As Long, ByVal AggVal_ As Long, ByVal EltVal_ As Long, ByVal Index_ As Long, ByVal Name_ As String) As Long
'Public Declare Function LLVMBuildIsNull Lib "llvm2.9.dll" (ByVal a1_ As Long, ByVal Val_ As Long, ByVal Name_ As String) As Long
'Public Declare Function LLVMBuildIsNotNull Lib "llvm2.9.dll" (ByVal a1_ As Long, ByVal Val_ As Long, ByVal Name_ As String) As Long
'Public Declare Function LLVMBuildPtrDiff Lib "llvm2.9.dll" (ByVal a1_ As Long, ByVal LHS_ As Long, ByVal RHS_ As Long, ByVal Name_ As String) As Long
'Public Declare Function LLVMCreateModuleProviderForExistingModule Lib "llvm2.9.dll" (ByVal M_ As Long) As Long
'Public Declare Sub LLVMDisposeModuleProvider Lib "llvm2.9.dll" (ByVal M_ As Long)
'Public Declare Function LLVMCreateMemoryBufferWithContentsOfFile Lib "llvm2.9.dll" (ByVal Path_ As String, ByRef OutMemBuf_ As Long, ByRef OutMessage_ As Long) As Long
'Public Declare Function LLVMCreateMemoryBufferWithSTDIN Lib "llvm2.9.dll" (ByRef OutMemBuf_ As Long, ByRef OutMessage_ As Long) As Long
'Public Declare Sub LLVMDisposeMemoryBuffer Lib "llvm2.9.dll" (ByVal MemBuf_ As Long)
'Public Declare Function LLVMGetGlobalPassRegistry Lib "llvm2.9.dll" () As Long
'Public Declare Function LLVMCreatePassManager Lib "llvm2.9.dll" () As Long
'Public Declare Function LLVMCreateFunctionPassManagerForModule Lib "llvm2.9.dll" (ByVal M_ As Long) As Long
'Public Declare Function LLVMCreateFunctionPassManager Lib "llvm2.9.dll" (ByVal MP_ As Long) As Long
'Public Declare Function LLVMRunPassManager Lib "llvm2.9.dll" (ByVal PM_ As Long, ByVal M_ As Long) As Long
'Public Declare Function LLVMInitializeFunctionPassManager Lib "llvm2.9.dll" (ByVal FPM_ As Long) As Long
'Public Declare Function LLVMRunFunctionPassManager Lib "llvm2.9.dll" (ByVal FPM_ As Long, ByVal F_ As Long) As Long
'Public Declare Function LLVMFinalizeFunctionPassManager Lib "llvm2.9.dll" (ByVal FPM_ As Long) As Long
'Public Declare Sub LLVMDisposePassManager Lib "llvm2.9.dll" (ByVal PM_ As Long)
'--- F:\Projects\llvm-2.9\include\llvm-c\test.h ---
'--- F:\Projects\llvm-2.9\include\llvm-c\Analysis.h ---
'Public Declare Function LLVMVerifyModule Lib "llvm2.9.dll" (ByVal M_ As Long, ByVal Action_ As LLVMVerifierFailureAction, ByRef OutMessage_ As Long) As Long
'Public Declare Function LLVMVerifyFunction Lib "llvm2.9.dll" (ByVal Fn_ As Long, ByVal Action_ As LLVMVerifierFailureAction) As Long
'Public Declare Sub LLVMViewFunctionCFG Lib "llvm2.9.dll" (ByVal Fn_ As Long)
'Public Declare Sub LLVMViewFunctionCFGOnly Lib "llvm2.9.dll" (ByVal Fn_ As Long)
'--- F:\Projects\llvm-2.9\include\llvm-c\test.h ---
'--- F:\Projects\llvm-2.9\include\llvm-c\BitReader.h ---
'Public Declare Function LLVMParseBitcode Lib "llvm2.9.dll" (ByVal MemBuf_ As Long, ByRef OutModule_ As Long, ByRef OutMessage_ As Long) As Long
'Public Declare Function LLVMParseBitcodeInContext Lib "llvm2.9.dll" (ByVal ContextRef_ As Long, ByVal MemBuf_ As Long, ByRef OutModule_ As Long, ByRef OutMessage_ As Long) As Long
'Public Declare Function LLVMGetBitcodeModuleInContext Lib "llvm2.9.dll" (ByVal ContextRef_ As Long, ByVal MemBuf_ As Long, ByRef OutM_ As Long, ByRef OutMessage_ As Long) As Long
'Public Declare Function LLVMGetBitcodeModule Lib "llvm2.9.dll" (ByVal MemBuf_ As Long, ByRef OutM_ As Long, ByRef OutMessage_ As Long) As Long
'Public Declare Function LLVMGetBitcodeModuleProviderInContext Lib "llvm2.9.dll" (ByVal ContextRef_ As Long, ByVal MemBuf_ As Long, ByRef OutMP_ As Long, ByRef OutMessage_ As Long) As Long
'Public Declare Function LLVMGetBitcodeModuleProvider Lib "llvm2.9.dll" (ByVal MemBuf_ As Long, ByRef OutMP_ As Long, ByRef OutMessage_ As Long) As Long
'--- F:\Projects\llvm-2.9\include\llvm-c\test.h ---
'--- F:\Projects\llvm-2.9\include\llvm-c\BitWriter.h ---
'Public Declare Function LLVMWriteBitcodeToFile Lib "llvm2.9.dll" (ByVal M_ As Long, ByVal Path_ As String) As Long
'Public Declare Function LLVMWriteBitcodeToFD Lib "llvm2.9.dll" (ByVal M_ As Long, ByVal FD_ As Long, ByVal ShouldClose_ As Long, ByVal Unbuffered_ As Long) As Long
'Public Declare Function LLVMWriteBitcodeToFileHandle Lib "llvm2.9.dll" (ByVal M_ As Long, ByVal Handle_ As Long) As Long
'--- F:\Projects\llvm-2.9\include\llvm-c\test.h ---
'--- F:\Projects\llvm-2.9\include\llvm-c\EnhancedDisassembly.h ---
'Public Declare Function EDGetDisassembler Lib "llvm2.9.dll" (ByRef disassembler_ As Long, ByVal triple_ As String, ByVal syntax_ As Long) As Long
'Public Declare Function EDGetRegisterName Lib "llvm2.9.dll" (ByRef regName_ As Long, ByRef disassembler_ As Any, ByVal regID_ As Long) As Long
'Public Declare Function EDRegisterIsStackPointer Lib "llvm2.9.dll" (ByRef disassembler_ As Any, ByVal regID_ As Long) As Long
'Public Declare Function EDRegisterIsProgramCounter Lib "llvm2.9.dll" (ByRef disassembler_ As Any, ByVal regID_ As Long) As Long
'Public Declare Function EDCreateInsts Lib "llvm2.9.dll" (ByRef insts_ As Long, ByVal count_ As Long, ByRef disassembler_ As Any, ByVal byteReader_ As Long, ByVal address_ As Currency, ByRef arg_ As Any) As Long
'Public Declare Sub EDReleaseInst Lib "llvm2.9.dll" (ByRef inst_ As Any)
'Public Declare Function EDInstByteSize Lib "llvm2.9.dll" (ByRef inst_ As Any) As Long
'Public Declare Function EDGetInstString Lib "llvm2.9.dll" (ByRef buf_ As Long, ByRef inst_ As Any) As Long
'Public Declare Function EDInstID Lib "llvm2.9.dll" (ByRef instID_ As Long, ByRef inst_ As Any) As Long
'Public Declare Function EDInstIsBranch Lib "llvm2.9.dll" (ByRef inst_ As Any) As Long
'Public Declare Function EDInstIsMove Lib "llvm2.9.dll" (ByRef inst_ As Any) As Long
'Public Declare Function EDBranchTargetID Lib "llvm2.9.dll" (ByRef inst_ As Any) As Long
'Public Declare Function EDMoveSourceID Lib "llvm2.9.dll" (ByRef inst_ As Any) As Long
'Public Declare Function EDMoveTargetID Lib "llvm2.9.dll" (ByRef inst_ As Any) As Long
'Public Declare Function EDNumTokens Lib "llvm2.9.dll" (ByRef inst_ As Any) As Long
'Public Declare Function EDGetToken Lib "llvm2.9.dll" (ByRef token_ As Long, ByRef inst_ As Any, ByVal index_ As Long) As Long
'Public Declare Function EDGetTokenString Lib "llvm2.9.dll" (ByRef buf_ As Long, ByRef token_ As Any) As Long
'Public Declare Function EDOperandIndexForToken Lib "llvm2.9.dll" (ByRef token_ As Any) As Long
'Public Declare Function EDTokenIsWhitespace Lib "llvm2.9.dll" (ByRef token_ As Any) As Long
'Public Declare Function EDTokenIsPunctuation Lib "llvm2.9.dll" (ByRef token_ As Any) As Long
'Public Declare Function EDTokenIsOpcode Lib "llvm2.9.dll" (ByRef token_ As Any) As Long
'Public Declare Function EDTokenIsLiteral Lib "llvm2.9.dll" (ByRef token_ As Any) As Long
'Public Declare Function EDTokenIsRegister Lib "llvm2.9.dll" (ByRef token_ As Any) As Long
'Public Declare Function EDTokenIsNegativeLiteral Lib "llvm2.9.dll" (ByRef token_ As Any) As Long
'Public Declare Function EDLiteralTokenAbsoluteValue Lib "llvm2.9.dll" (ByRef value_ As Currency, ByRef token_ As Any) As Long
'Public Declare Function EDRegisterTokenValue Lib "llvm2.9.dll" (ByRef registerID_ As Long, ByRef token_ As Any) As Long
'Public Declare Function EDNumOperands Lib "llvm2.9.dll" (ByRef inst_ As Any) As Long
'Public Declare Function EDGetOperand Lib "llvm2.9.dll" (ByRef operand_ As Long, ByRef inst_ As Any, ByVal index_ As Long) As Long
'Public Declare Function EDOperandIsRegister Lib "llvm2.9.dll" (ByRef operand_ As Any) As Long
'Public Declare Function EDOperandIsImmediate Lib "llvm2.9.dll" (ByRef operand_ As Any) As Long
'Public Declare Function EDOperandIsMemory Lib "llvm2.9.dll" (ByRef operand_ As Any) As Long
'Public Declare Function EDRegisterOperandValue Lib "llvm2.9.dll" (ByRef value_ As Long, ByRef operand_ As Any) As Long
'Public Declare Function EDImmediateOperandValue Lib "llvm2.9.dll" (ByRef value_ As Currency, ByRef operand_ As Any) As Long
'Public Declare Function EDEvaluateOperand Lib "llvm2.9.dll" (ByRef result_ As Currency, ByRef operand_ As Any, ByVal regReader_ As Long, ByRef arg_ As Any) As Long
'--- F:\Projects\llvm-2.9\include\llvm-c\test.h ---
'--- F:\Projects\llvm-2.9\include\llvm-c\ExecutionEngine.h ---
'--- F:\Projects\llvm-2.9\include\llvm-c\Target.h ---
'--- F:\Projects\llvm-2.9\vs2008\include\llvm\Config\llvm-config.h ---
Public Const LLVM_HOSTTRIPLE As String = "i686-pc-win32"
Public Const LLVM_PREFIX As String = "D:/Program Files/LLVM"
'--- F:\Projects\llvm-2.9\include\llvm-c\Target.h ---
'--- F:\Projects\llvm-2.9\vs2008\include\llvm\Config\Targets.def ---
'Public Declare Sub LLVMInitializeAlphaTargetInfo Lib "llvm2.9.dll" ()
'Public Declare Sub LLVMInitializeARMTargetInfo Lib "llvm2.9.dll" ()
'Public Declare Sub LLVMInitializeBlackfinTargetInfo Lib "llvm2.9.dll" ()
'Public Declare Sub LLVMInitializeCBackendTargetInfo Lib "llvm2.9.dll" ()
'Public Declare Sub LLVMInitializeCellSPUTargetInfo Lib "llvm2.9.dll" ()
'Public Declare Sub LLVMInitializeCppBackendTargetInfo Lib "llvm2.9.dll" ()
'Public Declare Sub LLVMInitializeMipsTargetInfo Lib "llvm2.9.dll" ()
'Public Declare Sub LLVMInitializeMBlazeTargetInfo Lib "llvm2.9.dll" ()
'Public Declare Sub LLVMInitializeMSP430TargetInfo Lib "llvm2.9.dll" ()
'Public Declare Sub LLVMInitializePowerPCTargetInfo Lib "llvm2.9.dll" ()
'Public Declare Sub LLVMInitializePTXTargetInfo Lib "llvm2.9.dll" ()
'Public Declare Sub LLVMInitializeSparcTargetInfo Lib "llvm2.9.dll" ()
'Public Declare Sub LLVMInitializeSystemZTargetInfo Lib "llvm2.9.dll" ()
'Public Declare Sub LLVMInitializeX86TargetInfo Lib "llvm2.9.dll" ()
'Public Declare Sub LLVMInitializeXCoreTargetInfo Lib "llvm2.9.dll" ()
'--- F:\Projects\llvm-2.9\include\llvm-c\Target.h ---
'Public Declare Sub LLVMInitializeAllTargetInfos Lib "llvm2.9.dll" ()
'Public Declare Sub LLVMInitializeAllTargets Lib "llvm2.9.dll" ()
'Public Declare Function LLVMInitializeNativeTarget Lib "llvm2.9.dll" () As Long
'Public Declare Function LLVMCreateTargetData Lib "llvm2.9.dll" (ByVal StringRep_ As String) As Long
'Public Declare Sub LLVMAddTargetData Lib "llvm2.9.dll" (ByVal a1_ As Long, ByVal a2_ As Long)
'Public Declare Function LLVMCopyStringRepOfTargetData Lib "llvm2.9.dll" (ByVal a1_ As Long) As Long 'Byte*
'Public Declare Function LLVMByteOrder Lib "llvm2.9.dll" (ByVal a1_ As Long) As LLVMByteOrdering
'Public Declare Function LLVMPointerSize Lib "llvm2.9.dll" (ByVal a1_ As Long) As Long
'Public Declare Function LLVMIntPtrType Lib "llvm2.9.dll" (ByVal a1_ As Long) As Long
'Public Declare Function LLVMSizeOfTypeInBits Lib "llvm2.9.dll" (ByVal a1_ As Long, ByVal a2_ As Long) As Currency
'Public Declare Function LLVMStoreSizeOfType Lib "llvm2.9.dll" (ByVal a1_ As Long, ByVal a2_ As Long) As Currency
'Public Declare Function LLVMABISizeOfType Lib "llvm2.9.dll" (ByVal a1_ As Long, ByVal a2_ As Long) As Currency
'Public Declare Function LLVMABIAlignmentOfType Lib "llvm2.9.dll" (ByVal a1_ As Long, ByVal a2_ As Long) As Long
'Public Declare Function LLVMCallFrameAlignmentOfType Lib "llvm2.9.dll" (ByVal a1_ As Long, ByVal a2_ As Long) As Long
'Public Declare Function LLVMPreferredAlignmentOfType Lib "llvm2.9.dll" (ByVal a1_ As Long, ByVal a2_ As Long) As Long
'Public Declare Function LLVMPreferredAlignmentOfGlobal Lib "llvm2.9.dll" (ByVal a1_ As Long, ByVal GlobalVar_ As Long) As Long
'Public Declare Function LLVMElementAtOffset Lib "llvm2.9.dll" (ByVal a1_ As Long, ByVal StructTy_ As Long, ByVal Offset_ As Currency) As Long
'Public Declare Function LLVMOffsetOfElement Lib "llvm2.9.dll" (ByVal a1_ As Long, ByVal StructTy_ As Long, ByVal Element_ As Long) As Currency
'Public Declare Sub LLVMInvalidateStructLayout Lib "llvm2.9.dll" (ByVal a1_ As Long, ByVal StructTy_ As Long)
'Public Declare Sub LLVMDisposeTargetData Lib "llvm2.9.dll" (ByVal a1_ As Long)
'--- F:\Projects\llvm-2.9\include\llvm-c\ExecutionEngine.h ---
'Public Declare Sub LLVMLinkInJIT Lib "llvm2.9.dll" ()
'Public Declare Sub LLVMLinkInInterpreter Lib "llvm2.9.dll" ()
'Public Declare Function LLVMCreateGenericValueOfInt Lib "llvm2.9.dll" (ByVal Ty_ As Long, ByVal N_ As Currency, ByVal IsSigned_ As Long) As Long
'Public Declare Function LLVMCreateGenericValueOfPointer Lib "llvm2.9.dll" (ByRef P_ As Any) As Long
'Public Declare Function LLVMCreateGenericValueOfFloat Lib "llvm2.9.dll" (ByVal Ty_ As Long, ByVal N_ As Double) As Long
'Public Declare Function LLVMGenericValueIntWidth Lib "llvm2.9.dll" (ByVal GenValRef_ As Long) As Long
'Public Declare Function LLVMGenericValueToInt Lib "llvm2.9.dll" (ByVal GenVal_ As Long, ByVal IsSigned_ As Long) As Currency
'Public Declare Function LLVMGenericValueToPointer Lib "llvm2.9.dll" (ByVal GenVal_ As Long) As Long 'Void*
'Public Declare Function LLVMGenericValueToFloat Lib "llvm2.9.dll" (ByVal TyRef_ As Long, ByVal GenVal_ As Long) As Double
'Public Declare Sub LLVMDisposeGenericValue Lib "llvm2.9.dll" (ByVal GenVal_ As Long)
'Public Declare Function LLVMCreateExecutionEngineForModule Lib "llvm2.9.dll" (ByRef OutEE_ As Long, ByVal M_ As Long, ByRef OutError_ As Long) As Long
'Public Declare Function LLVMCreateInterpreterForModule Lib "llvm2.9.dll" (ByRef OutInterp_ As Long, ByVal M_ As Long, ByRef OutError_ As Long) As Long
'Public Declare Function LLVMCreateJITCompilerForModule Lib "llvm2.9.dll" (ByRef OutJIT_ As Long, ByVal M_ As Long, ByVal OptLevel_ As Long, ByRef OutError_ As Long) As Long
'Public Declare Function LLVMCreateExecutionEngine Lib "llvm2.9.dll" (ByRef OutEE_ As Long, ByVal MP_ As Long, ByRef OutError_ As Long) As Long
'Public Declare Function LLVMCreateInterpreter Lib "llvm2.9.dll" (ByRef OutInterp_ As Long, ByVal MP_ As Long, ByRef OutError_ As Long) As Long
'Public Declare Function LLVMCreateJITCompiler Lib "llvm2.9.dll" (ByRef OutJIT_ As Long, ByVal MP_ As Long, ByVal OptLevel_ As Long, ByRef OutError_ As Long) As Long
'Public Declare Sub LLVMDisposeExecutionEngine Lib "llvm2.9.dll" (ByVal EE_ As Long)
'Public Declare Sub LLVMRunStaticConstructors Lib "llvm2.9.dll" (ByVal EE_ As Long)
'Public Declare Sub LLVMRunStaticDestructors Lib "llvm2.9.dll" (ByVal EE_ As Long)
'Public Declare Function LLVMRunFunctionAsMain Lib "llvm2.9.dll" (ByVal EE_ As Long, ByVal F_ As Long, ByVal ArgC_ As Long, ByRef ArgV_ As Long, ByRef EnvP_ As Long) As Long
'Public Declare Function LLVMRunFunction Lib "llvm2.9.dll" (ByVal EE_ As Long, ByVal F_ As Long, ByVal NumArgs_ As Long, ByRef Args_ As Long) As Long
'Public Declare Sub LLVMFreeMachineCodeForFunction Lib "llvm2.9.dll" (ByVal EE_ As Long, ByVal F_ As Long)
'Public Declare Sub LLVMAddModule Lib "llvm2.9.dll" (ByVal EE_ As Long, ByVal M_ As Long)
'Public Declare Sub LLVMAddModuleProvider Lib "llvm2.9.dll" (ByVal EE_ As Long, ByVal MP_ As Long)
'Public Declare Function LLVMRemoveModule Lib "llvm2.9.dll" (ByVal EE_ As Long, ByVal M_ As Long, ByRef OutMod_ As Long, ByRef OutError_ As Long) As Long
'Public Declare Function LLVMRemoveModuleProvider Lib "llvm2.9.dll" (ByVal EE_ As Long, ByVal MP_ As Long, ByRef OutMod_ As Long, ByRef OutError_ As Long) As Long
'Public Declare Function LLVMFindFunction Lib "llvm2.9.dll" (ByVal EE_ As Long, ByVal Name_ As String, ByRef OutFn_ As Long) As Long
'Public Declare Function LLVMRecompileAndRelinkFunction Lib "llvm2.9.dll" (ByVal EE_ As Long, ByVal Fn_ As Long) As Long 'Void*
'Public Declare Function LLVMGetExecutionEngineTargetData Lib "llvm2.9.dll" (ByVal EE_ As Long) As Long
'Public Declare Sub LLVMAddGlobalMapping Lib "llvm2.9.dll" (ByVal EE_ As Long, ByVal Global_ As Long, ByRef Addr_ As Any)
'Public Declare Function LLVMGetPointerToGlobal Lib "llvm2.9.dll" (ByVal EE_ As Long, ByVal Global_ As Long) As Long 'Void*
'--- F:\Projects\llvm-2.9\include\llvm-c\test.h ---
'--- F:\Projects\llvm-2.9\include\llvm-c\Initialization.h ---
'Public Declare Sub LLVMInitializeCore Lib "llvm2.9.dll" (ByVal R_ As Long)
'Public Declare Sub LLVMInitializeTransformUtils Lib "llvm2.9.dll" (ByVal R_ As Long)
'Public Declare Sub LLVMInitializeScalarOpts Lib "llvm2.9.dll" (ByVal R_ As Long)
'Public Declare Sub LLVMInitializeInstCombine Lib "llvm2.9.dll" (ByVal R_ As Long)
'Public Declare Sub LLVMInitializeIPO Lib "llvm2.9.dll" (ByVal R_ As Long)
'Public Declare Sub LLVMInitializeInstrumentation Lib "llvm2.9.dll" (ByVal R_ As Long)
'Public Declare Sub LLVMInitializeAnalysis Lib "llvm2.9.dll" (ByVal R_ As Long)
'Public Declare Sub LLVMInitializeIPA Lib "llvm2.9.dll" (ByVal R_ As Long)
'Public Declare Sub LLVMInitializeCodeGen Lib "llvm2.9.dll" (ByVal R_ As Long)
'Public Declare Sub LLVMInitializeTarget Lib "llvm2.9.dll" (ByVal R_ As Long)
'--- F:\Projects\llvm-2.9\include\llvm-c\test.h ---
'--- F:\Projects\llvm-2.9\include\llvm-c\LinkTimeOptimizer.h ---
'--- F:\Projects\llvm-2.9\include\llvm-c\test.h ---
'--- F:\Projects\llvm-2.9\include\llvm-c\lto.h ---
'Public Declare Function lto_get_version Lib "llvm2.9.dll" () As Long 'Byte*
'Public Declare Function lto_get_error_message Lib "llvm2.9.dll" () As Long 'Byte*
'Public Declare Function lto_module_is_object_file Lib "llvm2.9.dll" (ByVal path_ As String) As Byte
'Public Declare Function lto_module_is_object_file_for_target Lib "llvm2.9.dll" (ByVal path_ As String, ByVal target_triple_prefix_ As String) As Byte
'Public Declare Function lto_module_is_object_file_in_memory Lib "llvm2.9.dll" (ByRef mem_ As Any, ByVal length_ As Long) As Byte
'Public Declare Function lto_module_is_object_file_in_memory_for_target Lib "llvm2.9.dll" (ByRef mem_ As Any, ByVal length_ As Long, ByVal target_triple_prefix_ As String) As Byte
'Public Declare Function lto_module_create Lib "llvm2.9.dll" (ByVal path_ As String) As Long
'Public Declare Function lto_module_create_from_memory Lib "llvm2.9.dll" (ByRef mem_ As Any, ByVal length_ As Long) As Long
'Public Declare Function lto_module_create_from_fd Lib "llvm2.9.dll" (ByVal fd_ As Long, ByVal path_ As String, ByVal size_ As Long) As Long
'Public Declare Sub lto_module_dispose Lib "llvm2.9.dll" (ByVal mod_ As Long)
'Public Declare Function lto_module_get_target_triple Lib "llvm2.9.dll" (ByVal mod_ As Long) As Long 'Byte*
'Public Declare Sub lto_module_set_target_triple Lib "llvm2.9.dll" (ByVal mod_ As Long, ByVal triple_ As String)
'Public Declare Function lto_module_get_num_symbols Lib "llvm2.9.dll" (ByVal mod_ As Long) As Long
'Public Declare Function lto_module_get_symbol_name Lib "llvm2.9.dll" (ByVal mod_ As Long, ByVal index_ As Long) As Long 'Byte*
'Public Declare Function lto_module_get_symbol_attribute Lib "llvm2.9.dll" (ByVal mod_ As Long, ByVal index_ As Long) As lto_symbol_attributes
'Public Declare Function lto_codegen_create Lib "llvm2.9.dll" () As Long
'Public Declare Sub lto_codegen_dispose Lib "llvm2.9.dll" (ByVal a1_ As Long)
'Public Declare Function lto_codegen_add_module Lib "llvm2.9.dll" (ByVal cg_ As Long, ByVal mod_ As Long) As Byte
'Public Declare Function lto_codegen_set_debug_model Lib "llvm2.9.dll" (ByVal cg_ As Long, ByVal a2_ As lto_debug_model) As Byte
'Public Declare Function lto_codegen_set_pic_model Lib "llvm2.9.dll" (ByVal cg_ As Long, ByVal a2_ As lto_codegen_model) As Byte
'Public Declare Sub lto_codegen_set_cpu Lib "llvm2.9.dll" (ByVal cg_ As Long, ByVal cpu_ As String)
'Public Declare Sub lto_codegen_set_assembler_path Lib "llvm2.9.dll" (ByVal cg_ As Long, ByVal path_ As String)
'Public Declare Sub lto_codegen_set_assembler_args Lib "llvm2.9.dll" (ByVal cg_ As Long, ByRef args_ As Long, ByVal nargs_ As Long)
'Public Declare Sub lto_codegen_add_must_preserve_symbol Lib "llvm2.9.dll" (ByVal cg_ As Long, ByVal symbol_ As String)
'Public Declare Function lto_codegen_write_merged_modules Lib "llvm2.9.dll" (ByVal cg_ As Long, ByVal path_ As String) As Byte
'Public Declare Function lto_codegen_compile Lib "llvm2.9.dll" (ByVal cg_ As Long, ByRef length_ As Long) As Long 'Void*
'Public Declare Sub lto_codegen_debug_options Lib "llvm2.9.dll" (ByVal cg_ As Long, ByVal a2_ As String)
'--- F:\Projects\llvm-2.9\include\llvm-c\test.h ---
'--- F:\Projects\llvm-2.9\tools\clang\include\clang-c\Index.h ---
'Public Type CXUnsavedFile
' Filename As Long 'Byte*
' Contents As Long 'Byte*
' Length As Long
'End Type '12 bytes
'Public Type CXString
' data As Long 'Void*
' private_flags As Long
'End Type '8 bytes
'Public Declare Function clang_getCString Lib "llvm2.9.dll" (ByRef string_ As CXString) As Long 'Byte*
'Public Declare Sub clang_disposeString Lib "llvm2.9.dll" (ByRef string_ As CXString)
'Public Declare Function clang_createIndex Lib "llvm2.9.dll" (ByVal excludeDeclarationsFromPCH_ As Long, ByVal displayDiagnostics_ As Long) As Long 'Void*
'Public Declare Sub clang_disposeIndex Lib "llvm2.9.dll" (ByRef index_ As Any)
'Public Declare Sub clang_getFileName Lib "llvm2.9.dll" (ByRef SFile_ As Any, ByRef ret__ As CXString)
'Public Declare Function clang_getFileTime Lib "llvm2.9.dll" (ByRef SFile_ As Any) As Currency
'Public Declare Function clang_getFile Lib "llvm2.9.dll" (ByVal tu_ As Long, ByVal file_name_ As String) As Long 'Void*
'Public Type CXSourceLocation
' ptr_data(0 To 2 - 1) As Long 'Void*
' int_data As Long
'End Type '12 bytes
'Public Type CXSourceRange
' ptr_data(0 To 2 - 1) As Long 'Void*
' begin_int_data As Long
' end_int_data As Long
'End Type '16 bytes
'Public Declare Sub clang_getNullLocation Lib "llvm2.9.dll" (ByRef ret__ As CXSourceLocation)
'Public Declare Function clang_equalLocations Lib "llvm2.9.dll" (ByRef loc1_ As CXSourceLocation, ByRef loc2_ As CXSourceLocation) As Long
'Public Declare Sub clang_getLocation Lib "llvm2.9.dll" (ByVal tu_ As Long, ByRef file_ As Any, ByVal line_ As Long, ByVal column_ As Long, ByRef ret__ As CXSourceLocation)
'Public Declare Sub clang_getLocationForOffset Lib "llvm2.9.dll" (ByVal tu_ As Long, ByRef file_ As Any, ByVal offset_ As Long, ByRef ret__ As CXSourceLocation)
'Public Declare Sub clang_getNullRange Lib "llvm2.9.dll" (ByRef ret__ As CXSourceRange)
'Public Declare Sub clang_getRange Lib "llvm2.9.dll" (ByRef begin_ As CXSourceLocation, ByRef end_ As CXSourceLocation, ByRef ret__ As CXSourceRange)
'Public Declare Sub clang_getInstantiationLocation Lib "llvm2.9.dll" (ByRef location_ As CXSourceLocation, ByRef file_ As Long, ByRef line_ As Long, ByRef column_ As Long, ByRef offset_ As Long)
'Public Declare Sub clang_getSpellingLocation Lib "llvm2.9.dll" (ByRef location_ As CXSourceLocation, ByRef file_ As Long, ByRef line_ As Long, ByRef column_ As Long, ByRef offset_ As Long)
'Public Declare Sub clang_getRangeStart Lib "llvm2.9.dll" (ByRef range_ As CXSourceRange, ByRef ret__ As CXSourceLocation)
'Public Declare Sub clang_getRangeEnd Lib "llvm2.9.dll" (ByRef range_ As CXSourceRange, ByRef ret__ As CXSourceLocation)
'Public Declare Function clang_getNumDiagnostics Lib "llvm2.9.dll" (ByVal Unit_ As Long) As Long
'Public Declare Function clang_getDiagnostic Lib "llvm2.9.dll" (ByVal Unit_ As Long, ByVal Index_ As Long) As Long 'Void*
'Public Declare Sub clang_disposeDiagnostic Lib "llvm2.9.dll" (ByRef Diagnostic_ As Any)
'Public Declare Sub clang_formatDiagnostic Lib "llvm2.9.dll" (ByRef Diagnostic_ As Any, ByVal Options_ As Long, ByRef ret__ As CXString)
'Public Declare Function clang_defaultDiagnosticDisplayOptions Lib "llvm2.9.dll" () As Long
'Public Declare Function clang_getDiagnosticSeverity Lib "llvm2.9.dll" (ByRef a1_ As Any) As CXDiagnosticSeverity
'Public Declare Sub clang_getDiagnosticLocation Lib "llvm2.9.dll" (ByRef a1_ As Any, ByRef ret__ As CXSourceLocation)
'Public Declare Sub clang_getDiagnosticSpelling Lib "llvm2.9.dll" (ByRef a1_ As Any, ByRef ret__ As CXString)
'Public Declare Sub clang_getDiagnosticOption Lib "llvm2.9.dll" (ByRef Diag_ As Any, ByRef Disable_ As CXString, ByRef ret__ As CXString)
'Public Declare Function clang_getDiagnosticCategory Lib "llvm2.9.dll" (ByRef a1_ As Any) As Long
'Public Declare Sub clang_getDiagnosticCategoryName Lib "llvm2.9.dll" (ByVal Category_ As Long, ByRef ret__ As CXString)
'Public Declare Function clang_getDiagnosticNumRanges Lib "llvm2.9.dll" (ByRef a1_ As Any) As Long
'Public Declare Sub clang_getDiagnosticRange Lib "llvm2.9.dll" (ByRef Diagnostic_ As Any, ByVal Range_ As Long, ByRef ret__ As CXSourceRange)
'Public Declare Function clang_getDiagnosticNumFixIts Lib "llvm2.9.dll" (ByRef Diagnostic_ As Any) As Long
'Public Declare Sub clang_getDiagnosticFixIt Lib "llvm2.9.dll" (ByRef Diagnostic_ As Any, ByVal FixIt_ As Long, ByRef ReplacementRange_ As CXSourceRange, ByRef ret__ As CXString)
'Public Declare Sub clang_getTranslationUnitSpelling Lib "llvm2.9.dll" (ByVal CTUnit_ As Long, ByRef ret__ As CXString)
'Public Declare Function clang_createTranslationUnitFromSourceFile Lib "llvm2.9.dll" (ByRef CIdx_ As Any, ByVal source_filename_ As String, ByVal num_clang_command_line_args_ As Long, ByRef clang_command_line_args_ As Long, ByVal num_unsaved_files_ As Long, ByRef unsaved_files_ As CXUnsavedFile) As Long
'Public Declare Function clang_createTranslationUnit Lib "llvm2.9.dll" (ByRef a1_ As Any, ByVal ast_filename_ As String) As Long
'Public Declare Function clang_defaultEditingTranslationUnitOptions Lib "llvm2.9.dll" () As Long
'Public Declare Function clang_parseTranslationUnit Lib "llvm2.9.dll" (ByRef CIdx_ As Any, ByVal source_filename_ As String, ByRef command_line_args_ As Long, ByVal num_command_line_args_ As Long, ByRef unsaved_files_ As CXUnsavedFile, ByVal num_unsaved_files_ As Long, ByVal options_ As Long) As Long
'Public Declare Function clang_defaultSaveOptions Lib "llvm2.9.dll" (ByVal TU_ As Long) As Long
'Public Declare Function clang_saveTranslationUnit Lib "llvm2.9.dll" (ByVal TU_ As Long, ByVal FileName_ As String, ByVal options_ As Long) As Long
'Public Declare Sub clang_disposeTranslationUnit Lib "llvm2.9.dll" (ByVal a1_ As Long)
'Public Declare Function clang_defaultReparseOptions Lib "llvm2.9.dll" (ByVal TU_ As Long) As Long
'Public Declare Function clang_reparseTranslationUnit Lib "llvm2.9.dll" (ByVal TU_ As Long, ByVal num_unsaved_files_ As Long, ByRef unsaved_files_ As CXUnsavedFile, ByVal options_ As Long) As Long
'Public Type CXCursor
' kind As CXCursorKind
' data(0 To 3 - 1) As Long 'Void*
'End Type '16 bytes
'Public Declare Sub clang_getNullCursor Lib "llvm2.9.dll" (ByRef ret__ As CXCursor)
'Public Declare Sub clang_getTranslationUnitCursor Lib "llvm2.9.dll" (ByVal a1_ As Long, ByRef ret__ As CXCursor)
'Public Declare Function clang_equalCursors Lib "llvm2.9.dll" (ByRef a1_ As CXCursor, ByRef a2_ As CXCursor) As Long
'Public Declare Function clang_hashCursor Lib "llvm2.9.dll" (ByRef a1_ As CXCursor) As Long
'Public Declare Function clang_getCursorKind Lib "llvm2.9.dll" (ByRef a1_ As CXCursor) As CXCursorKind
'Public Declare Function clang_isDeclaration Lib "llvm2.9.dll" (ByVal a1_ As CXCursorKind) As Long
'Public Declare Function clang_isReference Lib "llvm2.9.dll" (ByVal a1_ As CXCursorKind) As Long
'Public Declare Function clang_isExpression Lib "llvm2.9.dll" (ByVal a1_ As CXCursorKind) As Long
'Public Declare Function clang_isStatement Lib "llvm2.9.dll" (ByVal a1_ As CXCursorKind) As Long
'Public Declare Function clang_isInvalid Lib "llvm2.9.dll" (ByVal a1_ As CXCursorKind) As Long
'Public Declare Function clang_isTranslationUnit Lib "llvm2.9.dll" (ByVal a1_ As CXCursorKind) As Long
'Public Declare Function clang_isPreprocessing Lib "llvm2.9.dll" (ByVal a1_ As CXCursorKind) As Long
'Public Declare Function clang_isUnexposed Lib "llvm2.9.dll" (ByVal a1_ As CXCursorKind) As Long
'Public Declare Function clang_getCursorLinkage Lib "llvm2.9.dll" (ByRef cursor_ As CXCursor) As CXLinkageKind
'Public Declare Function clang_getCursorAvailability Lib "llvm2.9.dll" (ByRef cursor_ As CXCursor) As CXAvailabilityKind
'Public Declare Function clang_getCursorLanguage Lib "llvm2.9.dll" (ByRef cursor_ As CXCursor) As CXLanguageKind
'Public Declare Function clang_createCXCursorSet Lib "llvm2.9.dll" () As Long
'Public Declare Sub clang_disposeCXCursorSet Lib "llvm2.9.dll" (ByVal cset_ As Long)
'Public Declare Function clang_CXCursorSet_contains Lib "llvm2.9.dll" (ByVal cset_ As Long, ByRef cursor_ As CXCursor) As Long
'Public Declare Function clang_CXCursorSet_insert Lib "llvm2.9.dll" (ByVal cset_ As Long, ByRef cursor_ As CXCursor) As Long
'Public Declare Sub clang_getCursorSemanticParent Lib "llvm2.9.dll" (ByRef cursor_ As CXCursor, ByRef ret__ As CXCursor)
'Public Declare Sub clang_getCursorLexicalParent Lib "llvm2.9.dll" (ByRef cursor_ As CXCursor, ByRef ret__ As CXCursor)
'Public Declare Sub clang_getOverriddenCursors Lib "llvm2.9.dll" (ByRef cursor_ As CXCursor, ByRef overridden_ As Long, ByRef num_overridden_ As Long)
'Public Declare Sub clang_disposeOverriddenCursors Lib "llvm2.9.dll" (ByRef overridden_ As CXCursor)
'Public Declare Function clang_getIncludedFile Lib "llvm2.9.dll" (ByRef cursor_ As CXCursor) As Long 'Void*
'Public Declare Sub clang_getCursor Lib "llvm2.9.dll" (ByVal a1_ As Long, ByRef a2_ As CXSourceLocation, ByRef ret__ As CXCursor)
'Public Declare Sub clang_getCursorLocation Lib "llvm2.9.dll" (ByRef a1_ As CXCursor, ByRef ret__ As CXSourceLocation)
'Public Declare Sub clang_getCursorExtent Lib "llvm2.9.dll" (ByRef a1_ As CXCursor, ByRef ret__ As CXSourceRange)
'Public Type CXType
' kind As CXTypeKind
' data(0 To 2 - 1) As Long 'Void*
'End Type '12 bytes
'Public Declare Sub clang_getCursorType Lib "llvm2.9.dll" (ByRef C_ As CXCursor, ByRef ret__ As CXType)
'Public Declare Function clang_equalTypes Lib "llvm2.9.dll" (ByRef A_ As CXType, ByRef B_ As CXType) As Long
'Public Declare Sub clang_getCanonicalType Lib "llvm2.9.dll" (ByRef T_ As CXType, ByRef ret__ As CXType)
'Public Declare Function clang_isConstQualifiedType Lib "llvm2.9.dll" (ByRef T_ As CXType) As Long
'Public Declare Function clang_isVolatileQualifiedType Lib "llvm2.9.dll" (ByRef T_ As CXType) As Long
'Public Declare Function clang_isRestrictQualifiedType Lib "llvm2.9.dll" (ByRef T_ As CXType) As Long
'Public Declare Sub clang_getPointeeType Lib "llvm2.9.dll" (ByRef T_ As CXType, ByRef ret__ As CXType)
'Public Declare Sub clang_getTypeDeclaration Lib "llvm2.9.dll" (ByRef T_ As CXType, ByRef ret__ As CXCursor)
'Public Declare Sub clang_getDeclObjCTypeEncoding Lib "llvm2.9.dll" (ByRef C_ As CXCursor, ByRef ret__ As CXString)
'Public Declare Sub clang_getTypeKindSpelling Lib "llvm2.9.dll" (ByVal K_ As CXTypeKind, ByRef ret__ As CXString)
'Public Declare Sub clang_getResultType Lib "llvm2.9.dll" (ByRef T_ As CXType, ByRef ret__ As CXType)
'Public Declare Sub clang_getCursorResultType Lib "llvm2.9.dll" (ByRef C_ As CXCursor, ByRef ret__ As CXType)
'Public Declare Function clang_isPODType Lib "llvm2.9.dll" (ByRef T_ As CXType) As Long
'Public Declare Function clang_isVirtualBase Lib "llvm2.9.dll" (ByRef a1_ As CXCursor) As Long
'Public Declare Function clang_getCXXAccessSpecifier Lib "llvm2.9.dll" (ByRef a1_ As CXCursor) As CX_CXXAccessSpecifier
'Public Declare Function clang_getNumOverloadedDecls Lib "llvm2.9.dll" (ByRef cursor_ As CXCursor) As Long
'Public Declare Sub clang_getOverloadedDecl Lib "llvm2.9.dll" (ByRef cursor_ As CXCursor, ByVal index_ As Long, ByRef ret__ As CXCursor)
'Public Declare Sub clang_getIBOutletCollectionType Lib "llvm2.9.dll" (ByRef a1_ As CXCursor, ByRef ret__ As CXType)
'Public Declare Function clang_visitChildren Lib "llvm2.9.dll" (ByRef parent_ As CXCursor, ByVal visitor_ As Long, ByRef client_data_ As Any) As Long
'Public Declare Sub clang_getCursorUSR Lib "llvm2.9.dll" (ByRef a1_ As CXCursor, ByRef ret__ As CXString)
'Public Declare Sub clang_constructUSR_ObjCClass Lib "llvm2.9.dll" (ByVal class_name_ As String, ByRef ret__ As CXString)
'Public Declare Sub clang_constructUSR_ObjCCategory Lib "llvm2.9.dll" (ByVal class_name_ As String, ByVal category_name_ As String, ByRef ret__ As CXString)
'Public Declare Sub clang_constructUSR_ObjCProtocol Lib "llvm2.9.dll" (ByVal protocol_name_ As String, ByRef ret__ As CXString)
'Public Declare Sub clang_constructUSR_ObjCIvar Lib "llvm2.9.dll" (ByVal name_ As String, ByRef classUSR_ As CXString, ByRef ret__ As CXString)
'Public Declare Sub clang_constructUSR_ObjCMethod Lib "llvm2.9.dll" (ByVal name_ As String, ByVal isInstanceMethod_ As Long, ByRef classUSR_ As CXString, ByRef ret__ As CXString)
'Public Declare Sub clang_constructUSR_ObjCProperty Lib "llvm2.9.dll" (ByVal property_ As String, ByRef classUSR_ As CXString, ByRef ret__ As CXString)
'Public Declare Sub clang_getCursorSpelling Lib "llvm2.9.dll" (ByRef a1_ As CXCursor, ByRef ret__ As CXString)
'Public Declare Sub clang_getCursorDisplayName Lib "llvm2.9.dll" (ByRef a1_ As CXCursor, ByRef ret__ As CXString)
'Public Declare Sub clang_getCursorReferenced Lib "llvm2.9.dll" (ByRef a1_ As CXCursor, ByRef ret__ As CXCursor)
'Public Declare Sub clang_getCursorDefinition Lib "llvm2.9.dll" (ByRef a1_ As CXCursor, ByRef ret__ As CXCursor)
'Public Declare Function clang_isCursorDefinition Lib "llvm2.9.dll" (ByRef a1_ As CXCursor) As Long
'Public Declare Sub clang_getCanonicalCursor Lib "llvm2.9.dll" (ByRef a1_ As CXCursor, ByRef ret__ As CXCursor)
'Public Declare Function clang_CXXMethod_isStatic Lib "llvm2.9.dll" (ByRef C_ As CXCursor) As Long
'Public Declare Function clang_getTemplateCursorKind Lib "llvm2.9.dll" (ByRef C_ As CXCursor) As CXCursorKind
'Public Declare Sub clang_getSpecializedCursorTemplate Lib "llvm2.9.dll" (ByRef C_ As CXCursor, ByRef ret__ As CXCursor)
'Public Type CXToken
' int_data(0 To 4 - 1) As Long
' ptr_data As Long 'Void*
'End Type '20 bytes
'Public Declare Function clang_getTokenKind Lib "llvm2.9.dll" (ByRef a1_ As CXToken) As CXTokenKind
'Public Declare Sub clang_getTokenSpelling Lib "llvm2.9.dll" (ByVal a1_ As Long, ByRef a2_ As CXToken, ByRef ret__ As CXString)
'Public Declare Sub clang_getTokenLocation Lib "llvm2.9.dll" (ByVal a1_ As Long, ByRef a2_ As CXToken, ByRef ret__ As CXSourceLocation)
'Public Declare Sub clang_getTokenExtent Lib "llvm2.9.dll" (ByVal a1_ As Long, ByRef a2_ As CXToken, ByRef ret__ As CXSourceRange)
'Public Declare Sub clang_tokenize Lib "llvm2.9.dll" (ByVal TU_ As Long, ByRef Range_ As CXSourceRange, ByRef Tokens_ As Long, ByRef NumTokens_ As Long)
'Public Declare Sub clang_annotateTokens Lib "llvm2.9.dll" (ByVal TU_ As Long, ByRef Tokens_ As CXToken, ByVal NumTokens_ As Long, ByRef Cursors_ As CXCursor)
'Public Declare Sub clang_disposeTokens Lib "llvm2.9.dll" (ByVal TU_ As Long, ByRef Tokens_ As CXToken, ByVal NumTokens_ As Long)
'Public Declare Sub clang_getCursorKindSpelling Lib "llvm2.9.dll" (ByVal Kind_ As CXCursorKind, ByRef ret__ As CXString)
'Public Declare Sub clang_getDefinitionSpellingAndExtent Lib "llvm2.9.dll" (ByRef a1_ As CXCursor, ByRef startBuf_ As Long, ByRef endBuf_ As Long, ByRef startLine_ As Long, ByRef startColumn_ As Long, ByRef endLine_ As Long, ByRef endColumn_ As Long)
'Public Declare Sub clang_enableStackTraces Lib "llvm2.9.dll" ()
'Public Declare Sub clang_executeOnThread Lib "llvm2.9.dll" (ByRef fn_ As Long, ByRef user_data_ As Any, ByVal stack_size_ As Long)
'Public Type CXCompletionResult
' CursorKind As CXCursorKind
' CompletionString As Long 'Void*
'End Type '8 bytes
'Public Declare Function clang_getCompletionChunkKind Lib "llvm2.9.dll" (ByRef completion_string_ As Any, ByVal chunk_number_ As Long) As CXCompletionChunkKind
'Public Declare Sub clang_getCompletionChunkText Lib "llvm2.9.dll" (ByRef completion_string_ As Any, ByVal chunk_number_ As Long, ByRef ret__ As CXString)
'Public Declare Function clang_getCompletionChunkCompletionString Lib "llvm2.9.dll" (ByRef completion_string_ As Any, ByVal chunk_number_ As Long) As Long 'Void*
'Public Declare Function clang_getNumCompletionChunks Lib "llvm2.9.dll" (ByRef completion_string_ As Any) As Long
'Public Declare Function clang_getCompletionPriority Lib "llvm2.9.dll" (ByRef completion_string_ As Any) As Long
'Public Declare Function clang_getCompletionAvailability Lib "llvm2.9.dll" (ByRef completion_string_ As Any) As CXAvailabilityKind
'Public Type CXCodeCompleteResults
' Results As Long 'CXCompletionResult*
' NumResults As Long
'End Type '8 bytes
'Public Declare Function clang_defaultCodeCompleteOptions Lib "llvm2.9.dll" () As Long
'Public Declare Function clang_codeCompleteAt Lib "llvm2.9.dll" (ByVal TU_ As Long, ByVal complete_filename_ As String, ByVal complete_line_ As Long, ByVal complete_column_ As Long, ByRef unsaved_files_ As CXUnsavedFile, ByVal num_unsaved_files_ As Long, ByVal options_ As Long) As Long 'CXCodeCompleteResults*
'Public Declare Sub clang_sortCodeCompletionResults Lib "llvm2.9.dll" (ByRef Results_ As CXCompletionResult, ByVal NumResults_ As Long)
'Public Declare Sub clang_disposeCodeCompleteResults Lib "llvm2.9.dll" (ByRef Results_ As CXCodeCompleteResults)
'Public Declare Function clang_codeCompleteGetNumDiagnostics Lib "llvm2.9.dll" (ByRef Results_ As CXCodeCompleteResults) As Long
'Public Declare Function clang_codeCompleteGetDiagnostic Lib "llvm2.9.dll" (ByRef Results_ As CXCodeCompleteResults, ByVal Index_ As Long) As Long 'Void*
'Public Declare Sub clang_getClangVersion Lib "llvm2.9.dll" (ByRef ret__ As CXString)
'Public Declare Sub clang_getInclusions Lib "llvm2.9.dll" (ByVal tu_ As Long, ByVal visitor_ As Long, ByRef client_data_ As Any)
'--- F:\Projects\llvm-2.9\include\llvm-c\test.h ---
'--- F:\Projects\llvm-2.9\include\llvm-c\test.i ---

Public Enum EDAssemblySyntax_t
'/*! @constant kEDAssemblySyntaxX86Intel Intel syntax for i386 and x86_64. */
  kEDAssemblySyntaxX86Intel = 0 ',
'/*! @constant kEDAssemblySyntaxX86ATT AT&T syntax for i386 and x86_64. */
  kEDAssemblySyntaxX86ATT = 1   ',
  kEDAssemblySyntaxARMUAL = 2
End Enum ';

'// Relocation model types.
Public Enum LLVMReloc
    LLVMReloc_Default
    LLVMReloc_Static
    LLVMReloc_PIC_
    LLVMReloc_DynamicNoPIC
End Enum

'// Code model types.
Public Enum LLVMCodeModel
    LLVMCodeModel_Default
    LLVMCodeModel_Small
    LLVMCodeModel_Kernel
    LLVMCodeModel_Medium
    LLVMCodeModel_Large
End Enum

'// Code generation optimization level.
Public Enum LLVMCodeGenOpt
    LLVMCodeGenOpt_None '        // -O0
    LLVMCodeGenOpt_Less '        // -O1
    LLVMCodeGenOpt_Default '     // -O2, -Os
    LLVMCodeGenOpt_Aggressive '  // -O3
End Enum

Public Enum LLVMSched
    LLVMSched_None ',             // No preference
    LLVMSched_Latency ',          // Scheduling for shortest total latency.
    LLVMSched_RegPressure ',      // Scheduling for lowest register pressure.
    LLVMSched_Hybrid ',           // Scheduling for both latency and register pressure.
    LLVMSched_ILP '               // Scheduling for ILP in low register pressure mode.
End Enum

Public Enum LLVMCodeGenFileType
    CGFT_AssemblyFile
    CGFT_ObjectFile
    CGFT_Null         '// Do not emit any output.
End Enum

Public Enum LLVMFloatABI
      LLVMFloatABI_Default ', // Target-specific (either soft of hard depending on triple, etc).
      LLVMFloatABI_Soft ', // Soft float.
      LLVMFloatABI_Hard '  // Hard float.
End Enum

Public Enum LLVMOpenMode
    LLVMOpenMode_in = &H1
    LLVMOpenMode_out = &H2
    LLVMOpenMode_ate = &H4
    LLVMOpenMode_app = &H8
    LLVMOpenMode_trunc = &H10
    LLVMOpenMode_Nocreate = &H40
    LLVMOpenMode_Noreplace = &H80
    LLVMOpenMode_binary = &H20
End Enum

Public Enum LLVMInliningPassThreshold
  LLVMInliningPassThreshold_AlwaysInliner = -2
  LLVMInliningPassThreshold_Default = -1
  LLVMInliningPassThreshold_NoInlining = 0
End Enum

Public Declare Sub LLVMPrintVersion Lib "llvm2.9.dll" ()
Public Declare Sub LLVMShutDown Lib "llvm2.9.dll" ()
Public Declare Sub LLVMInitializeAllAsmPrinters Lib "llvm2.9.dll" ()
Public Declare Sub LLVMInitializeAllAsmParsers Lib "llvm2.9.dll" ()
Public Declare Sub LLVMInitializeAllDisassemblers Lib "llvm2.9.dll" ()

Public Declare Function LLVMCreateAsmInfo Lib "llvm2.9.dll" (ByVal Triple_ As String) As Long
Public Declare Function LLVMCreateTargetMachine Lib "llvm2.9.dll" (ByVal Triple_ As String, ByVal Features_ As String) As Long
Public Declare Function LLVMCreateAsmBackend Lib "llvm2.9.dll" (ByVal Triple_ As String) As Long
Public Declare Function LLVMCreateAsmLexer Lib "llvm2.9.dll" (ByVal Triple_ As String, ByVal MCAsmInfo_ As Long) As Long
Public Declare Function LLVMCreateAsmParser Lib "llvm2.9.dll" (ByVal Triple_ As String, ByVal MCAsmParser_ As Long, ByVal TargetMachine_ As Long) As Long
Public Declare Function LLVMCreateAsmPrinter Lib "llvm2.9.dll" (ByVal Triple_ As String, ByVal TargetMachine_ As Long, ByVal MCStreamer_ As Long) As Long
Public Declare Function LLVMCreateMCDisassembler Lib "llvm2.9.dll" (ByVal Triple_ As String) As Long
Public Declare Function LLVMCreateMCInstPrinter Lib "llvm2.9.dll" (ByVal Triple_ As String, ByVal SyntaxVariant_ As EDAssemblySyntax_t, ByVal MCAsmInfo_ As Long) As Long
Public Declare Function LLVMCreateCodeEmitter Lib "llvm2.9.dll" (ByVal Triple_ As String, ByVal TargetMachine_ As Long, ByVal MCContext_ As Long) As Long
Public Declare Function LLVMCreateObjectStreamer Lib "llvm2.9.dll" (ByVal Triple_ As String, ByVal TT_ As String, ByVal MCContext_ As Long, ByVal TargetAsmBackend_ As Long, ByVal OS_ As Long, ByVal MCCodeEmitter_ As Long, ByVal RelaxAll_ As Byte, ByVal NoExecStack_ As Byte) As Long

Public Declare Sub LLVMDisposeMCAsmInfo Lib "llvm2.9.dll" (ByVal MCAsmInfo_ As Long)
Public Declare Sub LLVMDisposeTargetMachine Lib "llvm2.9.dll" (ByVal TargetMachine_ As Long)
Public Declare Sub LLVMDisposeTargetAsmBackend Lib "llvm2.9.dll" (ByVal TargetAsmBackend_ As Long)
Public Declare Sub LLVMDisposeTargetAsmLexer Lib "llvm2.9.dll" (ByVal TargetAsmLexer_ As Long)
Public Declare Sub LLVMDisposeTargetAsmParser Lib "llvm2.9.dll" (ByVal TargetAsmParser_ As Long)
Public Declare Sub LLVMDisposeAsmPrinter Lib "llvm2.9.dll" (ByVal AsmPrinter_ As Long)
Public Declare Sub LLVMDisposeMCDisassembler Lib "llvm2.9.dll" (ByVal MCDisassembler_ As Long)
Public Declare Sub LLVMDisposeMCInstPrinter Lib "llvm2.9.dll" (ByVal MCInstPrinter_ As Long)
Public Declare Sub LLVMDisposeMCCodeEmitter Lib "llvm2.9.dll" (ByVal MCCodeEmitter_ As Long)
Public Declare Sub LLVMDisposeMCStreamer Lib "llvm2.9.dll" (ByVal MCStreamer_ As Long)

Public Declare Sub LLVMTargetMachineSetMCRelaxAll Lib "llvm2.9.dll" (ByVal TargetMachine_ As Long, ByVal b As Byte)
Public Declare Function LLVMTargetMachineGetDataLayout Lib "llvm2.9.dll" (ByVal TargetMachine_ As Long, ByRef s As Any, ByVal m As Long) As Long 'ByVal s As String

Public Declare Function LLVMTargetMachineAddPassesToEmitFile Lib "llvm2.9.dll" (ByVal TargetMachine_ As Long, ByVal FPM_ As Long, ByVal FOS_ As Long, ByVal CFGT_ As LLVMCodeGenFileType, ByVal OptLevel_ As LLVMCodeGenOpt, ByVal DisableVerify_ As Byte) As Byte
Public Declare Function LLVMTargetMachineAddPassesToEmitMachineCode Lib "llvm2.9.dll" (ByVal TargetMachine_ As Long, ByVal FPM_ As Long, ByVal JITCodeEmitter_ As Long, ByVal OptLevel_ As LLVMCodeGenOpt, ByVal DisableVerify_ As Byte) As Byte
Public Declare Function LLVMTargetMachineAddPassesToEmitMC Lib "llvm2.9.dll" (ByVal TargetMachine_ As Long, ByVal FPM_ As Long, ByRef MCContext_ As Long, ByVal OptLevel_ As LLVMCodeGenOpt, ByVal DisableVerify_ As Byte) As Byte

Public Declare Function Util_CreateOStreamFromFile Lib "llvm2.9.dll" (ByVal FileName_ As String, ByVal OpenMode_ As LLVMOpenMode) As Long
Public Declare Function Util_GetCOut Lib "llvm2.9.dll" () As Long
Public Declare Function Util_GetCErr Lib "llvm2.9.dll" () As Long
Public Declare Sub Util_DisposeOStream Lib "llvm2.9.dll" (ByVal obj As Long)

Public Declare Function Util_CreateEmptyString Lib "llvm2.9.dll" () As Long
Public Declare Function Util_CreateString Lib "llvm2.9.dll" (ByVal s As String) As Long
Public Declare Function Util_CloneString Lib "llvm2.9.dll" (ByVal obj As Long) As Long
Public Declare Function Util_GetStringPointer Lib "llvm2.9.dll" (ByVal obj As Long) As Long
Public Declare Function Util_GetStringLength Lib "llvm2.9.dll" (ByVal obj As Long) As Long
Public Declare Sub Util_DisposeString Lib "llvm2.9.dll" (ByVal obj As Long)

Public Declare Function Util_CreateByteSVector Lib "llvm2.9.dll" (ByVal Size_ As Long) As Long
Public Declare Function Util_CloneByteSVector Lib "llvm2.9.dll" (ByVal obj As Long) As Long
Public Declare Function Util_GetByteSVectorPointer Lib "llvm2.9.dll" (ByVal obj As Long) As Long
Public Declare Function Util_GetByteSVectorLength Lib "llvm2.9.dll" (ByVal obj As Long) As Long
Public Declare Sub Util_DisposeByteSVector Lib "llvm2.9.dll" (ByVal obj As Long)

Public Declare Function LLVMCreateRaw_OS_OStream Lib "llvm2.9.dll" (ByVal obj As Long) As Long
Public Declare Function LLVMCreateRaw_String_OStream Lib "llvm2.9.dll" (ByVal obj As Long) As Long
Public Declare Function LLVMCreateRaw_SVector_OStream Lib "llvm2.9.dll" (ByVal obj As Long) As Long
Public Declare Function LLVMCreateFormattedRawOStream Lib "llvm2.9.dll" (ByVal obj As Long, ByVal bDeleteOld_ As Byte) As Long

Public Declare Sub LLVMDisposeRaw_OStream Lib "llvm2.9.dll" (ByVal obj As Long)

Public Type LLVMTargetOptions
  FloatABIType As LLVMFloatABI
  StackAlignment As Long
  RelocationModel As LLVMReloc
  CMModel As LLVMCodeModel
  '///
  PrintMachineCode As Byte
  NoFramePointerElim As Byte
  NoFramePointerElimNonLeaf As Byte
  LessPreciseFPMADOption As Byte
  NoExcessFPPrecision As Byte
  UnsafeFPMath As Byte
  NoInfsFPMath As Byte
  NoNaNsFPMath As Byte
  HonorSignDependentRoundingFPMathOption As Byte
  UseSoftFloat As Byte
  NoZerosInBSS As Byte
  JITExceptionHandling As Byte
  JITEmitDebugInfo As Byte
  JITEmitDebugInfoToDisk As Byte
  UnwindTablesMandatory As Byte
  GuaranteedTailCallOpt As Byte
  RealignStack As Byte
  DisableJumpTables As Byte
  EnableFastISel As Byte
  StrongPHIElim As Byte
  NoImplicitFloat As Byte
  AsmVerbosityDefault As Byte
  FunctionSections As Byte
  DataSections As Byte
End Type

Public Declare Sub LLVMGetTargetOptions Lib "llvm2.9.dll" (ByRef obj As LLVMTargetOptions)
Public Declare Sub LLVMSetTargetOptions Lib "llvm2.9.dll" (ByRef obj As LLVMTargetOptions)

Public Declare Sub LLVMAddFunctionInliningPassWithThreshold Lib "llvm2.9.dll" (ByVal PM_ As Long, ByVal Threshold_ As Long)
Public Declare Sub LLVMAddAlwaysInlinerPass Lib "llvm2.9.dll" (ByVal PM_ As Long)

Public Declare Sub LLVMCreateStandardFunctionPasses Lib "llvm2.9.dll" (ByVal FPM_ As Long, ByVal OptimizationLevel_ As LLVMCodeGenOpt)
Public Declare Sub LLVMCreateStandardModulePasses Lib "llvm2.9.dll" (ByVal PM_ As Long, ByVal OptimizationLevel_ As LLVMCodeGenOpt, ByVal OptimizeSize_ As Byte, ByVal UnitAtATime_ As Byte, ByVal UnrollLoops_ As Byte, ByVal SimplifyLibCalls_ As Byte, ByVal HaveExceptions_ As Byte, ByVal InliningPassThreshold_ As LLVMInliningPassThreshold)
Public Declare Sub LLVMCreateStandardLTOPasses Lib "llvm2.9.dll" (ByVal PM_ As Long, ByVal OptimizationLevel_ As LLVMCodeGenOpt, ByVal Internalize_ As Byte, ByVal RunInliner_ As Byte, ByVal VerifyEach_ As Byte)

Public Declare Sub LLVMAddBitcodeWriterPass Lib "llvm2.9.dll" (ByVal PM_ As Long, ByVal OS_ As Long)
Public Declare Sub LLVMAddPrintModulePass Lib "llvm2.9.dll" (ByVal PM_ As Long, ByVal OS_ As Long, ByVal DeleteStream_ As Byte, ByVal Banner_ As String)
Public Declare Sub LLVMAddPrintFunctionPass Lib "llvm2.9.dll" (ByVal PM_ As Long, ByVal OS_ As Long, ByVal DeleteStream_ As Byte, ByVal Banner_ As String)
