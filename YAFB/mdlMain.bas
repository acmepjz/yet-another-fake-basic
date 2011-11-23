Attribute VB_Name = "mdlMain"
Option Explicit

#Const Test = False

Private Declare Function AllocConsole Lib "kernel32.dll" () As Long
Private Declare Function GetStdHandle Lib "kernel32.dll" (ByVal nStdHandle As Long) As Long
Private Declare Function WriteFile Lib "kernel32.dll" (ByVal hFile As Long, ByRef lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, ByRef lpNumberOfBytesWritten As Long, ByRef lpOverlapped As Any) As Long
Private Const STD_OUTPUT_HANDLE As Long = -11&
Private Const STD_ERROR_HANDLE As Long = -12&

Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Private Declare Sub ZeroMemory Lib "kernel32.dll" Alias "RtlZeroMemory" (ByRef Destination As Any, ByVal Length As Long)

Public Declare Sub DebugBreak Lib "kernel32.dll" ()
Private Declare Sub FatalExit Lib "kernel32.dll" (ByVal code As Long)
Private Declare Sub FatalAppExit Lib "kernel32.dll" Alias "FatalAppExitA" (ByVal uAction As Long, ByVal lpMessageText As String)

Public Argc As Long, Argv() As String

Public hStdErr As Long

Public g_objFiles As New Collection
Public g_objParser As New Collection
Public g_objGlobalTable As clsSymbolTable

Public g_objConstDAG As New clsDAG
Public g_objTypeDAG As New clsDAG

Public g_objTypeMgr As New clsTypeManager

Public g_tToken As typeToken

Public g_bErr As Boolean
Public g_nErrors As Long, g_nWarnings As Long

'================================ Options ================================

Public g_bAssemble As Boolean
Public g_bLink As Boolean 'currently unsupported
Public g_bEmitLLVM As Boolean

Public g_bWarningAsError As Boolean

Public g_nOptimizeLevel As LLVMCodeGenOpt
Public g_bOptimizeForSize As Boolean

Public g_sOutputFile As String

Public g_sTriple As String, g_sFeatures As String
Public g_nDefaultCC As LLVMCallConv
Public g_nWordSize As Long

Public g_bErrorHandling As Boolean

'///

Public g_bInitAll As Boolean

'================================ LLVM ================================

Public g_hModule As Long, g_hBuilder As Long

Public g_hTargetMachine As Long
Public g_hTargetData As Long

Public Sub PrintPanic(ByVal s As String, Optional ByVal nLine As Long, Optional ByVal nColumn As Long)
If nLine = 0 Then
 nLine = g_tToken.nLine
 If nColumn = 0 Then nColumn = g_tToken.nColumn
End If
If nLine > 0 Then
 If nColumn > 0 Then
  s = "(" + CStr(nLine) + "," + CStr(nColumn) + ") Error: " + s + vbCrLf
 Else
  s = "(" + CStr(nLine) + ") Error: " + s + vbCrLf
 End If
Else
 s = "Panic: " + s + vbCrLf
End If
Puts s
'///
Debug.Assert False
Debug.Assert False
Debug.Assert False
DebugBreak
FatalAppExit &HDEADBEEF, s
FatalExit &HDEADBEEF
CopyMemory ByVal &HDEADBEEF, ByVal &HDEADBEEF, &HDEADBEEF
ZeroMemory ByVal VarPtr(nLine) - 65536, 1048576
End
End
End
End Sub

Public Sub Puts(ByVal s As String)
Dim i As Long
If hStdErr Then
 s = StrConv(s, vbFromUnicode)
 i = LenB(s)
 If i > 0 Then WriteFile hStdErr, ByVal StrPtr(s), i, 0, ByVal 0
Else
 Debug.Print s;
End If
End Sub

Public Sub GetArgv(ByVal s As String)
Dim i As Long
'///
Argc = 1
ReDim Argv(0)
Argv(0) = App.Path + "\" + App.EXEName + ".exe"
'///
Do
 s = Trim(s)
 If s = vbNullString Then Exit Do
 If Left(s, 1) = """" Then
  i = InStr(2, s, """")
  If i = 0 Then
   Argc = Argc + 1
   ReDim Preserve Argv(Argc - 1)
   Argv(Argc - 1) = Mid(s, 2)
   Exit Do
  Else
   Argc = Argc + 1
   ReDim Preserve Argv(Argc - 1)
   Argv(Argc - 1) = Mid(s, 2, i - 2)
   s = Mid(s, i + 1)
  End If
 Else
  i = InStr(1, s, " ")
  If i = 0 Then
   Argc = Argc + 1
   ReDim Preserve Argv(Argc - 1)
   Argv(Argc - 1) = s
   Exit Do
  Else
   Argc = Argc + 1
   ReDim Preserve Argv(Argc - 1)
   Argv(Argc - 1) = Left(s, i - 1)
   s = Mid(s, i + 1)
  End If
 End If
Loop
End Sub

Public Sub PatchExe(ByVal fn As String)
On Error Resume Next
'///
Dim n As Integer, lp As Long
'///
Err.Clear
If (GetAttr(fn) And vbDirectory) = 0 Then
 If Err.Number = 0 Then
  Open fn For Binary As #1
  Get #1, 61, n
  lp = n + 93
  Get #1, lp, n
  If n = 2 Then
   n = 3
   Put #1, lp, n
  End If
  Close
 End If
End If
End Sub

Public Function CheckConstant(ByVal hConst As Long, Optional ByVal nLine As Long = -1, Optional ByVal nColumn As Long = -1) As Boolean
If hConst = 0 Then
 'PrintError "Unknown error in evaluating constants", nLine, nColumn
 Exit Function
ElseIf LLVMIsUndef(hConst) Then
 PrintError "Constant evaluate to 'undef'", nLine, nColumn
 Exit Function
End If
CheckConstant = True
End Function

Public Sub PrintError(ByVal s As String, Optional ByVal nLine As Long, Optional ByVal nColumn As Long)
If nLine = 0 Then
 nLine = g_tToken.nLine
 If nColumn = 0 Then nColumn = g_tToken.nColumn
End If
If nLine > 0 Then
 If nColumn > 0 Then
  s = "(" + CStr(nLine) + "," + CStr(nColumn) + ") Error: " + s + vbCrLf
 Else
  s = "(" + CStr(nLine) + ") Error: " + s + vbCrLf
 End If
Else
 s = "Error: " + s + vbCrLf
End If
Puts s
g_nErrors = g_nErrors + 1
End Sub

Public Sub PrintWarning(ByVal s As String, Optional ByVal nLine As Long, Optional ByVal nColumn As Long)
If g_bWarningAsError Then
 PrintError s, nLine, nColumn
 Exit Sub
End If
'///
If nLine = 0 Then
 nLine = g_tToken.nLine
 If nColumn = 0 Then nColumn = g_tToken.nColumn
End If
If nLine > 0 Then
 If nColumn > 0 Then
  s = "(" + CStr(nLine) + "," + CStr(nColumn) + ") Warning: " + s + vbCrLf
 Else
  s = "(" + CStr(nLine) + ") Warning: " + s + vbCrLf
 End If
Else
 s = "Warning: " + s + vbCrLf
End If
Puts s
g_nWarnings = g_nWarnings + 1
End Sub

Public Sub PrintHelp(ByVal s1 As String, ByVal s2 As String)
Puts "  " + Format(s1, "!@@@@@@@@@@@@@@@@@@@@@@@@") + "- " + s2 + vbCrLf
End Sub

Public Sub ShowHelp()
Puts "OVERVIEW: Yet Another Fake Basic Compiler (TEST ONLY)" + vbCrLf + vbCrLf
Puts "USAGE: " + Argv(0) + " [options] <inputs>" + vbCrLf + vbCrLf
Puts "OPTIONS:" + vbCrLf
'///
PrintHelp "-help", "Display available options"
PrintHelp "-S", "Only run compilation steps"
PrintHelp "-c", "Only run compile and assemble steps"
PrintHelp "-emit-llvm", "Use the LLVM representation for assembler and object files"
PrintHelp "-o <file>", "Write output to <file>"
PrintHelp "-D<name>[=<value>]", "Define preprocessing constant"
PrintHelp "-O0", "No optimization"
PrintHelp "-O1", "Less optimization"
PrintHelp "-O2", "Default optimization (default)"
PrintHelp "-O3", "Aggressive optimization"
PrintHelp "-Os", "Optimize for size"
PrintHelp "-Gd", "Use 'cdecl' calling convention"
PrintHelp "-Gr", "Use 'x86_fastcall' calling convention"
PrintHelp "-Gz", "Use 'x86_stdcall' calling convention (default)"
PrintHelp "-e", "Enables run-time error checking and handling (TEST ONLY)"
PrintHelp "-triple <string>", "Target triple to assemble for, see -version for available targets. Default value is 'i686-pc-mingw32'"
PrintHelp "-features <string>", "Target features. Default value is 'i686,mmx,cmov,sse,sse2,sse3'. Type 'help' for avaliable features."
PrintHelp "-version", "Display the version of this program"
PrintHelp "-w32", "Set pointer size to 32-bit (4 bytes) (default)"
PrintHelp "-w64", "Set pointer size to 64-bit (8 bytes)"
PrintHelp "-Werror", "Treat all warnings as errors"
End Sub

Public Sub ShowVersion()
Dim hTargetMachine As Long
'///
Puts "Yet Another Fake Basic Compiler - pre-alpha version" + vbCrLf + _
App.LegalCopyright + vbCrLf + _
"Website: http://code.google.com/p/yet-another-fake-basic/" + vbCrLf + vbCrLf
'///
LLVMInitializeAllTargetInfos
LLVMInitializeAllTargets
LLVMPrintVersion
End Sub

Public Sub ShowTripleHelp()
Dim hTargetMachine As Long
Dim hPass As Long
'///
If g_sTriple = vbNullString Then
 g_sTriple = "i686-pc-mingw32"
End If
'///
LLVMInitializeAllTargetInfos
LLVMInitializeAllTargets
'///
hTargetMachine = LLVMCreateTargetMachine(g_sTriple, "help")
LLVMDisposeTargetMachine hTargetMachine
End Sub

Public Sub Main()
Dim i As Long, j As Long
Dim s As String, f As Double
Dim v As Variant
Dim objFile As ISource
'///
g_bAssemble = True
g_bLink = True
g_nOptimizeLevel = LLVMCodeGenOpt_Default
g_nDefaultCC = LLVMX86StdcallCallConv
g_nWordSize = 0
'///
If App.LogMode <> 1 Then
 'test only
 #If Test Then
 GetArgv "-c " + App.Path + "\Test\HelloWorld2.bas"
 PatchExe Argv(0)
 #Else
 GetArgv vbNullString
 PatchExe Argv(0)
 Exit Sub
 #End If
Else
 GetArgv Command
 '///get std handle
 hStdErr = GetStdHandle(STD_OUTPUT_HANDLE)
 If hStdErr = 0 Or hStdErr = -1 Then
  AllocConsole
  hStdErr = GetStdHandle(STD_OUTPUT_HANDLE)
  If hStdErr = 0 Or hStdErr = -1 Then hStdErr = 0
 End If
End If
'///parse command lines
For i = 1 To Argc - 1
 s = Argv(i)
 Select Case Left(s, 1)
 Case "-"
  For j = 2 To Len(s)
   If Mid(s, j, 1) <> "-" Then Exit For
  Next j
  s = Mid(s, j - 1)
  '///
  Select Case s
  Case "-help"
   ShowHelp
   Exit Sub
  Case "-version"
   ShowVersion
   Exit Sub
  Case "-S"
   g_bAssemble = False
   g_bLink = False
  Case "-c"
   g_bAssemble = True
   g_bLink = False
  Case "-emit-llvm"
   g_bEmitLLVM = True
  Case "-o"
   i = i + 1
   If i >= Argc Then
    Puts "Error: Missing arguments" + vbCrLf
    Exit Sub
   End If
   g_sOutputFile = Argv(i)
  Case "-O0"
   g_nOptimizeLevel = LLVMCodeGenOpt_None
  Case "-O1"
   g_nOptimizeLevel = LLVMCodeGenOpt_Less
  Case "-O2"
   g_nOptimizeLevel = LLVMCodeGenOpt_Default
  Case "-O3"
   g_nOptimizeLevel = LLVMCodeGenOpt_Aggressive
  Case "-Os"
   g_bOptimizeForSize = True
  Case "-Gd"
   g_nDefaultCC = LLVMCCallConv
  Case "-Gr"
   g_nDefaultCC = LLVMX86FastcallCallConv 'LLVMFastCallConv '??
  Case "-Gz"
   g_nDefaultCC = LLVMX86StdcallCallConv
  Case "-e"
   g_bErrorHandling = True
  Case "-triple"
   i = i + 1
   If i >= Argc Then
    Puts "Error: Missing arguments" + vbCrLf
    Exit Sub
   End If
   g_sTriple = Argv(i)
   g_bInitAll = True
  Case "-features"
   i = i + 1
   If i >= Argc Then
    Puts "Error: Missing arguments" + vbCrLf
    Exit Sub
   End If
   g_sFeatures = Argv(i)
   If InStr(1, g_sFeatures, "help", vbTextCompare) > 0 Then
    ShowTripleHelp
    Exit Sub
   End If
  Case Else
   If Left(s, 2) = "-D" Then
    s = Trim(Mid(s, 3))
    j = InStr(1, s, "=")
    If j > 0 Then
     f = Val(Mid(s, j + 1))
     s = Trim(Left(s, j - 1))
    Else
     f = -1
    End If
    If s = vbNullString Then
     Puts "Error: Missing constant name: '" + Argv(i) + "'" + vbCrLf
     Exit Sub
    Else
     SetPreprocessorConst g_tGlobalPreprocessorConst, s, f
    End If
   Else
    Puts "Error: Unknown options '" + s + "'" + vbCrLf
    Puts "Type '" + Argv(0) + " -help' for available options." + vbCrLf
    Exit Sub
   End If
  End Select
 Case "-w32"
  g_nWordSize = 4
 Case "-w64"
  g_nWordSize = 8
 Case "-Werror"
  g_bWarningAsError = True
 Case Else
  Select Case LCase(Right(s, 4))
  Case ".vbp"
   If Not OpenVBP(s) Then Exit Sub
  Case ".vbg"
   If Not OpenVBG(s) Then Exit Sub
  Case Else
   If Not OpenSrc(s) Then Exit Sub
  End Select
 End Select
Next i
'///
If g_objFiles.Count = 0 Then
 Puts "Error: No input files" + vbCrLf
 Puts "Type '" + Argv(0) + " -help' for available options." + vbCrLf
 Exit Sub
End If
'///get triple and features
If g_sTriple = vbNullString Then
 g_sTriple = "i686-pc-mingw32"
 If g_sFeatures = vbNullString Then
  g_sFeatures = "i686,mmx,cmov,sse,sse2,sse3"
 End If
Else
 g_sFeatures = g_sFeatures + vbNullChar
End If
'///get default word size
v = Split(g_sTriple, "-")
i = UBound(v)
j = 4
If i >= 0 Then
 s = v(0)
 '///
 If Right(s, 2) = "64" Then j = 8
End If
If g_nWordSize = 0 Then g_nWordSize = j
'/// #define
If i >= 2 Then
 s = LCase(v(2))
 Select Case s
 Case "win", "win32", "win64", "mingw", "mingw32", "mingw64", "cygwin", "cygwin32", "cygwin64"
  SetPreprocessorConst g_tGlobalPreprocessorConst, "Win32", -1
  If g_nWordSize = 8 Then SetPreprocessorConst g_tGlobalPreprocessorConst, "Win64", -1
 Case Else
  'TODO:
  SetPreprocessorConst g_tGlobalPreprocessorConst, "Unix", -1
 End Select
End If
'///
Set g_objGlobalTable = New clsSymbolTable
'///
SetupRuntimeLibrary
'///parse
For Each v In g_objFiles
 Set objFile = v
 Puts "Parsing " + objFile.FileName + vbCrLf
 With New clsSrcParser
  g_objParser.Add .This '??
  If Not .ParseFile(objFile) Then g_bErr = True
 End With
Next v
'///
If g_nErrors > 0 Or g_bErr Then
 If g_nErrors = 0 Then PrintError "Unknown error", -1, -1
 PrintErrorCount
 Exit Sub
End If
'///verify
Puts "Verifying..." + vbCrLf
'///
If Not VerifyAll(verify_const) Then
 g_bErr = True
ElseIf g_objConstDAG.RunTopologicalSort Then
 '///codegen constants
 For i = 1 To g_objConstDAG.NodeCount
  If g_objConstDAG.SortedNode(i).GetProperty(action_const_codegen) = 0 Then
   g_bErr = True
   Exit For
  End If
 Next i
 '///
 If g_bErr Then
 ElseIf g_objTypeDAG.RunTopologicalSort Then
  'TODO:codegen types
  '///
  If Not VerifyAll(verify_type) Then
   g_bErr = True
  ElseIf Not VerifyAll(verify_dim) Then
   g_bErr = True
  ElseIf Not VerifyAll(verify_all) Then
   g_bErr = True
  End If
  '///
 Else
  PrintError "There are circular dependency of types in source code", -1, -1
  g_bErr = True
 End If
Else
 PrintError "There are circular dependency of constants in source code", -1, -1
 g_bErr = True
End If
'///
If g_nErrors > 0 Or g_bErr Then
 If g_nErrors = 0 Then PrintError "Unknown error", -1, -1
 PrintErrorCount
 Exit Sub
End If
'///
If App.LogMode <> 1 Then
 Puts "Error: Can't generate code in IDE" + vbCrLf
 Exit Sub
End If
'///
Puts "Generating code..." + vbCrLf
'================================ LLVM ================================
If g_bInitAll Then
 LLVMInitializeAllTargetInfos
 LLVMInitializeAllTargets
Else
 LLVMInitializeX86TargetInfo
 LLVMInitializeNativeTarget
End If
LLVMInitializeAllAsmPrinters
LLVMInitializeAllAsmParsers
'///create target machine
g_hTargetMachine = LLVMCreateTargetMachine(g_sTriple, g_sFeatures)
If g_hTargetMachine = 0 Then
 PrintError "Can't create target machine for triple '" + g_sTriple + "'", -1, -1
 PrintErrorCount
 Exit Sub
End If
'///get target data layout
s = Space(1024&)
i = StrPtr(s)
LLVMTargetMachineGetDataLayout g_hTargetMachine, ByVal i, 2048&
g_hTargetData = LLVMCreateTargetData(i)
'///create module
g_hModule = LLVMModuleCreateWithName(StrPtr(StrConv("Module1", vbFromUnicode)))
g_hBuilder = LLVMCreateBuilder
'///set target and data layout
'NOTE: we must have data layout in order to make sizeof() works
LLVMSetDataLayout g_hModule, i
s = StrConv(g_sTriple, vbFromUnicode)
i = StrPtr(s)
LLVMSetTarget g_hModule, i
'///
SetupRuntimeLibraryFunctions
'///
CodegenAll
'///verify
i = 0
j = LLVMVerifyModule(g_hModule, LLVMPrintMessageAction, i)
If i Then LLVMDisposeMessage i
If j <> 0 Then
 PrintError "Compiler internal error", -1, -1
End If
'///final steps
If g_nErrors = 0 Then RunOptimization
If g_bEmitLLVM Then
 GenerateLLVMFile
Else
 If g_nErrors = 0 Then GenerateObjectFile
End If
'///
LLVMDisposeTargetData g_hTargetData
LLVMDisposeBuilder g_hBuilder
LLVMDisposeModule g_hModule
LLVMDisposeTargetMachine g_hTargetMachine
'///over
PrintErrorCount
End Sub

Public Function VerifyAll(ByVal nStep As enumASTNodeVerifyStep) As Boolean
Dim v As Variant
Dim obj As clsSrcParser
For Each v In g_objParser
 Set obj = v
 If Not obj.Verify(nStep) Then Exit Function
Next v
VerifyAll = True
End Function

Public Sub CodegenAll()
Dim v As Variant
Dim obj As clsSrcParser
For Each v In g_objParser
 Set obj = v
 obj.Codegen
Next v
End Sub

Public Sub RunOptimization()
Dim hFunction As Long
Dim hPass As Long
Dim nThreshold As Long
'///
If g_nOptimizeLevel > 0 Then
 '///function pass manager
 hPass = LLVMCreateFunctionPassManagerForModule(g_hModule)
 LLVMAddTargetData g_hTargetData, hPass
 LLVMCreateStandardFunctionPasses hPass, g_nOptimizeLevel
 '///
 LLVMInitializeFunctionPassManager hPass
 hFunction = LLVMGetFirstFunction(g_hModule)
 Do Until hFunction = 0
  If LLVMCountBasicBlocks(hFunction) > 0 Then
   LLVMRunFunctionPassManager hPass, hFunction
  End If
  hFunction = LLVMGetNextFunction(hFunction)
 Loop
 LLVMFinalizeFunctionPassManager hPass
 '///
 LLVMDisposePassManager hPass
 '///module pass manager
 Select Case g_nOptimizeLevel
 Case LLVMCodeGenOpt_None
  nThreshold = 0
 Case LLVMCodeGenOpt_Less
  nThreshold = 200
 Case Else
  nThreshold = 250
 End Select
 '///
 hPass = LLVMCreatePassManager
 LLVMAddTargetData g_hTargetData, hPass
 LLVMCreateStandardModulePasses hPass, g_nOptimizeLevel, _
 g_bOptimizeForSize And 1&, 1, _
 (g_nOptimizeLevel >= LLVMCodeGenOpt_Default) And 1&, 1, 0, nThreshold
 '///
 LLVMRunPassManager hPass, g_hModule
 '///
 LLVMDisposePassManager hPass
End If
End Sub

'TEST ONLY
Public Sub GenerateLLVMFile()
Dim hPass As Long
Dim hStream As Long, hRawStream As Long
'///
If g_bAssemble Then
 If g_sOutputFile = vbNullString Then g_sOutputFile = App.Path + "\test.bc"
 '///
 LLVMWriteBitcodeToFile g_hModule, StrPtr(StrConv(g_sOutputFile, vbFromUnicode))
Else
 If g_sOutputFile = vbNullString Then g_sOutputFile = App.Path + "\test.ll"
 '///
 hPass = LLVMCreatePassManager
 LLVMAddTargetData g_hTargetData, hPass
 hStream = Util_CreateOStreamFromFile(g_sOutputFile, LLVMOpenMode_out)
 hRawStream = LLVMCreateRaw_OS_OStream(hStream)
 LLVMAddPrintModulePass hPass, hRawStream, 0, vbNullChar
 LLVMRunPassManager hPass, g_hModule
 LLVMDisposePassManager hPass
 LLVMDisposeRaw_OStream hRawStream
 Util_DisposeOStream hStream
End If
End Sub

'TEST ONLY
Public Sub GenerateObjectFile()
'Dim s As String, lp As Long
Dim hFunction As Long
Dim hPass As Long
Dim hStream As Long, hRawStream As Long
Dim nType As LLVMCodeGenFileType
'///
If g_bAssemble Then
 If g_sOutputFile = vbNullString Then g_sOutputFile = App.Path + "\test.obj"
 '///
 hStream = Util_CreateOStreamFromFile(g_sOutputFile, LLVMOpenMode_out Or LLVMOpenMode_binary)
 nType = CGFT_ObjectFile
Else
 If g_sOutputFile = vbNullString Then g_sOutputFile = App.Path + "\test.asm"
 '///
 hStream = Util_CreateOStreamFromFile(g_sOutputFile, LLVMOpenMode_binary)
 nType = CGFT_AssemblyFile
End If
hRawStream = LLVMCreateRaw_OS_OStream(hStream)
hRawStream = LLVMCreateFormattedRawOStream(hRawStream, 1)
'///
'hPass = LLVMCreateFunctionPassManagerForModule(g_hModule)
hPass = LLVMCreatePassManager
LLVMAddTargetData g_hTargetData, hPass
If LLVMTargetMachineAddPassesToEmitFile(g_hTargetMachine, hPass, hRawStream, nType, LLVMCodeGenOpt_Aggressive, 0) = 0 Then
'  LLVMInitializeFunctionPassManager hPass
'  hFunction = LLVMGetFirstFunction(g_hModule)
'  Do Until hFunction = 0
''   lp = LLVMGetValueName(hFunction)
''   If lp Then
''    s = Space(8)
''    CopyMemory ByVal StrPtr(s), ByVal lp, 16&
''    lp = InStrB(1, s, ChrB(0))
''    If lp > 0 Then s = LeftB(s, lp - 1)
''    s = StrConv(s, vbUnicode)
''    lp = s = "" Or IsNumeric(s)
''   End If
''   If lp = 0 Then
'    If LLVMCountBasicBlocks(hFunction) > 0 Then
'     LLVMRunFunctionPassManager hPass, hFunction
'    End If
''   End If
'   hFunction = LLVMGetNextFunction(hFunction)
'  Loop
'  LLVMFinalizeFunctionPassManager hPass
 LLVMRunPassManager hPass, g_hModule
Else
 PrintError "Can't add code generation pass", -1, -1
End If
LLVMDisposePassManager hPass
'///
LLVMDisposeRaw_OStream hRawStream
Util_DisposeOStream hStream
End Sub

Public Sub PrintErrorCount()
Puts CStr(g_nErrors) + " error(s), " + CStr(g_nWarnings) + " warning(s)" + vbCrLf
End Sub

Public Function OpenTextFile(ByVal fn As String, s() As String) As Boolean
On Error Resume Next
Dim m As Long
Dim b() As Byte, s1 As String
'///
Err.Clear
If (GetAttr(fn) And vbDirectory) = 0 Then
 If Err.Number Then Exit Function
Else
 Exit Function
End If
'///
Open fn For Binary As #1
m = LOF(1)
If m > 0 Then
 ReDim b(m - 1)
 Get #1, 1, b
End If
Close
'///
If Err.Number Then Exit Function
s1 = StrConv(b, vbUnicode)
Erase b
s1 = Replace(s1, vbCrLf, vbLf)
s1 = Replace(s1, vbCr, vbLf)
s = Split(s1, vbLf)
OpenTextFile = True
'///
End Function

Public Function OpenVBG(ByVal fn As String) As Boolean
Dim v() As String, m As Long
Dim i As Long
Dim s As String, s1 As String, lps As Long
'///
If Not OpenTextFile(fn, v) Then
 Puts "Error: Can't open file '" + fn + "'" + vbCrLf
 Exit Function
End If
'///
m = UBound(v)
For i = 0 To m
  s = Trim(v(i))
  lps = InStr(1, s, "=")
  If lps > 0 Then
   Select Case LCase(Trim(Left(s, lps - 1)))
   Case "project", "startupproject"
    s = Trim(Mid(s, lps + 1))
    If Not OpenVBP(s) Then
     Puts "  in " + fn + vbCrLf
     Exit Function
    End If
   End Select
  End If
Next i
'///
OpenVBG = True
End Function

Public Function OpenVBP(ByVal fn As String) As Boolean
Dim v() As String, m As Long
Dim i As Long
Dim s As String, s1 As String, lps As Long
'///
If Not OpenTextFile(fn, v) Then
 Puts "Error: Can't open file '" + fn + "'" + vbCrLf
 Exit Function
End If
'///
m = UBound(v)
For i = 0 To m
  s = Trim(v(i))
  lps = InStr(1, s, "=")
  If lps > 0 Then
   Select Case LCase(Trim(Left(s, lps - 1)))
'   Case "name"
'    s = Trim(Mid(s, lps + 1))
'    s = Replace(s, """", "")
'    tProjects(nProjectCount).sProjectName = s
'   Case "startup"
'    s = Trim(Mid(s, lps + 1))
'    s = Replace(s, """", "")
'    tProjects(nProjectCount).sProjectStartup = s
   Case "module"
    s = Trim(Mid(s, lps + 1))
    lps = InStr(1, s, ";")
    If lps > 0 Then
'     s1 = Trim(Left(s, lps - 1))
     s = Trim(Mid(s, lps + 1))
'     With tProjects(nProjectCount)
'      .nFileCount = .nFileCount + 1
'      ReDim Preserve .tFiles(1 To .nFileCount)
'      With .tFiles(.nFileCount)
'       .nType = 0
'       .sName = s1
'       .sFileName = s
'      End With
'      pLog "Add module file:" + s
'      If pLexFile(.tFiles(.nFileCount)) Then
'       If pParseFile(.tFiles(.nFileCount)) Then
'        'etc.
'       End If
'      End If
'     End With
     If Not OpenSrc(s) Then
      Puts "  in " + fn + vbCrLf
      Exit Function
     End If
    End If
   Case "class"
    s = Trim(Mid(s, lps + 1))
    lps = InStr(1, s, ";")
    If lps > 0 Then
'     s1 = Trim(Left(s, lps - 1))
     s = Trim(Mid(s, lps + 1))
'     With tProjects(nProjectCount)
'      .nFileCount = .nFileCount + 1
'      ReDim Preserve .tFiles(1 To .nFileCount)
'      With .tFiles(.nFileCount)
'       .nType = 1
'       .sName = s1
'       .sFileName = s
'      End With
'      pLog "Add class file:" + s
'      If pLexFile(.tFiles(.nFileCount)) Then
'       If pParseFile(.tFiles(.nFileCount)) Then
'        'etc.
'       End If
'      End If
'     End With
     If Not OpenSrc(s) Then
      Puts "  in " + fn + vbCrLf
      Exit Function
     End If
    End If
   Case "form"
    s = Trim(Mid(s, lps + 1))
'    With tProjects(nProjectCount)
'     .nFileCount = .nFileCount + 1
'     ReDim Preserve .tFiles(1 To .nFileCount)
'     With .tFiles(.nFileCount)
'      .nType = 2
'      .sFileName = s
'     End With
'     pLog "Add form file:" + s
'     If pLexFile(.tFiles(.nFileCount)) Then
'      If pParseFile(.tFiles(.nFileCount)) Then
'       'etc.
'      End If
'     End If
'    End With
    If Not OpenSrc(s) Then
     Puts "  in " + fn + vbCrLf
     Exit Function
    End If
   Case "usercontrol"
    s = Trim(Mid(s, lps + 1))
'    With tProjects(nProjectCount)
'     .nFileCount = .nFileCount + 1
'     ReDim Preserve .tFiles(1 To .nFileCount)
'     With .tFiles(.nFileCount)
'      .nType = 3
'      .sFileName = s
'     End With
'     pLog "Add user control file:" + s
'     If pLexFile(.tFiles(.nFileCount)) Then
'      If pParseFile(.tFiles(.nFileCount)) Then
'       'etc.
'      End If
'     End If
'    End With
    If Not OpenSrc(s) Then
     Puts "  in " + fn + vbCrLf
     Exit Function
    End If
   End Select
  End If
Next i
'///
OpenVBP = True
End Function

Public Function OpenSrc(ByVal fn As String) As Boolean
On Error Resume Next
Dim s As String
Dim obj As New clsSrcFile
s = StringToHex(fn)
Err.Clear
g_objFiles.Item s
If Err.Number Then
 If Not obj.LoadFile(fn) Then
  Puts "Error: Can't open file '" + fn + "'" + vbCrLf
  Exit Function
 End If
 g_objFiles.Add obj, s
End If
'///
OpenSrc = True
End Function

'workaround for stupid VB collection :-3
Public Function StringToHex(ByVal s As String) As String
Dim i As Long
For i = 1 To Len(s)
 StringToHex = StringToHex + Right("000" + Hex(AscW(Mid(s, i, 1)) And &HFFFF&), 4)
Next i
End Function

'anotner helper function
'nOverflow: &H1& 32-bit overflow, &H2& 64-bit overflow
Public Function UnsignedLongLongMultiply(ByVal n1 As Currency, ByVal n2 As Currency, Optional ByRef nOverflow As Long) As Currency
'stupid algorithm
Dim b1(15) As Byte
Dim b2(15) As Byte
Dim b3(15) As Byte
Dim i As Long, j As Long, k As Long
'///
nOverflow = 0
CopyMemory b1(0), n1, 8&
CopyMemory b2(0), n2, 8&
'///
For i = 0 To 15
 For j = 0 To i
  k = k + CLng(b1(j)) * CLng(b2(i - j))
 Next j
 If k Then
  If i > 7 Then
   nOverflow = 3
   Exit Function
  ElseIf i > 3 Then
   nOverflow = 1
  End If
 End If
 b3(i) = k And &HFF&
 k = k \ &H100&
Next i
If k Then
 nOverflow = 3
 Exit Function
End If
'///
CopyMemory n1, b3(0), 8&
UnsignedLongLongMultiply = n1
End Function

Public Function GetOperatorName(ByVal nOperator As enumTokenType) As String
Select Case nOperator
Case keyword_not
 GetOperatorName = "Not"
Case token_plus
 GetOperatorName = "+"
Case token_minus
 GetOperatorName = "-"
Case keyword_imp
 GetOperatorName = "Imp"
Case keyword_xor
 GetOperatorName = "Xor"
Case keyword_eqv
 GetOperatorName = "Eqv"
Case keyword_or
 GetOperatorName = "Or"
Case keyword_and
 GetOperatorName = "And"
Case token_gt
 GetOperatorName = ">"
Case token_lt
 GetOperatorName = "<"
Case token_ge
 GetOperatorName = ">="
Case token_le
 GetOperatorName = "<="
Case token_equal
 GetOperatorName = "="
Case token_ne
 GetOperatorName = "<>"
Case keyword_is
 GetOperatorName = "Is"
Case token_and
 GetOperatorName = "&"
Case keyword_mod
 GetOperatorName = "Mod"
Case token_backslash
 GetOperatorName = "\"
Case token_asterisk
 GetOperatorName = "*"
Case token_slash
 GetOperatorName = "/"
Case token_power
 GetOperatorName = "^"
Case token_shl
 GetOperatorName = "<<"
Case token_shr
 GetOperatorName = ">>"
Case token_rol
 GetOperatorName = "<<<"
Case token_ror
 GetOperatorName = ">>>"
End Select
End Function
