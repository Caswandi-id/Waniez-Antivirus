Attribute VB_Name = "Modprotec"
Option Explicit

Private Const MEM_DECOMMIT = &H4000
Private Const MEM_RELEASE = &H8000
Private Const MEM_COMMIT = &H1000
Private Const MEM_RESERVE = &H2000
Private Const MEM_RESET = &H80000
Private Const MEM_TOP_DOWN = &H100000
Private Const PAGE_READONLY = &H2
Private Const PAGE_READWRITE = &H4
Private Const PAGE_EXECUTE = &H10
Private Const PAGE_EXECUTE_READ = &H20
Private Const PAGE_EXECUTE_READWRITE = &H40
Private Const PAGE_GUARD = &H100
Private Const PAGE_NOACCESS = &H1
Private Const PAGE_NOCACHE = &H200
Private Declare Function VirtualAlloc Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal flAllocationType As Long, ByVal flProtect As Long) As Long
Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)
Private Declare Function CreateThread Lib "kernel32" (lpThreadAttributes As Any, ByVal dwStackSize As Long, lpStartAddress As Long, lpParameter As Any, ByVal dwCreationFlags As Long, lpThreadId As Long) As Long
Function InstallAntiDebugger() As Long
Dim ThreadID As Long
Dim ThreadEntryPoint As Long
Dim ThreadCode As String
Dim ThreadCodeByte() As Byte
Dim ModuleHandle As Long
Dim ProcIDPAddr As Long
Dim ProcGCPAddr As Long
Dim ProcTPAddr As Long
Dim ProcSPAddr As Long

'This is the assembler code to check when your application is beeing debugged
'----------------------------------------------------------------------------
'00401FBC      BF B1F5577C   MOV EDI,KERNEL32.IsDebuggerPresent
'00401FC1      FFD7          CALL EDI
'00401FC3      83F8 01       CMP EAX,1
'00401FC6      75 0F         JNZ SHORT 00401FD7
'00401FC8      BF 2579597C   MOV EDI,KERNEL32.GetCurrentProcess
'00401FCD      FFD7          CALL EDI
'00401FCF      50            PUSH EAX
'00401FD0      BF 6D6A597C   MOV EDI,KERNEL32.TerminateProcess
'00401FD5      FFD7          CALL EDI
'00401FD7      BF 91A2597C   MOV EDI,KERNEL32.Sleep
'00401FDC      B8 10270000   MOV EAX,2710                           ;Sleep 10 seconds before check the debugger again
'00401FE1      50            PUSH EAX
'00401FE2      FFD7          CALL EDI
'00401FE4    ^ EB D6         JMP SHORT 00401FBC

'Get the module entry point for the kernel32.dll
ModuleHandle = LoadLibrary("Kernel32.dll")
If ModuleHandle = 0 Then
    InstallAntiDebugger = 0
Else
    'Get the function address
    ProcIDPAddr = GetProcAddress(ModuleHandle, "IsDebuggerPresent")
    ProcGCPAddr = GetProcAddress(ModuleHandle, "GetCurrentProcess")
    ProcTPAddr = GetProcAddress(ModuleHandle, "TerminateProcess")
    ProcSPAddr = GetProcAddress(ModuleHandle, "Sleep")
    
    'Build the assembler code (opcodes)
    ThreadCode = "BF" & AlignDWORD(ProcIDPAddr) & _
                 "FFD7" & _
                 "83F801" & _
                 "750F" & _
                 "BF" & AlignDWORD(ProcGCPAddr) & _
                 "FFD7" & _
                 "50" & _
                 "BF" & AlignDWORD(ProcTPAddr) & _
                 "FFD7" & _
                 "BF" & AlignDWORD(ProcSPAddr) & _
                 "B810270000" & _
                 "50" & _
                 "FFD7" & _
                 "EBD6"
    'Transform the string into a byte array
    ConvHEX2ByteArray ThreadCode, ThreadCodeByte
    
    'Allocate virtual memory to install our code
    ThreadEntryPoint = VirtualAlloc(0, UBound(ThreadCodeByte) - LBound(ThreadCodeByte) + 1, MEM_COMMIT, PAGE_EXECUTE_READWRITE)
    If ThreadEntryPoint <> 0 Then
        'Copy the assembler codes from our array into the new allocated virtual memory
        CopyMemory ByVal ThreadEntryPoint, ByVal VarPtr(ThreadCodeByte(LBound(ThreadCodeByte))), ByVal UBound(ThreadCodeByte) - LBound(ThreadCodeByte) + 1
        
        'Start the new thread, using as entry point then start of allocated virtual memory
        CreateThread ByVal 0&, ByVal 0&, ByVal ThreadEntryPoint, ByVal 0&, ByVal 0&, ThreadID
        
        'Return the threadid for future uses on ur main program (like suspendthread, resumethread, etc)
        InstallAntiDebugger = ThreadID
    Else
        InstallAntiDebugger = 0
    End If

End If
End Function


Sub ConvHEX2ByteArray(pStr As String, pByte() As Byte)
Dim i As Long
Dim J As Long
ReDim pByte(1 To Len(pStr))
For i = 1 To Len(pStr) Step 2
    J = J + 1
    pByte(J) = CByte("&H" & Mid(pStr, i, 2))
Next
End Sub



Private Function AlignDWORD(pParam As Long) As String
Dim HiW As Integer
Dim LoW As Integer

Dim HiBHiW As Byte
Dim HiBLoW As Byte

Dim LoBHiW As Byte
Dim LoBLoW As Byte

HiW = hiword(pParam)
LoW = loword(pParam)

HiBHiW = HiByte(HiW)
HiBLoW = HiByte(LoW)

LoBHiW = LoByte(HiW)
LoBLoW = LoByte(LoW)

AlignDWORD = IIf(Len(Hex(LoBLoW)) = 1, "0" & Hex(LoBLoW), Hex(LoBLoW)) & _
         IIf(Len(Hex(HiBLoW)) = 1, "0" & Hex(HiBLoW), Hex(HiBLoW)) & _
         IIf(Len(Hex(LoBHiW)) = 1, "0" & Hex(LoBHiW), Hex(LoBHiW)) & _
         IIf(Len(Hex(HiBHiW)) = 1, "0" & Hex(HiBHiW), Hex(HiBHiW))

End Function

Private Function HiByte(ByVal wParam As Integer) As Byte

    HiByte = (wParam And &HFF00&) \ (&H100)

End Function

Private Function hiword(DWord As Long) As Integer
   hiword = (DWord And &HFFFF0000) \ &H10000
End Function

Private Function LoByte(ByVal wParam As Integer) As Byte

  LoByte = wParam And &HFF&

End Function

Private Function loword(DWord As Long) As Integer
   If DWord And &H8000& Then ' &H8000& = &H00008000
      loword = DWord Or &HFFFF0000
   Else
      loword = DWord And &HFFFF&
   End If
End Function





