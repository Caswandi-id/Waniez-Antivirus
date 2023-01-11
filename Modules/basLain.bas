Attribute VB_Name = "basLain"
 
' Module untuk penanganan fungsi-fungsi tambahan gak jelas :D
'
Public Declare Function Beep Lib "kernel32" (ByVal dwFreq As Long, _
ByVal dwDuration As Long) As Long

Private Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
End Type

Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long

Declare Function GetDriveType& Lib "kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String)
' Gak unicode ga papa :D
Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Declare Function MessageBoxW Lib "user32" (ByVal hwnd As Long, ByVal lpText As Long, ByVal lpCaption As Long, ByVal wType As Long) As Long
Private Declare Sub ReleaseCapture Lib "user32" ()
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function SHGetSpecialFolderPath Lib "shell32.dll" Alias "SHGetSpecialFolderPathA" (ByVal hwnd As Long, ByVal pszPath As String, ByVal CSIDL As Long, ByVal fCreate As Boolean) As Boolean

Private Declare Function InternetGetConnectedState Lib "wininet.dll" (ByRef _
    lpSFlags As Long, ByVal dwReserved As Long) As Long
'new
Public Enum SHGFI
    C_LARGEICON = &H0
    C_USEFILEATTRIBUTES = &H10
    C_TYPENAME = &H400
    C_SMALLICON = &H1
    C_SYSICONINDEX = &H4000
    C_SHELLICONSIZE = &H4
    C_DISPLAYNAME = &H200
    C_EXETYPE = &H2000
    C_BASIC_FLAGS = SHGFI.C_TYPENAME Or SHGFI.C_SHELLICONSIZE Or SHGFI.C_SYSICONINDEX Or SHGFI.C_DISPLAYNAME Or SHGFI.C_EXETYPE
End Enum
Const WM_NCLBUTTONDOWN = &HA1 'Saat tombol kiri mouse ditekan.
Const HTCAPTION = 2
Public Const fKiri = 2400, fAtas = 1200, fTinggi = 6615, fLebar = 8295
Private Const HWND_NOTOPMOST = -2
Private Const HWND_TOPMOST = -1
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOSIZE = &H1
Private Const TOPFLAGS = SWP_NOMOVE Or SWP_NOSIZE
Private cCtlVir  As New gComCtl
'============
Const INTERNET_CONNECTION_MODEM = 1
Const INTERNET_CONNECTION_LAN = 2
Const INTERNET_CONNECTION_PROXY = 4
Const INTERNET_CONNECTION_MODEM_BUSY = 8


Public Enum IDFolder
    ALL_USER_STARTUP = &H18
    WINDOWS_DIR = &H24
    SYSTEM_DIR = &H25
    PROGRAM_FILE = &H26
    USER_DOC = &H5
    USER_STARTUP = &H7
    RECENT_DOC = &H8
    DEKSTOP_PATH = &H19
End Enum


Private Function GetWindowsVersion() As String
Dim OSINFO As OSVERSIONINFO
Dim ret As Integer
OSINFO.dwOSVersionInfoSize = 148
OSINFO.szCSDVersion = Space$(128)
ret = GetVersionEx(OSINFO)

With OSINFO
Select Case .dwPlatformId
    Case 2 ' udah NT
        If .dwMajorVersion = 3 Then
           GetWindowsVersion = "Win NT 3.51"
        ElseIf .dwMajorVersion = 4 Then
            GetWindowsVersion = "Win NT 4.0"
        ElseIf .dwMajorVersion = 5 Then
            If .dwMinorVersion = 0 Then
                GetWindowsVersion = "Win 2000"
            Else
                GetWindowsVersion = "WinXP"
            End If
        End If
     Case Else
        GetWindowsVersion = "Gak Tau - Gak Penting"
     End Select
End With
End Function
Public Function IsWinXPOS() As Boolean
Dim RetString As String
RetString = GetWindowsVersion
If RetString = "WinXP" Then IsWinXPOS = True Else IsWinXPOS = False
End Function

' Fungsi Utama daptakan Path Folder Spesial
Public Function GetSpecFolder(ByVal lpCSIDL As IDFolder) As String

Dim sPath As String
Dim lRet As Long
    
    sPath = String$(255, 0)
    
    lRet = SHGetSpecialFolderPath(0&, sPath, lpCSIDL, False)
    
    If lRet <> 0 Then
        GetSpecFolder = FixBuffer(sPath)
    End If
    
End Function

Private Function FixBuffer(ByVal sBuffer As String) As String

Dim NullPos As Long
    
    NullPos = InStr(sBuffer, Chr$(0))
    
    If NullPos > 0 Then
        FixBuffer = Left$(sBuffer, NullPos - 1)
    End If
    
End Function

' Mendapatkan drive flash dan mutual
Public Function GetDrive(ByRef sDrive() As String) As Long
Dim PKey    As Byte
Dim nDrve  As Long
Dim dType   As Long
ReDim sDrive(8) As String
For PKey = 0 To 8 ' Mulai dari Drive C:\ --> K:\ aja
    dType = GetDriveType(Chr$(67 + PKey) & ":\")
    If dType = 2 Or dType = 3 Then
        nDrve = nDrve + 1
        sDrive(nDrve) = Chr$(67 + PKey) & ":\"
    End If
Next
GetDrive = nDrve
End Function
' Extract Resource data
Public Function ExtractRes(Pathnya As String, id As Integer, sTipe As String) ' Tidak Mendukung Unicode [ektraknya ke folder2 ansi]
On Error Resume Next
    HapusFile Pathnya
    WriteUnicodeFile Pathnya, 1, LoadResData(id, sTipe)
End Function
Public Function BuangSpaceAwal(ByVal sKar As String) As String
If Left(sKar, 1) = Chr$(32) Then
    BuangSpaceAwal = Mid(sKar, 2)
Else
    BuangSpaceAwal = sKar
End If

End Function
Public Function CariIndekItemTerpilih(LvInput As ucListView) 'base1
Dim CNTA As Long
On Error Resume Next
For CNTA = 1 To LvInput.ListItems.Count
    If LvInput.ListItems.Item(CNTA).Selected = True Then
       CariIndekItemTerpilih = CNTA
       Exit For
    End If
Next
End Function

' Membedakan path dari registry dan path dari file biasa (untuk keperluan kecil ajh)
Public Function ApakahPathRegistry(sPathNormal As String) As Boolean
' jika data ke 2 sebanyak 2 kar adalah ":\" = path file
If Mid$(sPathNormal, 2, 2) = ":\" Then
   ApakahPathRegistry = False
Else
   ApakahPathRegistry = True ' anggap tue
End If
End Function
Public Function MaulanjutScan(LvChecked As ucListView) As Boolean
If AdakahYangBelumDiFix(LvChecked) = True Then
   If MsgBox(j_bahasa(19) & Chr$(13) & j_bahasa(20), vbExclamation + vbYesNo) = vbYes Then
      MaulanjutScan = True
   Else
      MaulanjutScan = False
      Call scancomputer
   End If
Else
   MaulanjutScan = True
End If

End Function
Public Function LookPropertyLink(TheFullPath As String, Wanted As String) As String

    Dim LinkShell As New WshShell
    'Dim LinkShortCut As New WshShortcut
    Set LinkShortCut = LinkShell.CreateShortCut(TheFullPath)


    Select Case UCase$(Wanted)
        Case "TARGET"
        LookPropertyLink = LinkShortCut.TargetPath
        Case "NAME"
        LookPropertyLink = LinkShortCut.FullName
        Case "ICON"
        LookPropertyLink = LinkShortCut.IconLocation
        Case "START"
        LookPropertyLink = LinkShortCut.WorkingDirectory
        Case "KEY"
        LookPropertyLink = LinkShortCut.Hotkey 'if any
        Case Else
    End Select

Set LinkShell = Nothing
Set LinkShortCut = Nothing
End Function

Public Function GetTargetLink(ByRef TheFullPath As String, ByVal WithArgumen As Boolean) As String

    Dim LinkShell As New WshShell
    'Dim LinkShortCut As New WshShortcut
    Set LinkShortCut = LinkShell.CreateShortCut(TheFullPath)

   GetTargetLink = LinkShortCut.TargetPath
   ' klo ada argumen VBS ambil argumenya aja
   If WithArgumen = True Then
     If UCase$(Left(LinkShortCut.Arguments, 12)) = "//E:VBSCRIPT" Then
        GetTargetLink = ArgumenToPath(LinkShortCut.Arguments, TheFullPath)
     End If
End If

Set LinkShell = Nothing
Set LinkShortCut = Nothing
End Function

Private Function ArgumenToPath(ByRef sArgumenFull As String, ByRef ScPath As String) As String
If InStr(sArgumenFull, Chr$(34)) > 0 Then
   ArgumenToPath = Mid$(sArgumenFull, InStr(sArgumenFull, " ") + 1)
   ArgumenToPath = Left(ArgumenToPath, InStr(ArgumenToPath, " ") - 1)
Else
   ArgumenToPath = Mid$(sArgumenFull, InStr(sArgumenFull, " ") + 1)
End If
   If Mid$(ArgumenToPath, 2, 1) <> ":" Then ArgumenToPath = GetFilePath(ScPath) & "\" & ArgumenToPath

End Function

Public Function CreateShortCut(ByVal vsFileNameAndPath As String, ByVal vsShortCutPath As String, Optional ByVal vsWorkingDir As String = "", Optional ByVal vsArguments As String = "") As Boolean
   Dim bRetVal As Boolean
   Dim oFso As FileSystemObject
   
   Dim oShortCut As IWshRuntimeLibrary.IWshShortcut
   Dim oShell As New WshShell
   Set oFso = New FileSystemObject
   
   If oFso.FolderExists(vsShortCutPath) And oFso.FileExists(vsFileNameAndPath) Then
       Set oShortCut = oShell.CreateShortCut(vsShortCutPath & "\Wan'iez Antivirus.lnk") '& oFSO.GetFileName(vsFileNameAndPath) & ".lnk")
              

       With oShortCut
           .TargetPath = vsFileNameAndPath
           .WindowStyle = 1
           .Description = "Wan'iez Antivirus"
           If Len(vsWorkingDir) > 0 Then
               If oFso.FolderExists(vsWorkingDir) Then
                   .WorkingDirectory = vsWorkingDir
               End If
           End If
           .IconLocation = vsFileNameAndPath & ", 0"
           .Arguments = vsArguments
           .Save
       End With
      bRetVal = True
   Else
      bRetVal = False
   End If

   CreateShortCut = bRetVal
End Function

Public Sub LayOnDekstop()
    CreateShortCut App_FullPathW, GetSpecFolder(DEKSTOP_PATH)
End Sub
Public Function IsConnectedToInternet(Optional connectMode As Integer) As Boolean
    Dim Flags As Long
    ' this ASPI function does it all
    IsConnectedToInternet = InternetGetConnectedState(Flags, 0)
    ' return the flag through the optional argument
    connectMode = Flags
End Function
Public Sub DragForm(frm As Form)
    Call ReleaseCapture
    Call SendMessage(frm.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
End Sub
