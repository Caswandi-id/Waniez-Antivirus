Attribute VB_Name = "basRTPapi"
Option Explicit

Private Declare Function SendMessage Lib _
"user32" Alias "SendMessageA" _
(ByVal hwnd As Long, ByVal wMsg As Long, _
ByVal wParam As Long, lParam As Any) As Long
Public Declare Function FindWindow Lib _
"user32" Alias "FindWindowA" _
(ByVal lpClassName As String, _
ByVal lpWindowName As String) As Long
Private Declare Function FindWindowEx Lib _
"user32" Alias "FindWindowExA" _
(ByVal hWnd1 As Long, ByVal hWnd2 As Long, _
ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
'tes uni

Private Declare Function GetForegroundWindow Lib "user32" () As Long
Const WM_GETTEXT = &HD
Private Declare Function PathIsDirectory Lib "shlwapi.dll" Alias "PathIsDirectoryW" (ByVal pszPath As String) As Long
Rem; buat special folder, seperti my document, dkk.. ;)
Private Declare Function SHGetSpecialFolderPath Lib "shell32.dll" Alias "SHGetSpecialFolderPathA" (ByVal hwnd As Long, ByVal pszPath As String, ByVal CSIDL As Long, ByVal fCreate As Boolean) As Boolean

Rem; const Special Folder nih..
Rem; Cari sendiri ya const utk My Music, dkk.. [Seven]
Private Enum IDFolder
    ALL_USER_STARTUP = &H18
    WINDOWS_DIR = &H24
    SYSTEM_DIR = &H25
    PROGRAM_FILE = &H26
    USER_DOC = &H5
    USER_STARTUP = &H7
    RECENT_DOC = &H8
    DEKSTOP_PATH = &H19
End Enum

Dim hand1, hand2, hand3, hand4, hand5, hand6 As Long
Dim Temp As String * 256
Private Function CariWindow(strClass As String, strClassEx As String) As Long
    Dim Urut()        As String, tampung As Variant
    Dim Hand(10)      As Long, i As Integer
    i = 0
    Hand(i) = FindWindow(strClass, vbNullString)
    If InStr(1, strClassEx, "\", vbTextCompare) <> 0 Then
        Urut = Split(strClassEx, "\")
        For Each tampung In Urut()
            i = i + 1
            Hand(i) = FindWindowEx(Hand(i - 1), 0&, tampung, vbNullString)
        Next
    Else
        i = i + 1
        Hand(i) = FindWindowEx(Hand(i - 1), 0&, strClassEx, vbNullString)
    End If
    CariWindow = Hand(i)
End Function

Private Function GetSpecFolder(ByVal lpCSIDL As IDFolder) As String
Rem; buat ambil special folder nih...
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
Public Function GetAlamat() As String
    Dim Temp            As String * 256
    Dim AlamatFile      As String
    Dim OS              As String '-> seleksi OS yang digunakan
    
    Const SevenBar = "WorkerW\ReBarWindow32\" & _
    "Address Band Root\msctls_progress32\ComboBoxEx32\ComboBox\Edit"
        
    Rem; DEFAULT SEVEN GUNAKAN STRUKTUR WINDOW CLASS INI
    Const Seven = "WorkerW\RebarWindow32\Address Band Root\msctls_progress32\" & _
    "Breadcrumb Parent\ToolbarWindow32"
    
    Rem; XP mah standard aja... ;)
    Const XPBar = "WorkerW\ReBarWindow32\" & _
    "ComboBoxEx32\ComboBox\Edit"

    'dapatkan string pada address bar u vistaw
    If CariWindow("CabinetWClass", XPBar) <> 0 Then
        OS = "XP" 'buat seleksi nanti di special folder
        SendMessage CariWindow("CabinetWClass", XPBar), WM_GETTEXT, 200, ByVal Temp
    ElseIf CariWindow("CabinetWClass", Seven) <> 0 Then
        OS = "Seven"
        SendMessage CariWindow("CabinetWClass", Seven), WM_GETTEXT, 200, ByVal Temp
    End If
    
    AlamatFile = Mid$(Temp, 1, InStr(Temp, Chr$(0)) - 1)
    
    Rem; karena nilai balik dari ..\Breadcrumb Parent\ToolbarWindow32 adalah
    Rem; "Address: C:\Gary Keren", maka hilangkan kata "Address: "
    AlamatFile = Replace(AlamatFile, "Address: ", "")
    
    Rem; kondisi special folder
    If OS = "Seven" Then
        Select Case AlamatFile
        Case "Desktop":
            AlamatFile = GetSpecFolder(DEKSTOP_PATH)
        Case "Recent Places":
            AlamatFile = GetSpecFolder(RECENT_DOC)
        Case "Libraries\Documents":
            AlamatFile = GetSpecFolder(USER_DOC)
        End Select
    ElseIf OS = "XP" Then
        Select Case AlamatFile
        Case "My Documents":
            AlamatFile = GetSpecFolder(USER_DOC)
        Case "Desktop":
            AlamatFile = GetSpecFolder(DEKSTOP_PATH)
        End Select
    End If
    
    GetAlamat = AlamatFile
    
    Rem; tidak diuji di Vista! saya rasa hampir sama dengan Seven.
End Function

Private Function GetWindowsExplorerWindowFolder(hwnd As Long) As String
    On Error GoTo endd
    Dim myShell As Shell
    Dim myExplorerWindow As WebBrowser
    Set myShell = New Shell
    
    For Each myExplorerWindow In myShell.Windows
    On Error Resume Next
        'DoEvents
        If myExplorerWindow.hwnd = hwnd Then
            If err = 0 Then GetWindowsExplorerWindowFolder = myExplorerWindow.Document.Folder.Self.path
        End If
    On Error GoTo 0
    Next
endd:
End Function

Public Function lihat() As String
    Dim hand1 As Long
    Dim hand2 As Long
    Dim hand3 As Long
    Dim hand4 As Long
    Dim hand5 As Long
    Dim hand6 As Long
    Dim hand7 As Long
    Dim dam   As String
    
'Dapatkan Handle pertama / Parent Window dari Class Name
    hand1 = FindWindow("ExploreWClass", vbNullString)
    hand2 = FindWindow("CabinetWClass", vbNullString)
    If hand1 = GetForegroundWindow Then
         dam = GetWindowsExplorerWindowFolder(hand1)
    ElseIf hand2 = GetForegroundWindow Then
         dam = GetWindowsExplorerWindowFolder(hand2)
    End If

    If dam <> "" Then lihat = dam
End Function






Public Function TemuWindow(ClassApa As String, Caption As String) As Long
TemuWindow = FindWindow(ClassApa, Caption)
End Function

