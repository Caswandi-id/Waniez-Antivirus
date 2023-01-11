Attribute VB_Name = "ModRegMain"
Option Explicit

Private Const REG_SZ = 1&
Private Const REG_DWORD = 4
Private Const ERROR_SUCCESS = 0&
Private Const KEY_SET_VALUE = &H2&
Private Const KEY_NOTIFY = &H10&
Private Const READ_CONTROL = &H20000
Private Const STANDARD_RIGHTS_READ = READ_CONTROL

Private Enum REG
    HKEY_CLASSES_ROOT = &H80000000
    HKEY_CURRENT_USER = &H80000001
    HKEY_LOCAL_MACHINE = &H80000002
    HKEY_USERS = &H80000003
    HKEY_CURRENT_CONFIG = &H80000005
    HKEY_PERFORMANCE_DATA = &H80000004
    HKEY_DYN_DATA = &H80000006
End Enum

Private Const KEY_ALL_ACCESS = &HF003F 'Permission for all types of access.
Private Const KEY_ENUMERATE_SUB_KEYS = &H8 'Permission to enumerate subkeys.
Private Const KEY_WRITE = &H20006 'Permission for general write access.
Private Const KEY_QUERY_VALUE = &H1 'Permission to query subkey data.
Private Const KEY_READ = STANDARD_RIGHTS_READ Or _
                        KEY_QUERY_VALUE Or _
                        KEY_ENUMERATE_SUB_KEYS Or _
                        KEY_NOTIFY 'Permission for general read access.

'-- import/export registry key
Private Const TOKEN_QUERY As Long = &H8&
Private Const TOKEN_ADJUST_PRIVILEGES As Long = &H20&
Private Const SE_PRIVILEGE_ENABLED As Long = &H2
Private Const SE_BACKUP_NAME = "SeBackupPrivilege"
Private Const SE_RESTORE_NAME = "SeRestorePrivilege" 'Important for what we're trying to accomplish
Private Const REG_FORCE_RESTORE As Long = 8& 'Permission to overwrite a registry key

'-- digunakan untuk menulis key registry
Private Type SECURITY_ATTRIBUTES
  nLength As Long
  lpSecurityDescriptor As Long
  bInheritHandle As Boolean
End Type

'-- enumerating registrykeys
Private Type FILETIME
  dwLowDateTime As Long
  dwHighDateTime As Long
End Type

'-- import/export registry key
Private Type LUID
  lowpart As Long
  highpart As Long
End Type

Private Type LUID_AND_ATTRIBUTES
  pLuid As LUID
  Attributes As Long
End Type

Private Type TOKEN_PRIVILEGES
  PrivilegeCount As Long
  Privileges As LUID_AND_ATTRIBUTES
End Type

Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegSetValue Lib "advapi32.dll" Alias "RegSetValueA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal dwType As Long, ByVal lpData As String, ByVal cbData As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long         ' Note that if you declare the lpData parameter as String, you must pass it By Value.
Private Declare Function RegSetValueExA Lib "advapi32.dll" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByRef lpData As Long, ByVal cbData As Long) As Long
Private Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, lpSecurityAttributes As SECURITY_ATTRIBUTES, phkResult As Long, lpdwDisposition As Long) As Long
Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long         ' Note that if you declare the lpData parameter as String, you must pass it By Value.
Private Declare Function RegQueryValueExA Lib "advapi32.dll" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByRef lpData As Long, lpcbData As Long) As Long
Private Declare Function RegQueryInfoKey Lib "advapi32.dll" Alias "RegQueryInfoKeyA" (ByVal hKey As Long, ByVal lpClass As String, lpcbClass As Long, ByVal lpReserved As Long, lpcSubKeys As Long, lpcbMaxSubKeyLen As Long, lpcbMaxClassLen As Long, lpcValues As Long, lpcbMaxValueNameLen As Long, lpcbMaxValueLen As Long, lpcbSecurityDescriptor As Long, lpftLastWriteTime As Long) As Long
Private Declare Function RegSaveKey Lib "advapi32.dll" Alias "RegSaveKeyA" (ByVal hKey As Long, ByVal lpFile As String, lpSecurityAttributes As Any) As Long
Private Declare Function RegRestoreKey Lib "advapi32.dll" Alias "RegRestoreKeyA" (ByVal hKey As Long, ByVal lpFile As String, ByVal dwFlags As Long) As Long
Private Declare Function RegEnumKey Lib "advapi32.dll" Alias "RegEnumKeyA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, ByVal cbName As Long) As Long
Private Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, lpcbValueName As Long, ByVal lpReserved As Long, lpType As Long, lpData As Byte, lpcbData As Long) As Long
Private Declare Function XRegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, lpcbValueName As Long, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long

'Export/Import Registry Keys
Private Declare Function RegFlushKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Private Declare Function AdjustTokenPrivileges Lib "advapi32.dll" (ByVal TokenHandle As Long, ByVal DisableAllPriv As Long, newState As TOKEN_PRIVILEGES, ByVal BufferLength As Long, PreviousState As TOKEN_PRIVILEGES, ReturnLength As Long) As Long
Private Declare Function LookupPrivilegeValue Lib "advapi32.dll" Alias "LookupPrivilegeValueA" (ByVal lpSystemName As Any, ByVal lpName As String, lpLuid As LUID) As Long

'-- konstanta
Private Const HKLM = &H80000002
Private Const HKCU = &H80000001
Private Const R = "Software\Microsoft\Windows\CurrentVersion\Run"
Private Const RO = "Software\Microsoft\Windows\CurrentVersion\RunOnce"
Private Const ROX = "Software\Microsoft\Windows\CurrentVersion\RunOnceEx"
Private Const RS = "Software\Microsoft\Windows\CurrentVersion\RunServices"
Private Const RSO = "Software\Microsoft\Windows\CurrentVersion\RunServicesOnce"
Private Const PR = "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer\Run"

Dim valname As String
Dim vallen As Long
Dim datatype As Long
Dim Data(0 To 128) As Byte
Dim datalen As Long
Dim handle As Long
Dim Index As Long
Dim rval As Long
Dim strbuff As String
Dim MainKeyHandle As REG
Dim rtn As Long, lBuffer As Long, sBuffer As String
Dim lBufferSize As Long
Dim lDataSize As Long
Dim lKey As Long

Private lIndex As Long

Private Function CreateRegistryKey(hKey As REG, sSubKey As String)
    On Error Resume Next
    Dim lReg As Long
    RegCreateKey hKey, sSubKey, lReg
    RegCloseKey lReg
End Function

Private Function GetDWORDValue(hKey As REG, SubKey As String, Entry As String)
    Dim ret As Long
    rtn = RegOpenKeyEx(hKey, SubKey, 0, KEY_READ, ret)
    If rtn = ERROR_SUCCESS Then
        rtn = RegQueryValueExA(ret, Entry, 0, REG_DWORD, lBuffer, 4)
        If rtn = ERROR_SUCCESS Then
            rtn = RegCloseKey(ret)
            GetDWORDValue = lBuffer
        Else
            GetDWORDValue = "Error"
        End If
    Else
        GetDWORDValue = "Error"
    End If
End Function

Private Function GetSTRINGValue(hKey As REG, SubKey As String, Entry As String)
    Dim ret As Long
    rtn = RegOpenKeyEx(hKey, SubKey, 0, KEY_READ, ret)
    If rtn = ERROR_SUCCESS Then
        sBuffer = Space$(255)
        lBufferSize = Len(sBuffer)
        rtn = RegQueryValueEx(ret, Entry, 0, REG_SZ, sBuffer, lBufferSize)
        If rtn = ERROR_SUCCESS Then
            rtn = RegCloseKey(ret)
            sBuffer = Trim$(sBuffer)
            GetSTRINGValue = Left$(sBuffer, Len(sBuffer) - 1)
        Else
            GetSTRINGValue = "Error"
        End If
    Else
        GetSTRINGValue = "Error"
    End If
End Function

 Private Function CreateDwordValue(hKey As REG, SubKey As String, strValueName As String, dwordData As Long) As Long
    On Error Resume Next
    Dim ret As Long
    RegCreateKey hKey, SubKey, ret
    CreateDwordValue = RegSetValueEx(ret, strValueName, 0, REG_DWORD, dwordData, 4)
    RegCloseKey ret
End Function

Private Function CreateStringValue(hKey As REG, SubKey As String, strValueName As String, strData As String) As Long
    On Error Resume Next
    Dim ret As Long
    RegCreateKey hKey, SubKey, ret
    CreateStringValue = RegSetValueEx(ret, strValueName, 0, REG_SZ, ByVal strData, Len(strData))
    RegCloseKey ret
End Function

Private Function DeleteValue(hKey As REG, SubKey As String, lpValName As String) As Long
    On Error Resume Next
    Dim ret As Long
    RegOpenKey hKey, SubKey, ret
    DeleteValue = RegDeleteValue(ret, lpValName)
    RegCloseKey ret
End Function

Private Function DeleteKey(hKey As REG, SubKey As String, lpValName As String) As Long
    On Error Resume Next
    Dim ret As Long
    RegOpenKey hKey, SubKey, ret
    DeleteKey = RegDeleteKey(ret, lpValName)
    RegCloseKey ret
End Function

Private Function ReadValue(hKey As REG, SubKey As String, strValueName As String) As String
    Dim rootKey As Long
    Dim isi As String
    Dim ldatabufsize As Long
    Dim lValueType As Long
    Dim x
    Dim ret As Long
    On Error Resume Next
    x = RegOpenKey(hKey, SubKey, rootKey)
    ret = RegQueryValueEx(rootKey, strValueName, 0, lValueType, 0, ldatabufsize)
    isi = String$(ldatabufsize, Chr$(0))
    ret = RegQueryValueEx(rootKey, strValueName, 0, 0, ByVal isi, ldatabufsize)
    ReadValue = Left$(isi, InStr(1, isi, Chr$(0)) - 1)
    RegCloseKey rootKey
End Function


' rutin hapus data registry key
Private Function DeleteRegKeyValue(KeyRoot As REG, KeyPath As String, Optional SubKey As String = "") As Boolean
    On Error Resume Next
    Dim hKey As Long
    Dim ReturnValue As Long
    ReturnValue = RegOpenKeyEx(KeyRoot, KeyPath, 0, KEY_ALL_ACCESS, hKey)
    If ReturnValue <> 0 Then
        DeleteRegKeyValue = False
        ReturnValue = RegCloseKey(hKey)
        Exit Function
    End If
    If SubKey = "" Then SubKey = KeyPath
    ReturnValue = RegDeleteValue(hKey, SubKey)
        If ReturnValue = 0 Then
            DeleteRegKeyValue = True
        Else
        DeleteRegKeyValue = False
    End If
    ReturnValue = RegCloseKey(hKey)
End Function



Private Function FixRegistry()
    On Error Resume Next
    CreateStringValue HKEY_CLASSES_ROOT, "exefile", "", "Application"
    CreateStringValue HKEY_CLASSES_ROOT, "exefile\shell\open\command", "", Chr$(&H22) & "%1" & Chr$(&H22) & " %*"
    CreateStringValue HKEY_CLASSES_ROOT, "lnkfile\shell\open\command", "", Chr$(&H22) & "%1" & Chr$(&H22) & " %*"
    CreateStringValue HKEY_CLASSES_ROOT, "piffile\shell\open\command", "", Chr$(&H22) & "%1" & Chr$(&H22) & " %*"
    CreateStringValue HKEY_CLASSES_ROOT, "batfile\shell\open\command", "", Chr$(&H22) & "%1" & Chr$(&H22) & " %*"
    CreateStringValue HKEY_CLASSES_ROOT, "comfile\shell\open\command", "", Chr$(&H22) & "%1" & Chr$(&H22) & " %*"
    CreateStringValue HKEY_CLASSES_ROOT, "cmdfile\shell\open\command", "", Chr$(&H22) & "%1" & Chr$(&H22) & " %*"
    CreateStringValue HKEY_CLASSES_ROOT, "scrfile\shell\open\command", "", Chr$(&H22) & "%1" & Chr$(&H22) & " /S"
    CreateStringValue HKEY_CLASSES_ROOT, "regfile\shell\open\command", "", "regedit.exe %1"
    CreateStringValue HKEY_CLASSES_ROOT, "vbsfile\shell\open\command", "", "%SystemRoot%\System32\WScript.exe ""%1"" %*"
    CreateStringValue HKEY_CURRENT_USER, "Control Panel\Desktop\", "SCRNSAVE.EXE", ""
    CreateStringValue HKEY_LOCAL_MACHINE, "SYSTEM\ControlSet001\Control\SafeBoot\", "AlternateShell", "cmd.exe"
    CreateStringValue HKEY_LOCAL_MACHINE, "SYSTEM\ControlSet002\Control\SafeBoot\", "AlternateShell", "cmd.exe"
    CreateStringValue HKEY_LOCAL_MACHINE, "SYSTEM\ControlSet003\Control\SafeBoot\", "AlternateShell", "cmd.exe"
    CreateStringValue HKEY_LOCAL_MACHINE, "SYSTEM\CurrentControlSet\Control\SafeBoot\", "AlternateShell", "cmd.exe"
    CreateStringValue HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\AeDebug", "Auto", "0"
    CreateStringValue HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon", "Shell", "Explorer.exe"
    CreateStringValue HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon", "Userinit", "C:\WINDOWS\system32\userinit.exe,"
    CreateStringValue HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon", "LegalNoticeText", ""
    CreateStringValue HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon", "LegalNoticeCaption", ""
    CreateStringValue HKEY_CURRENT_USER, "Control Panel\International\", "s1159", "AM"
    CreateStringValue HKEY_CURRENT_USER, "Control Panel\International\", "s2359", "PM"
    CreateDwordValue HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Advanced\Folder\SuperHidden\", "CheckedValue", "0"
    CreateDwordValue HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Advanced\Folder\SuperHidden\", "UncheckedValue", "1"
    CreateDwordValue HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Advanced\Folder\HideFileExt\", "CheckedValue", "1"
    CreateDwordValue HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Advanced\Folder\HideFileExt\", "UncheckedValue", "0"
    
    'NO - NO
    DeleteValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoWinKeys"
    DeleteValue HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoWinKeys"
    DeleteValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer\", "NoActiveDesktop"
    DeleteValue HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer\", "NoActiveDesktop"
    DeleteValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoClose"
    DeleteValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoDesktop"
    DeleteValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoRun"
    DeleteValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoLogoff"
    DeleteValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoFind"
    DeleteValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoTrayContextMenu"
    DeleteValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoViewContextMenu"
    DeleteValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer\", "NoStarMenuMorePrograms"
    DeleteValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\System", "NoDispSettingsPage"
    DeleteValue HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Policies\System", "NoDispSettingsPage"
    DeleteValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\System", "NoDispBackgroundPage"
    DeleteValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\System", "NoScrSavPage"
    DeleteValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\System", "NoDispApprearancePage"
    DeleteValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\System", "NoDispCpl"
    DeleteValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer\", "NoSMHelp"
    DeleteValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer\", "NoDrives"
    DeleteValue HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer\", "NoDrives"
    DeleteValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer\", "NoDrive"
    DeleteValue HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer\", "NoDrive"
    DeleteValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer\", "NoViewOnDrive\"
    DeleteValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer\", "NoSaveSettings"
    DeleteValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer\", "NoSetFolders"
    DeleteValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer\", "NoSetTaskbar"
    DeleteValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer\", "NoTrayContextMenu"
    DeleteValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer\", "NoViewContextMenu"
    DeleteValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer\", "NoUserNameInStartMenu"
    DeleteValue HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer\", "NoPrinters"
    DeleteValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer\", "NoDesktop"
    DeleteValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer\", "NoThemesTab"
    DeleteValue HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer\", "NoThemesTab"
    DeleteValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer\", "NoInstrumentation"
    DeleteValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer\", "NoPrinters"
    DeleteValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer\", "NoSMHelp"
    DeleteValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer\", "NoStartMenuMorePrograms"
    DeleteValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer\", "NoThemesTabNoThemesTab"
    
    'DISABLE
    DeleteValue HKEY_LOCAL_MACHINE, "SOFTWARE\Policies\Microsoft\Windows\Installer", "DisableMSI"
    DeleteValue HKEY_LOCAL_MACHINE, "SOFTWARE\Policies\Microsoft\Windows NT\SystemRestore\", "DisableSR"
    DeleteValue HKEY_LOCAL_MACHINE, "SOFTWARE\Policies\Microsoft\Windows NT\SystemRestore\", "DisableConfig"
    DeleteValue HKEY_CURRENT_USER, "Software\Policies\Microsoft\Windows\Task Scheduler5.0\", "Execution"
    DeleteValue HKEY_LOCAL_MACHINE, "Software\Policies\Microsoft\Windows\Task Scheduler5.0\", "Execution"
    DeleteValue HKEY_CURRENT_USER, "Software\Policies\Microsoft\Windows\Task Scheduler5.0\", "Property Pages"
    DeleteValue HKEY_LOCAL_MACHINE, "Software\Policies\Microsoft\Windows\Task Scheduler5.0\", "Property Pages"
    DeleteValue HKEY_CURRENT_USER, "Software\Policies\Microsoft\Windows\Task Scheduler5.0\", "Task Creation"
    DeleteValue HKEY_CURRENT_USER, "Software\Policies\Microsoft\Windows\Task Scheduler5.0\", "Task Deletion"
    DeleteValue HKEY_LOCAL_MACHINE, "SOFTWARE\Policies\Microsoft\Windows\Installer", "LimitSystemRestoreCheckpointing"
    DeleteValue HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer\", "shutdownwithoutlogon"
    DeleteValue HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer\", "undockwithoutlogon"
    DeleteKey HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies", "WinOldApp"
    DeleteKey HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Policies", "WinOldApp"
    DeleteValue HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\policies\System", "legalnoticetext"
    DeleteValue HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\policies\System\", "legalnoticecaption"
    DeleteValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer\", "StartMenuLogOff"
    DeleteValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\", "ActiveDesktop"
    DeleteValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer\", "DisallowRun"
    DeleteValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer\", "ClassicShell"
    DeleteValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer\", "HideClock"
    
    'SystemOptimizer
    CreateStringValue HKEY_CURRENT_USER, "Control Panel\Desktop", "AutoEndTasks", "1"
    CreateStringValue HKEY_CURRENT_USER, "Control Panel\Desktop", "MenuShowDelay", "1"
    CreateStringValue HKEY_USERS, ".DEFAULT\Control Panel\Desktop\", "MenuShowDelay", "1"
    CreateStringValue HKEY_LOCAL_MACHINE, "SYSTEM\ControlCurrentSet\Control\CrashControl", "AutoReboot", "1"
    CreateDwordValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer", "DesktopProcess", "1"
    CreateStringValue HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\", "AlwaysUnloadDLL", "1"
    DeleteValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer\System", "DisableRegistryTools"
    DeleteValue HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer\System", "DisableRegistryTools"
    DeleteValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\System\", "DisableTaskMgr"
    DeleteValue HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System\", "DisableTaskMgr"
    DeleteKey HKEY_CURRENT_USER, "Software\Policies\Microsoft\Windows", "system"
    DeleteValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoFolderOptions"
    DeleteValue HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoFolderOptions"
    'DoEvents
End Function

Public Sub SystemEditor()
Dim x
Dim lDword
'================================================================================================================================'
'Start Menu
'================================================================================================================================'
'NoFind
x = ReadValue(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoFind")
If x <> "" Then frmMain.Check7.Value = 0 Else: frmMain.Check7.Value = 1
x = ReadValue(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\policies\Explorer", "NoFind")
If x <> "" Then frmMain.Check7.Value = 0 Else: frmMain.Check7.Value = 1
        
'LogOff
x = ReadValue(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoLogOff")
If x <> "" Then frmMain.Check8.Value = 0 Else: frmMain.Check8.Value = 1
x = ReadValue(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoLogOff")
If x <> "" Then frmMain.Check8.Value = 0 Else: frmMain.Check8.Value = 1
        
'Recent Documents
x = ReadValue(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoRecentDocsMenu")
If x <> "" Then frmMain.Check9.Value = 0 Else: frmMain.Check9.Value = 1
x = ReadValue(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoRecentDocsMenu")
If x <> "" Then frmMain.Check9.Value = 0 Else: frmMain.Check9.Value = 1
        
'Show Run
x = ReadValue(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer\", "NoRun")
If x <> "" Then frmMain.Check10.Value = 0 Else: frmMain.Check10.Value = 1
x = ReadValue(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\policies\Explorer\", "NoRun")
If x <> "" Then frmMain.Check10.Value = 0 Else: frmMain.Check10.Value = 1
        
'Show Shutdown
x = ReadValue(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoClose")
If x <> "" Then frmMain.Check11.Value = 0 Else: frmMain.Check11.Value = 1
x = ReadValue(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\policies\Explorer", "NoClose")
If x <> "" Then frmMain.Check11.Value = 0 Else: frmMain.Check11.Value = 1
        
'Show Help
x = ReadValue(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoSMHelp")
If x <> "" Then frmMain.Check12.Value = 0 Else: frmMain.Check12.Value = 1
x = ReadValue(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\policies\Explorer", "NoSMHelp")
If x <> "" Then frmMain.Check12.Value = 0 Else: frmMain.Check12.Value = 1

'================================================================================================================================='
'System
'================================================================================================================================='
'Enable Display Setting
x = ReadValue(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\System", "NoDispCPL")
If x <> "" Then frmMain.Check13.Value = 0 Else: frmMain.Check13.Value = 1
x = ReadValue(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\policies\system", "NoDispCPL")
If x <> "" Then frmMain.Check13.Value = 0 Else: frmMain.Check13.Value = 1
        
'Enable Registry Editor
x = ReadValue(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\System", "DisableRegistryTools")
If x <> "" Then frmMain.Check14.Value = 0 Else: frmMain.Check14.Value = 1
x = ReadValue(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\policies\system", "DisableRegistryTools")
If x <> "" Then frmMain.Check14.Value = 0 Else: frmMain.Check14.Value = 1
        
'Enable Task Manager
x = ReadValue(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\System", "DisableTaskMgr")
If x <> "" Then frmMain.Check15.Value = 0 Else: frmMain.Check15.Value = 1
x = ReadValue(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\policies\system", "DisableTaskMgr")
If x <> "" Then frmMain.Check15.Value = 0 Else: frmMain.Check15.Value = 1
        
'Enable Command Prompt(CMD)
x = ReadValue(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\System", "DisableCMD")
If x <> "" Then frmMain.Check16.Value = 0 Else: frmMain.Check16.Value = 1
x = ReadValue(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\policies\system", "DisableCMD")
If x <> "" Then frmMain.Check16.Value = 0 Else: frmMain.Check16.Value = 1
        
'Enable Windows Hotkeys
x = ReadValue(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoWinKeys")
If x <> "" Then frmMain.Check17.Value = 0 Else: frmMain.Check17.Value = 1
x = ReadValue(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\policies\Explorer", "NoWinKeys")
If x <> "" Then frmMain.Check17.Value = 0 Else: frmMain.Check17.Value = 1
        
'================================================================================================================================='
'Windows Explorer
'================================================================================================================================='
'Enable Folder Options
lDword = GetDWORDValue(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer\", "NoFolderOptions")
If Val(lDword) = 1 Then frmMain.Check18.Value = 0 Else: frmMain.Check18.Value = 1
lDword = GetDWORDValue(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\policies\Explorer\", "NoFolderOptions")
If Val(lDword) = 1 Then frmMain.Check18.Value = 0 Else: frmMain.Check18.Value = 1
lDword = GetDWORDValue(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\policies\", "NoFolderOptions")
If Val(lDword) = 1 Then frmMain.Check18.Value = 0 Else: frmMain.Check18.Value = 1
        
'Enable Control Panel
lDword = GetDWORDValue(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoControlPanel")
If Val(lDword) = 1 Then frmMain.Check19.Value = 0 Else: frmMain.Check19.Value = 1
lDword = GetDWORDValue(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\policies\Explorer", "NoControlPanel")
If Val(lDword) = 1 Then frmMain.Check19.Value = 0 Else: frmMain.Check19.Value = 1
        
'Enable Explorer's Context Menu
lDword = GetDWORDValue(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoViewContextMenu")
If Val(lDword) = 1 Then frmMain.Check20.Value = 0 Else: frmMain.Check20.Value = 1
lDword = GetDWORDValue(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\policies\Explorer", "NoViewContextMenu")
If Val(lDword) = 1 Then frmMain.Check20.Value = 0 Else: frmMain.Check20.Value = 1
        
'Enable Taskbar's Context Menu
lDword = GetDWORDValue(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoTrayContextMenu")
If Val(lDword) = 1 Then frmMain.Check21.Value = 0 Else: frmMain.Check21.Value = 1
lDword = GetDWORDValue(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\policies\Explorer", "NoTrayContextMenu")
If Val(lDword) = 1 Then frmMain.Check21.Value = 0 Else: frmMain.Check21.Value = 1
        
'Enable Taskbar Setting
lDword = GetDWORDValue(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoSetTaskbar")
If Val(lDword) = 1 Then frmMain.Check22.Value = 0 Else: frmMain.Check22.Value = 1
lDword = GetDWORDValue(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\policies\Explorer", "NoSetTaskbar")
If Val(lDword) = 1 Then frmMain.Check22.Value = 0 Else: frmMain.Check22.Value = 1
        
'Enable Desktop Item
lDword = GetDWORDValue(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoDesktop")
If Val(lDword) = 1 Then frmMain.Check23.Value = 0 Else: frmMain.Check23.Value = 1
lDword = GetDWORDValue(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\policies\Explorer", "NoDesktop")
If Val(lDword) = 1 Then frmMain.Check23.Value = 0 Else: frmMain.Check23.Value = 1
        
'SuperHidden
lDword = GetDWORDValue(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced", "ShowSuperHidden")
If lDword = 1 Then frmMain.Check24.Value = 1 Else: frmMain.Check24.Value = 0
        
'Hidden
lDword = GetDWORDValue(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced\", "Hidden")
If lDword = 1 Then frmMain.Check25.Value = 1 Else: frmMain.Check25.Value = 0
        
'HideFileExt
lDword = GetDWORDValue(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced", "HideFileExt")
If lDword = 1 Then frmMain.Check26.Value = 0 Else: frmMain.Check26.Value = 1
        
'Full Path in Addres Bar
lDword = GetDWORDValue(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\CabinetState", "FullPathAddress")
If lDword = 1 Then frmMain.Check27.Value = 1 Else: frmMain.Check27.Value = 0
        
'Full Path in Title Bar
lDword = GetDWORDValue(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\CabinetState", "FullPath")
If lDword = 1 Then frmMain.Check28.Value = 1 Else: frmMain.Check28.Value = 0

'StopAutoplay
lDword = GetDWORDValue(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoDriveTypeAutoRun")
If Val(lDword) = 255 Then frmMain.Check29.Value = 1 Else: frmMain.Check29.Value = 0
lDword = GetDWORDValue(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoDriveTypeAutoRun")
If Val(lDword) = 255 Then frmMain.Check29.Value = 1 Else: frmMain.Check29.Value = 0
End Sub
'terapkan configurasi
Public Sub Check10M()
'Run
If frmMain.Check10.Value = 1 Then DeleteValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoRun": DeleteValue HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoRun" Else: CreateDwordValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoRun", "1": CreateDwordValue HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoRun", "1"
End Sub

Public Sub Check11M()
'Shutdown
If frmMain.Check11.Value = 1 Then DeleteValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoClose": DeleteValue HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoClose" Else: CreateDwordValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoClose", "1": CreateDwordValue HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoClose", "1"
End Sub

Public Sub Check12M()
'Help
If frmMain.Check12.Value = 1 Then DeleteValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoSMHelp": DeleteValue HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoSMHelp" Else: CreateDwordValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoSMHelp", "1": CreateDwordValue HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoSMHelp", "1"
End Sub

Public Sub Check13M()
'Display Setting
If frmMain.Check13.Value = 1 Then DeleteValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\System", "NoDispCPL": DeleteValue HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Policies\System", "NoDispCPL" Else: CreateDwordValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\System", "NoDispCPL", "1": CreateDwordValue HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Policies\System", "NoDispCPL", "1"
End Sub

Public Sub Check14M()
'Registry Editor
If frmMain.Check14.Value = 1 Then DeleteValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\System", "DisableRegistryTools": DeleteValue HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Policies\System", "DisableRegistryTools" Else: CreateDwordValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\System", "DisableRegistryTools", "1": CreateDwordValue HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Policies\System", "DisableRegistryTools", "1"
End Sub

Public Sub Check15M()
'task manager
If frmMain.Check15.Value = 1 Then DeleteValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\System", "DisableTaskMgr": DeleteValue HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Policies\System", "DisableTaskMgr" Else: CreateDwordValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\System", "DisableTaskMgr", "1": CreateDwordValue HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Policies\System", "DisableTaskMgr", "1"
End Sub

Public Sub Check16M()
'CMD
If frmMain.Check16.Value = 1 Then DeleteValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\System", "DisableCMD": DeleteValue HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Policies\System", "DisableCMD" Else: CreateDwordValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\System", "DisableCMD", "1": CreateDwordValue HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Policies\System", "DisableCMD", "1"
End Sub

Public Sub Check17M()
'Windows Hotkeys
If frmMain.Check17.Value = 1 Then DeleteValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoWinKeys": DeleteValue HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoWinKeys" Else: CreateDwordValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoWinKeys", "1": CreateDwordValue HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoWinKeys", "1"
End Sub

Public Sub Check18M()
'FolderOptions
If frmMain.Check18.Value = 1 Then DeleteValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer\", "NoFolderOptions": DeleteValue HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\Explorer\", "NoFolderOptions": DeleteValue HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Policies", "NoFolderOptions" Else: CreateDwordValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer\", "NoFolderOptions", "1": CreateDwordValue HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\Explorer\", "NoFolderOptions", "1": CreateDwordValue HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Policies", "NoFolderOptions", "1"
End Sub

Public Sub Check19M()
'ControlPanel
If frmMain.Check19.Value = 1 Then DeleteValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer\", "NoControlPanel": DeleteValue HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\Explorer\", "NoControlPanel" Else: CreateDwordValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer\", "NoControlPanel", "1": CreateDwordValue HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\Explorer\", "NoControlPanel", "1"
End Sub

Public Sub Check20M()
'ViewContextMenu
If frmMain.Check20.Value = 1 Then DeleteValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer\", "NoViewContextMenu": DeleteValue HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\Explorer\", "NoViewContextMenu" Else: CreateDwordValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer\", "NoViewContextMenu", "1": CreateDwordValue HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\Explorer\", "NoViewContextMenu", "1"
End Sub

Public Sub Check21M()
'TrayContextMenu
If frmMain.Check21.Value = 1 Then DeleteValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer\", "NoTrayContextMenu": DeleteValue HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\Explorer\", "NoTrayContextMenu" Else: CreateDwordValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer\", "NoTrayContextMenu", "1": CreateDwordValue HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\Explorer\", "NoTrayContextMenu", "1"
End Sub

Public Sub Check22M()
'SetTaskbar
If frmMain.Check22.Value = 1 Then DeleteValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer\", "NoSetTaskbar": DeleteValue HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\Explorer\", "NoSetTaskbar" Else: CreateDwordValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer\", "NoSetTaskbar", "1": CreateDwordValue HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\Explorer\", "NoSetTaskbar", "1"
End Sub

Public Sub Check23M()
'Desktop
If frmMain.Check23.Value = 1 Then DeleteValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer\", "NoDesktop": DeleteValue HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\Explorer\", "NoDesktop" Else: CreateDwordValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer\", "NoDesktop", "1": CreateDwordValue HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\Explorer\", "NoDesktop", "1"
End Sub

Public Sub Check24M()
'Show superhidden
If frmMain.Check24.Value = 1 Then CreateDwordValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced", "ShowSuperHidden", "1" Else: CreateDwordValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced", "ShowSuperHidden", "0"
End Sub

Public Sub Check25M()
'Show Hidden file
If frmMain.Check25.Value = 1 Then CreateDwordValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced", "Hidden", "1" Else: CreateDwordValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced", "Hidden", "2"
End Sub

Public Sub Check26M()
'Show Extension file
If frmMain.Check26.Value = 1 Then CreateDwordValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced", "HideFileExt", "0" Else: CreateDwordValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced", "HideFileExt", "1"
End Sub

Public Sub Check27M()
'Show Full Path in Address Bar
If frmMain.Check27.Value = 1 Then CreateDwordValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\CabinetState", "FullPathAddress", "1" Else: CreateDwordValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\CabinetState", "FullPathAddress", "0"
End Sub

Public Sub Check28M()
'Show Full Path in Title Bar
If frmMain.Check28.Value = 1 Then CreateDwordValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\CabinetState", "FullPath", "1" Else: CreateDwordValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\CabinetState", "FullPath", "0"
End Sub

Public Sub Check29M()
'Stop Autoplay Drive
If frmMain.Check29.Value = 1 Then CreateDwordValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer\", "NoDriveTypeAutoRun", "255": CreateDwordValue HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer\", "NoDriveTypeAutoRun", "255" Else: CreateDwordValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer\", "NoDriveTypeAutoRun", "145": CreateDwordValue HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer\", "NoDriveTypeAutoRun", "145"
End Sub

Public Sub Check7M()
'Find
If frmMain.Check7.Value = 1 Then DeleteValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoFind": DeleteValue HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoFind" Else: CreateDwordValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoFind", "1": CreateDwordValue HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoFind", "1"
End Sub

Public Sub Check8M()
'LogOff
If frmMain.Check8.Value = 1 Then DeleteValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoLogOff": DeleteValue HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoLogOff" Else: CreateDwordValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoLogOff", "1": CreateDwordValue HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoLogOff", "1"
End Sub

Public Sub Check9M()
'RecentDocs
If frmMain.Check9.Value = 1 Then DeleteValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoRecentDocsMenu": DeleteValue HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoRecentDocsMenu" Else: CreateDwordValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoRecentDocsMenu", "1": CreateDwordValue HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoRecentDocsMenu", "1"
End Sub 'cek all


