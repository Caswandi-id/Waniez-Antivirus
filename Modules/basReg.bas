Attribute VB_Name = "basReg"
Option Explicit

Private lReg            As Long
Private KeyHandle       As Long
Private lResult         As Long
Private lValueType      As Long
Private ldatabufsize    As Long

Private Const ERROR_SUCCESS = 0&
Private Const REG_SZ = 1
Private Const REG_DWORD = 4
Const KEY_READ = ((&H20000 Or &H1 Or &H8 Or &H10) And (Not &H100000))

' API yang berhubungan dengan Registry
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegCreateKeyW Lib "advapi32.dll" (ByVal hKey As Long, ByVal lpSubKey As Long, ByVal pz_phkResult As Long) As Long '---###FIX!IT###---:phkResult--->ByVal pz_phkResult As Long
Private Declare Function RegDeleteKeyW Lib "advapi32.dll" (ByVal hKey As Long, ByVal lpSubKey As Long) As Long
Private Declare Function RegDeleteValueW Lib "advapi32.dll" (ByVal hKey As Long, ByVal pz_lpValueName As Long) As Long
Private Declare Function RegEnumValueW Lib "advapi32.dll" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As Long, ByVal pz_lpcbValueName As Long, ByVal lpReserved As Long, ByVal pz_lpType As Long, ByVal pz_lpData As Long, ByVal pz_lpcbData As Long) As Long '---###FIX!IT###---:variabel yang ditambah diawali dgn "pz_" diubah menjadi "ByVal * As Long" (sebagai pointer).
Private Declare Function RegOpenKeyW Lib "advapi32.dll" (ByVal hKey As Long, ByVal lpSubKey As Long, ByVal pz_phkResult As Long) As Long '---###FIX!IT###---:variabel yang ditambah diawali dgn "pz_" diubah menjadi "ByVal * As Long" (sebagai pointer).
Private Declare Function RegOpenKeyExW Lib "advapi32.dll" (ByVal hKey As Long, ByVal lpSubKey As Long, ByVal ulOptions As Long, ByVal samDesired As Long, ByVal pz_phkResult As Long) As Long '---###FIX!IT###---:variabel yang ditambah diawali dgn "pz_" diubah menjadi "ByVal * As Long" (sebagai pointer).
Private Declare Function RegQueryValueExW Lib "advapi32.dll" (ByVal hKey As Long, ByVal lpValueName As Long, ByVal lpReserved As Long, ByVal pz_lpType As Long, ByVal pz_lpData As Long, ByVal pz_lpcbData As Long) As Long '---###FIX!IT###---:variabel yang ditambah diawali dgn "pz_" diubah menjadi "ByVal * As Long" (sebagai pointer).
Private Declare Function RegSetValueExW Lib "advapi32.dll" (ByVal hKey As Long, ByVal lpValueName As Long, ByVal Reserved As Long, ByVal dwType As Long, ByVal pz_lpData As Long, ByVal cbData As Long) As Long '---###FIX!IT###---:variabel yang ditambah diawali dgn "pz_" diubah menjadi "ByVal * As Long" (sebagai pointer).
Public Function CreateKeyReg(ByVal hKey As Long, ByRef sPath As String) As Long '---?tanya?:nggak kepakai? '---###FIX!IT###---:optimisasi variabel string,ganti ByRef aja,kalau cuman buat dibaca, bukan ditulis / dimodifikasi ulang.
    
    lReg = RegCreateKeyW(hKey, StrPtr(sPath), VarPtr(KeyHandle)) '---###FIX!IT###---:KeyHandle--->VarPtr(KeyHandle)
    lReg = RegCloseKey(KeyHandle)
    
End Function

Public Function GetSTRINGValue(ByVal hKey As Long, ByRef sPath As String, ByRef sValue As String) As String '---###FIX!IT###---:optimisasi variabel string,ganti ByRef aja,kalau cuman buat dibaca, bukan ditulis / dimodifikasi ulang.
    Dim sBuff As String
    Dim intZeroPos As Integer
    
    lReg = RegOpenKeyW(hKey, StrPtr(sPath), VarPtr(KeyHandle)) '---###FIX!IT###---:KeyHandle--->VarPtr(KeyHandle)
    lResult = RegQueryValueExW(KeyHandle, StrPtr(sValue), 0&, VarPtr(lValueType), 0&, VarPtr(ldatabufsize)) '---###FIX!IT###---:jadikan pointer-ke-value dari variabel.

    If lValueType = REG_SZ Then
        sBuff = String$(ldatabufsize / 2, Chr$(32)) '---###FIX!IT###---:character unicode adalah 2 bytes perchar,fungsi string(param1,param2)menghitung char dengan: 1count=1char (di kode),1char=2bytes (di memori).
        lResult = RegQueryValueExW(KeyHandle, StrPtr(sValue), 0&, VarPtr(lValueType), StrPtr(sBuff), VarPtr(ldatabufsize)) '---###FIX!IT###---:lValueType harus sama.untuk fungsi "RegQueryValueExW" yang diminta adalah ukuran buffer dalam bytes,bukan chars.
        If lResult = ERROR_SUCCESS Then
            '---###FIX!IT###---:kalau ukuran buffer yang disyaratkan dan yang dialokasikan sama, sepertinya fungsi trimnullchars di bawah ini tidak terpakai lagi---:
            intZeroPos = InStr(sBuff, Chr$(0))
            If intZeroPos > 0 Then '---?tanya?TrimNullChars.
                GetSTRINGValue = Replace(sBuff, Chr$(0), "") 'Left$(sBuff, intZeroPos - 1)
            Else
                GetSTRINGValue = sBuff
            End If
            '------------------;
            'GetStringValue = sBuff
        End If
    End If
    'MsgBox "TEST_TRIMMER:[" & GetStringValue & "]"
End Function

Public Function SetStringValue(ByVal hKey As Long, ByRef sPath As String, ByRef sValue As String, ByRef sData As String) As Long '---###FIX!IT###---:optimisasi variabel string,ganti ByRef aja,kalau cuman buat dibaca, bukan ditulis / dimodifikasi ulang.
    
    lReg = RegCreateKeyW(hKey, StrPtr(sPath), VarPtr(KeyHandle)) '---###FIX!IT###---:KeyHandle--->VarPtr(KeyHandle)
    lReg = RegSetValueExW(KeyHandle, StrPtr(sValue), 0, REG_SZ, StrPtr(sData), LenB(sData)) '---###FIX!IT###---:jadikan pointer-ke-value dari variabel.Len(?) jadi LenB(?),untuk fungsi "RegSetValueExW" yang diminta adalah ukuran buffer dalam bytes, bukan chars.
    lReg = RegCloseKey(KeyHandle)
    
End Function


Function GetDWORDValue(ByVal hKey As Long, ByRef sPath As String, ByRef sValueName As String) As Long '---###FIX!IT###---:optimisasi variabel string,ganti ByRef aja,kalau cuman buat dibaca, bukan ditulis / dimodifikasi ulang.
    
    Dim lBuff As Long
    
    lReg = RegOpenKeyW(hKey, StrPtr(sPath), VarPtr(KeyHandle)) '---###FIX!IT###---:KeyHandle--->VarPtr(KeyHandle)
    ldatabufsize = 4 '?info?DWORD.
    lResult = RegQueryValueExW(KeyHandle, StrPtr(sValueName), 0&, VarPtr(lValueType), VarPtr(lBuff), VarPtr(ldatabufsize)) '---###FIX!IT###---:lBuff adalah variant-variable,bukan string.

    If lResult = ERROR_SUCCESS Then
        If lValueType = REG_DWORD Then
            GetDWORDValue = lBuff
        End If
    End If
    
    lReg = RegCloseKey(KeyHandle)
    
End Function

Public Function SetDwordValue(ByVal hKey As Long, ByRef sPath As String, ByRef sValueName As String, ByVal lData As Long) As Long '---###FIX!IT###---:optimisasi variabel string,ganti ByRef aja,kalau cuman buat dibaca, bukan ditulis / dimodifikasi ulang.
    
    lReg = RegCreateKeyW(hKey, StrPtr(sPath), VarPtr(KeyHandle)) '---###FIX!IT###---:KeyHandle--->VarPtr(KeyHandle)
    lResult = RegSetValueExW(KeyHandle, StrPtr(sValueName), 0&, REG_DWORD, VarPtr(lData), 4) '---###FIX!IT###---:jadikan pointer-ke-value dari variabel.
    lReg = RegCloseKey(KeyHandle)
    
End Function

Public Function DeleteKey(ByVal hKey As Long, ByRef sKey As String) As Long '---###FIX!IT###---:optimisasi variabel string,ganti ByRef aja,kalau cuman buat dibaca, bukan ditulis / dimodifikasi ulang.
    lReg = RegDeleteKeyW(hKey, StrPtr(sKey))
End Function

Public Function DeleteValue(ByVal hKey As Long, ByRef sPath As String, ByRef sValue As String) As Long '---###FIX!IT###---:optimisasi variabel string,ganti ByRef aja,kalau cuman buat dibaca, bukan ditulis / dimodifikasi ulang.
    lReg = RegOpenKeyW(hKey, StrPtr(sPath), VarPtr(KeyHandle)) '---###FIX!IT###---:KeyHandle--->VarPtr(KeyHandle)
    lReg = RegDeleteValueW(KeyHandle, StrPtr(sValue))
    lReg = RegCloseKey(KeyHandle)
End Function


Public Function RegEnumStr(ByVal MainKey As Long, ByRef sPath As String, ByRef sValue() As String, ByRef sData() As String) As Long
On Error Resume Next
Dim iHKey As Long, iHasil As Long, Num As Long, vallen As Long '<---biasakan inisialisasi variabel dengan mencantumkan tipe variabel yang jelas.
Dim nOutDataLength  As Long
Dim szOutDataValue As String
Dim StrValue As String
Dim pOpType As Long

iHasil = RegOpenKeyExW(MainKey, StrPtr(sPath), 0, KEY_READ, VarPtr(iHKey))
If iHasil <> 0& Then ' 0& = ERROR_SUCCESS
    Exit Function
End If
Num = 0
ReDim sValue(100) As String
ReDim sData(100) As String

pOpType = REG_SZ
Do
    vallen = 2048  ' Penampung Panjang Value Maximal aja
    StrValue = String$(vallen, 0)
    nOutDataLength = 2048
    szOutDataValue = String$(nOutDataLength, 0)
    iHasil = RegEnumValueW(iHKey, Num, StrPtr(StrValue), VarPtr(vallen), 0&, VarPtr(pOpType), StrPtr(szOutDataValue), VarPtr(nOutDataLength))
    If iHasil = 0& Then
        Num = Num + 1
        StrValue = Left$(StrValue, vallen) '---dalam chars.
        szOutDataValue = Left$(szOutDataValue, (nOutDataLength / 2) - 1) '---dalam bytes,nullchars dibuang.
        sValue(Num) = StrValue
        sData(Num) = szOutDataValue
        '---------------------------;
    End If
Loop While iHasil = 0& ' 0& = ERROR_SUCCESS
    RegEnumStr = Num
    StrValue = ""
    szOutDataValue = ""
Call RegCloseKey(iHKey)
End Function

'---Catatan
' Belum bisa buffer RUNDLL.EXE PathDll,Param -> ah kurang penting untuk mendapatkan startup virus di reg
Public Function BufferStartupPath(SFILE As String) As String
Dim sTmp        As String
Dim sSpecial    As String
Dim nNum        As Long
Dim iCount      As Long
If ValidFile(SFILE) = False Then
    ' dapatkan awal dari drive:\
    If InStr(SFILE, ":\") > 0 Then sTmp = Mid$(SFILE, InStr(SFILE, ":\") - 1)
    sTmp = Replace(sTmp, Chr$(34), "")
    If ValidFile(sTmp) = True Then GoTo KLIMAKS
        
    ' Hilangkan /[param] --- contoh C:\Memeil\Jelek.exe /s
    nNum = InStr(SFILE, "/")
    If nNum > 0 Then
       sTmp = Left$(SFILE, nNum - 1)
    Else
       ' Hilangkan -[param] --- contoh C:\Memeil\Jelek.exe -start
       nNum = InStr(StrReverse(SFILE), "-") ' ambil terkanan pertama [karena namafile atau folder boleh -]
       If nNum > 0 Then
          sTmp = Left$(SFILE, Len(SFILE) - nNum)
       Else
          sTmp = SFILE
       End If
    End If
    
    Do
        'jika ada spasi terkanan hilangkan
        If Right$(sTmp, 1) = Chr$(32) Then
            sTmp = Left$(sTmp, Len(sTmp) - 1)
        Else
            sTmp = sTmp
        End If
    Loop While Right$(sTmp, 1) = Chr$(32) ' hapus sampai char terkanan bukan spasi

    ' klo ada Chr$(34) / [""] -- buang aj
    sTmp = Replace(sTmp, Chr$(34), "")
    
    '-------Lalu buat jaga-jaga kalo auto path misal SOUNDMAN.EXE
    If InStr(sTmp, "\") = 0 Then
       sSpecial = GetSpecFolder(WINDOWS_DIR) ' coba di windows dulu
       If ValidFile(sSpecial & "\" & sTmp) = True Then
          sTmp = sSpecial & "\" & sTmp
       Else
          sSpecial = GetSpecFolder(SYSTEM_DIR) ' coba di system32 sekarang
          If ValidFile(sSpecial & "\" & sTmp) = True Then
             sTmp = sSpecial & "\" & sTmp
          End If
       End If
    End If
    If ValidFile(sTmp) = True Then sTmp = sTmp Else sTmp = "Gagal ;("
Else
    sTmp = SFILE
End If

KLIMAKS:
' Klimaks :D
BufferStartupPath = sTmp
End Function


' Buat Enum startup
Public Function EnumRegStartup(ByRef sFileStart() As String, bWithCommon As Boolean) As Long
Dim nJum         As Long
Dim nLong        As Long
Dim nStart       As Long
Dim nCount       As Long
Dim sName        As String
Dim SFILE        As String
Dim ArrFile()    As String
Dim sPathReg(7)  As String
Dim sKeyRegN(7)  As String
Dim sValueName() As String
Dim sValueData() As String


ReDim sFileStart(100) As String ' karena blum tahu secara pasti berap Startup-nya

sKeyRegN(0) = "HKCU"
sPathReg(0) = "Software\Microsoft\Windows\CurrentVersion\Run"

sKeyRegN(1) = "HKLM"
sPathReg(1) = "Software\Microsoft\Windows\CurrentVersion\Run"

sKeyRegN(2) = "HKLM"
sPathReg(2) = "Software\Microsoft\Windows\CurrentVersion\Run-"

sKeyRegN(3) = "HKLM"
sPathReg(3) = "Software\Microsoft\Windows\CurrentVersion\RunOnce"

sKeyRegN(4) = "HKLM"
sPathReg(4) = "Software\Microsoft\Windows\CurrentVersion\RunOnceEx"


For nStart = 0 To 4
    nJum = RegEnumStr(SingkatanKey(sKeyRegN(nStart)), sPathReg(nStart), sValueName(), sValueData())
    For nLong = 1 To nJum
        SFILE = BufferStartupPath(sValueData(nLong))
        If ValidFile(SFILE) = True Then
           sFileStart(nCount) = SFILE
           nCount = nCount + 1
        End If
    Next
    ' hayoo habis dipakai diset ulang dulu
    nLong = 1
    Erase sValueName
    Erase sValueData
Next
' Ditambah Reg Start-Up Singgle
sPathReg(5) = "SOFTWARE\microsoft\Windows NT\CurrentVersion\Winlogon"
sKeyRegN(5) = "HKLM"
SFILE = GetSTRINGValue(SingkatanKey(sKeyRegN(5)), sPathReg(5), "Shell")
If UCase$(SFILE) <> "EXPLORER.EXE" Then ' berarti ada tuh
   SFILE = Mid$(SFILE, InStr(SFILE, Chr$(32)) + 1)
   If ValidFile(SFILE) = True Then
      sFileStart(nCount) = SFILE
      nCount = nCount + 1
   End If
End If

sPathReg(6) = "SOFTWARE\microsoft\Windows NT\CurrentVersion\Winlogon"
sKeyRegN(6) = "HKLM"
SFILE = GetSTRINGValue(SingkatanKey(sKeyRegN(6)), sPathReg(6), "Userinit")
sName = GetSpecFolder(SYSTEM_DIR) & "\userinit.exe"
If UCase$(SFILE) <> UCase$(sName) Then  ' berarti ada tuh
   SFILE = Replace(UCase$(SFILE), UCase$(sName) & ",", "")
   SFILE = BuangSpaceAwal(SFILE)
   If ValidFile(SFILE) = True Then
      sFileStart(nCount) = SFILE
      nCount = nCount + 1
   End If
End If

sPathReg(7) = "SOFTWARE\microsoft\Windows NT\CurrentVersion\Windows"
sKeyRegN(7) = "HKLM"
SFILE = GetSTRINGValue(SingkatanKey(sKeyRegN(7)), sPathReg(7), "Load")
If ValidFile(SFILE) = True Then
   sFileStart(nCount) = SFILE
   nCount = nCount + 1
End If

If bWithCommon = True Then
   nJum = GetFile(GetSpecFolder(USER_STARTUP), ArrFile)
   For nLong = 1 To nJum
       sFileStart(nCount) = ArrFile(nLong - 1)
       nCount = nCount + 1
   Next
   nLong = 1 ' reset
   nJum = GetFile(GetSpecFolder(ALL_USER_STARTUP), ArrFile)
   For nLong = 1 To nJum
       sFileStart(nCount) = ArrFile(nLong - 1)
       nCount = nCount + 1
   Next

End If

EnumRegStartup = nCount

End Function


Public Sub InstalInReg(sAppPath As String, sParam As String)
    SetStringValue SingkatanKey("HKLM"), "Software\Microsoft\Windows\CurrentVersion\Run", "Wan'iez Antivirus", sAppPath & sParam
End Sub

Public Sub UnInstalInReg(sValueStartup As String)
    DeleteValue SingkatanKey("HKLM"), "Software\Microsoft\Windows\CurrentVersion\Run", sValueStartup
End Sub
Public Function Install_CMenuDB(CMenuName As String) ' shell menu
    SetStringValue &H80000000, "*\shell\" & CMenuName & "\command", "", App_FullPathW(False) & " -K %1"
End Function
Public Function UnInstall_CMenuDB(CMenuName As String) ' shell menu
Static Detx As Long

Detx = DeleteKey(&H80000000, "*\shell\" & CMenuName & "\command")
Detx = DeleteKey(&H80000000, "*\shell\" & CMenuName)
End Function

Public Function Install_CMenu(CMenuName As String) ' shell menu
    SetStringValue &H80000002, "Software\Classes\Folder\shell\" & CMenuName & "\command", "", App_FullPathW(False) & " -K %1"
End Function

Public Function unInstall_CMenu(CMenuName As String) ' cabut shell menu
Static Detx As Long
    Detx = DeleteKey(&H80000002, "Software\Classes\Folder\shell\" & CMenuName & "\command")
    Detx = DeleteKey(&H80000002, "Software\Classes\Folder\shell\" & CMenuName)
End Function
Public Function Regis_EXT_AV()
Dim sPath As String
sPath = GetFilePath(App_FullPathW(False))

End Function

