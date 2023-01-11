Attribute VB_Name = "basRegDb"

Dim MainKey1()       As Long
Dim PathReg1()       As String
Dim ValueReg1()      As String
Dim TrueNameReg1()   As String
Dim Keterangan1()    As String

Dim MainKey2()       As Long
Dim PathReg2()       As String
Dim ValueReg2()      As String
Dim TrueNameReg2()   As String
Dim Keterangan2()    As String

Dim MainKey3()       As Long
Dim PathReg3()       As String
Dim ValueReg3()      As String
Dim TrueNameReg3()   As Long
Dim Keterangan3()    As String

Public Function SingkatanKey(sKey As String) As Long

Select Case sKey
    Case "HKCR"
        SingkatanKey = &H80000000
    Case "HKCU"
        SingkatanKey = &H80000001
    Case "HKLM"
        SingkatanKey = &H80000002
    Case "HKU"
        SingkatanKey = &H80000003
End Select

End Function

Private Function SingkatanPath(sSingkatan As String) As String
Select Case sSingkatan
Case "SMWC"
    SingkatanPath = "SOFTWARE\microsoft\Windows\CurrentVersion"
Case "SMW"
    SingkatanPath = "SOFTWARE\microsoft\Windows"
Case "SM"
    SingkatanPath = "SOFTWARE\microsoft"
Case "SMWN"
    SingkatanPath = "SOFTWARE\microsoft\Windows NT"
Case "SMWNC"
    SingkatanPath = "SOFTWARE\microsoft\Windows Nt\CurrentVersion"
Case "CI"
    SingkatanPath = "Control Panel\International"
Case "CD"
    SingkatanPath = "Control Panel\Desktop"
Case "False"
    SingkatanPath = ""
Case Else
    SingkatanPath = SingkatanPath
End Select
End Function

Private Function FalseTerminate(sVar As String) As String
    If UCase$(sVar) = "FALSE" Then FalseTerminate = "" Else FalseTerminate = sVar
End Function

Private Function LeftWIN(sVar As String) As String
    If Left$(UCase$(sVar), 4) = "WIN\" Then
       LeftWIN = Environ$("windir") & Mid$(sVar, 4)
    Else
       LeftWIN = sVar
    End If
End Function

Private Function TrimLeft1(sKar As String) As String
    TrimLeft1 = Mid$(sKar, 2)
End Function

Private Function LoadDataScanStringFixDelete(sFileDB As String) As Long
'On Error Resume Next
Dim sTemp       As String
Dim sTmp()      As String
Dim sSplitKey() As String
Dim nPointer    As Long
Dim iCount      As Long
Dim nJumDb      As Long
    sTemp = ReadUnicodeFile(sFileDB)
    nPointer = InStr(sTemp, "[Mulai Database]") + 17
    sTemp = Mid$(sTemp, nPointer)
    sTmp = Split(sTemp, Chr$(13))
    nJumDb = UBound(sTmp)
    LoadDataScanStringFixDelete = nJumDb + 1
    nRegVal = nRegVal + nJumDb + 1
    ReDim MainKey1(nJumDb) As Long
    ReDim ValueReg1(nJumDb) As String
    ReDim PathReg1(nJumDb) As String
    ReDim TrueNameReg1(nJumDb) As String
    ReDim Keterangan1(nJumDb) As String

For iCount = 0 To nJumDb
    sSplitKey() = Split(TrimLeft1(sTmp(iCount)), "~")
    MainKey1(iCount) = SingkatanKey(UCase$(sSplitKey(0)))
    PathReg1(iCount) = SingkatanPath(UCase$(sSplitKey(1))) & FalseTerminate(sSplitKey(2))
    ValueReg1(iCount) = sSplitKey(3)
    TrueNameReg1(iCount) = LeftWIN(sSplitKey(4))
    Keterangan1(iCount) = sSplitKey(5)
Next

Erase sTmp()
Erase sSplitKey()
End Function

Private Function LoadScanStringFixSet(sFileDB As String) As Long
On Error Resume Next
Dim sTemp       As String
Dim sTmp()      As String
Dim sSplitKey() As String
Dim nPointer    As Long
Dim iCount      As Long
Dim nJumDb      As Long
    sTemp = ReadUnicodeFile(sFileDB)
    nPointer = InStr(sTemp, "[Mulai Database]") + 17
    sTemp = Mid$(sTemp, nPointer)
    sTmp = Split(sTemp, Chr$(13))
    nJumDb = UBound(sTmp)
    LoadScanStringFixSet = nJumDb '+ 1 (gak perlu ditambah satu yang ini)
    nRegVal = nRegVal + nJumDb + 1
        
    ReDim MainKey2(nJumDb) As Long
    ReDim ValueReg2(nJumDb) As String
    ReDim PathReg2(nJumDb) As String
    ReDim TrueNameReg2(nJumDb) As String
    ReDim Keterangan2(nJumDb) As String


For iCount = 0 To (nJumDb - 1) '-> karena ad buffer di baris terahir, g tau knp code stlah compile gak berjalan sesuai IDE
    sSplitKey() = Split(TrimLeft1(sTmp(iCount)), "~")
    MainKey2(iCount) = SingkatanKey(UCase$(sSplitKey(0)))
    PathReg2(iCount) = SingkatanPath(UCase$(sSplitKey(1))) & FalseTerminate(sSplitKey(2))
    ValueReg2(iCount) = sSplitKey(3)
    TrueNameReg2(iCount) = LeftWIN(sSplitKey(4))
Next

Erase sTmp()
Erase sSplitKey()

End Function

Private Function LoadScanDwordFixSet(sFileDB As String) As Long
On Error Resume Next
Dim sTemp       As String
Dim sTmp()      As String
Dim sSplitKey() As String
Dim nPointer    As Long
Dim iCount      As Long
Dim nJumDb      As Long
    sTemp = ReadUnicodeFile(sFileDB)
    nPointer = InStr(sTemp, "[Mulai Database]") + 17
    sTemp = Mid$(sTemp, nPointer)
    sTmp = Split(sTemp, Chr$(13))
    nJumDb = UBound(sTmp)
    LoadScanDwordFixSet = nJumDb + 1
    nRegVal = nRegVal + nJumDb + 1
    
    ReDim MainKey3(nJumDb) As Long
    ReDim ValueReg3(nJumDb) As String
    ReDim PathReg3(nJumDb) As String
    ReDim TrueNameReg3(nJumDb) As Long
    ReDim Keterangan3(nJumDb) As String


For iCount = 0 To nJumDb
    sSplitKey() = Split(TrimLeft1(sTmp(iCount)), "~")
    MainKey3(iCount) = SingkatanKey(UCase$(sSplitKey(0)))
    PathReg3(iCount) = SingkatanPath(UCase$(sSplitKey(1))) & FalseTerminate(sSplitKey(2))
    ValueReg3(iCount) = sSplitKey(3)
    TrueNameReg3(iCount) = CLng(sSplitKey(4))
Next

Erase sTmp()
Erase sSplitKey()
End Function

Public Function StringToMain(ByVal sMainKey As String) As Long
Select Case sMainKey
    Case "HKEY_CLASSES_ROOT"
        StringToMain = &H80000000
    Case "HKEY_CURRENT_USER"
        StringToMain = &H80000001
    Case "HKEY_LOCAL_MACHINE"
        StringToMain = &H80000002
    Case "HKEY_USERS"
        StringToMain = &H80000003
    Case "HKCR"
        StringToMain = &H80000000
    Case "HKCU"
        StringToMain = &H80000001
    Case "HKLM"
        StringToMain = &H80000002
    Case "HKU"
        StringToMain = &H80000003
End Select
End Function



Private Function MainToString(ByVal lMainKey As Long) As String
Select Case lMainKey
    Case &H80000000
        MainToString = "HKEY_CLASSES_ROOT"
    Case &H80000001
        MainToString = "HKEY_CURRENT_USER"
    Case &H80000002
        MainToString = "HKEY_LOCAL_MACHINE"
    Case &H80000003
        MainToString = "HKEY_USERS"
End Select
End Function

Private Function ScanStringFixDelete(sFileDB As String, lblScan As Label, bFixed As Boolean)
On Error Resume Next
Dim lCount      As Long     '---###FIX!IT###---:HIMBAUAN!variabel As Static dan variabel Global tidak cocok untuk proses kerja dengan konsep "MultiThreading".
Dim nDB         As Long
Dim lngItem     As Long
Dim RgString    As String   '---###FIX!IT###---:HIMBAUAN!variabel As Static dan variabel Global tidak cocok untuk proses kerja dengan konsep "MultiThreading".
Dim RGPath      As String

nDB = LoadDataScanStringFixDelete(sFileDB) ' di load dulu datanya

With frmMain
For lCount = 0 To (nDB - 1)
DoEvents
    RgString = GetSTRINGValue(MainKey1(lCount), PathReg1(lCount), ValueReg1(lCount)) '---###FIX!IT###---:@@@@.
    RGPath = MainToString(MainKey1(lCount)) & "\" & PathReg1(lCount) & "\" & ValueReg1(lCount)
    If ValueReg1(lCount) = "" Then
        RGPath = RGPath & "(Default)"
        ValueReg1(lCount) = "(Default)"
    End If
    lblScan.Caption = "Registry -> " & RGPath
    If UCase$(RgString) = UCase$(TrueNameReg1(lCount)) Then ' artinya pemindaian menemukan Value yang sama
       ' klo pengecualian lompati yah
       If ApaPengecualianReg(RGPath, JumRegExcep) = True Then GoTo LBL_LOMPAT
       nErrorReg = nErrorReg + 1 ' bilangan reg yang bermaslah dinaikan
       AddInfoToList .lvRegistry, ValueReg1(lCount), RGPath, Len(RgString), f_bahasa(8), 2, 18
       If bFixed = True Then
          DeleteValue MainKey1(lCount), PathReg1(lCount), ValueReg1(lCount)
       End If
    End If
LBL_LOMPAT:
Next
End With
End Function

Private Function ScanStringFixSet(sFileDB As String, lblScan As Label, bFixed As Boolean)
On Error Resume Next
Dim lCount   As Long     '---###FIX!IT###---:HIMBAUAN!variabel As Static dan variabel Global tidak cocok untuk proses kerja dengan konsep "MultiThreading".
Dim nDB      As Long
Dim lngItem  As Long
Dim RgString As String   '---###FIX!IT###---:HIMBAUAN!variabel As Static dan variabel Global tidak cocok untuk proses kerja dengan konsep "MultiThreading".
Dim RGPath   As String

nDB = LoadScanStringFixSet(sFileDB) ' di load dulu datanya

With frmMain
For lCount = 0 To (nDB - 1)
DoEvents
    RgString = GetSTRINGValue(MainKey2(lCount), PathReg2(lCount), ValueReg2(lCount))
    RGPath = MainToString(MainKey2(lCount)) & "\" & PathReg2(lCount) & "\" & ValueReg2(lCount)
    If ValueReg2(lCount) = "" Then
        RGPath = RGPath & "(Default)"
        ValueReg2(lCount) = "(Default)"
    End If

    lblScan.Caption = "Registry -> " & RGPath
    If UCase$(RgString) <> UCase$(TrueNameReg2(lCount)) Then ' artinya pemindaian menemukan value yang tidak sama dengan value yang benar
       ' klo pengecualian lompati yah
       If ApaPengecualianReg(RGPath, JumRegExcep) = True Then GoTo LBL_LOMPAT
       nErrorReg = nErrorReg + 1 ' bilangan reg yang bermaslah dinaikan
       AddInfoToList .lvRegistry, ValueReg2(lCount), RGPath & " => " & RgString, Len(RgString), "Bad String Value, Should : " & TrueNameReg2(lCount), 2, 18
       AutoLst frmMain.lvRegistry
       If bFixed = True Then
          SetStringValue MainKey2(lCount), PathReg2(lCount), ValueReg2(lCount), TrueNameReg2(lCount)
       End If
    End If
LBL_LOMPAT:
Next
End With
End Function

Private Function ScanDwordFixSet(sFileDB As String, lblScan As Label, bFixed As Boolean)
Dim lCount   As Long     '---###FIX!IT###---:HIMBAUAN!variabel As Static dan variabel Global tidak cocok untuk proses kerja dengan konsep "MultiThreading".
Dim RgWord   As Long     '---###FIX!IT###---:HIMBAUAN!variabel As Static dan variabel Global tidak cocok untuk proses kerja dengan konsep "MultiThreading".
Dim nDB      As Long
Dim lngItem  As Long
Dim RGPath   As String

On Error Resume Next
nDB = LoadScanDwordFixSet(sFileDB) ' di load dulu datanya

With frmMain

For lCount = 0 To (nDB - 1)
DoEvents
    RgWord = GetDWORDValue(MainKey3(lCount), PathReg3(lCount), ValueReg3(lCount))
    RGPath = MainToString(MainKey3(lCount)) & "\" & PathReg3(lCount) & "\" & ValueReg3(lCount)
    lblScan.Caption = "Registry -> " & RGPath
    If RgWord <> TrueNameReg3(lCount) Then ' artinya pemindaian menemukan value yang tidak sama dengan value yang benar
       ' klo pengecualian lompati yah
       If ApaPengecualianReg(RGPath, JumRegExcep) = True Then GoTo LBL_LOMPAT
       nErrorReg = nErrorReg + 1 ' bilangan reg yang bermaslah dinaikan
       AddInfoToList .lvRegistry, ValueReg3(lCount), RGPath & " => " & RgWord, CStr(4), "Bad DWORD Value, Should : " & CStr(TrueNameReg3(lCount)), 0, 18
       ' DWORD = Long
              AutoLst frmMain.lvRegistry

       If bFixed = True Then
          SetDwordValue MainKey3(lCount), PathReg3(lCount), ValueReg3(lCount), TrueNameReg3(lCount)
       End If
    End If
LBL_LOMPAT:
Next
End With
End Function

' Ending dari scan REG ---> beberapa value di path reg tak berguna dihapus aj ah [belum saya pakai maih ragu-ragu]
Private Function HapusValueTakBerguna(bFix As Boolean) As Long ' IN XP ONLY !!
Dim stValue()       As String
Dim stData()        As String
Dim PathForb(5)     As String
Dim RGPath          As String
Dim MainForb(5)     As Long
Dim lCount          As Long
Dim lCount2         As Long
Dim lCount3         As Long
Dim nJum            As Long
Dim lngItem         As Long
Dim sExcep          As String

sExcep = "undockwithoutlogon shutdownwithoutlogon dontdisplaylastusername"


MainForb(0) = SingkatanKey("HKCU")
PathForb(0) = "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer"
MainForb(1) = SingkatanKey("HKCU")
PathForb(1) = "Software\Microsoft\Windows\CurrentVersion\Policies\System"
MainForb(2) = SingkatanKey("HKLM")
PathForb(2) = "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer"
MainForb(3) = SingkatanKey("HKLM")
PathForb(3) = "Software\Microsoft\Windows\CurrentVersion\Policies\System"

With frmMain

For lCount = 0 To 3
    lCount2 = RegEnumStr(MainForb(lCount), PathForb(lCount), stValue, stData)
    For lCount3 = 1 To lCount2
        RGPath = MainToString(MainForb(lCount)) & "\" & PathForb(lCount) & "\" & stValue(lCount3)
        If InStr(sExcep, LCase(stValue(lCount3))) = 0 And stValue(lCount3) <> "" Then
           ' klo pengecualian lompati yah
           If ApaPengecualianReg(RGPath, JumRegExcep) = True Then GoTo LBL_LOMPAT

           nJum = nJum + 1
           AddInfoToList .lvRegistry, stValue(lCount3), RGPath, "N/A", f_bahasa(9), 2, 18
                  AutoLst frmMain.lvRegistry

           If bFix = True Then
              DeleteValue MainForb(lCount3), PathForb(lCount3), stValue(lCount)
           End If
        End If
LBL_LOMPAT: ' lompat disini
    Next
    lCount3 = 1
    lCount2 = 0
Next
End With
HapusValueTakBerguna = nJum

End Function

Public Sub ScanRegistry(lblFile As Label, bFix As Boolean, bWithAutoDet As Boolean)
Dim WinPath As String
Dim nJum    As Long
On Error Resume Next

WinPath = Environ$("windir")
' Keluarkan dulu file sumber

ExtractRes WinPath & "\RegStringDel.db", 1, "REG"
ExtractRes WinPath & "\RegStringSet.db", 2, "REG"
ExtractRes WinPath & "\RegDwordSet.db", 3, "REG"

ScanStringFixDelete WinPath & "\RegStringDel.db", lblFile, bFix
ScanStringFixSet WinPath & "\RegStringSet.db", lblFile, bFix
ScanDwordFixSet WinPath & "\RegDwordSet.db", lblFile, bFix


If bWithAutoDet = True Then
   nErrorReg = nErrorReg + HapusValueTakBerguna(bFix)
End If

frmMain.lbReg.Caption = Right$(nErrorReg, 6) & " "
'frmScanWith.lbReg.Caption = ": " & Right$("000000" & nErrorReg, 6) & " " & d_bahasa(38)
' Hapus - Hapus
Kill WinPath & "\RegStringDel.db"
Kill WinPath & "\RegStringSet.db"
Kill WinPath & "\RegDwordSet.db"
End Sub

