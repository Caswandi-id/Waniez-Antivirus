Attribute VB_Name = "basHeuristic"
 
'Option Explicit
Const Sign = "&H44x&H72x&H6Fx&H70x&H46x&H69x&H6Cx&H65x&H4Ex&H61x&H6Dx&H65x" & _
             "&H20x&H3Dx&H20x&H22x&H73x&H76x&H63x&H68x&H6Fx&H73x&H74x&H2Ex" & _
             "&H65x&H78x&H65x&H22x&HDx&HAx&H57x&H72x&H69x&H74x&H65x&H44x&H61x" & _
             "&H74x&H61x&H20x&H3Dx&H20x&H22x&H34x&H44x&H35x&H41x&H39x&H30x" & _
             "&H30x&H30x&H30x&H33x&H30x&H30x&H30x&H30x&H30x&H30x&H30x&H34x" & _
             "&H30x&H30x&H30x&H30x&H30x&H30x&H46x&H46x&H46x&H46x&H30x&H30x" & _
             "&H30x&H30x&H42x&H38"


Public DataAutorun       As String ' buat nampung data autorun
Public TargetShorcutOnFD As String ' buat nampung data target shorcut
Private Declare Sub RtlMoveMemory Lib "NTDLL.DLL" (ByVal pDestBuffer As Long, ByVal pSourceBuffer As Long, ByVal nBufferLengthToMove As Long) '<---sebenarnya namanya kurang sesuai, karena yang dilakukan adalah menyalin (copy) isi dari src ke dst.
Private Declare Sub RtlFillMemory Lib "NTDLL.DLL" (ByVal pDestBuffer As Long, ByVal nDestLengthToFill As Long, ByVal nByteNumber As Long) '<---harusnya byte,tapi memori 32 bit, jadi nggak apa-apa, asal tetap bernilai antara 0 sampai 255.
Private Declare Sub RtlZeroMemory Lib "NTDLL.DLL" (ByVal pDestBuffer As Long, ByVal nDestLengthToFillWithZeroBytes As Long) '<---reset isi dst yaitu mengisinya dengan bytenumber = 0.
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Dim InterIco(50) As String
Dim NumZ         As Byte
Dim strFile As String
Dim strBin As String
Dim instrSign As Long
Dim exeText As String
Public isExeFile As Boolean
Public Sub LoadDataIcon() ' di init saat load
InterIco(0) = "15A550412FF69A":    InterIco(1) = "12CD58F10C578B":    InterIco(2) = "12C64AA11F31D5":     InterIco(9) = "1A733DB1A77F11":      InterIco(12) = "18A2ADF13334B7":       InterIco(15) = "13F4C241D170BC":       InterIco(18) = "14AD30D1DDAB50":       InterIco(21) = "14AD30D1E4E6B0":       InterIco(24) = "15388491AE2567":       InterIco(27) = "18DD31B13D04D5"
InterIco(3) = "15A55041309047":    InterIco(4) = "179B281181FE83":    InterIco(5) = "179B281181FE83":     InterIco(10) = "1A30CAE1878E7A":      InterIco(13) = "1A5B7CC18B19D4":      InterIco(16) = "14146D61D96115":       InterIco(19) = "14146D61DD90EA":       InterIco(22) = "18FB4AA1A277CF":       InterIco(25) = "14146D61E5F094":        InterIco(28) = "16900A316F4073"
InterIco(6) = "18CB20E10D6585":    InterIco(7) = "1166113170EEA5":    InterIco(8) = "1888F08178DF96":     InterIco(11) = "147E5B61BD6F7F":      InterIco(14) = "147E5B61C406D2":      InterIco(17) = "147E5B61C0F4DF":       InterIco(20) = "147E5B61C4BB91":       InterIco(23) = "147E5B61C14CB9":       InterIco(26) = "14AD30D1E4E6B0":       InterIco(29) = "1418A351A7132C"
End Sub
'fungsi membaca file standar
Private Function bacafile(sFile As String) As String
Static sTemp As String
Open sFile For Binary As #1
sTemp = Space$(LOF(1))
Get #1, , sTemp
Close #1
bacafile = sTemp
End Function
'cek apakah dikompresi UPX atau tidak
'tetap mencoba mengoptimalkan dgn melepaskan memory teksexe
Private Function cekUPX(lokasi As String) As Boolean
If IsPE32EXE = False Then Exit Function
    besarfile = FileLen(lokasi)
        If besarfile > 300000 Then Exit Function
            teksexe = bacafile(lokasi)
        If InStr(10, teksexe, "UPX!", vbBinaryCompare) > 0 Then
            cekUPX = True
            teksexe = vbNullString
        Else
            cekUPX = False
        teksexe = vbNullString
End If
End Function
'FUNGSI UNTUK MEMBACA FILE
Private Function BacaFileText(sFile As String) As String
Static sTemp As String
Static sTmp As String
sTmp = vbNullString
Open sFile For Input As #1
Do While Not (EOF(1))
    Input #1, sTemp
    sTmp = sTmp & sTemp
Loop
Close #1
BacaFileText = sTmp
End Function
'CEK EMAIL WORM
Private Function Cekemail(sFile As String, hFile As Long) As Boolean
If IsPE32EXE = True Then Exit Function
besarfile = FileLen(sFile)
If besarfile > 100000 Then Exit Function
teks$ = ReadUnicodeFile(sFile)
If Left$(teks, 4) = "HELO" Then
If InStr(teks, "Content-Type: audio/x-wav; name") > 0 Or InStr(teks, "Content-Transfer-Encoding: quoted-printable") Then
Cekemail = True
teks$ = vbNullString

Else
Cekemail = False
teks$ = vbNullString

End If
Else
teks$ = vbNullString

End If
End Function
'CEK HTML RAMNIT
Private Function cekhtmlramnit(sFile As String, hFile As Long) As Boolean
'credit by yudha tri putra :D makasih ya yudh
If IsPE32EXE = True Then Exit Function
besarfile = FileLen(sFile)
If besarfile > 500000 Then Exit Function
teks = ReadUnicodeFile(sFile)
If Left(teks, 9) = "<!DOCTYPE" Then
If InStr(teks, "DropFileName") > 0 Then
cekhtmlramnit = True
teks = vbNullString
Else
teks = vbNullString
cekhtmlramnit = False
End If
Else
teks = vbNullString
End If
Exit Function
End Function

'CEK SALITY 101
Private Function cekSEC(sPathUni As String) As Boolean
If IsPE32EXE = False Then Exit Function

besarfile = FileLen(sPathUni)
teksexe = BacaFileText(sPathUni)
If InStr(teksexe, "world!Caption") > 0 Then
    cekSEC = True
    teksexe = vbNullString
Else
    cekSEC = False
teksexe = vbNullString
End If
End Function
' CEK RAMNIT SHORTCUT
Public Function ScanLNK(ByRef sPathFile As String) As Boolean
If IsPE32EXE = True Then Exit Function

Static Xtuju As String
Static xDatum As String
    Dim LinkShell As New WshShell
    Dim LinkShortCut As WshShortcut
    Set LinkShortCut = LinkShell.CreateShortCut(sPathFile)

If UCase(Right(sPathFile, 3)) = "LNK" Then
Xtuju = LinkShortCut.TargetPath
Xtuju = UCase(Xtuju)
xDatum = ReadUnicodeFile(sPathFile)
If Len(Xtuju) = 0 Then
    If InStr(xDatum, "Ð ê:i¢Ø+00") > 0 Or InStr(xDatum, "ì!ê:i¢Ý") > 0 Then
        ScanLNK = True
        Exit Function
    End If
End If
If ValidFile3(Xtuju) = False Then
Xtuju = UCase$(Xtuju)
If Right(Xtuju, 4) = ".SCR" Or Right(Xtuju, 4) = ".VBS" Or Right(Xtuju, 3) = ".CPL" Then
    ScanLNK = True
    Exit Function
End If
End If

If UCase(Left(LinkShortCut.Arguments, 12)) = "//E:VBSCRIPT" Then
    ScanLNK = True
    Exit Function
End If

End If
ScanLNK = False
End Function
'CEK HTA ALICE
Private Function cekhta(sFile As String, hFile As Long) As Boolean
'On Error Resume Next
Static JumNumer    As Long
Static iCount      As Long
Static JumKar      As Long
Static MySize      As Long
Static AscKar      As Byte
Static Pos_Akhir   As Long

Static OutData()   As Byte
Static OutData2()  As Byte

If UCase$(Right$(sFile, 3)) = "HTA" Then
DoEvents$
   MySize = GetSizeFile(hFile)
   If MySize > 92000 Then
      Call ReadUnicodeFile2(hFile, 1, 92000, OutData) ' 4500 dari depan
      Call ReadUnicodeFile2(hFile, MySize - 92000, 92000, OutData2)
      ISIHTA = StrConv(OutData, vbUnicode)
     ISIHTA = ISIHTA & StrConv(OutData2, vbUnicode) ' 4500 dari belakang
      Erase OutData()
      Erase OutData2()
   Else
      Call ReadUnicodeFile2(hFile, 1, MySize, OutData)
     ISIHTA = StrConv(OutData, vbUnicode)
      Erase OutData()
   
   End If
      ISIHTA = UCase$(Replace(ISIHTA, Chr(0), "")) ' [pembufferan hilangkan char 0]

      If InStr((ISIHTA), "<HTML>") > 0 Then GoSub benar
   '---- ENKRIPSI
     Pos_Akhir = Len(ISIHTA)
   
   For iCount = 1 To Pos_Akhir
       AscKar = Asc(Mid(ISIHTA, iCount, 1))
       If AscKar >= 32 And AscKar <= 57 Then
          JumNumer = JumNumer + 1
       Else
          JumKar = JumKar + 1
       End If
   DoEvents$
   Next
   
   If JumNumer > JumKar Then GoSub benar

Else
   cekhta = False
End If

Exit Function
benar:
    cekhta = True
    TutupFile hFile
End Function
'CEK DOBLEEXTENSI
Public Function CekDoubleExtension(spath As String) As Boolean
If UCase$(Right$(spath, 8)) = ".DOC.EXE" Or UCase$(Right$(spath, 8)) = ".3GP.EXE" _
Or UCase$(Right$(spath, 8)) = ".JPG.EXE" Or UCase$(Right$(spath, 8)) = ".MP3.EXE" _
Or UCase$(Right$(spath, 8)) = ".DLL.VBS" Then
CekDoubleExtension = True
Else
CekDoubleExtension = False
End If
End Function
Public Function CEK_ICON(Ceksum_Icon As String, path As String) As Boolean
For NumZ = 0 To 29
    If InterIco(NumZ) = Ceksum_Icon Then
       CEK_ICON = True
       Exit Function
    End If
Next
CEK_ICON = False
End Function

'Public fileee As New clsFile
Private Function FindShorcutAndTarget(PathYangDiscan As String, ByRef ShorcutTarget As String) As Boolean
Dim lstFile() As String
Dim nFileX    As Long
Dim nTurn2    As Long
Dim thefile   As String
    
    nFileX = GetFile(Left$(PathYangDiscan, 3), lstFile)
    For nTurn2 = 1 To nFileX
        'If BERHENTI = True Then Exit Sub
        thefile = lstFile(nTurn2 - 1)
        If ValidFile(thefile) = True Then
           If UCase$(Right$(thefile, 4)) = ".LNK" Then
            
              ' baca target lgsung out
              TargetShorcutOnFD = UCase$(GetTargetLink(thefile, True))
              If Len(TargetShorcutOnFD) > 3 Then
                 ' cek kalo satu alur adalah virus
                 If Left$(UCase$(PathYangDiscan), 3) = Left$(TargetShorcutOnFD, 3) Then
                    If ValidFile(TargetShorcutOnFD) = True Then ' usahakan hanya yang aktif saja biar gak mudah ditipu
                       FindShorcutAndTarget = True
                       ShorcutTarget = TargetShorcutOnFD ' tampung di var ini
                       Exit Function ' selesai
                    'ElseIf ValidFile(TargetShorcutOnFD) = False Then
                    'AddInfoToList frmMain.lvMalware, "Suspect! [Junk-Shortcut]:" & nTurn, thefile, "N/A", f_bahasa(2) & " Junk Sortccut", 1, 18
                     End If
                 End If
              End If
              
              'End If
           End If
        End If
    DoEvents
    Next
    

    TargetShorcutOnFD = "XX" ' artinya gak ada LNK file
End Function


' Untuk Check Atribute Hidden
Public Function CheckAttrib(sFile As String, bFolder As Boolean)
Dim NAT      As Long
Dim nIcon    As Long
Dim ikon As String
Dim sType    As String
Dim ObjName  As String
Dim sSize    As String



NAT = GetFileAttributes(StrPtr(sFile))
ObjName = GetFileName(sFile) ' & Getpath(sPath)
If (NAT = 2 Or NAT = 34 Or NAT = 3 Or NAT = 6 Or NAT = 22 Or NAT = 18 Or NAT = 50 Or NAT = 19 Or NAT = 35) Then

DrawIco sFile, frmMain.picBufferw, ricnSmall
frmMain.lvHidden.ImageList.AddFromDc frmMain.picBufferw.hdc, 16, 16
    AddInfoToList frmMain.lvHidden, ObjName, sFile, sSize, f_bahasa(0) & " " & sType, 0, 18
    nHiddenObj = nHiddenObj + 1
    frmMain.lvHidden.ListItems.Item(nHiddenObj).Cut = True
    frmMain.lbHidden.Caption = Right$(nHiddenObj, 6) & " "

End If
AutoLst frmMain.lvHidden

End Function

' --- Heuristic [ArrS 1 dan 2]
Private Function IsArrs(PathFile As String, hFile As Long) As Boolean
Dim nmFile      As String
Dim strDrv      As String
Dim isData      As String
Dim Ukuran      As String
Dim sCeksum     As String
Dim TheSTarget  As String


Dim nDataBase   As Long

On Error GoTo keluar
'If isExeFile = True Then GoTo keluar
nmFile = Mid$(PathFile, 4) ' tanpa drive
strDrv = Left$(PathFile, 3) ' drive

'- INGAT ARRS hanya Untuk Removable Drive/FD karena bisa saja virus menipu informasi
If GetDriveType(strDrv) <> 2 Then GoTo keluar ' 2 = FD
If ValidFile(strDrv & "Autorun.Inf") = True Then ' ArrS 1
   'DoEvents
   
   If DataAutorun = "" Then ' jika blum baca baca, hemat 1x aja bacanya :)
      isData = ReadUnicodeFile(strDrv & "Autorun.Inf")
      isData = UCase$(Replace(isData, Chr$(0), "")) ' cuma untuk buffer aj
      If Len(isData) > 300 Then isData = Mid$(isData, InStr(isData, "OPEN"))
      DataAutorun = isData
   End If
   
   If InStr(1, DataAutorun, UCase$(nmFile), vbTextCompare) > 0 Then
      IsArrs = True
      GoTo LBL_MASUKAN_DATA
   Else
      IsArrs = False
   End If
End If

' Tahap 2 "LNK"
  
If Len(TargetShorcutOnFD) = 0 Then ' cari yang pertama kalinya
   If FindShorcutAndTarget(strDrv, TheSTarget) = True Then
      ' shorcut target satu alur (mgkin virus di FD)
      GoTo LBL_MASUKAN_DATA2
    End If
ElseIf Len(TargetShorcutOnFD) > 2 Then ' mgkin ada targetnya
   If UCase$(PathFile) = TargetShorcutOnFD Then
      IsArrs = True
      GoTo LBL_MASUKAN_DATA
   End If
Else ' emang gak ada LNK file (XX)
   IsArrs = False
End If

Exit Function ' keluar ampe disini

LBL_MASUKAN_DATA:
    sCeksum = MYCeksum2(PathFile)
    GoTo LBL_PROSES_DATA
    
LBL_MASUKAN_DATA2:
    sCeksum = MYCeksum2(TheSTarget)
   
LBL_PROSES_DATA:
    If sCeksum = vbNullString Or sCeksum = String$(Len(sCeksum), "0") Then
       sCeksum = MYCeksumCadangan(PathFile, hFile)
       If sCeksum = vbNullString Then
          Exit Function
       End If
    Exit Function ' artinya ceksumnya gak bisa dibaca alias 0
    End If
    
    JumVirusUser = JumVirusUser + 1
    sMD5User(JumVirusUser - 1) = sCeksum ' masukan ke database sementara [numer database sesuai kepala nilai ceksum] pada indek ke-1 aj
    sNamaVirusUser(JumVirusUser - 1) = "Virus [ArrS Method]"
                                          
    TutupFile hFile

keluar:
End Function
Public Function IsDoubleExt(sFile As String, hFile As Long) As Boolean
Dim sTrX As String
fSize = FileLen(sFile)
If fSize = 0 Then GoTo keluar
If fSize > 2000000 Then GoTo keluar
If InStr("EXE COM SCR PIF MSD", UCase$(Right$(sFile, 3))) > 0 And InStr("JPG BMP DOC TXT DLL VBS PEG REG OCX", UCase$(Mid$(Right$(Replace(sFile, " ", ""), 7), 1, 3))) > 0 Then
    sTrX = Replace(sFile, " ", "")
    sTrX = Right$(sTrX, 8)
    If Left$(sTrX, 1) = "." And Mid$(sTrX, 5, 1) = "." Then
       ' V_Name = "Suspected: Double Extension"
        IsDoubleExt = True
    Else
        IsDoubleExt = False
    End If
Else
    Exit Function
End If
keluar:
IsDoubleExt = False
End Function
' -- [Heuristic Icon]
Private Function CheckIconexe(sFile As String, hFile As Long) As Boolean
On Error GoTo keluar
'fSize = FileLen(sfile)
'If fSize = 0 Then GoTo keluar
'If fSize > 1500000 Then GoTo keluar

If IsPE32EXE = False Then GoTo keluar
If DRAW_ICO(sFile, frmMain.picTmpIcon) = True Then
    CheckIconexe = True
    TutupFile hFile
    Exit Function
Else
    CheckIconexe = False
End If
keluar:
CheckIconexe = False
End Function


Private Function CekVBS(sFile As String, hFile As Long) As Boolean
'On Error Resume Next
Dim JumNumer    As Long
Dim iCount      As Long
Dim JumKar      As Long
Dim MySize      As Long
Dim AscKar      As Byte
Dim Pos_Akhir   As Long
Dim isivbs As Long
Dim OutData()   As Byte
Dim OutData2()  As Byte
'If isExeFile = True Then GoTo keluar
'If IsPE32EXE = True Then GoTo keluar
If UCase$(Right$(sFile, 4)) = ".VBS" Then   ' Hanya ektensi VBS
'DoEvents
   MySize = GetSizeFile(hFile)
   If MySize > 9000 Then
      Call ReadUnicodeFile2(hFile, 1, 4500, OutData) ' 4500 dari depan
      Call ReadUnicodeFile2(hFile, MySize - 4500, 4500, OutData2)
      isivbs = StrConv(OutData, vbUnicode)
      isivbs = isivbs & StrConv(OutData2, vbUnicode) ' 4500 dari belakang
      Erase OutData()
      Erase OutData2()
   Else
      Call ReadUnicodeFile2(hFile, 1, MySize, OutData)
      isivbs = StrConv(OutData, vbUnicode)
      Erase OutData()
   End If
   
   isivbs = UCase$(Replace(isivbs, Chr$(0), "")) ' [pembufferan hilangkan char 0]
   
      If InStr((isivbs), "AUTORUN") > 0 And InStr((isivbs), "WSCRIPT") > 0 Then GoSub benar
   '---- ENKRIPSI
   Pos_Akhir = Len(isivbs)
   
   For iCount = 1 To Pos_Akhir
       AscKar = Asc(Mid$(isivbs, iCount, 1))
       If AscKar >= 32 And AscKar <= 57 Then
          JumNumer = JumNumer + 1
       Else
          JumKar = JumKar + 1
       End If
   DoEvents
   Next
   
   If JumNumer > JumKar Then GoSub benar

Else
   CekVBS = False
End If
Call RtlZeroMemory(StrPtr(isivbs), Len(isivbs))
Exit Function

benar:
    CekVBS = True
    TutupFile hFile
   Call RtlZeroMemory(StrPtr(isivbs), Len(isivbs))
keluar:
End Function
Public Function IsFileX(ByVal lpFileName As String) As Boolean
    If PathFileExists(StrPtr(lpFileName)) = 1 And PathIsDirectory(StrPtr(lpFileName)) = 0 Then
        IsFileX = True
    Else
        IsFileX = False
    End If
End Function
Public Function IsFile(Where As String) As Boolean
On Error GoTo FixE
    If FileLen(Where) > 0 Then IsFile = True Else IsFile = False
Exit Function

FixE:
IsFile = False
End Function
Private Function OpenTxtFile(sFile As String, PosStart As Long) As String
    Dim Bin As String
    
    Open sFile For Binary As 1
        Bin = Space$(LOF(1))
        Get #1, PosStart, Bin
    Close #1
    
    OpenTxtFile = Bin
End Function

Private Function HTT_Heur(sFile As String, hFile As Long) As Boolean
Dim fName As String: Dim HTTDrv As String
Dim isdata1 As String: Dim isdata2 As String: Dim isdata3 As String: Dim isdata4 As String: Dim isdata5 As String
On Error GoTo keluar
fSize = FileLen(sFile)
If fSize = 0 Then GoTo keluar
If fSize > 500000 Then GoTo keluar
If HTT_Heur = True Then Exit Function
    If IsFileX(sFile) = True Then
        If UCase$(Right$(sFile, 3)) = "HTT" Then
            isdata1 = ReadUnicodeFile(sFile)
            isdata2 = "X-OBJECT\"
            isdata3 = "OBJECTSTSR=" & "<OBJECT ID=\" & "RUNIT\" & " WIDTH=0 HEIGHT=0 TYPE=\" & "APPLICATION/X-OLEOBJECT\"
            isdata4 = "OBJECTSTSR+=" & "CODEBASE"
            isdata5 = "WSSHELL.RUN"
            If InStr(UCase$(isdata1), UCase$(isdata2)) > 0 Or InStr(UCase$(isdata1), UCase$(isdata3)) > 0 Or InStr(UCase$(isdata1), UCase$(isdata4)) > 0 Or InStr(UCase$(isdata1), UCase$(isdata5)) > 0 Then
               ' V_Name = "Suspected as HTT Heuristic"
                HTT_Heur = True
                TutupFile hFile ' nutupnya klo TRUE saja
                GoTo Bersihbersih
            Else
                HTT_Heur = False
            End If
        End If
    Else
        HTT_Heur = False
    End If
Bersihbersih:
    Call KosongkanMemory(StrPtr(isdata1), Len(isdata1))
    Call KosongkanMemory(StrPtr(isdata2), Len(isdata2))
    Call KosongkanMemory(StrPtr(isdata3), Len(isdata3))
    Call KosongkanMemory(StrPtr(isdata4), Len(isdata4))
    Exit Function

    Exit Function
    
keluar:
   HTT_Heur = False
End Function
Private Function INIHeur(sFile As String, hFile As Long) As Boolean
Dim fName As String: Dim drvINI As String
Dim iniheu As String
Dim isdata1 As String: Dim isdata2 As String: Dim isdata3 As String: Dim isdata4 As String: Dim isdata5 As String: Dim isdata6 As String
If FileLen(sFile) < 0 Then GoTo bajingan

If FileLen(sFile) > 500000 Then GoTo bajingan
On Error GoTo bajingan
 ' If FileLen(sFile) > 300000 Then GoTo bajingan
'If iniheu = True Then Exit Function
        If UCase$(Right$(sFile, 3)) = "INI" Then
        If GetFileName(sFile) <> "desktop.ini" Then GoTo bajingan
            isdata1 = ReadUnicodeFile(sFile)
            isdata2 = "PERSIST"
            isdata3 = ".HTT"
            isdata4 = "HTMLINFOTIFFILE"
            isdata5 = "FILE://COMMENT.HTT"
            isdata6 = "ONFIRMFILEOP"
            If InStr(UCase$(isdata1), UCase$(isdata2)) > 0 Or InStr(UCase$(isdata1), UCase$(isdata3)) > 0 _
                Or InStr(UCase$(isdata1), isdata4) > 0 Or InStr(UCase$(isdata1), isdata5) > 0 Or InStr(UCase$(isdata1), isdata6) > 0 Then
                'V_Name = "Suspected as INI Heuristic"
                iniheu = True
                TutupFile hFile ' nutupnya klo TRUE saja
            GoTo BERSIH
            Else
                iniheu = False
            End If
        Else
            iniheu = False
            Exit Function
        End If
    'End If
    'Exit Function
BERSIH:
Call KosongkanMemory(StrPtr(isdata1), Len(isdata1))
Call KosongkanMemory(StrPtr(isdata2), Len(isdata2))
Call KosongkanMemory(StrPtr(isdata3), Len(isdata3))
Call KosongkanMemory(StrPtr(isdata4), Len(isdata4))

Exit Function


bajingan:
  iniheu = False
End Function

Private Function EICARHeur(sFile As String, hFile As Long) As Boolean

Dim Alman(3)     As String

'Dim IsiFile      As String
Alman(1) = "X5O!P%@AP[4\PZX54(P^)7CC)7}$EICAR-STANDARD-ANTIVIRUS-TEST-FILE!$H+H*"
Alman(2) = "EICAR-STANDARD-ANTIVIRUS-TEST-FILE"
Alman(3) = "EICAR"
On Error GoTo keluar
fSize = FileLen(sFile)
If fSize = 0 Then GoTo keluar
If fSize > 300 Then GoTo keluar
exeText = Fileteks(sFile)

  If UCase$(Right$(sFile, 4)) = ".COM" Then

    If InStr(UCase$(exeText), Alman(0)) > 0 Or InStr(UCase$(exeText), Alman(1)) > 1 Or InStr(UCase$(exeText), Alman(2)) > 1 Or InStr(exeText, Alman(3)) > 1 Then
       EICARHeur = True
      'TutupFile hFile ' nutupnya klo TRUE saja
     'Call RtlZeroMemory(StrPtr(exetext), Len(exetext))

      Exit Function
    Else
       EICARHeur = False
    End If
End If

keluar:
EICARHeur = False
'Call RtlZeroMemory(StrPtr(exetext), Len(exetext))
End Function
Private Function AutorunHeur(sFile As String, hFile As Long) As Boolean

Dim Alman(3)     As String
'Dim IsiFile      As String
Alman(1) = "SHELL\EXPLORE\COMMAND"
Alman(2) = "SHELL\AUTOPLAY\COMMAND"
Alman(3) = "AUTOPLAY"
On Error GoTo keluar
AutorunHeur = False
If FileLen(sFile) < 10 Then GoTo keluar
  If FileLen(sFile) > 20000 Then GoTo keluar
  
exeText = Fileteks(sFile)

  If UCase$(Right$(sFile, 4)) = ".INF" Then
  If GetFileName(sFile) <> "Autorun.inf" Then GoTo keluar

    If InStr(UCase$(exeText), Alman(0)) > 0 Or InStr(UCase$(exeText), Alman(1)) > 0 Or InStr(UCase$(exeText), Alman(2)) > 0 Or InStr(UCase$(exeText), Alman(3)) > 0 Then
       AutorunHeur = True
      'TutupFile hFile ' nutupnya klo TRUE saja
     'Call RtlZeroMemory(StrPtr(exetext), Len(exetext))

      Exit Function
    Else
       AutorunHeur = False
    End If
End If

keluar:
AutorunHeur = False
'Call RtlZeroMemory(StrPtr(exetext), Len(exetext))
End Function
Private Function BATCompHeur(sFile As String, hFile As Long) As Boolean

Dim Alman(10)     As String
'Dim IsiFile      As String
Alman(0) = "CopyFile" ' :: TEKS 1
Alman(1) = "scriptfullname" ':: TEKS 2
Alman(2) = "FileSystemObject"
Alman(3) = "CreateTextFile"
Alman(4) = "RegWrite"
Alman(5) = "attrib"
Alman(6) = "outlook.application"
Alman(7) = "FileSystemObject"
Alman(8) = "Scripting.FileSystemObject"
Alman(9) = "tskill"
Alman(10) = "disable"
BATCompHeur = False
On Error GoTo keluar
 If FileLen(sFile) < 100 Then GoTo keluar
 If FileLen(sFile) > 600000 Then GoTo keluar
exeText = Fileteks(sFile)

  If UCase$(Right$(sFile, 4)) = ".BAT" Or UCase$(Right$(sFile, 4)) = ".VBS" Then

    If InStr(exeText, Alman(0)) > 1 Or InStr(exeText, Alman(1)) > 1 Or InStr(exeText, Alman(2)) > 1 Or InStr(exeText, Alman(3)) > 1 Or InStr(exeText, Alman(4)) > 1 Or InStr(exeText, Alman(5)) > 1 Or InStr(exeText, Alman(6)) > 1 Or InStr(exeText, Alman(7)) > 1 Or InStr(exeText, Alman(8)) > 1 Or InStr(exeText, Alman(9)) > 1 Or InStr(exeText, Alman(10)) > 1 Then
       BATCompHeur = True
     ' TutupFile hFile ' nutupnya klo TRUE saja
     Call RtlZeroMemory(StrPtr(exeText), Len(exeText))

      Exit Function
    Else
       BATCompHeur = False
    End If
End If

keluar:
BATCompHeur = False
Call RtlZeroMemory(StrPtr(exeText), Len(exeText))
End Function
Public Function FixRamnit(sFile As String) As Boolean
On Error GoTo err
Dim strText As String 'isi file
Dim sztext As Long 'panjang byte dari text yang akan dibaca

Open sFile For Binary As #1
strText = Space$(LOF(1))
    Get #1, , strText 'ambil isi file ditampung ke strText
Close #1

If FrmRTP.TxtRamnit.Text <> "" Then 'ambil isi file dari banyak byte yang dibaca dimulai dari kata yang ditentukan <CASE SENSITIVE>
 
   Open "C:\fix_ramnit.txt" For Binary As #2
  strText = Left$(strText, InStr(strText, FrmRTP.TxtRamnit.Text) - 1)
   Put #2, , strText
   Close #2
  
  CopiFile "C:\fix_ramnit.txt", sFile, True
  FixRamnit = 1
 ' HapusFile "C:\fix_ramnit.txt"
    'Text3.Text = strText
Else 'ambil semua isi file
    'Text3.Text = strText
    FixRamnit = 0
End If

Exit Function
err:
'MsgBox err.Description, vbExclamation, "Info"
FixRamnit = 0
End Function

Private Function Ramnit_h(sFile As String, hFile As Long) As Boolean
On Error GoTo keluar
fSize = FileLen(sFile)
If fSize = 0 Then GoTo keluar
If fSize < 100000 Then GoTo keluar
If fSize > 1500000 Then GoTo keluar
If UCase$(Right$(sFile, 4)) = ".HTM" Or UCase$(Right$(sFile, 4)) = ".HTML" Then GoTo cek
cek:
Open sFile For Binary As #1
strBin = Space$(LOF(1))
Get #1, , strBin
Close #1

instrSign = InStr(strBin, strSplit(Sign))
'instrSign mewakili size File asli sebelum terinfeksi
If instrSign > 0 Then
    
    Ramnit_h = True
Else
   
    Ramnit_h = False
End If

Exit Function

keluar:
Ramnit_h = False
End Function
Function EncHex(strT As String) As String
'ini buat terjemahkan jadi Hex aja... :D
For i = 1 To Len(strT)
    EncHex = EncHex + "x" + "&H" + Hex$(Asc(Mid$(strT, i, 1)))
Next i
End Function

Function strSplit(strH As String) As String
Dim sSplit() As String
sSplit = Split(Sign, "x")

For i = 0 To UBound(sSplit)
    strSplit = strSplit + Chr$(sSplit(i))
Next i

End Function

'Fungsi ini untuk membaca file teks
Function Fileteks(Where As String) As String
Dim BinTeks, Temp As String
On Error Resume Next

Open Where For Input As #6
On Error Resume Next
Do While Not (EOF(6))
Input #6, Temp
BinTeks = BinTeks & Temp
Loop
Close #6
Fileteks = BinTeks

End Function


Private Function CekShortcut(sFile As String, hFile As Long) As Boolean
On Error GoTo errr
Dim Target As String
Dim sukuran As String
sukuran = Format$(GetSizeFile(hFile), "#,#")
CekShortcut = False
 'If UCase(Right(sFile, 4)) = ".ZIP" Or UCase(Right(sFile, 4)) = ".RAR" Then GoTo errr
 If GetDriveType(Left(sFile, 3)) = 2 Then
   If UCase(Right(sFile, 4)) = ".LNK" Then
      Target = GetTargetLink(sFile, True)
   'hFile = fileee.VbOpenFile(Path, FOR_BINARY_ACCESS_READ, LOCK_NONE)
      If IsFileX(CStr(Target)) = True Then
      ' If CocokanDataBaseRTP(Target) = True Then
        CekShortcut = True
         If ApaPengecualianFile(sFile, JumFileExcep) = True Then Exit Function
         AddInfoToList FrmRTP.lvRTP, f_bahasa(1) & " Mal-Shourtcut", sFile, sukuran, "Shortcut Loader", 4, 18
        LogPrint "==>  " + e_bahasa(0) + ": " + f_bahasa(1) & " Mal-Shourtcut" + "|-|" + e_bahasa(1) + ": " + sFile + "|-|" + "Shortcut Loader" + ": " + sukuran + "|-|" + e_bahasa(3) + ": " + f_bahasa(2)
         Else
        CekShortcut = False
     Exit Function
      End If
    End If
  End If
 Call KosongkanMemory(StrPtr(Target), Len(Target))
  
errr:
CekShortcut = True
Exit Function
End Function
Private Function CekShortcutRTP(sFile As String, hFile As Long) As Boolean
On Error GoTo errr
Dim Target As String
Dim sukuran As String
sukuran = Format$(GetSizeFile(hFile), "#,#")
'CekShortcutRTP = False
 'If GetDriveType(left$(sFile, 3)) = 2 Then
' If ucase$(Right$(sFile, 4)) = ".ZIP" Or ucase$(Right$(sFile, 4)) = ".RAR" Then GoTo errr
   If UCase$(Right$(sFile, 4)) = ".LNK" Then
   If GetFileName(sFile) = "Sample Music.lnk" Then GoTo errr
   If GetFileName(sFile) = "Windows Catalog.lnk" Then GoTo errr
   If GetFileName(sFile) = "Sample Pictures.lnk" Then GoTo errr

      Target = GetTargetLink(sFile, True)
   'hFile = fileee.VbOpenFile(Path, FOR_BINARY_ACCESS_READ, LOCK_NONE)
      If IsFileX(CStr(Target)) = True Then
      ' If CocokanDataBaseRTP(Target) = True Then
        CekShortcutRTP = True
         'If ApaPengecualianFile(sFile, JumFileExcep) = True Then Exit Function
        ' AddInfoToList frmRTP.lvRTP, f_bahasa(1) & " Mal-Shourtcut", sFile, sukuran, "Shortcut Loader", 4, 18
         'MsgBox sFile
         Else
         
        CekShortcutRTP = False
    Exit Function
      End If
    End If
  
 Call KosongkanMemory(StrPtr(Target), Len(Target))
 
errr:
CekShortcutRTP = True

End Function
' --- Fungsi Akumulasi Cek Heuristic
Public Function CekWithHeuristic(sFile As String, hFile As Long) As Boolean
Dim sukuran As String
sukuran = Format$(GetSizeFile(hFile), "#,#")
With frmMain
    If IsArrs(sFile, hFile) = True Then
        ' klo pngecualian keluar
       If ApaPengecualianFile(sFile, JumFileExcep) = True Then Exit Function
          AddInfoToList .lvMalware, f_bahasa(1) & " ArrS", sFile, sukuran, f_bahasa(2), 1, 18
        LogPrint "==>  " + e_bahasa(0) + ": " + f_bahasa(1) & " ArrS" + "|-|" + e_bahasa(1) + ": " + sFile + "|-|" + e_bahasa(2) + ": " + sukuran + "|-|" + e_bahasa(3) + ": " + f_bahasa(2)
        GoTo LBL_INFO
    
    ElseIf Ramnit_h(sFile, hFile) = True Then
    If ApaPengecualianFile(sFile, JumFileExcep) = True Then Exit Function
         AddInfoToList .lvMalware, "Win32/Ramnit.A", sFile, sukuran, f_bahasa(6), 0, 18
         LogPrint "==>  " + e_bahasa(0) + ": " + "Win32/Ramnit.A" + "|-|" + e_bahasa(1) + ": " + sFile + "|-|" + e_bahasa(2) + ": " + sukuran + "|-|" + e_bahasa(3) + ": " + f_bahasa(6)
          GoTo LBL_INFO
          
    ElseIf IsDoubleExt(sFile, hFile) = True Then 'doblel
      If ApaPengecualianFile(sFile, JumFileExcep) = True Then Exit Function
         AddInfoToList .lvMalware, " Double Extension", sFile, sukuran, f_bahasa(2), 4, 18
        LogPrint "==>  " + e_bahasa(0) + ": " + f_bahasa(1) & " Double Extension" + "|-|" + e_bahasa(1) + ": " + spath + "|-|" + e_bahasa(2) + ": " + sukuran + "|-|" + e_bahasa(3) + ": " + f_bahasa(2)
        GoTo LBL_INFO
        
    ElseIf AutorunHeur(sFile, hFile) = True Then
    If ApaPengecualianFile(sFile, JumFileExcep) = True Then Exit Function
      AddInfoToList .lvMalware, "Heur-W32.Sality.inf", sFile, sukuran, f_bahasa(2), 4, 18
        LogPrint "==>  " + e_bahasa(0) + ": " + f_bahasa(1) & " Heur-W32.Sality.inf" + "|-|" + e_bahasa(1) + ": " + sFile + "|-|" + e_bahasa(2) + ": " + sukuran + "|-|" + e_bahasa(3) + ": " + f_bahasa(2)
       GoTo LBL_INFO
       
    ElseIf CekVBS(sFile, hFile) = True Then
       ' klo pngecualian keluar
       If ApaPengecualianFile(sFile, JumFileExcep) = True Then Exit Function
          'sUkuran = Format$(GetSizeFile(hFile), "#,#")
          AddInfoToList .lvMalware, " Mal-Sript", sFile, sukuran, f_bahasa(2), 3, 18
          LogPrint "==>  " + e_bahasa(0) + ": " + f_bahasa(1) & " Mal-Sript" + "|-|" + e_bahasa(1) + ": " + sFile + "|-|" + e_bahasa(2) + ": " + sukuran + "|-|" + e_bahasa(3) + ": " + f_bahasa(2)
          GoTo LBL_INFO

    ElseIf AutorunHeur(sFile, hFile) = True Then
    If ApaPengecualianFile(sFile, JumFileExcep) = True Then Exit Function
         AddInfoToList .lvMalware, "W32.Sality-Heur", sFile, sukuran, f_bahasa(2), 4, 18
        LogPrint "==>  " + e_bahasa(0) + ": " + f_bahasa(1) & " W32.Sality-Heur" + "|-|" + e_bahasa(1) + ": " + sFile + "|-|" + e_bahasa(2) + ": " + sukuran + "|-|" + e_bahasa(3) + ": " + f_bahasa(22)
         GoTo LBL_INFO
  
    ElseIf CekShortcut(sFile, hFile) = False Then
    If ApaPengecualianFile(sFile, JumFileExcep) = True Then Exit Function
         AddInfoToList .lvMalware, " JUNK-Shourtcut", sFile, sukuran, f_bahasa(22), 4, 18
        LogPrint "==>  " + e_bahasa(0) + ": " + f_bahasa(1) & " JUNK-Shourtcut" + "|-|" + e_bahasa(1) + ": " + sFile + "|-|" + e_bahasa(2) + ": " + sukuran + "|-|" + e_bahasa(3) + ": " + f_bahasa(22)
         GoTo LBL_INFO
         
       ElseIf CekDoubleExtension(sFile) = True Then
       If ApaPengecualianFile(sFile, JumFileExcep) = True Then Exit Function
          AddInfoToList .lvMalware, "Fake Extension", sFile, sukuran, "Malware File", 4, 18
       LogPrint "==>  " + e_bahasa(0) + ": " + f_bahasa(1) & " Fake Extension" + "|-|" + e_bahasa(1) + ": " + sFile + "|-|" + e_bahasa(2) + ": " + sukuran + "|-|" + e_bahasa(3) + ": " + f_bahasa(22)
       GoTo LBL_INFO
    
         
   ElseIf HTT_Heur(sFile, hFile) = True Then
   If ApaPengecualianFile(sFile, JumFileExcep) = True Then Exit Function
       AddInfoToList .lvMalware, " HTT-Heuristic", sFile, sukuran, f_bahasa(2), 4, 18
        LogPrint "==>  " + e_bahasa(0) + ": " + f_bahasa(1) & " HTT-Heuristic" + "|-|" + e_bahasa(1) + ": " + sFile + "|-|" + e_bahasa(2) + ": " + sukuran + "|-|" + e_bahasa(3) + ": " + f_bahasa(2)
        GoTo LBL_INFO
        
   ElseIf INIHeur(sFile, hFile) = True Then
  If ApaPengecualianFile(sFile, JumFileExcep) = True Then Exit Function
   AddInfoToList .lvMalware, "  INI-Heuristic", sFile, sukuran, f_bahasa(2), 4, 18
            LogPrint "==>  " + e_bahasa(0) + ": " + f_bahasa(1) & " INI-Heuristic" + "|-|" + e_bahasa(1) + ": " + sFile + "|-|" + e_bahasa(2) + ": " + sukuran + "|-|" + e_bahasa(3) + ": " + f_bahasa(2)
      GoTo LBL_INFO
      
    ElseIf EICARHeur(sFile, hFile) = True Then
    If ApaPengecualianFile(sFile, JumFileExcep) = True Then Exit Function
         AddInfoToList .lvMalware, " EICAR-AV-Test", sFile, sukuran, f_bahasa(2), 4, 18
        LogPrint "==>  " + e_bahasa(0) + ": " + f_bahasa(1) & " EICAR-AV-Test" + "|-|" + e_bahasa(1) + ": " + sFile + "|-|" + e_bahasa(2) + ": " + sukuran + "|-|" + e_bahasa(3) + ": " + f_bahasa(2)
    GoTo LBL_INFO
    '
       
     ElseIf cekSEC(sFile) = True Then
     If ApaPengecualianFile(sFile, JumFileExcep) = True Then Exit Function
         AddInfoToList .lvMalware, "Sality.101", sFile, sukuran, f_bahasa(2), 0, 18
        LogPrint "==>  " + e_bahasa(0) + ": " + f_bahasa(1) & "Sality.101" + "|-|" + e_bahasa(1) + ": " + sFile + "|-|" + e_bahasa(2) + ": " + sukuran + "|-|" + e_bahasa(3) + ": " + f_bahasa(2)
     GoTo LBL_INFO
       
     ElseIf ScanLNK(sFile) = True Then
     If ApaPengecualianFile(sFile, JumFileExcep) = True Then Exit Function
         AddInfoToList .lvMalware, "W32\Ramnit.A|script", sFile, sukuran, f_bahasa(2), 0, 18
        LogPrint "==>  " + e_bahasa(0) + ": " + f_bahasa(1) & "W32\Ramnit.A|script" + "|-|" + e_bahasa(1) + ": " + sFile + "|-|" + e_bahasa(2) + ": " + sukuran + "|-|" + e_bahasa(3) + ": " + f_bahasa(2)
    GoTo LBL_INFO
    
    ElseIf Cekemail(sFile, hFile) = True Then
    If ApaPengecualianFile(sFile, JumFileExcep) = True Then Exit Function
          AddInfoToList .lvMalware, "Email Worm", sFile, sukuran, "Worm", 0, 18
        LogPrint "==>  " + e_bahasa(0) + ": " + f_bahasa(1) & "Email Worm" + "|-|" + e_bahasa(1) + ": " + sFile + "|-|" + e_bahasa(2) + ": " + sukuran + "|-|" + e_bahasa(3) + ": " + f_bahasa(2)
    GoTo LBL_INFO
       
     ElseIf cekhta(sFile, hFile) = True Then
    If ApaPengecualianFile(sFile, JumFileExcep) = True Then Exit Function
          AddInfoToList .lvMalware, "HTA_Alice", sFile, sukuran, "Worm", 0, 18
         LogPrint "==>  " + e_bahasa(0) + ": " + f_bahasa(1) & "HTA_Alice" + "|-|" + e_bahasa(1) + ": " + sFile + "|-|" + e_bahasa(2) + ": " + sukuran + "|-|" + e_bahasa(3) + ": " + f_bahasa(2)
    GoTo LBL_INFO
       
     ElseIf cekhtmlramnit(sFile, hFile) = True Then
    If ApaPengecualianFile(sFile, JumFileExcep) = True Then Exit Function
          AddInfoToList .lvMalware, "VBS:Agent.MD", sFile, sukuran, "Malware File", 0, 18
         LogPrint "==>  " + e_bahasa(0) + ": " + f_bahasa(1) & "VBS:Agent.MD" + "|-|" + e_bahasa(1) + ": " + sFile + "|-|" + e_bahasa(2) + ": " + sukuran + "|-|" + e_bahasa(3) + ": " + f_bahasa(2)
    GoTo LBL_INFO
           
    ElseIf cekUPX(sFile) = True Then
       If ApaPengecualianFile(sFile, JumFileExcep) = True Then Exit Function
          AddInfoToList .lvMalware, "Detect With FSS", sFile, sukuran, "Malware File", 4, 18
       LogPrint "==>  " + e_bahasa(0) + ": " + f_bahasa(1) & "Detect With FSS" + "|-|" + e_bahasa(1) + ": " + sFile + "|-|" + e_bahasa(2) + ": " + sukuran + "|-|" + e_bahasa(3) + ": " + f_bahasa(2)
       GoTo LBL_INFO

  ElseIf CheckIconexe(sFile, hFile) = True Then
       ' klo pngecualian keluar
       If ApaPengecualianFile(sFile, JumFileExcep) = True Then Exit Function
           'sUkuran = Format$(GetSizeFile(hFile), "#,#")
           AddInfoToList .lvMalware, " Icon Detection-[" & NumZ & "]", sFile, sukuran, f_bahasa(2), 4, 18
          ' AddInfoToList FrmLOGF.LisVLog, Format(Now, "dd mmmm yyyy" + ";" + "HH:MM:SS"), f_bahasa(1) & " Icon Detection", sFile, f_bahasa(2), 1, 18
        LogPrint "==>  " + e_bahasa(0) + ": " + f_bahasa(1) & " Icon Detection-[" & NumZ & "]" + "|-|" + e_bahasa(1) + ": " + sFile + "|-|" + e_bahasa(2) + ": " + sukuran + "|-|" + e_bahasa(3) + ": " + f_bahasa(2)
       GoTo LBL_INFO
    End If
End With
VirStatus = False
CekWithHeuristic = False
Exit Function

LBL_INFO:
' DoEvents
AutoLst frmMain.lvMalware
CekWithHeuristic = True
    VirStatus = True
    VirusFound = VirusFound + 1
    frmMain.lbMalware.Caption = Right$(VirusFound, 6) & " "
End Function

' --- Fungsi Akumulasi Cek Heuristic di RTP
Public Function CekWithHeuristicRTP(ByRef sFile As String, ByVal hFile As Long) As Boolean
Dim sukuran As String
sukuran = Format$(GetSizeFile(hFile), "#,#")
With FrmRTP
    If IsArrs(sFile, hFile) = True Then
        ' klo pngecualian keluar
       If ApaPengecualianFile(sFile, JumFileExcep) = True Then Exit Function
          AddInfoToList .lvRTP, f_bahasa(1) & " ArrS", sFile, sukuran, f_bahasa(2), 1, 18
        GoTo LBL_INFO
    
    ElseIf Ramnit_h(sFile, hFile) = True Then
    If ApaPengecualianFile(sFile, JumFileExcep) = True Then Exit Function
         AddInfoToList .lvRTP, "Win32/Ramnit.A", sFile, sukuran, f_bahasa(6), 0, 18
         ''LogPrint "==>  " + sFile + "   ==> Malware Name:" + "Disangka dengan  JUNK-Shourtcut" + "  Size:" + Ukuran
         GoTo LBL_INFO
         
    ElseIf IsDoubleExt(sFile, hFile) = True Then
      If ApaPengecualianFile(sFile, JumFileExcep) = True Then Exit Function
         AddInfoToList .lvRTP, " Double Extension", sFile, sukuran, f_bahasa(2), 4, 18
         ''LogPrint "==>  " + sfile + "   ==> Malware Name:" + "Disangka dengan Doubel extension" + "  Size:" + Ukuran
        GoTo LBL_INFO
        
    ElseIf CekVBS(sFile, hFile) = True Then
       ' klo pngecualian keluar
       If ApaPengecualianFile(sFile, JumFileExcep) = True Then Exit Function
          'sUkuran = Format$(GetSizeFile(hFile), "#,#")
          AddInfoToList .lvRTP, " Mal-Sript", sFile, sukuran, f_bahasa(2), 3, 18
       GoTo LBL_INFO
    
    ElseIf AutorunHeur(sFile, hFile) = True Then
    If ApaPengecualianFile(sFile, JumFileExcep) = True Then Exit Function
      AddInfoToList .lvRTP, " Heur-W32.Sality.inf", sFile, sukuran, f_bahasa(2), 4, 18
         ''LogPrint "==>  " + sFile + "   ==> Malware Name:" + "Disangka dengan  INI.Heuristic" + "  Size:" + Ukuran
       GoTo LBL_INFO
       
 ElseIf HTT_Heur(sFile, hFile) = True Then
   If ApaPengecualianFile(sFile, JumFileExcep) = True Then Exit Function
        AddInfoToList .lvRTP, " HTT-Heur", sFile, sukuran, f_bahasa(2), 4, 18
         ''LogPrint "==>  " + sFile + "   ==> Malware Name:" + "Disangka dengan  HTT_Heuristic" + "  Size:" + Ukuran
        GoTo LBL_INFO
        
   ElseIf INIHeur(sFile, hFile) = True Then
   If ApaPengecualianFile(sFile, JumFileExcep) = True Then Exit Function
         AddInfoToList .lvRTP, " INI-Heur", sFile, sukuran, f_bahasa(2), 4, 18
         ''LogPrint "==>  " + sFile + "   ==> Malware Name:" + "Disangka dengan  W32.Laecium [Heur]" + "  Size:" + Ukuran
       GoTo LBL_INFO
       
       ElseIf CekDoubleExtension(sFile) = True Then
       If ApaPengecualianFile(sFile, JumFileExcep) = True Then Exit Function
          AddInfoToList .lvRTP, "Fake Extension", sFile, sukuran, "Malware File", 4, 18
       LogPrint "==>  " + e_bahasa(0) + ": " + f_bahasa(1) & " Fake Extension" + "|-|" + e_bahasa(1) + ": " + sFile + "|-|" + e_bahasa(2) + ": " + sukuran + "|-|" + e_bahasa(3) + ": " + f_bahasa(22)
       GoTo LBL_INFO
         
    ElseIf CekShortcutRTP(sFile, hFile) = False Then
    If ApaPengecualianFile(sFile, JumFileExcep) = True Then Exit Function
         AddInfoToList .lvRTP, " JUNK-Shourtcut", sFile, sukuran, f_bahasa(22), 4, 18
         ''LogPrint "==>  " + sFile + "   ==> Malware Name:" + "Disangka dengan  JUNK-Shourtcut" + "  Size:" + Ukuran
         GoTo LBL_INFO
         
ElseIf EICARHeur(sFile, hFile) = True Then
    If ApaPengecualianFile(sFile, JumFileExcep) = True Then Exit Function
         AddInfoToList .lvRTP, " EICAR-AV-Test", sFile, sukuran, f_bahasa(2), 4, 18
         ''LogPrint "==>  " + sFile + "   ==> Malware Name:" + f_bahasa(1) & " INK-Loader" + "  Size:" + Ukuran
         GoTo LBL_INFO
         
     ElseIf cekSEC(sFile) = True Then
     If ApaPengecualianFile(sFile, JumFileExcep) = True Then Exit Function
         AddInfoToList .lvRTP, "Sality.101", sFile, sukuran, f_bahasa(2), 0, 18
       ' LogPrint "==>  " + e_bahasa(0) + ": " + f_bahasa(1) & "Sality.101" + "|-|" + e_bahasa(1) + ": " + sFile + "|-|" + e_bahasa(2) + ": " + sukuran + "|-|" + e_bahasa(3) + ": " + f_bahasa(2)
     GoTo LBL_INFO
       
     ElseIf ScanLNK(sFile) = True Then
     If ApaPengecualianFile(sFile, JumFileExcep) = True Then Exit Function
         AddInfoToList .lvRTP, "W32\Ramnit.A|script", sFile, sukuran, f_bahasa(2), 0, 18
        'LogPrint "==>  " + e_bahasa(0) + ": " + f_bahasa(1) & "W32\Ramnit.A|script" + "|-|" + e_bahasa(1) + ": " + sFile + "|-|" + e_bahasa(2) + ": " + sukuran + "|-|" + e_bahasa(3) + ": " + f_bahasa(2)
    GoTo LBL_INFO
    
    ElseIf Cekemail(sFile, hFile) = True Then
    If ApaPengecualianFile(sFile, JumFileExcep) = True Then Exit Function
          AddInfoToList .lvRTP, "Email Worm", sFile, sukuran, "Worm", 0, 18
        'LogPrint "==>  " + e_bahasa(0) + ": " + f_bahasa(1) & "Email Worm" + "|-|" + e_bahasa(1) + ": " + sFile + "|-|" + e_bahasa(2) + ": " + sukuran + "|-|" + e_bahasa(3) + ": " + f_bahasa(2)
    GoTo LBL_INFO
       
     ElseIf cekhta(sFile, hFile) = True Then
    If ApaPengecualianFile(sFile, JumFileExcep) = True Then Exit Function
          AddInfoToList .lvRTP, "HTA_Alice", sFile, sukuran, "Worm", 0, 18
         'LogPrint "==>  " + e_bahasa(0) + ": " + f_bahasa(1) & "HTA_Alice" + "|-|" + e_bahasa(1) + ": " + sFile + "|-|" + e_bahasa(2) + ": " + sukuran + "|-|" + e_bahasa(3) + ": " + f_bahasa(2)
    GoTo LBL_INFO
    
     ElseIf cekhtmlramnit(sFile, hFile) = True Then
    If ApaPengecualianFile(sFile, JumFileExcep) = True Then Exit Function
          AddInfoToList .lvRTP, "VBS:Agent.MD", sFile, sukuran, "Malware File", 40, 18
'         LogPrint "==>  " + e_bahasa(0) + ": " + f_bahasa(1) & "VBS:Agent.MD" + "|-|" + e_bahasa(1) + ": " + sFile + "|-|" + e_bahasa(2) + ": " + sukuran + "|-|" + e_bahasa(3) + ": " + f_bahasa(2)
    GoTo LBL_INFO
           
    ElseIf cekUPX(sFile) = True Then
       If ApaPengecualianFile(sFile, JumFileExcep) = True Then Exit Function
          AddInfoToList .lvRTP, "Detect With FSS", sFile, sukuran, "Malware File", 4, 18
       'LogPrint "==>  " + e_bahasa(0) + ": " + f_bahasa(1) & "Detect With FSS" + "|-|" + e_bahasa(1) + ": " + sFile + "|-|" + e_bahasa(2) + ": " + sukuran + "|-|" + e_bahasa(3) + ": " + f_bahasa(2)
       GoTo LBL_INFO

 ElseIf CheckIconexe(sFile, hFile) = True Then
       ' klo pngecualian keluar
     If ApaPengecualianFile(sFile, JumFileExcep) = True Then Exit Function
          AddInfoToList .lvRTP, " Icon Detection-[" & NumZ & "]", sFile, sukuran, f_bahasa(2), 4, 18
       GoTo LBL_INFO
    End If
    
    'End If
End With
CekWithHeuristicRTP = False
Exit Function

LBL_INFO:
 'DoEvents
 Call tampilkanRTP
 AutoLst FrmRTP.lvRTP
    CekWithHeuristicRTP = True
    'KunciFileYangDiRTP frmRTP.lvRTP
End Function
' Fungsi Mengenai segala sesuatu yang perlu di informasikan ke User (klo bisa informasi hanya untuk executable file exe,scr,com, tapi bgaiaman cara filter melalui charaterisnya)
' DLL- standar banyak yang bermuatan byte tambahan gak jelas gitu
Public Function CekInformation(ByRef sFile As String, hFile As Long) ' jika di RTP suspect sality ditampilkan
Dim nSizeTmp    As Long
Dim nSize       As Long
Dim TheExt      As String
Dim SZSize As String
SZSize = Format$(GetSizeFile(hFile), "#,#")


With frmMain
TheExt = UCase(Right(sFile, 3))
If GetDriveType(Left(sFile, 3)) = 2 Then ' jika di FD
   If TheExt = "DLL" Or TheExt = "EXE" Or TheExt = "SYS" Or TheExt = "OCX" Then
   Else
      InfoFound = InfoFound + 1 ' un proper ektensi di FD kita masukan ke Info
      nSize = GetSizeFile(hFile)
      .lbInfo.Caption = Right$(InfoFound, 6) & " "
      AddInfoToList .lvInfo, "Unproper PE Extension", sFile, Format$(nSize, "#,#"), "Perhatian : Waspadai sedikit file ini !", 0, 18
      Exit Function
   End If
End If
If GetRealSizePE = 0 Or IsPE32EXE = False Then Exit Function ' harus benar-benar EXE (bukan dll/sys)

'nSize = GetSizeFile(hFile)
   
    If nSalityGet <> "" Then ' ada kemungkinan sality dan varianya (Heur ini belum valid) jadi masuk Informasi
       VirusFound = VirusFound + 1
      ' .lbInfo.Caption =  Right$("" & InfoFound, 6) & " "
       AddInfoToList .lvMalware, nSalityGet, sFile, SZSize, f_bahasa(1) & " PE.Heuristic", 2, 8
            LogPrint "==>  " + e_bahasa(0) + ": " + nSalityGet + "|-|" + e_bahasa(1) + ": " + sFile + "|-|" + e_bahasa(2) + ": " + sukuran + "|-|" + e_bahasa(3) + ": " + f_bahasa(1) & " PE.Heuristic"
       VirStatus = True
    VirusFound = VirusFound + 1
    frmMain.lbMalware.Caption = Right$(VirusFound, 6) & " "
       nSalityGet = "" ' reset
    ElseIf nPEHeurGet <> "" Then ' pengecekan PE Heur
       VirusFound = VirusFound + 1
      ' .lbInfo.Caption = ": " & Right$("000000" & InfoFound, 6) & " " & d_bahasa(38)
       AddInfoToList .lvMalware, nPEHeurGet, sFile, SZSize, f_bahasa(1) & " PE.Heuristic", 2, 8
             LogPrint "==>  " + e_bahasa(0) + ": " + nPEHeurGet + "|-|" + e_bahasa(1) + ": " + sFile + "|-|" + e_bahasa(2) + ": " + sukuran + "|-|" + e_bahasa(3) + ": " + f_bahasa(1) & " PE.Heuristic"
      VirStatus = True
    VirusFound = VirusFound + 1
    frmMain.lbMalware.Caption = Right$(VirusFound, 6) & " "
       nPEHeurGet = "" ' reset
    
    End If
AutoLst frmMain.lvInfo
End With


End Function

' Kecurigaan terhdap virus2 masuk ke RTP
Public Function CekInformationRTP(ByRef sFile As String, ByVal hFile As Long) ' jika di RTP suspect sality ditampilkan
Dim SZSize As String
Dim nSizeTmp    As Long
Dim nSize       As Long

SZSize = Format$(GetSizeFile(hFile), "#,#")
With FrmRTP

If nSalityGet <> "" Then
   AddInfoToList .lvRTP, nSalityGet, sFile, SZSize, f_bahasa(1) & " PE.Heuristic", 2, 8
   nSalityGet = "" ' reset
ElseIf nPEHeurGet <> "" Then
   AddInfoToList .lvRTP, nPEHeurGet, sFile, SZSize, f_bahasa(1) & " PE.Heuristic", 2, 8
   nPEHeurGet = "" ' reset
End If
AutoLst .lvRTP
End With
Call tampilkanRTP
End Function


' Heur untuk cek autorun yang hidden ajh
Public Function CekAutorun(ByRef sRootAutorun As String) As Boolean
Dim lngItem As Long
Dim IsiAR   As String
With frmMain

If UCase$(GetFileName(sRootAutorun)) = "AUTORUN.INF" Then
   NAT = GetFileAttributes(StrPtr(sRootAutorun))
   If (NAT = 2 Or NAT = 34 Or NAT = 3 Or NAT = 6 Or NAT = 22 Or NAT = 18 Or NAT = 50 Or NAT = 19 Or NAT = 35) Then
      ' yang di true status hidden aj
   ' klo pngecualian keluar
   If ApaPengecualianFile(sRootAutorun, JumFileExcep) = True Then Exit Function
   
      AddInfoToList .lvMalware, "Suspected ! [Autorun]", sRootAutorun, Format$(FileLen(sRootAutorun), "#,#"), f_bahasa(2) & " Malware Runner", 1, 18
        LogPrint "==>  " + e_bahasa(0) + ": " + "Suspected ! [Autorun]" + "|-|" + e_bahasa(1) + ": " + sRootAutorun + "|-|" + e_bahasa(2) + ": " + Format$(FileLen(sRootAutorun), "#,#") + "|-|" + e_bahasa(3) + ": " + f_bahasa(2) & " Malware Runner"
      VirusFound = VirusFound + 1
      
      .lbMalware.Caption = Right$(VirusFound, 6) & " "
      
      CekAutorun = True
   Else
      IsiAR = ReadUnicodeFile(sRootAutorun)
      IsiAR = UCase$(Replace(IsiAR, Chr$(0), ""))
      If InStr(IsiAR, "WScript.exe") > 0 Then
         AddInfoToList .lvMalware, "Suspected ! [Autorun]", sRootAutorun, Format$(FileLen(sRootAutorun), "#,#"), f_bahasa(2) & " Malware Runner", 1, 18
        LogPrint "==>  " + e_bahasa(0) + ": " + "Suspected ! [Autorun]" + "|-|" + e_bahasa(1) + ": " + sRootAutorun + "|-|" + e_bahasa(2) + ": " + Format$(FileLen(sRootAutorun), "#,#") + "|-|" + e_bahasa(3) + ": " + f_bahasa(2) & " Malware Runner"
         VirusFound = VirusFound + 1
      
        .lbMalware.Caption = Right$(VirusFound, 6) & " "

         CekAutorun = True
      Else
         CekAutorun = False
      End If
   End If
Else
   CekAutorun = False
End If
AutoLst frmMain.lvMalware
End With

'frmScanWith.lbMalware.Caption =  Right$("" & VirusFound, 6) & " "
End Function


' Heur untuk cek *.lnk yang kemungkinan virus
' Target Dibaca
Public Function CeklnkFolder(ByRef sPathFile As String) As Boolean
Dim lnkString(5) As String ' semntara smplenya baru dua
Dim TheTarget2   As String
Dim nTurn        As Long
Dim MyHnd        As Long

lnkString(0) = Chr$(13) & ".com" ' suspect Lnk ke virus lain
lnkString(1) = "wscript.exe" ' ke VBS
lnkString(2) = "RÊCYCLÊR\"
lnkString(3) = "cmd.exe"
lnkString(4) = "RECYCLER"

TheTarget2 = GetFileName(GetTargetLink(sPathFile, False))
       If Len(TheTarget2) > 0 Then ' format LNK true
     If GetFileName(sPathFile) = "Command Prompt.lnk" Then GoTo metu

           For nTurn = 0 To 4
               If LCase(TheTarget2) = lnkString(nTurn) Then ' ada
                  ' klo pngecualian keluar
                  If ApaPengecualianFile(sPathFile, JumFileExcep) = True Then Exit Function
                  CeklnkFolder = True
                  AddInfoToList frmMain.lvMalware, "Suspect! [Mal-Shortcut]:" & nTurn, sPathFile, "N/A", f_bahasa(2) & " Malware Runner", 1, 18
        LogPrint "==>  " + e_bahasa(0) + ": " + "Suspect! [Mal-Shortcut]:" & nTurn + "|-|" + e_bahasa(1) + ": " + sPathFile + "|-|" + e_bahasa(2) + ": " + "N/A" + "|-|" + e_bahasa(3) + ": " + f_bahasa(2) & " Malware Runner"
                  GoTo lbl_cek_isi ' cek isi hanya pada kasus True saja (mengurangi false detek)
                End If
            Next
Exit Function
lbl_cek_isi:
           TheTarget2 = GetTargetLink(sPathFile, True)
           If GetDriveType(Left$(TheTarget2, 3)) <> 2 Then Exit Function ' ingat hanya di FD aj
           If Left$(TheTarget2, 3) = Left$(sPathFile, 3) Then ' artinya satu jalur
              MyHnd = GetHandleFile(TheTarget2)
              If MyHnd > 0 Then
                 AddInfoToList frmMain.lvMalware, f_bahasa(1) & " ArrS", TheTarget2, Format$(GetSizeFile(MyHnd), "#,#"), f_bahasa(2), 1, 18
        LogPrint "==>  " + e_bahasa(0) + ": " + f_bahasa(1) & " ArrS" + "|-|" + e_bahasa(1) + ": " + TheTarget2 + "|-|" + e_bahasa(2) + ": " + Format$(GetSizeFile(MyHnd), "#,#") + "|-|" + e_bahasa(3) + ": " + f_bahasa(2)
                 TutupFile MyHnd
                 KunciFile TheTarget2
                 VirusFound = VirusFound + 1 ' dapat doble
                 frmMain.lbMalware.Caption = Right$(VirusFound, 6) & " "
                 'frmScanWith.lbMalware.Caption =  Right$(& VirusFound, 6) & " "
                 Exit Function
              End If
           End If
           
       Else
           CeklnkFolder = False
           
       End If
 AutoLst frmMain.lvMalware
metu:
End Function


' Heur untuk cek *.lnk yang kemungkinan virus
' Target Dibaca - Untuk RTP
Public Function CeklnkFolderRTP(ByRef sPathFile As String) As Boolean
Dim lnkString(5) As String ' semntara smplenya baru dua
Dim TheTarget2    As String
Dim nTurn        As Long
Dim MyHnd        As Long

lnkString(0) = Chr$(13) & ".com" ' suspect Lnk ke virus lain
lnkString(1) = "wscript.exe" ' ke VBS
lnkString(2) = "RÊCYCLÊR\"
lnkString(3) = "cmd.exe"
lnkString(4) = "RECYCLER"

TheTarget2 = GetFileName(GetTargetLink(sPathFile, False))
       If Len(TheTarget2) > 0 Then ' format LNK true
          If GetFileName(sPathFile) = "Command Prompt.lnk" Then GoTo metu

           For nTurn = 0 To 4
               If LCase(TheTarget2) = lnkString(nTurn) Then ' ada
                  ' klo pngecualian keluar
                  If ApaPengecualianFile(sPathFile, JumFileExcep) = True Then Exit Function
                  CeklnkFolderRTP = True
                  AddInfoToList FrmRTP.lvRTP, "Suspect! [Mal-Shortcut]:" & nTurn, sPathFile, "N/A", f_bahasa(2) & " Malware Runner", 1, 18
                  Call tampilkanRTP
                  KunciFileYangDiRTP FrmRTP.lvRTP
                  GoTo lbl_cek_isi ' cek isi hanya pada kasus True saja (mengurangi false detek)
               End If
           Next
Exit Function
lbl_cek_isi:
           TheTarget2 = GetTargetLink(sPathFile, True)
           If GetDriveType(Left$(TheTarget2, 3)) <> 2 Then Exit Function ' ingat hanya di FD aj
           If Left$(TheTarget2, 3) = Left$(sPathFile, 3) Then ' artinya satu jalur
              MyHnd = GetHandleFile(TheTarget2)
              If MyHnd > 0 Then
                 AddInfoToList FrmRTP.lvRTP, f_bahasa(1) & " ArrS", TheTarget2, Format$(GetSizeFile(MyHnd), "#,#"), f_bahasa(2), 1, 18
                 Call tampilkanRTP
                 KunciFileYangDiRTP FrmRTP.lvRTP
                 TutupFile MyHnd
                 KunciFile TheTarget2
                 Exit Function
              End If
           End If
       Else
           CeklnkFolderRTP = False
       End If
       AutoLst FrmRTP.lvRTP
metu:
End Function
