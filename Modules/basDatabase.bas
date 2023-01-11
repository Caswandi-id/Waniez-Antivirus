Attribute VB_Name = "basDatabase"
' Module untuk penanganan akses Database
Public PathScan As String
Public JumlahVirusINT As Long
Public JumlahVirusOUT As Long
Public virusradar As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public Function BacaDatabase()
Dim sTemp       As String
Dim sTmp()      As String
Dim sTmp2()     As String
Dim SignBack    As String
Dim pisah       As String
Dim ResPath     As String
Dim sPath       As String
Dim iCount      As Integer
Dim ITemp       As Integer
Dim iTurn       As Byte

On Error Resume Next ' Redimensi dulu
ReDim JumlahVirus(15) As Long
ReDim JumlahVirusNonPE(15) As Long

JumVirus = 0 'init

ResPath = GetSpecFolder(WINDOWS_DIR)

SignBack = "_a.upx" ' untuk PE 0x.dat

' Baca DB PE
For iTurn = 0 To UBound(JumlahVirus) ' 0-15
    ' inisialisasi
    iCount = 0
    
    sPath = GetFilePath(App_FullPathW(False)) & "\upx\" & Hex$(iTurn) & SignBack
   ' spath = "E:\VBA\C.M.C PH#3\sign\" & Hex(iTurn) & SignBack
    sTemp = ReadDatabaseAV(sPath, 200)
    
    pisah = Chr$(13)
    
    
    If sTemp = "" Then GoTo LBL_GAWAT ' gagal baca

    sTmp() = Split(sTemp, pisah)
    ITemp = UBound(sTmp())  ' untuk jumlah virus
    For iCount = 1 To ITemp
        sTmp2() = Split(sTmp(iCount), "=")
        sMD5(iTurn, iCount) = Mid$(sTmp2(0), 2)
        sNamaVirus(iTurn, iCount) = sTmp2(1)
    Next
    JumlahVirus(iTurn) = ITemp     ' jumlah virus pada dbx
    JumVirus = JumVirus + (ITemp)  ' jumlah virus pada db0-15
LEWAT:
Next
'Erase sTmp()
iTurn = 0
SignBack = "_x.upx" ' untuk non PE 0z.dat


' Baca DB non PE
For iTurn = 0 To UBound(JumlahVirus) ' 0-15
    ' inisialisasi
    iCount = 0
    'ENKRIP NON PE
    'spath = "E:\VBA\C.M.C PH#3\signx\no enkrip\" & Hex(iTurn) & SignBack
    'EnkripDB spath, 9, "E:\VBA\C.M.C PH#3\signx\" & Hex(iTurn) & SignBack
    'GoTo LEWAT2
    
    sPath = GetFilePath(App_FullPathW(False)) & "\upx\" & Hex$(iTurn) & SignBack
    'spath = "E:\VBA\C.M.C PH#3\signx\" & Hex(iTurn) & SignBack
    sTemp = ReadDatabaseAV(sPath, 200)
    
    pisah = Chr$(13)
    
    
    If sTemp = "" Then GoTo LBL_GAWAT ' gagal baca

    sTmp() = Split(sTemp, pisah)
    ITemp = UBound(sTmp())  ' untuk jumlah virus
    For iCount = 1 To ITemp
        sTmp2() = Split(sTmp(iCount), "=")
        sMD5nonPE(iTurn, iCount) = Mid$(sTmp2(0), 2)
        sNamaVirusnonPE(iTurn, iCount) = sTmp2(1)
    Next
    JumlahVirusNonPE(iTurn) = ITemp     ' jumlah virus pada dbx
    JumVirus = JumVirus + (ITemp)  ' jumlah virus pada db0-15
LEWAT2:
Next
'Erase sTmp()
frmMain.LbExDB.Caption = "Exstrnal Database " & CStr(JumVirus) ' + JumlahVirusM31 + 125
FrmAbout.lbWorm.Caption = ": " & CStr(JumVirus)
Exit Function
LBL_GAWAT: ' klo ada yang gagal baca
   MsgBox j_bahasa(28) & " ( " & Hex$(iTurn) & SignBack & " )", vbCritical
    'End
'frmMain.lbktotDB.Caption = "Signature database: " & CStr(JumVirus) + JumlahVirusINT + JumlahVirusOUT & " Virus + Heuristic": DoEvents
End Function

Public Function SelectDB(ByRef ceksum As String) As Long
Select Case Left$(ceksum, 1)
    Case "1": SelectDB = 1
    Case "2": SelectDB = 2
    Case "3": SelectDB = 3
    Case "4": SelectDB = 4
    Case "5": SelectDB = 5
    Case "6": SelectDB = 6
    Case "7": SelectDB = 7
    Case "8": SelectDB = 8
    Case "9": SelectDB = 9
    Case "A": SelectDB = 10
    Case "B": SelectDB = 11
    Case "C": SelectDB = 12
    Case "D": SelectDB = 13
    Case "E": SelectDB = 14
    Case "F": SelectDB = 15
    Case "0": SelectDB = 0
End Select
End Function

' Disini pusat pencocokan baik virus, worm, dan informasi
' RTP tidak masuk sini
Public Function CocokanDataBase(ByRef sPath As String) As Boolean ' Memakai turboloop based hex [perulanganya DB di hemat]
Dim iCount       As Integer
Dim Ukuran       As String
Dim CeksumFile   As String
Dim CeksumVirus  As String
Dim nDataBase    As Byte
Dim RetPE        As Long
Dim TmpHGlobal   As Long
Dim RetVirus     As String
Dim RetInfect As String
Dim RetHeur      As Boolean
Dim hFilePE As String
Dim CounTx As Long
Dim Splitter() As String
Dim A As String
On Error GoTo LBL_AKHIR
'IsCocokMeiPattern (sPath)
TmpHGlobal = GetHandleFile(sPath)

With frmMain

If TmpHGlobal <= 0 Then GoTo LBL_AKHIR
'fSize = FileLen(sPath)
'If fSize = 0 Then Exit Function
'Cek *.lnk dulu
If UCase$(Right$(sPath, 4)) = ".LNK" Then
    If CeklnkFolder(sPath) = True Then
           VirusFound = VirusFound + 1
       VirStatus = True
       .lbMalware.Caption = Right$("" & VirusFound, 6) & " "
       'frmScanWith.lbMalware.Caption = ": " & Right$("00000000" & VirusFound, 6) & " " & d_bahasa(38)
       GoTo LBL_AKHIR ' akhiri aj
    End If
End If

RetPE = IsValidPE32(TmpHGlobal) ' fungsi balik IsValidPE32 adalah AddresOfNewHeader

If RetPE > 64 Then ' PE - Ternyata DLL kan bisa diinjek juga gitu
    ' Cek dengan Database Virus dulu jika file PE 32 exe
    RetVirus = GetDataEP(TmpHGlobal, 40, RetPE)
    If RetVirus <> "" Then
       ' klo pengecualian keluar
       If ApaPengecualianFile(sPath, JumFileExcep) = True Then GoTo LBL_AKHIR
          VirusFound = VirusFound + 1
          VirStatus = True
          Ukuran = Format$(GetSizeFile(TmpHGlobal), "#,#") ' ukuran dalam string
          .lbMalware.Caption = Right$(VirusFound, 6) & " "
           'frmScanWith.lbMalware.Caption = ": " & Right$("00000000" & VirusFound, 6) & " "
          If Left$(RetVirus, 3) = "PW:" Then ' artinya hanya WormPoli
             AddInfoToList .lvMalware, Mid$(RetVirus, 4), sPath, Ukuran, f_bahasa(7), 0, 18
             LogPrint "==>  " + e_bahasa(0) + ": " + Mid(RetVirus, 4) + "|-|" + e_bahasa(1) + ": " + sPath + "|-|" + e_bahasa(2) + ": " + Ukuran + "|-|" + e_bahasa(3) + ": " + f_bahasa(7)
          Else
             AddInfoToList .lvMalware, RetVirus, sPath, Ukuran, f_bahasa(6), 2, 18
             LogPrint "==>  " + e_bahasa(0) + ": " + RetVirus + "|-|" + e_bahasa(1) + ": " + sPath + "|-|" + e_bahasa(2) + ": " + Ukuran + "|-|" + e_bahasa(3) + ": " + f_bahasa(7)
          End If
       GoTo LBL_AKHIR ' akhiri aj
    End If
  End If

'fix by.heru
'If GetPE3264Type(hFile) > 0 Then
 RetVirus = CekEntryPoint(TmpHGlobal, sPath)
If RetVirus <> "" Then
  If ApaPengecualianFile(sPath, JumFileExcep) = True Then GoTo LBL_AKHIR
 VirusFound = VirusFound + 1
 VirStatus = True
 Ukuran = Format$(GetSizeFile(TmpHGlobal), "#,#") ' ukuran dalam string
 .lbMalware.Caption = Right$(VirusFound, 6) & " "
 'frmScanWith.lbMalware.Caption = ": " & Right$("00000000" & VirusFound, 6) & " " & d_bahasa(38)
AddInfoToList .lvMalware, RetVirus, sPath, Ukuran, f_bahasa(6), 2, 18
LogPrint "==>  " + e_bahasa(0) + ": " + RetVirus + "|-|" + e_bahasa(1) + ": " + sPath + "|-|" + e_bahasa(2) + ": " + Ukuran + "|-|" + e_bahasa(3) + ": " + f_bahasa(6)

GoTo LBL_AKHIR '
End If
'End If
'If fSize > 700000 Then GoTo LBL_AKHIR
If RetPE > 0 Then ' tergolong PE

       CeksumFile = MYCeksum(sPath, TmpHGlobal)
       
       If CeksumFile = String$(Len(CeksumFile), "0") Then
          CeksumFile = MYCeksumCadangan(sPath, TmpHGlobal)
       End If

       nDataBase = SelectDB(CeksumFile)
           
       'Ceksumer PE
       For iCount = 1 To JumlahVirus(nDataBase)
         If sMD5(nDataBase, iCount) = CeksumFile Then  ' jika virus didapet
           ' klo pngecualian keluar
           If ApaPengecualianFile(sPath, JumFileExcep) = True Then GoTo LBL_AKHIR
           VirusFound = VirusFound + 1
           VirStatus = True
           Ukuran = Format$(GetSizeFile(TmpHGlobal), "#,#") ' ukuran dalam string
           AddInfoToList .lvMalware, sNamaVirus(nDataBase, iCount), sPath, Ukuran, f_bahasa(7), 0, 18
             LogPrint "==>  " + e_bahasa(0) + ": " + sNamaVirus(nDataBase, iCount) + "|-|" + e_bahasa(1) + ": " + sPath + "|-|" + e_bahasa(2) + ": " + Ukuran + "|-|" + e_bahasa(3) + ": " + f_bahasa(7)
           .lbMalware.Caption = Right$(VirusFound, 6) & " "
            'frmScanWith.lbMalware.Caption = ": " & Right$("00000000" & VirusFound, 6) & " " & d_bahasa(38)
            GoTo LBL_AKHIR
         End If
         'DoEvents
      Next
  
Else

   CeksumFile = MYCeksum(sPath, TmpHGlobal)
       
   If CeksumFile = String$(Len(CeksumFile), "0") Then
      CeksumFile = MYCeksumCadangan(sPath, TmpHGlobal)
   End If
   
   nDataBase = SelectDB(CeksumFile)
    
   ' Ceksumer nonPE
    For iCount = 1 To JumlahVirusNonPE(nDataBase)
        If sMD5nonPE(nDataBase, iCount) = CeksumFile Then  ' jika virus didapet
           ' klo pngecualian keluar
           If ApaPengecualianFile(sPath, JumFileExcep) = True Then GoTo LBL_AKHIR
           VirusFound = VirusFound + 1
           VirStatus = True
           Ukuran = Format$(GetSizeFile(TmpHGlobal), "#,#") ' ukuran dalam string
           AddInfoToList .lvMalware, sNamaVirusnonPE(nDataBase, iCount), sPath, Ukuran, f_bahasa(7), 0, 18
             LogPrint "==>  " + e_bahasa(0) + ": " + sNamaVirusnonPE(nDataBase, iCount) + "|-|" + e_bahasa(1) + ": " + sPath + "|-|" + e_bahasa(2) + ": " + Ukuran + "|-|" + e_bahasa(3) + ": " + f_bahasa(7)
           .lbMalware.Caption = Right$(VirusFound, 6) & " "
           'frmScanWith.lbMalware.Caption = ": " & Right$("00000000" & VirusFound, 6) & " " & d_bahasa(38)
            GoTo LBL_AKHIR
        End If
        'DoEvents
    Next
End If

'If frmMain.lvM31.ListItems.Count > 0 Then
'If ValidFile(App.Path & "\USER.DAT") = True Then 'ReadUDB lvM31
   
 CeksumFile = MYCeksum(sPath, TmpHGlobal)
       
       If CeksumFile = String$(Len(CeksumFile), "0") Then
       CeksumFile = MYCeksumCadangan(sPath, TmpHGlobal)
       End If
              
 For CounTx = 0 To (2 + BanyakUDB) ' 125 banyaknya Db internal dari 0
    On Error GoTo LBL_AKHIR
    Splitter = Split(InternalDb(CounTx), "|")
    If Splitter(0) = CeksumFile Then 'MeiPattern(spath, 202) Then
       'VirusFound = VirusFound + 1
       Ukuran = Format$(GetSizeFile(TmpHGlobal), "#,#")
       If Splitter(1) = "Suspected With ArrS" Then
       If ApaPengecualianFile(sPath, JumFileExcep) = True Then GoTo LBL_AKHIR
        AddInfoToList frmMain.lvMalware, Splitter(1), sPath, Ukuran, f_bahasa(24), 2, 18
        LogPrint "==>  " + e_bahasa(0) + ": " + Splitter(1) + "|-|" + e_bahasa(1) + ": " + sPath + "|-|" + e_bahasa(2) + ": " + Ukuran + "|-|" + e_bahasa(3) + ": " + f_bahasa(24)
        .lbMalware.Caption = Right$(VirusFound, 6) & " "
        VirusFound = VirusFound + 1
       VirStatus = True
        GoTo LBL_AKHIR
       Else
       VirusFound = VirusFound + 1
       VirStatus = True
       If ApaPengecualianFile(sPath, JumFileExcep) = True Then GoTo LBL_AKHIR
        AddInfoToList frmMain.lvMalware, Splitter(1), sPath, Ukuran, f_bahasa(24), 2, 18
        LogPrint "==>  " + e_bahasa(0) + ": " + Splitter(1) + "|-|" + e_bahasa(1) + ": " + sPath + "|-|" + e_bahasa(2) + ": " + Ukuran + "|-|" + e_bahasa(3) + ": " + f_bahasa(24)
       .lbMalware.Caption = Right$(VirusFound, 6) & " "
       GoTo LBL_AKHIR
       End If
      
    End If
Next
'End If
  
VirStatus = False ' set false

If FrmConfig.ck2.Value = 1 Then RetHeur = CekWithHeuristic(sPath, TmpHGlobal)


' jika heuristic tidak nemu, option di set dan PE valid
If RetHeur = False And RetPE > 0 Then
   
   Call CekInformation(sPath, TmpHGlobal) ' tidak diakai RTP
   
Else
   GoTo LBL_AKHIR
End If

End With

TutupFile TmpHGlobal ' jaga-jaga aja

Exit Function
LBL_AKHIR: ' kalo error/udah dapet lngsung akhiri pemindaian file saat ini
    TutupFile TmpHGlobal
    nSalityGet = "" 'biar gak bahaya
    AutoLst frmMain.lvMalware
End Function


' Menampilkan Daftar Virus yang ada (ingat kondisi dari tab pemicu ini (Wan'iez info) ini jangan sampe terbuka dulu sblum db dibaca)
Public Sub ListVirus(OutObjek As ListBox)
Dim ITemp       As Integer
Dim nDB         As Byte
Dim lngItem     As Integer

On Error Resume Next
With OutObjek
    .Clear

For nDB = 0 To 15 ' jumlah db=16 --> PE
    For ITemp = 1 To JumlahVirus(nDB)
        lngItem = lngItem + 1
        .AddItem sNamaVirus(nDB, ITemp)
    Next
    ITemp = 0 ' reset
Next
nDB = 0
For nDB = 0 To 15 ' jumlah db=16 --> Non PE
    For ITemp = 1 To JumlahVirusNonPE(nDB)
        lngItem = lngItem + 1
        .AddItem sNamaVirusnonPE(nDB, ITemp)
    Next
    ITemp = 0 ' reset
Next

For ITemp = 1 To JumVirus
   .List(ITemp - 1) = Right$("00000" & CStr(ITemp), 5) & " - " & .List(ITemp - 1)
Next

End With
End Sub
Public Function AddVirusTemp(sFile As String, namavirus As String) As Boolean
Dim sCeksum     As String
Dim nDB         As Long
Dim nJumVirus   As Long
Dim MyHandle    As Long
Dim FalsCek     As String

MyHandle = GetHandleFile(sFile)

sCeksum = MYCeksum(sFile, MyHandle)
FrmRTP.Text1.Text = MYCeksum(sFile, MyHandle)
TutupFile MyHandle

FalsCek = String$(Len(sCeksum), "0")

' pakai cadangan jika perlu
If FalsCek = sCeksum Or sCeksum = vbNullString Then
   sCeksum = MYCeksumCadangan(sFile, MyHandle)
   FrmRTP.Text1.Text = MYCeksumCadangan(sFile, MyHandle)
   
End If

If ValidFile(sFile) = False Or sCeksum = vbNullString Or sCeksum = FalsCek Then
   AddVirusTemp = False
Else
  ' JumVirusUser = JumVirusUser + 1
 '  nJumVirus = JumVirusUser - 1
   ' Lalu tambahkan nama virus dan ceksum ke DB kusus user virus
 '  sMD5User(nJumVirus) = sCeksum ' masukan ke database sementara
  
  ' sNamaVirusUser(nJumVirus) = namavirus ' nama virusnya
    'MsgBox sCeksum + NamaVirus
   AddVirusTemp = True
   
   
End If
End Function
'analisa file
Public Function AnalisisPE(sFile As String)
Dim Ukuran As String
Dim RetPE As Long
Dim TmpHGlobal As Long
Dim CeksumFile   As String
Dim CekMuatanByte As String
Dim CRC As String
On Error GoTo LBL_AKHIR
TmpHGlobal = GetHandleFile(sFile)
   Dim m_CRC As clsCRC
   Set m_CRC = New clsCRC


With frmMain
 DrawIco sFile, .PicAnalisis, ricnLarge  ', , ricnSmall
 .PicAnalisis.Visible = True
.TxtNama.Text = GetFileTitle(sFile)
'cek ukuran
'Ukuran = FileLen(sfile)
Ukuran = GetSizeFile(TmpHGlobal)
'.TxtSize.Text = Ukuran
.TxtSize.Text = FormatSize(Ukuran)
'cek type
RetPE = IsValidPE32(TmpHGlobal)
If RetPE > 0 Then
.txtType.Text = "File PE"
   CekMuatanByte = DeteksiMuatanPE32(sFile)
   If CekMuatanByte > 0 Then
   .txtCeck.Text = "Caution this PE file contain " & FormatSize(CekMuatanByte) & ", Additional code." ' It may be infected by virus"
   Else
   .txtCeck.Text = "File PE ini tidak mengandung byte tambahan"
   End If
Else
.txtType.Text = "Not PE File"
.txtCeck.Text = "-"
End If
'cek ceksum
        CeksumFile = MYCeksum(sFile, TmpHGlobal)
       If CeksumFile = String$(Len(CeksumFile), "0") Then
       CeksumFile = MYCeksumCadangan(sFile, TmpHGlobal)
       End If
' CRC = Hex(m_CRC.CalculateFile(sfile))
 .txtCecksum.Text = CeksumFile '& vbNewLine & "CRC32 : " & CRC
 'If isVBa(sfile) = True Then
' .txtCompiler = "Visual Basic"
' Else
' .txtCompiler = "Bukan Visual Basic"
' End If
' .txtPacker = get_Packer(sfile)
End With
TutupFile TmpHGlobal

LBL_AKHIR:
End Function
'cek compiler program

' untuk mencocokan User Virus (dipanggil jika JumVirusUser>0)
' ArrS juga masuk sini
Private Function CocokanVirusUser(ByRef MyHash As String) As String
Dim MyCounter As Long

For MyCounter = 1 To JumVirusUser
    If sMD5User(MyCounter - 1) = MyHash Then
       CocokanVirusUser = sNamaVirusUser(MyCounter - 1)
    End If
Next MyCounter
End Function



' Membaca DB Wan'iez di folder (sign) - DecCode harus 9
Private Function ReadDatabaseAV(sFileDatabBase As String, DecCode As Byte) As String
Dim DataKeluar()   As Byte
Dim SignUkuran     As Long
Dim SizeFDB        As Long
Dim hFileDB        As Long
Dim iCount         As Long
Dim PenampungStr   As String
OpenFileNow sFileDatabBase ' hGlobal handlenya

SizeFDB = nSizeGlobal
hFileDB = hGlobal

If hFileDB > 0 Then
   Call ReadUnicodeFile2(hFileDB, 1, 10, DataKeluar)
   PenampungStr = StrConv(DataKeluar, vbUnicode)
   If Left$(PenampungStr, 2) = "KM" Then ' header benar
     SignUkuran = CLng(Mid$(PenampungStr, 3, 6))
     If (SizeFDB - 10) = SignUkuran Then ' ukuran data disamakan
        Erase DataKeluar
        Call ReadUnicodeFile2(hFileDB, 11, SignUkuran, DataKeluar)
        For iCount = 0 To UBound(DataKeluar)
            DataKeluar(iCount) = DataKeluar(iCount) Xor DecCode ' dekripsi
        Next
        PenampungStr = StrConv(DataKeluar, vbUnicode)
        ReadDatabaseAV = PenampungStr
     Else
       ReadDatabaseAV = "" ' udah gugur
     End If
   Else
     ReadDatabaseAV = "" ' udah gugur
   End If
   TutupFile hFileDB
Else
   ReadDatabaseAV = ""
End If
End Function

' Untuk Enkripsi database
Public Sub EnkripDB(sFileEnk As String, EnkCode As Byte, sFileOut As String)
Dim PenampungStr   As String
Dim DataKeluar()   As Byte
Dim SignUkuran     As Long
Dim SizeFDB        As Long
Dim hFileDB        As Long
Dim iCount         As Long

OpenFileNow sFileEnk ' hGlobal handlenya

SizeFDB = nSizeGlobal
hFileDB = hGlobal

If hFileDB > 0 Then
   Call ReadUnicodeFile2(hFileDB, 1, SizeFDB, DataKeluar)
   
   TutupFile hFileDB

   For iCount = 0 To UBound(DataKeluar)
       DataKeluar(iCount) = DataKeluar(iCount) Xor EnkCode ' dekripsi
   Next
   PenampungStr = "KM" & Right("000000" & CStr(SizeFDB), 6) & "##"
   PenampungStr = PenampungStr & StrConv(DataKeluar, vbUnicode)
   
   Erase DataKeluar
   
   If ValidFile(sFileOut) = True Then HapusFile sFileOut
   
   WriteFileUniSim sFileOut, PenampungStr

End If
End Sub


'...................... COCOKAN TAPI MILIK RTP
Public Function CocokanDataBaseRTP(sPath As String) As Boolean ' Memakai turboloop based hex [perulanganya DB di hemat]
Dim iCount       As Integer
Dim Ukuran       As String
Dim CeksumFile   As String
Dim CeksumVirus  As String
Dim nDataBase    As Byte
Dim RetPE        As Long
Dim TmpHGlobal   As Long
Dim RetVirus     As String
Dim RetHeur      As Boolean
Dim CounTx As Long
Dim Splitter() As String
'Dim A As String
On Error GoTo LBL_AKHIR

TmpHGlobal = GetHandleFile(sPath)

With FrmRTP

If TmpHGlobal <= 0 Then GoTo LBL_AKHIR
'fSize = FileLen(sPath)
'If fSize = 0 Then Exit Function
'Cek *.lnk dulu
If UCase$(Right$(sPath, 4)) = ".LNK" Then
    If CeklnkFolderRTP(sPath) = True Then
       GoTo LBL_AKHIR ' akhiri aj
    End If
End If

RetPE = IsValidPE32(TmpHGlobal) ' fungsi balik IsValidPE32 adalah AddresOfNewHeader

If RetPE > 64 Then ' Ternyata DLL jga bisa diinjek virus kan
    ' Cek dengan Database Virus dulu jika file PE 32 exe
   ' CekEntryPointRTP TmpHGlobal, sPath
    RetVirus = GetDataEP(TmpHGlobal, 40, RetPE)
    If RetVirus <> "" Then
       ' klo pengecualian keluar
       If ApaPengecualianFile(sPath, JumFileExcep) = True Then GoTo LBL_AKHIR
          Ukuran = Format$(GetSizeFile(TmpHGlobal), "#,#") ' ukuran dalam string
          If Left$(RetVirus, 3) = "PW:" Then ' artinya hanya WormPoli
             AddInfoToList .lvRTP, Mid$(RetVirus, 4), sPath, Ukuran, f_bahasa(7), 0, 18
          Call tampilkanRTP
          Else
             AddInfoToList .lvRTP, RetVirus, sPath, Ukuran, f_bahasa(6), 2, 18
          Call tampilkanRTP
          End If
       GoTo LBL_AKHIR ' akhiri aj
    End If
  
End If
'fix by.Heru
'If GetPE3264Type(hFile) > 0 Then
 RetVirus = CekEntryPoint(TmpHGlobal, sPath)
If RetVirus <> "" Then
  If ApaPengecualianFile(sPath, JumFileExcep) = True Then GoTo LBL_AKHIR
 Ukuran = Format$(GetSizeFile(TmpHGlobal), "#,#") ' ukuran dalam string
 AddInfoToList .lvRTP, RetVirus, sPath, Ukuran, f_bahasa(6), 2, 18
GoTo LBL_AKHIR '
End If
'If fSize > 700000 Then GoTo LBL_AKHIR
If RetPE > 0 Then ' tergolong PE
       CeksumFile = MYCeksum(sPath, TmpHGlobal)
       
       If CeksumFile = String$(Len(CeksumFile), "0") Then
          CeksumFile = MYCeksumCadangan(sPath, TmpHGlobal)
       End If
       
       nDataBase = SelectDB(CeksumFile)
    
       'Ceksumer PE
       For iCount = 1 To JumlahVirus(nDataBase)
         If sMD5(nDataBase, iCount) = CeksumFile Then  ' jika virus didapet
           ' klo pngecualian keluar
           If ApaPengecualianFile(sPath, JumFileExcep) = True Then GoTo LBL_AKHIR
           Ukuran = Format$(GetSizeFile(TmpHGlobal), "#,#") ' ukuran dalam string
           AddInfoToList .lvRTP, sNamaVirus(nDataBase, iCount), sPath, Ukuran, f_bahasa(7), 0, 18
           Call tampilkanRTP
           GoTo LBL_AKHIR
         End If
         DoEvents
       Next
Else
   CeksumFile = MYCeksum(sPath, TmpHGlobal)
   
   If CeksumFile = String$(Len(CeksumFile), "0") Then
      CeksumFile = MYCeksumCadangan(sPath, TmpHGlobal)
   End If
       
   nDataBase = SelectDB(CeksumFile)
    
   ' Ceksumer nonPE
    For iCount = 1 To JumlahVirusNonPE(nDataBase)
        If sMD5nonPE(nDataBase, iCount) = CeksumFile Then  ' jika virus didapet
           ' klo pngecualian keluar
           If ApaPengecualianFile(sPath, JumFileExcep) = True Then GoTo LBL_AKHIR
           Ukuran = Format$(GetSizeFile(TmpHGlobal), "#,#") ' ukuran dalam string
           AddInfoToList .lvRTP, sNamaVirusnonPE(nDataBase, iCount), sPath, Ukuran, f_bahasa(7), 0, 18
          Call tampilkanRTP
          GoTo LBL_AKHIR
        End If
        DoEvents
    Next
 
End If

'If ValidFile(App.Path & "\USER.DAT") = True Then
   
 CeksumFile = MYCeksum(sPath, TmpHGlobal)
       
       If CeksumFile = String$(Len(CeksumFile), "0") Then
       CeksumFile = MYCeksumCadangan(sPath, TmpHGlobal)
       End If
              
 For CounTx = 0 To (2 + BanyakUDB) ' 125 banyaknya Db internal dari 0
    On Error GoTo LBL_AKHIR
    Splitter = Split(InternalDb(CounTx), "|")
    If Splitter(0) = CeksumFile Then 'MeiPattern(spath, 202) Then
       'VirusFound = VirusFound + 1
       Ukuran = Format$(GetSizeFile(TmpHGlobal), "#,#")
       If Splitter(1) = "Suspected With ArrS" Then
       If ApaPengecualianFile(sPath, JumFileExcep) = True Then GoTo LBL_AKHIR
        AddInfoToList .lvRTP, Splitter(1), sPath, Ukuran, f_bahasa(24), 2, 18
       ' .lbMalware.Caption = ": " & right$("000000" & VirusFound, 6) & " " & d_bahasa(38)
       ' VirusFound = VirusFound + 1
      ' VirStatus = True
        GoTo LBL_AKHIR
       Else
      ' VirusFound = VirusFound + 1
       'VirStatus = True
       If ApaPengecualianFile(sPath, JumFileExcep) = True Then GoTo LBL_AKHIR
        AddInfoToList .lvRTP, Splitter(1), sPath, Ukuran, f_bahasa(24), 2, 18
       '.lbMalware.Caption = ": " & right$("000000" & VirusFound, 6) & " " & d_bahasa(38)
       GoTo LBL_AKHIR
       End If
      
    End If
Next
'End If

VirStatus = False ' set false
If FrmConfig.ck2.Value = 1 Then RetHeur = CekWithHeuristicRTP(sPath, TmpHGlobal)

' jika heuristic tidak nemu, option di set dan PE valid
If RetHeur = False And RetPE > 0 Then
   Call CekInformationRTP(sPath, TmpHGlobal)
   
Else
   GoTo LBL_AKHIR
End If
End With

TutupFile TmpHGlobal ' jaga-jaga aja

Exit Function
LBL_AKHIR: ' kalo error/udah dapet lngsung akhiri pemindaian file saat ini
Call tampilkanRTP

AutoLst FrmRTP.lvRTP
    TutupFile TmpHGlobal
    nSalityGet = "" 'biar gak bahaya
End Function
' disini untuk rtp hookingnya by heru
Public Function CocokanDataBaseHook(ByRef sPath As String) As Boolean ' Memakai turboloop based hex [perulanganya DB di hemat]
Dim iCount       As Integer
Dim Ukuran       As String
Dim CeksumFile   As String
Dim CeksumVirus  As String
Dim nDataBase    As Byte
Dim RetPE        As Long
Dim TmpHGlobal   As Long
Dim RetVirus     As String
Dim RetHeur      As Boolean
Dim RetSuOrNot As String
Dim CounTx As Long
Dim Splitter() As String
'Dim A As String

On Error GoTo LBL_AKHIR

TmpHGlobal = GetHandleFile(sPath)

With FrmRTP

If TmpHGlobal <= 0 Then GoTo LBL_AKHIR
'Cek *.lnk dulu
If UCase$(Right$(sPath, 4)) = ".LNK" Then
    If CeklnkFolderRTP(sPath) = True Then
       GoTo LBL_AKHIR ' akhiri aj
    End If
End If
RetPE = IsValidPE32(TmpHGlobal) ' fungsi balik IsValidPE32 adalah AddresOfNewHeader

If RetPE > 64 Then ' Ternyata DLL jga bisa diinjek virus kan
    ' Cek dengan Database Virus dulu jika file PE 32 exe
   ' CekEntryPointRTP TmpHGlobal, sPath
    RetVirus = GetDataEP(TmpHGlobal, 40, RetPE)
    If RetVirus <> "" Then
       ' klo pengecualian keluar
       If ApaPengecualianFile(sPath, JumFileExcep) = True Then GoTo LBL_AKHIR
          Ukuran = Format$(GetSizeFile(TmpHGlobal), "#,#") ' ukuran dalam string
          If Left$(RetVirus, 3) = "PW:" Then ' artinya hanya WormPoli
             AddInfoToList .lvRTP, Mid$(RetVirus, 4), sPath, Ukuran, f_bahasa(7), 0, 18
                virusradar = virusradar + 1
         ' frmMain.Label13.Caption = "Threads detected : " & Right$("00000000" & virusradar, 8) & " " & d_bahasa(38)
  
          Else
             AddInfoToList .lvRTP, RetVirus, sPath, Ukuran, f_bahasa(6), 2, 18
      virusradar = virusradar + 1
          'frmMain.Label13.Caption = "Threads detected : " & Right$("00000000" & virusradar, 8) & " " & d_bahasa(38)
  
          End If
       GoTo LBL_AKHIR ' akhiri aj
    End If
    
End If
'If GetPE3264Type(hFile) > 0 Then
RetVirus = CekEntryPoint(TmpHGlobal, sPath)
If RetVirus <> "" Then
  If ApaPengecualianFile(sPath, JumFileExcep) = True Then GoTo LBL_AKHIR
 Ukuran = Format$(GetSizeFile(TmpHGlobal), "#,#") ' ukuran dalam string
 AddInfoToList .lvRTP, RetVirus, sPath, Ukuran, f_bahasa(6), 2, 18
virusradar = virusradar + 1
          'frmMain.Label13.Caption = "Threads detected : " & Right$("00000000" & virusradar, 8) & " " & d_bahasa(38)
  
GoTo LBL_AKHIR '
End If
'End If
If RetPE > 0 Then ' tergolong PE
       CeksumFile = MYCeksum(sPath, TmpHGlobal)
       
       If CeksumFile = String$(Len(CeksumFile), "0") Then
          CeksumFile = MYCeksumCadangan(sPath, TmpHGlobal)
       End If
       
       nDataBase = SelectDB(CeksumFile)
    
       'Ceksumer PE
       For iCount = 1 To JumlahVirus(nDataBase)
         If sMD5(nDataBase, iCount) = CeksumFile Then  ' jika virus didapet
           ' klo pngecualian keluar
           If ApaPengecualianFile(sPath, JumFileExcep) = True Then GoTo LBL_AKHIR
           Ukuran = Format$(GetSizeFile(TmpHGlobal), "#,#") ' ukuran dalam string
           AddInfoToList .lvRTP, sNamaVirus(nDataBase, iCount), sPath, Ukuran, f_bahasa(7), 0, 18
           virusradar = virusradar + 1
          'frmMain.Label13.Caption = "Threads detected : " & Right$("00000000" & virusradar, 8) & " " & d_bahasa(38)
  
           GoTo LBL_AKHIR
         End If
         DoEvents
       Next

Else
   CeksumFile = MYCeksum(sPath, TmpHGlobal)
   
   If CeksumFile = String$(Len(CeksumFile), "0") Then
      CeksumFile = MYCeksumCadangan(sPath, TmpHGlobal)
   End If
       
   nDataBase = SelectDB(CeksumFile)
    
   ' Ceksumer nonPE
    For iCount = 1 To JumlahVirusNonPE(nDataBase)
        If sMD5nonPE(nDataBase, iCount) = CeksumFile Then  ' jika virus didapet
           ' klo pngecualian keluar
           If ApaPengecualianFile(sPath, JumFileExcep) = True Then GoTo LBL_AKHIR
           Ukuran = Format$(GetSizeFile(TmpHGlobal), "#,#") ' ukuran dalam string
           AddInfoToList .lvRTP, sNamaVirusnonPE(nDataBase, iCount), sPath, Ukuran, f_bahasa(7), 0, 18
           virusradar = virusradar + 1
        '  frmMain.Label13.Caption = "Threads detected : " & Right$("00000000" & virusradar, 8) & " " & d_bahasa(38)
  
           GoTo LBL_AKHIR
        End If
        DoEvents
    Next
End If

'If ValidFile(App.Path & "\USER.DAT") = True Then
   
        CeksumFile = MYCeksum(sPath, TmpHGlobal)
       
       If CeksumFile = String$(Len(CeksumFile), "0") Then
       CeksumFile = MYCeksumCadangan(sPath, TmpHGlobal)
       End If
              
 For CounTx = 0 To (2 + BanyakUDB) ' 125 banyaknya Db internal dari 0
    On Error GoTo LBL_AKHIR
    Splitter = Split(InternalDb(CounTx), "|")
    If Splitter(0) = CeksumFile Then 'MeiPattern(spath, 202) Then
       'VirusFound = VirusFound + 1
       'VirStatus = True
       If Splitter(1) = "Suspected With ArrS" Then
       If ApaPengecualianFile(sPath, JumFileExcep) = True Then GoTo LBL_AKHIR
        AddInfoToList .lvRTP, Splitter(1), sPath, FileLen(sPath), f_bahasa(24), 2, 18
       virusradar = virusradar + 1
          'frmMain.Label13.Caption = "Threads detected : " & Right$("00000000" & virusradar, 8) & " " & d_bahasa(38)
  
       ' .lbMalware.Caption = ": " & right$("000000" & VirusFound, 6) & " " & d_bahasa(38)
       ' VirusFound = VirusFound + 1
      ' VirStatus = True
        GoTo LBL_AKHIR
       Else
      ' VirusFound = VirusFound + 1
       'VirStatus = True
       If ApaPengecualianFile(sPath, JumFileExcep) = True Then GoTo LBL_AKHIR
        AddInfoToList .lvRTP, Splitter(1), sPath, FileLen(sPath), f_bahasa(24), 2, 18
virusradar = virusradar + 1
         ' frmMain.Label13.Caption = "Threads detected : " & Right$("00000000" & virusradar, 8) & " " & d_bahasa(38)
         GoTo LBL_AKHIR
       End If
      
    End If
Next
'End If
VirStatus = False ' set false
'If frmMain.ck2.Value = 1 Then
'RetHeur = CekWithHeuristicRTPUSB(sPath, TmpHGlobal)
If RetHeur = False And RetPE > 0 Then
   Call CekInformationRTP(sPath, TmpHGlobal)
   
Else
   GoTo LBL_AKHIR
End If
End With
' jika heuristic tidak nemu, option di set dan PE valid


TutupFile TmpHGlobal ' jaga-jaga aja

Exit Function
LBL_AKHIR: ' kalo error/udah dapet lngsung akhiri pemindaian file saat ini
Call tampilkanRTP
AutoLst FrmRTP.lvRTP
    TutupFile TmpHGlobal
    
    nSalityGet = "" 'biar gak bahaya
         
End Function

