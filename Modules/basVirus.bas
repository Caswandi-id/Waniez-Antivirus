Attribute VB_Name = "basVirus"
Option Explicit

' Segala sesuatu tentang virus.... :D

' Kalo mau pake fungsi-fungsi disini harus yakinkan Valid PE

' Memperkenalkan Heuristic Baru atau Ceksum Untuk Virus
' Namanya HBI LX - Heuristical Byte Identification LX - saya pake 30 sementara
' Namanya ngawur ja :D, X menandakan panjang byte untuk sample ceksum virus

Public TmpCeksumPE As String ' menampung ceksum PE
' -- local ajah sama kaya di basPE
Private Declare Function SetFilePointer Lib "kernel32" (ByVal hFile As Long, ByVal lDistanceToMove As Long, ByVal pv6432_lpDistancetoMoveHigh As Long, ByVal dwMoveMethod As Long) As Long
Private Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, ByVal pv_lpbuffer As Long, ByVal nNumberOfBytesToRead As Long, ByVal pv_lpNumberOfBytesRead As Long, ByVal pv_lpOverlapped As Long) As Long
Private Declare Sub RtlZeroMemory Lib "NTDLL.DLL" (ByVal pDestBuffer As Long, ByVal nDestLengthToFillWithZeroBytes As Long) '<---reset isi dst yaitu mengisinya dengan bytenumber = 0.


Dim PHVirus(100)           As String
Dim PHNameVirus(100)       As String


Public Sub InitPHPattern() ' Polimorphic Worm Masuk Sini

    ' Worm Poli (depanya ada koma itu wajib)
    PHVirus(0) = ",53,83,EC,44,B8,23,10,40,0,B9,0,0,0,0" ' ini worm tapi Poli ceksumnya
    PHNameVirus(0) = "Win32/Mabezat.A"
  
    ' Mulai Virus
    PHVirus(1) = "60,E8,XX,XX,XX,XX,33,C9,8B,2C,24,90,81,XX,XX,XX,XX,XX,81"
    PHNameVirus(1) = "Win32/Sality.A" ' ini pakai cara kusus, karena membuat section baru, biar lebih baik detektornya

    PHVirus(2) = "60,E8,E6,19,0,0,8B,XX,XX,XX,E8,XX,XX,XX,XX,61,68"
    PHNameVirus(2) = "Chirb@mm"
    
    
    PHVirus(3) = "52,60,B9,XX,XX,XX,XX,E8,0,0,0,0,5F,4F,66" ' XX,XX,66,XX,XX,XX,XX - BYTE relatifnya ad yang dimasukan, karena ada yang sama terus (mgkin krn 0 kali)
    PHNameVirus(3) = "Win32/Gaelicum.A"
    
    PHVirus(4) = "53,60,83,XX,XX,54,5B,53,E8,XX,XX,XX,XX,33,XX,E8"
    PHNameVirus(4) = "Win32/Downloader.NAE"
    
    PHVirus(5) = "BB,XX,XX,XX,XX,93,E9"
    PHNameVirus(5) = "Win32/Mabezat"
    
    PHVirus(6) = "60,E8,XX,XX,XX,XX,61,E9"
    PHNameVirus(6) = "Win32/Expiro"
      
    PHVirus(7) = "F5,E8,1E,XX,XX,XX,5D"
    PHNameVirus(7) = "Win32/Virut.NBH"
       
    PHVirus(8) = "E8,XX,XX,XX,XX,55"
    PHNameVirus(8) = "Win32/Virut.S"
    
    PHVirus(9) = "FC,E8,1A,XX,XX,XX,B6"
    PHNameVirus(9) = "Win32/Virut.BT"
    
  '  PHVirus(10) = "6A,XX,68,XX,15,0,XX,E8,XX,XX,0,0,BF,94,0,0,0,8B,C7,E8"
   ' PHNameVirus(10) = "HEUR.Worm.W32.Generic"
    
    PHVirus(11) = "60,E8,0,0,0,0,5D,8B,C5,81,ED"
    PHNameVirus(11) = "W32/Ramnit.H"
    
    PHVirus(12) = ",55,8B,EC,81,C4,F4,FE,FF,FF"
    PHNameVirus(12) = "Win32/Ramnit.CPL"
    
    PHVirus(13) = ",55,8B,EC,83,EC,3C,FF,35,A4,46"
    PHNameVirus(13) = "Win32/Ramnit.Ax"
  
    PHVirus(14) = ",55,8B,EC,83,EC,2C,81,65,F4"
    PHNameVirus(14) = "Win32/Ramnit.A"
    
    PHVirus(15) = ",60,E8,0,0,0,0,5D,8B"
    PHNameVirus(15) = "Win32/Ramnit.G"
    
    PHVirus(16) = ",83,C4,E0,E8"
    PHNameVirus(16) = "Win32/Vitro"
    
    PHVirus(17) = ",55,8B,EC,6A,FF,68,XX,40,40,0,68,XX,XX,40,0,64,A1,0,0,0,0,50,64,89,25,0,0,0,0,83,EC,58,53,56,57,89,65,E8,FF,15"
    PHNameVirus(17) = "Lyzapo"
    
    PHVirus(18) = ",55,8B,EC,53,8B,5D,8,56,8B,75,C,57,8B,7D,10,85,F6,75,9,83,3D,XX,XX,1,10,0,EB,26,83,FE,1,74,5,83,FE,2,75,22,A1,XX"
    PHNameVirus(18) = "Conficker"
    
    PHVirus(19) = "60,E8,XX,0,0,0,8D,BD,0,10,40,0,68,XX,XX"
    PHNameVirus(19) = "Win32/Sality"
    
    PHVirus(20) = "55,8B,EC,83,EC,XX,81,65,XX,0,0,0,0"
    PHNameVirus(20) = "Win32/WaterMark"
    
    PHVirus(21) = "BB,XX,XX,XX,XX,93,E9"
    PHNameVirus(21) = "Win32/Mabezat"
    
    PHVirus(22) = "60,E8,XX,XX,XX,XX,61,E9"
    PHNameVirus(22) = "Win32/Expiro"
    
    PHVirus(23) = ",60,E8,0,0,0,0,5D,8B,C5,81,ED,A8,A6,1,20,2B,85,F,AE,1,20,89,85,B,AE,1,20,B0,0,86,85,40,B0,1,20,3C,1,F,85,BC"
    PHNameVirus(23) = "Win32/Ramnit"
    
    PHVirus(24) = ",E9,25,E4,FF,FF,0,0,0,2A,EB,45,7E,1E,9C,E,0,0,0,0,0,0,0,0,0,3E,9C,E,0,2E,9C,E,0,26,9C,E"
    PHNameVirus(24) = "RontokBro"
    
    PHVirus(25) = ",80,7C,24,8,1,F,85,C2,1,0,0,60,BE,0,60,0,10,8D,BE,0,B0,FF,FF"
    PHNameVirus(25) = "Conficker.G"
    
    PHVirus(26) = ",68,34,C4,40,0,E8,F0,FF,FF,FF,0,0,0,0,0,0,30,0,0,0,38,0,0,0,0,0,0,0,96,A2,51,D0,75,75,D2,11,82,60,44,45"
    PHNameVirus(26) = "AMG"
    
    PHVirus(27) = ",55,8B,EC,53,8B,5D,8,56,8B,75,C,57,8B,7D,10,85,F6,75,9,83,3D,54,8,1,10,0,EB,26,83,FE,1,74,5,83,FE,2,75,22,A1,68"
    PHNameVirus(27) = "Fanny"

End Sub

Public Function CocokanVirusWithPHPattern(ByVal DataEPVirus As String) As String
Dim iCount As Byte
' Worm Poli Dulu
  
  If Left$(DataEPVirus, Len(PHVirus(0))) = PHVirus(0) Then ' harus tpt 100%
     CocokanVirusWithPHPattern = "PW:" & PHNameVirus(iCount) ' Prefik PW arinta PoliWorm
     Exit Function
  End If
'
   If Left(DataEPVirus, Len(PHVirus(1))) = PHVirus(1) Then ' harus tpt 100%
     CocokanVirusWithPHPattern = "PW:" & PHNameVirus(1) ' Prefik PW arinta PoliWorm
     Exit Function
  End If
  
     If Left(DataEPVirus, Len(PHVirus(2))) = PHVirus(2) Then ' harus tpt 100%
     CocokanVirusWithPHPattern = "PW:" & PHNameVirus(2) ' Prefik PW arinta PoliWorm
     Exit Function
  End If
  
     If Left(DataEPVirus, Len(PHVirus(3))) = PHVirus(3) Then ' harus tpt 100%
     CocokanVirusWithPHPattern = "PW:" & PHNameVirus(3) ' Prefik PW arinta PoliWorm
     Exit Function
  End If
  
     If Left(DataEPVirus, Len(PHVirus(4))) = PHVirus(4) Then ' harus tpt 100%
     CocokanVirusWithPHPattern = "PW:" & PHNameVirus(4) ' Prefik PW arinta PoliWorm
     Exit Function
  End If
  
     If Left(DataEPVirus, Len(PHVirus(5))) = PHVirus(5) Then ' harus tpt 100%
     CocokanVirusWithPHPattern = "PW:" & PHNameVirus(5) ' Prefik PW arinta PoliWorm
     Exit Function
  End If
  
     If Left(DataEPVirus, Len(PHVirus(6))) = PHVirus(6) Then ' harus tpt 100%
     CocokanVirusWithPHPattern = "PW:" & PHNameVirus(6) ' Prefik PW arinta PoliWorm
     Exit Function
  End If
  
   If Left(DataEPVirus, Len(PHVirus(7))) = PHVirus(7) Then ' harus tpt 100%
     CocokanVirusWithPHPattern = "PW:" & PHNameVirus(7) ' Prefik PW arinta PoliWorm
     Exit Function
  End If
  
   If Left(DataEPVirus, Len(PHVirus(8))) = PHVirus(8) Then ' harus tpt 100%
     CocokanVirusWithPHPattern = "PW:" & PHNameVirus(8) ' Prefik PW arinta PoliWorm
     Exit Function
  End If
  
   If Left(DataEPVirus, Len(PHVirus(9))) = PHVirus(9) Then ' harus tpt 100%
     CocokanVirusWithPHPattern = "PW:" & PHNameVirus(9) ' Prefik PW arinta PoliWorm
     Exit Function
  End If
  
   'If Left(DataEPVirus, Len(PHVirus(10))) = PHVirus(10) Then ' harus tpt 100%
    ' CocokanVirusWithPHPattern = "PW:" & PHNameVirus(10) ' Prefik PW arinta PoliWorm
    ' Exit Function
 ' End If
  
   If Left(DataEPVirus, Len(PHVirus(11))) = PHVirus(11) Then ' harus tpt 100%
     CocokanVirusWithPHPattern = "PW:" & PHNameVirus(11) ' Prefik PW arinta PoliWorm
     Exit Function
  End If
  
     If Left(DataEPVirus, Len(PHVirus(12))) = PHVirus(12) Then ' harus tpt 100%
     CocokanVirusWithPHPattern = "PW:" & PHNameVirus(12) ' Prefik PW arinta PoliWorm
     Exit Function
  End If
  
     If Left(DataEPVirus, Len(PHVirus(13))) = PHVirus(13) Then ' harus tpt 100%
     CocokanVirusWithPHPattern = "PW:" & PHNameVirus(13) ' Prefik PW arinta PoliWorm
     Exit Function
  End If
  
     If Left(DataEPVirus, Len(PHVirus(14))) = PHVirus(14) Then ' harus tpt 100%
     CocokanVirusWithPHPattern = "PW:" & PHNameVirus(14) ' Prefik PW arinta PoliWorm
     Exit Function
  End If
  
   If Left(DataEPVirus, Len(PHVirus(15))) = PHVirus(15) Then ' harus tpt 100%
     CocokanVirusWithPHPattern = "PW:" & PHNameVirus(15) ' Prefik PW arinta PoliWorm
     Exit Function
  End If
  
     If Left(DataEPVirus, Len(PHVirus(16))) = PHVirus(16) Then ' harus tpt 100%
     CocokanVirusWithPHPattern = "PW:" & PHNameVirus(16) ' Prefik PW arinta PoliWorm
     Exit Function
  End If
  
     If Left(DataEPVirus, Len(PHVirus(17))) = PHVirus(17) Then ' harus tpt 100%
     CocokanVirusWithPHPattern = "PW:" & PHNameVirus(17) ' Prefik PW arinta PoliWorm
     Exit Function
  End If
  
     If Left(DataEPVirus, Len(PHVirus(18))) = PHVirus(18) Then ' harus tpt 100%
     CocokanVirusWithPHPattern = "PW:" & PHNameVirus(18) ' Prefik PW arinta PoliWorm
     Exit Function
  End If
  
     If Left(DataEPVirus, Len(PHVirus(19))) = PHVirus(19) Then ' harus tpt 100%
     CocokanVirusWithPHPattern = "PW:" & PHNameVirus(19) ' Prefik PW arinta PoliWorm
     Exit Function
  End If
  
     If Left(DataEPVirus, Len(PHVirus(20))) = PHVirus(20) Then ' harus tpt 100%
     CocokanVirusWithPHPattern = "PW:" & PHNameVirus(20) ' Prefik PW arinta PoliWorm
     Exit Function
  End If
  
     If Left(DataEPVirus, Len(PHVirus(21))) = PHVirus(21) Then ' harus tpt 100%
     CocokanVirusWithPHPattern = "PW:" & PHNameVirus(21) ' Prefik PW arinta PoliWorm
     Exit Function
  End If
  
     If Left(DataEPVirus, Len(PHVirus(22))) = PHVirus(22) Then ' harus tpt 100%
     CocokanVirusWithPHPattern = "PW:" & PHNameVirus(22) ' Prefik PW arinta PoliWorm
     Exit Function
  End If
  
     If Left(DataEPVirus, Len(PHVirus(23))) = PHVirus(23) Then ' harus tpt 100%
     CocokanVirusWithPHPattern = "PW:" & PHNameVirus(23) ' Prefik PW arinta PoliWorm
     Exit Function
  End If
  
     If Left(DataEPVirus, Len(PHVirus(24))) = PHVirus(24) Then ' harus tpt 100%
     CocokanVirusWithPHPattern = "PW:" & PHNameVirus(24) ' Prefik PW arinta PoliWorm
     Exit Function
  End If
  
     If Left(DataEPVirus, Len(PHVirus(25))) = PHVirus(25) Then ' harus tpt 100%
     CocokanVirusWithPHPattern = "PW:" & PHNameVirus(25) ' Prefik PW arinta PoliWorm
     Exit Function
  End If
  
     If Left(DataEPVirus, Len(PHVirus(26))) = PHVirus(26) Then ' harus tpt 100%
     CocokanVirusWithPHPattern = "PW:" & PHNameVirus(26) ' Prefik PW arinta PoliWorm
     Exit Function
  End If
  
     If Left(DataEPVirus, Len(PHVirus(27))) = PHVirus(27) Then ' harus tpt 100%
     CocokanVirusWithPHPattern = "PW:" & PHNameVirus(27) ' Prefik PW arinta PoliWorm
     Exit Function
  End If
  
 For iCount = 1 To 10
  If HRInstr(DataEPVirus, PHVirus(iCount), 100) > 0 Then
     CocokanVirusWithPHPattern = PHNameVirus(iCount)
     Exit Function
  End If
Next

For iCount = 12 To 27 ' virus kecil data EP nya
  If HRInstr(Left$(DataEPVirus, 30), PHVirus(iCount), 100) > 0 Then
     CocokanVirusWithPHPattern = PHNameVirus(iCount)
     Exit Function
  End If
Next

CocokanVirusWithPHPattern = ""

End Function

' disini sambil menyelam sambil minum air (nanti di fungsi ini say ganti menjadi mendapatkan deretan byte sbanyak 256)
' Public Function GetHIBCeksum(hFilePE As Long, nBased As Long, AddNewHeaderBase0 As Long) As String ' return ke string
' virus juga di cek lgsung disini biar lbih ngebut
Public Function GetDataEP(hFilePE As Long, nBased As Long, AddNewHeaderBase0 As Long) As String ' return ke string
Dim INTH32              As IMAGE_NT_HEADERS_32
Dim ISECH()             As IMAGE_SECTION_HEADER
Dim RetFunct            As Long
Dim nNumberBytesOpsRet  As Long
Dim nSection            As Long
Dim PPhysicEP           As Long
Dim iCount              As Integer
Dim OutData()           As Byte
Dim Sec2(1)             As String ' cadangan deteksi sality dan tanatos
Dim nFisik              As Long
Dim nVirtual            As Long
Dim BiggestSectionOff   As Long
Dim SectionToSize       As Long
Dim OPTurnA             As Long
Dim OPTurnB             As Long
Dim btCPattern()        As Byte
Dim KePEHeur            As Boolean ' apakah layak untuk di proses ke PE Heur/Tidak


Dim StrSecAlman         As String

On Error GoTo LBL_AKHIRI

Call SetFilePointer(hFilePE, AddNewHeaderBase0, 0, 0)  '---Base0. lgsung menuju target
RetFunct = ReadFile(hFilePE, VarPtr(INTH32), Len(INTH32), VarPtr(nNumberBytesOpsRet), 0)

  
  nSection = INTH32.FileHeader.NumberOfSections
  
  If nSection <= 0 Then ' masak 0
     GoTo LBL_AKHIRI
  End If
    
  '---cek sectionheader:
  ReDim ISECH(nSection - 1) As IMAGE_SECTION_HEADER
  Call SetFilePointer(hFilePE, AddNewHeaderBase0 + Len(INTH32), 0, 0) '---Base0. INTH32=248 Bytes, set pointernya
  RetFunct = ReadFile(hFilePE, VarPtr(ISECH(0)), Len(ISECH(0)) * nSection, VarPtr(nNumberBytesOpsRet), 0) ' yang akan dibaca ukuran type section (40bytes) x jumlah section
  For iCount = 0 To nSection - 1
      If (INTH32.OptionalHeader.AddressOfEntryPoint >= ISECH(iCount).VirtualAddress) And (INTH32.OptionalHeader.AddressOfEntryPoint < (ISECH(iCount).VirtualAddress + ISECH(iCount).VirtualSize)) Then
          PPhysicEP = ISECH(iCount).PointerToRawData + (INTH32.OptionalHeader.AddressOfEntryPoint - ISECH(iCount).VirtualAddress)
          ' LogPrint (pPhysicEP)
            '---EP-di-file-fisik-ya ketemu,deh!
          If iCount = nSection - 1 And iCount > 1 Then KePEHeur = True Else KePEHeur = False ' layak untuk ke PE Heur karena EP pada section akhir dan section lebih dari 2
          Call ReadUnicodeFile2(hFilePE, PPhysicEP + 1, nBased, OutData)
          StrSecAlman = TATAByte(OutData) ' pinjam variablenya yah
          'LogPrint TATAByte(OutData) 'KETEMU JAWABANNYA HERU GITULHO
       End If
              If iCount = nSection - 1 Then
          xNamaSectionAkhir = StrConv(ISECH(iCount).SectionName, vbUnicode)
          xSectionAkhir = Hex$(ISECH(iCount).Characteristics)
       End If
       
       If iCount = nSection - 2 Then
          xNamaSectionAkhir2 = StrConv(ISECH(iCount).SectionName, vbUnicode)
          xSectionAkhir2 = Hex$(ISECH(iCount).Characteristics)
       End If
       If iCount > 0 Then
          If ISECH(iCount).PointerToRawData > BiggestSectionOff Then
             BiggestSectionOff = ISECH(iCount).PointerToRawData ' biasanya section terakhir
             SectionToSize = ISECH(iCount).SizeOfRawData
          End If
       Else
            BiggestSectionOff = ISECH(iCount).PointerToRawData ' awalnya baygkan terbesar ada yang pertama
            SectionToSize = ISECH(iCount).SizeOfRawData
       End If
       
   Next
   
   'Sekalian disini ngecek ukuran Real dari EXE :D
   nRealSizePE = BiggestSectionOff + SectionToSize
   ' Cek Virus dulu
   GetDataEP = CocokanVirusWithPHPattern(StrSecAlman)
   
   If InStr(GetDataEP, "ficker") > 0 Then
    If GetSizeFile(hFilePE) > 150000 Then
        If GetSizeFile(hFilePE) < 200000 Then
        Else
            GetDataEP = ""
        End If
    Else
        GetDataEP = ""
    End If
   End If
   If GetDataEP <> "" Then ' dapat virus
      Exit Function ' gak usah proses lagi
   End If
   GetDataEP = CekHeaderPE(xSectionAkhir, xNamaSectionAkhir)
   
   If GetDataEP <> "" Then ' dapat virus
      Exit Function ' gak usah proses lagi
   End If
    
    nFisik = ISECH(nSection - 1).SizeOfRawData ' ukuran section fisik terakhir

   ' Jika OP Code EP pertama adalah &H60 : PUSHAD
   If Left$(StrSecAlman, 3) = ",60" Then
      If (ISECH(nSection - 1).Characteristics And &H20000000) = &H20000000 Then ' pastikan sectiony Executable
         'Mainkan sality Awal dulu
         If HRInstr(StrSecAlman, PHVirus(1), 100) > 0 Then
            GetDataEP = PHNameVirus(1)
            Exit Function
         End If
         
         Call ReadUnicodeFile2(hFilePE, ISECH(nSection - 1).PointerToRawData + 1, ISECH(nSection - 1).SizeOfRawData, btCPattern)
         ' Cek Tanatos.M virus poly morphic banyak sampah gak cuma di section, tapi di luar section jg (biar lambat kali cekernya)
              OPTurnA = 0
              For OPTurnA = 0 To (ISECH(nSection - 1).SizeOfRawData - 1)
                    If btCPattern(OPTurnA) = &H8A Then '---8A 44 05 00 = MOV AL,BYTE PTR SS:[EBP+EAX]
                        If btCPattern(OPTurnA + 1) = &H44 Then
                            If btCPattern(OPTurnA + 2) = &H5 Then
                                If btCPattern(OPTurnA + 3) = &H0 Then
                                    OPTurnB = 0 '---preset.
                                    For OPTurnB = (OPTurnA + 4) To (ISECH(nSection - 1).SizeOfRawData - 1)
                                        If btCPattern(OPTurnB) = &H30 Then '---30 07 = XOR BYTE PTR DS:[EDI],AL
                                            If btCPattern(OPTurnB + 1) = &H7 Then
                                                OPTurnB = -1 '---maksimalkan value OPTurnB, yg berarti terlampaui (sudah dapat).
                                                Exit For
                                            End If
                                        End If
                                    Next
                                    If OPTurnB = -1 Then
                                        OPTurnA = -1 '---maksimalkan value OPTurnB, yg berarti terlampaui (sudah dapat).
                                        Exit For '---sudah,jangan terlalu lama berputar, 'ntar pusing :)
                                    End If
                                End If
                            End If
                        End If
                    End If
                Next
                
                If OPTurnA = -1 Then ' berhasil dapat Tanatos.M
                   GetDataEP = "Win32/Tanatos.M"
                   Exit Function
                End If
 
                
         If nSection > 1 And (ISECH(nSection - 1).Characteristics And &H20) = &H20 Then ' pastikan berisi Code
          ' 2 syarat sudah memenuhi bisa dianggap sality dengan heur
          Sec2(0) = TrimNull0(StrConv(ISECH(1).SectionName(), vbUnicode)) ' nama section ke2
          Sec2(1) = TrimNull0(StrConv(ISECH(nSection - 1).SectionName(), vbUnicode)) ' nama section terakhir
          nVirtual = ISECH(nSection - 1).VirtualSize ' ukuran virtual
          Call CekKemungkinanSality(Sec2(0), Sec2(1), nFisik, nVirtual)  ' lalu cek kemungkinan sality
         End If
      End If
   End If
       
LBL_ALMAN:
   If nFisik >= 36000 Then ' ukuran section fisik terakhir (alman masuk sini)
      Call ReadUnicodeFile2(hFilePE, (ISECH(nSection - 1).PointerToRawData + 1) + (nFisik - 36865), 12000, OutData) ' dari section terakhir offsetnya, 5000 bytes dari kanan
      StrSecAlman = StrConv(OutData, vbUnicode)
      If CekAlman("�������������h���������������������", StrSecAlman) = True Then
         GetDataEP = "Win32/Alman.A"
         Exit Function
      Else
         If CekAlman("4x� 35��PC���n���t�t(Z�����ԯ���/�", StrSecAlman) = True Then 'string alman ke 2
            GetDataEP = "Win32/Alman.B"
            Exit Function
         End If
      End If
   Else
      GetDataEP = ""
   End If
'
'If KePEHeur = True Then
'   If (ISECH(nSection - 1).Characteristics And &H20) = &H20 And (ISECH(nSection - 1).Characteristics And &H20000000) = &H20000000 Then ' contain code and executable
'      nPEHeurGet = "Suspect.PEHeur.2"
'   If (ISECH(nSection - 1).Characteristics And &H20000000) = &H20000000 Then ' executable
'      nPEHeurGet = "Suspect.PEHeur.1"
'   ElseIf (ISECH(nSection - 1).Characteristics And &H20) = &H20 Then ' contain code
'      nPEHeurGet = "Suspect.PEHeur.1"
'   Else
'   nPEHeurGet = "" ' bebaskan aj lah
'   End If
'Else
'   nPEHeurGet = ""
'End If
'If CekEntryPoint(hFilePE, path) = "positif" Then GoTo LBL_AKHIRI
'
Erase OutData
Erase btCPattern
Exit Function

LBL_AKHIRI:
    GetDataEP = ""
    nRealSizePE = 0
End Function

' Hanya bekerja setelah fungsi GetDataEP di proses
Public Function GetRealSizePE() As Long
    GetRealSizePE = nRealSizePE
End Function


' dipertajam ah
Function CekKemungkinanSality(nSec1 As String, nSecAkhir As String, SizeSecAkhirFisik As Long, SizeSecAkhirVirt As Long) As String
nSec1 = HilangkanTitik(nSec1)
nSecAkhir = HilangkanTitik(nSecAkhir)
If Mid$(nSecAkhir, 2) = nSec1 And SizeSecAkhirVirt = SizeSecAkhirFisik Then
   Select Case SizeSecAkhirFisik
          Case Is > 60000: CekKemungkinanSality = "70% Suspect Tanatos"
          Case Is > 20000: CekKemungkinanSality = "Win32/Ramnit.H"
          Case Is > 10000: CekKemungkinanSality = "Win32/Ramnit.H"
          Case Else: CekKemungkinanSality = ""
   End Select
ElseIf Mid$(nSecAkhir, 2) = nSec1 Then
   If SizeSecAkhirFisik > 60000 Then
      CekKemungkinanSality = "Win32/Ramnit.H"
   ElseIf SizeSecAkhirFisik > 10000 Then
      CekKemungkinanSality = "Win32/Ramnit.H"
   End If
ElseIf SizeSecAkhirVirt = SizeSecAkhirFisik Then
   If SizeSecAkhirFisik > 60000 Then
      CekKemungkinanSality = "Win32/Ramnit.H"
   ElseIf SizeSecAkhirFisik > 10000 Then
      CekKemungkinanSality = "Win32/Ramnit.H"
   End If
ElseIf SizeSecAkhirFisik > 60000 Then
   CekKemungkinanSality = "Win32/Ramnit.H"
ElseIf SizeSecAkhirFisik > 20000 Then
   CekKemungkinanSality = "Win32/Ramnit.H"
Else
   CekKemungkinanSality = ""
End If
nSalityGet = CekKemungkinanSality
End Function

' Hanya bekerja setelah fungsi GetDataEP di proses
Private Function CekAlman(ByRef StrInSect As String, ByRef sDataSection As String) As Boolean
If InStr(sDataSection, StrInSect) > 0 Then
    CekAlman = True
Else
    CekAlman = False
End If
End Function


' Untuk EP
Function TATAByte(sbyte() As Byte) As String
Dim i As Integer
For i = 1 To UBound(sbyte) + 1
    TATAByte = TATAByte & "," & Hex$(sbyte(i - 1))
Next
End Function


' Buffer
Private Function TrimNull0(sKar As String) As String
TrimNull0 = Left$(sKar, InStr(sKar, Chr$(0)) - 1)
End Function

' Buffer untuk suspect sality (hilangkan titik section)
Private Function HilangkanTitik(sKarBertitik As String) As String
If Left$(sKarBertitik, 1) = "." Then sKarBertitik = Mid$(sKarBertitik, 2)
HilangkanTitik = sKarBertitik
End Function


' InSTR Spesial gak pake Telur (dibuat seakurat dan secepat mungkin)
' Fungsi Balik bukan berarti posisi substring pada deretan Hex, jika pola cocok fungsi akan mnghasilkan lebih>0 dan sebaliknya (untuk optimalisasi aja)

Public Function HRInstr(ByVal DeretanHex As String, SubString As String, nProsenSensitif As Byte) As Long
' contoh 29,C0,FE,08,C0,74,XX,75,XX,EB -> XX wajib diberikan pada pattern walpun polanya sudah dpt panjang tanpa XX
Dim MyPos1    As Integer
Dim MyPos2    As Integer
Dim CutString As String
Dim ByHead    As String
Dim TmpPos    As Integer

' Ambil Header dari SubString seblum byte XX
ByHead = GetByteHeader(SubString)

Do
    MyPos1 = InStr(DeretanHex, ByHead)
    If MyPos1 > 0 Then
        If CocokanPolaPendek(SubString, Mid$(DeretanHex, MyPos1)) >= nProsenSensitif Then
           TmpPos = MyPos1
           GoTo BROAD_SUCCES
        End If
    Else
        GoTo BROAD_SUCCES
    End If
    DeretanHex = Mid$(DeretanHex, MyPos1 + 3) ' + 3 wajib banget
Loop While MyPos1 > 0

HRInstr = 0
Exit Function

BROAD_SUCCES:
    HRInstr = TmpPos

End Function


Private Function GetByteHeader(DeretanByte As String) As String
    GetByteHeader = Left$(DeretanByte, InStr(DeretanByte, "XX") - 1)
End Function

' bagi yang pencocokan ByteHeader sebelum XX sudah cocok panggil fungsi sini (meghasilkan prosentasi)
' dengan ini kita bisa milih berapa prosen pola yang cocok (XX tidakdihiraukan) saipa tahu aja pas masukin polanya terlalu pnjang jadi dengan ini bisa diantisipasi
Private Function CocokanPolaPendek(SubStringPola As String, DeretanHexTerpotong As String) As Long
Dim TmpHexTerpotong As String
Dim SubSplitter()   As String
Dim HexSplitter()   As String
Dim HexCocok        As Byte
Dim LengSub         As Byte
Dim MyCount         As Integer
Dim Penambah        As Byte ' penambah karena byte XX gak dihitung

On Error GoTo LBL_FALSE ' eror ya 0

LengSub = Len(SubStringPola)
TmpHexTerpotong = Left$(DeretanHexTerpotong, LengSub)

HexSplitter = Split(TmpHexTerpotong, ",") ' byte deteran hex yang sudah disesuakan ukuranya dengan pola virus
SubSplitter = Split(SubStringPola, ",")

For MyCount = 0 To UBound(SubSplitter)
  If HexSplitter(MyCount) = SubSplitter(MyCount) Then HexCocok = HexCocok + 1
  If SubSplitter(MyCount) = "XX" Then Penambah = Penambah + 1 ' artinya XX gak dihiraukan
Next

CocokanPolaPendek = ((HexCocok + Penambah) / (UBound(SubSplitter) + 1)) * 100 ' prosentasinya ketemu deh

Exit Function
LBL_FALSE:
    CocokanPolaPendek = 0
End Function
Public Function CekWithString(hFile As Long) As String
Static OutDat() As Byte
Static sData As String

Call ReadUnicodeFile2(hFile, 1, 70, OutDat)
sData = StrConv(OutDat, vbUnicode)
Erase OutDat
'Mulai Cek
If InStr(sData, "���X") > 0 Then
CekWithString = "Win32/Alman.B"
Exit Function
End If

Call ReadUnicodeFile2(hFile, 70, 176, OutDat)
sData = StrConv(OutDat, vbUnicode)
Erase OutDat

If InStr(sData, "D�Rich_D") > 0 Or InStr(sData, "D�Rich�D�") > 0 Then
CekWithString = "Win32/Oliga"
Exit Function
End If

Call ReadUnicodeFile2(hFile, 176, 208, OutDat)
sData = StrConv(OutDat, vbUnicode)
Erase OutDat

If InStr(sData, "(�I") > 0 Then
CekWithString = "Win32/Sysyer"
Exit Function
End If

Call ReadUnicodeFile2(hFile, 496, 540, OutDat)
sData = StrConv(OutDat, vbUnicode)
Erase OutDat

If InStr(sData, "�htnrnog") > 0 Then
CekWithString = "Win32/Vitro"
Exit Function
End If

Call ReadUnicodeFile2(hFile, 176, 200, OutDat)
sData = StrConv(OutDat, vbUnicode)
Erase OutDat

If InStr(sData, "�@��\�Rich�\�") > 0 Then
CekWithString = "Win32/Ramnit.L"

Call ReadUnicodeFile2(hFile, 1024, 1040, OutDat)
sData = StrConv(OutDat, vbUnicode)
Erase OutDat

If InStr(sData, "RRSP�") > 0 Then
CekWithString = "Win32/TeddyBear"
Exit Function
End If

If GetSizeFile(hFile) < 1048576 Then
Call ReadUnicodeFile2(hFile, 1, GetSizeFile(hFile), OutDat)
sData = StrConv(OutDat, vbUnicode)
Erase OutDat

If InStr(sData, "�e��=9�l��n�dY�l") > 0 Then
CekWithString = "Win32/Alman.B"
Exit Function
End If

If InStr(sData, "�h�7T�v�?") > 0 Then
CekWithString = "Win32/Service"
Exit Function
End If

If InStr(sData, "`�`oM�!5�=") > 0 Then
CekWithString = "Win32/Bumercia"
Exit Function
End If

If InStr(sData, ",$YG�F#G�(") > 0 Then
CekWithString = "Win32/KSplood"
Exit Function
End If

If InStr(sData, "x�T�h�L�") > 0 Then
CekWithString = "Win32/Service"
Exit Function
End If

If InStr(sData, "賶���E�3�") > 0 Then
CekWithString = "Win32/Spooler"
Exit Function
End If

If InStr(sData, "X5O!P%@AP") > 0 Then
CekWithString = "Eicar Not Virus!!"
Exit Function
End If

If InStr(sData, "#dq���") > 0 Then
CekWithString = "Virus Mow"
Exit Function
End If

If InStr(sData, "H��h�") > 0 Then
CekWithString = "Virut.CE"
Exit Function
End If

If InStr(sData, "ivofa.write") > 0 Then
CekWithString = "Vbs Alice.A.HTA"
Exit Function
End If

If InStr(sData, "�!�:i") > 0 Then
CekWithString = "Ramnit.Shortcut"
Exit Function
End If


If InStr(sData, "Q��XR") > 0 Then
CekWithString = "Sality.AT"
Exit Function
End If

If InStr(sData, "��I") > 0 Then
CekWithString = "W32.Vitro"
Exit Function
End If

If InStr(sData, "~�q=") > 0 Then
CekWithString = "W32.Vitro.A"
Exit Function
End If

If InStr(sData, "�QO") > 0 Then
CekWithString = "New Autorun Decrypt"
Exit Function
End If

If InStr(sData, "��?��{�◲�r��LhV]b�}�]�;m[��u�n����i�R� $=�ĭ") > 0 Then
CekWithString = "Ramnit.C"
Exit Function
End If

If InStr(sData, "��1�����)�_>3=�") > 0 Then
CekWithString = "Variant Virus Hidden FD"
Exit Function
End If

If InStr(sData, "readme.eml") > 0 Then
CekWithString = "W32.Runouce.A"
Exit Function
End If

If InStr(sData, "�!�:i") > 0 Then
CekWithString = "RECYCLER"
Exit Function
End If

If InStr(sData, "U��ĬjDE�P�����j") > 0 Then
CekWithString = "Ramnit.CPL"
Exit Function
End If
End If

If InStr(sData, "H") = 529 And InStr(sData, "w") = 541 And InStr(sData, "!") = 551 And InStr(sData, "C") = 557 Then
CekWithString = "Win32/Sality"
End If
Exit Function
End If
End Function