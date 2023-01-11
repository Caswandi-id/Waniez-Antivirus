Attribute VB_Name = "basVirus2"

Option Explicit
Public Declare Function SetFilePointer Lib "kernel32" (ByVal hFile As Long, ByVal lDistanceToMove As Long, ByVal pv6432_lpDistancetoMoveHigh As Long, ByVal dwMoveMethod As Long) As Long
Public Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, ByVal pv_lpbuffer As Long, ByVal nNumberOfBytesToRead As Long, ByVal py_lpNumberofBytesRead As Long, ByVal pv_lpOverlapped As Long) As Long
Public Declare Sub KosongkanMemory Lib "NTDLL.DLL" Alias "RtlZeroMemory" (ByVal pDestBuffer As Long, ByVal nDestLengthToFillWithZeroBytes As Long) '<---reset isi dst yaitu mengisinya dengan bytenumber = 0.
Public Heurvb As String
' Untuk PE Heuristic Redirected.
Const JMP1 = &HEB
Const JMP2 = &HEA
Const JMP3 = &HE9
Const JMP4 = &HFF
Const JMP5 = &HE8

'Public CekEntryPoint As Boolean


Function TATAByte(sbyte() As Byte) As String
Dim i As Integer
For i = 1 To UBound(sbyte) + 1
  TATAByte = TATAByte & "-" & Hex$(sbyte(i - 1))
Next
End Function

' Check PE File dengan metode FirstByte maupun heuristic PE

Public Function CekEntryPoint(hFilePE As Long, path As String) As String                        ' return ke string
Dim IDOSH                           As IMAGE_DOS_HEADER
Dim INTH32                          As IMAGE_NT_HEADERS_32
Dim ISECH()                         As IMAGE_SECTION_HEADER
Dim AddNewHeaderBase0               As Long
Dim nNumberBytesOpsRet              As Long
Dim nSection                        As Long
'Dim hFilePE                         As Long
Dim nFileLen                        As Long
Dim RetFunct                        As Long
Dim UkuranAsli                      As Long
Dim PPhysicEP                       As Long
Dim CekLengkap, BadPE, StrAlm       As String
Dim iCount                          As Integer
Dim TotalUkuranSection              As Long
Dim OutData()                       As Byte
Dim Saldata()                       As Byte
Dim AlmanData()                     As Byte
Dim i, Pointer                      As Integer
Dim CariString, CariLagi            As Long
Dim NS, NS2, CS, DS                 As String
Dim PosisiEP                        As Long
Dim BerapaText, PosisiText, Posdata As Integer
Dim FlagsEP                         As String
Dim nFisik                          As Long
Dim SalityBack                      As String
Dim Heurvb                          As String
Dim Ukuran                          As String
   ' CekEntryPoint = ""
    'MultipleThreat = False
    'If ReadDB(path) = True Then MultipleThreat = True
    'IconMalware = 2
    ' If FileLen(hFile) > 10000000 Then GoTo LBL_GAGAL

  ' If IsValidPE32(hFile) = 0 Then GoTo LBL_GAGAL ' Pastikan Valid PE 32 atoo 64 bit
    
    Call SetFilePointer(hFilePE, 0, 0, 0)  '---Base0, set ke pointer pertama
    RetFunct = ReadFile(hFilePE, VarPtr(IDOSH), Len(IDOSH), VarPtr(nNumberBytesOpsRet), 0) ' base1
   
    If RetFunct = 0 Then
        GoTo LBL_GAGAL
    End If
    
    ' cek header DOS
    If IDOSH.e_magic <> &H5A4D Then '---"MZ" mungkin udah DOS valid tapi.... cek lagi gak ya :D
        GoTo LBL_GAGAL
    End If
    
    AddNewHeaderBase0 = IDOSH.e_lfanew

    Call SetFilePointer(hFilePE, AddNewHeaderBase0, 0, 0) '---Base0.
    
    RetFunct = ReadFile(hFilePE, VarPtr(INTH32), Len(INTH32), VarPtr(nNumberBytesOpsRet), 0)
       
    If RetFunct = 0 Then
        GoTo LBL_GAGAL
    End If
    
    nSection = INTH32.FileHeader.NumberOfSections
  
    If nSection <= 0 Then ' masak 0 jumlah sectionya
       GoTo LBL_GAGAL
    End If
    
    TotalUkuranSection = 0
    ReDim ISECH(nSection - 1) As IMAGE_SECTION_HEADER
    Call SetFilePointer(hFilePE, AddNewHeaderBase0 + Len(INTH32), 0, 0) '---Base0. INTH32=248 Bytes, set pointernya - lebih irit
    RetFunct = ReadFile(hFilePE, VarPtr(ISECH(0)), Len(ISECH(0)) * nSection, VarPtr(nNumberBytesOpsRet), 0) ' yang akan dibaca ukuran type section (40bytes) x jumlah section
          
    PosisiEP = -1
    BerapaText = 0
    PosisiText = -1
    For iCount = 0 To nSection - 1
          'DoEvents
          TotalUkuranSection = TotalUkuranSection + (ISECH(iCount).VirtualAddress + ISECH(iCount).VirtualSize)
          'Cari dimana Posisi EP pada section
          If (INTH32.OptionalHeader.AddressOfEntryPoint >= ISECH(iCount).VirtualAddress) And (INTH32.OptionalHeader.AddressOfEntryPoint < (ISECH(iCount).VirtualAddress + ISECH(iCount).VirtualSize)) Then PosisiEP = iCount
          'Ambil isi Flags dari EP
          If PosisiEP > -1 Then FlagsEP = Hex$(ISECH(PosisiEP).Characteristics)
          'Ambil Flags dari section .text (Code Section)
          If TrimNull0(StrConv(ISECH(iCount).SectionName(), vbUnicode)) = ".text" Then
            CS = Left$(Hex$(ISECH(iCount).Characteristics), 1)
            BerapaText = BerapaText + 1
            PosisiText = iCount
            If BerapaText > 1 Then PosisiText = iCount
          End If
          ' Ambil posisi dan flags dari DS (Data Section)
          If TrimNull0(StrConv(ISECH(iCount).SectionName(), vbUnicode)) = ".data" Then
            DS = Left$(Hex$(ISECH(iCount).Characteristics), 1)
            Posdata = iCount
          End If
            
     Next
    
     nFisik = ISECH(nSection - 1).SizeOfRawData
     If PosisiEP > -1 Then
       PPhysicEP = ISECH(PosisiEP).PointerToRawData + (INTH32.OptionalHeader.AddressOfEntryPoint - ISECH(PosisiEP).VirtualAddress)
     Else
       GoTo LBL_GAGAL
     End If
     
     Call ReadFileADV(hFilePE, PPhysicEP, ISECH(PosisiEP).SizeOfRawData, OutData)
     If StrConv(OutData, vbUnicode) = "" Then GoTo LBL_GAGAL
  '   Debug.Print InfectChecksum(OutData, 1, 18) & " " & PathScan
     
     
      ' Deteksi W32.Polyene -- Update 29/11/2010
       If PosisiEP = nSection - 1 Then ' Cek apakah posisi EP berada pada section terakhir
        If InStr(FlagsEP, "E") > 0 Then ' Cek apakah section tujuan EP - WRITEABLE ??
         If InfectChecksum(OutData, 1, 18) = "50ED1" Then
           
              CekEntryPoint = "W32.Polyene"   ' YES, ketemu . .
          
          'CekEntryPointRTP = "Positif"
           GoTo LBL_SUDAH
         End If
       End If
    End If
    
    
    'Virut.O -- 30/11/2010
    If PosisiEP = nSection - 1 Then
      If InStr(FlagsEP, "E") > 0 Then
        For CariString = 1 To 88
          If OutData(CariString) = &HFF Then
            If OutData(CariString + 1) = &H15 Then
              If OutData(CariString + 4) = &H41 Then
              If OutData(CariString + 5) = &H0 Then
                
                  CekEntryPoint = "W32.Virut.O"
               
               ' CekEntryPointRTP = "Positif"
                GoTo LBL_SUDAH
              End If
             End If
            End If
          End If
        Next
       End If
    End If
    
    ' W32.Lethic.AA -- 9/12/2010
      If OutData(1) = &H60 Then
       If OutData(2) = &HC7 Then
        CariString = 0
       
        For CariString = 1 To ISECH(PosisiEP).SizeOfRawData
          If OutData(CariString) = &H61 Then ' POP AD
            ' JMP [ XXXX ]
            If OutData(CariString + 1) = &HE9 Then
             For CariLagi = CariString + 1 To ISECH(PosisiEP).SizeOfRawData
               If OutData(CariLagi) = &HFF Then
                
                If OutData(CariLagi + 1) = &HFF Then
                  
                     CekEntryPoint = "W32.Lethic.AA"
                  
                 ' CekEntryPointRTP = "Positif"
                  GoTo LBL_SUDAH
                 End If
                End If
               Next
              End If
            End If
          'End If
        Next
        End If
      End If
     
     ' Sality.NAO -- 29/11/2010
      If OutData(1) = &H60 Then
        'Call ReadFileADV(hFile, PPhysicEP + 1, 50, OutData())
      '  Debug.Print TATAByte(OutData())
        CariString = 0
        For CariString = 1 To 30
          If OutData(CariString) = &H6A Then ' PUSH EBP
            If OutData(CariString + 1) = &H0 Then ' MOV EBP, ESP
              If OutData(CariString + 2) = &HFF Then
            '    If OutData(CariString + 3) = &H6A Then 'PUSH FFFFFFFFH
            '     If OutData(CariString + 4) = &HFF Then
                    
                      CekEntryPoint = "W32.Sality.NAO"
                   
                    'CekEntryPointRTP = "Positif"
                    GoTo LBL_SUDAH
            '     End If
            '    End If
              End If
            End If
          End If
        Next
      End If
    
    ' Sality.NBA[ESED] -- 27/3/2011
      If OutData(1) = &H40 Then
        'Call ReadFileADV(hFile, PPhysicEP + 1, 50, OutData())
      '  Debug.Print TATAByte(OutData())
        CariString = 0
        For CariString = 1 To 20
          If OutData(CariString) = &HF6 Then ' PUSH EBP
            If OutData(CariString + 1) = &HC5 Then ' MOV EBP, ESP
              If OutData(CariString + 2) = &H38 Then
            '    If OutData(CariString + 3) = &H6A Then 'PUSH FFFFFFFFH
            '     If OutData(CariString + 4) = &HFF Then
                    
                      CekEntryPoint = "W32.Sality.NBA"
                   
                   ' CekEntryPointRTP = "Positif"
                    GoTo LBL_SUDAH
            '     End If
            '    End If
              End If
            End If
          End If
        Next
      End If
' One.Pruduction -- 30/8/2011
      If OutData(1) = &HE8 Then
        'Call ReadFileADV(hFile, PPhysicEP + 1, 50, OutData())
      '  Debug.Print TATAByte(OutData())
        CariString = 0
        For CariString = 1 To 30
          If OutData(CariString) = &H0 Then  ' PUSH EBP
            If OutData(CariString + 1) = &H0 Then  ' MOV EBP, ESP
              If OutData(CariString + 2) = &H0 Then
                If OutData(CariString + 3) = &H0 Then 'PUSH FFFFFFFFH
                 If OutData(CariString + 4) = &H0 Then
                    If OutData(CariString + 5) = &HE9 Then
                    If OutData(CariString + 6) = &H16 Then
                    If OutData(CariString + 7) = &HFE Then
                    If OutData(CariString + 8) = &HFF Then
                    If OutData(CariString + 9) = &HFF Then
                    If OutData(CariString + 10) = &HFF Then
                      CekEntryPoint = "W32.ONE.PRODUC"
                   
                   ' CekEntryPointRTP = "Positif"
                    GoTo LBL_SUDAH
                    End If
                End If
              End If
            End If
          End If
          End If
                 End If
                End If
              End If
            End If
          End If
        Next
      End If

    'virut NBP[eESED] 27/3/2011
      If OutData(1) = &H83 Then
        'Call ReadFileADV(hFile, PPhysicEP + 1, 50, OutData())
      '  Debug.Print TATAByte(OutData())
        CariString = 0
        For CariString = 1 To 30
          If OutData(CariString) = &H3C Then ' PUSH EBP
            If OutData(CariString + 1) = &H24 Then ' MOV EBP, ESP
              If OutData(CariString + 2) = &HFF Then
                If OutData(CariString + 3) = &HF Then   'PUSH FFFFFFFFH
                 If OutData(CariString + 4) = &H84 Then
                    
                      CekEntryPoint = "W32.Virut.NBP"
                   
                    ' = "Positif"
                    GoTo LBL_SUDAH
                 End If
                End If
              End If
            End If
          End If
        Next
      End If
      'virut NBU[eESED] 27/5/2011
      If OutData(1) = &H83 Then
        'Call ReadFileADV(hFile, PPhysicEP + 1, 50, OutData())
      '  Debug.Print TATAByte(OutData())
        CariString = 0
        For CariString = 1 To 30
          If OutData(CariString) = &HC4 Then ' PUSH EBP
            If OutData(CariString + 1) = &HE0 Then ' MOV EBP, ESP
              If OutData(CariString + 2) = &HE8 Then
                If OutData(CariString + 3) = &HBC Then   'PUSH FFFFFFFFH
                 If OutData(CariString + 4) = &HFF Then
                    
                      CekEntryPoint = "W32.Virut.NBU"
                   
                    ' = "Positif"
                    GoTo LBL_SUDAH
                 End If
                End If
              End If
            End If
          End If
        Next
      End If
      
       'SALITY NBA [eESED] 27/5/2011
      If OutData(1) = &H6A Then
        'Call ReadFileADV(hFile, PPhysicEP + 1, 50, OutData())
      '  Debug.Print TATAByte(OutData())
        CariString = 0
        For CariString = 1 To 30
          If OutData(CariString) = &H60 Then ' PUSH EBP
            If OutData(CariString + 1) = &H68 Then ' MOV EBP, ESP
              If OutData(CariString + 2) = &HC8 Then
                If OutData(CariString + 3) = &H20 Then   'PUSH FFFFFFFFH
                 If OutData(CariString + 4) = &H5A Then
                    
                      CekEntryPoint = "W32.Sality.NBA"
                   
                    ' = "Positif"
                    GoTo LBL_SUDAH
                 End If
                End If
              End If
            End If
          End If
        Next
      End If
           
      'W32.Virut.AN 19 mei 2011 '83,3C,24,FE,77,FE,XX,XX,8D,64,24,CC,60,83,EC,DC,E8,D7,FD,FF,FF 
     If OutData(1) = &H83 Then
        'Call ReadFileADV(hFile, PPhysicEP + 1, 50, OutData())
      '  Debug.Print TATAByte(OutData())
        CariString = 0
        For CariString = 1 To 30
          If OutData(CariString) = &H3C Then ' PUSH EBP
            If OutData(CariString + 1) = &H24 Then ' MOV EBP, ESP
              If OutData(CariString + 2) = &HFE Then
                If OutData(CariString + 3) = &H77 Then  'PUSH FFFFFFFFH
                 If OutData(CariString + 4) = &HFE Then
                    If OutData(CariString + 5) = &H8D Then
                     If OutData(CariString + 6) = &H64 Then
                      If OutData(CariString + 7) = &H24 Then
                      If OutData(CariString + 8) = &HCC Then
                      If OutData(CariString + 9) = &H60 Then
                      CekEntryPoint = "W32.Virut.AN"
                   
                    ' = "Positif"
                    GoTo LBL_SUDAH
                     End If
                     End If
                     End If
                    End If
                    End If
                 End If
                End If
              End If
            End If
          End If
        Next
      End If
    
    
    If BerapaText > 1 Then
    If PosisiEP = PosisiText Then
    If InfectChecksum(OutData, 16, 30) = "3DF56" Then
    
       CekEntryPoint = "W32.Resur"    ' YES, ketemu . .
     
    ' CekEntryPointRTP = "Positif"
     GoTo LBL_SUDAH
    End If
    End If
    End If
     
     ' Deteksi variant Virut.NBP dan Virut.O (ESET)
     ' Butuh variant Baru :)
     If Left$(Hex$(ISECH(nSection - 1).Characteristics), 1) = "E" Or Left$(Hex$(ISECH(nSection - 1).Characteristics), 1) = "A" Then  ' Ciri infeksi Virut >> Section terakhir mempunyai attribut 'Writeable' (0x800000000)
       If PosisiEP = nSection - 1 Then
        CariString = 0
        For CariString = 0 To 15
          ' DoEvents
          
            If OutData(CariString) = &HE8 Then  ' Cari HEX : E8 XX 00 XX -> Berisi perintah CALL [Address] yang merupakan alamat dari tubuh virus
              If OutData(CariString + 2) = &H0 Then
                  For CariLagi = (CariString + 5) To 30
                    If OutData(CariLagi) = &H60 Then   ' 60 -> PUSHAD
                        
                            CekEntryPoint = "W32.Virut.NBP.Variant"    ' YES, ketemu . . .
                        
                      '  CekEntryPointRTP = "Positif"
                        GoTo LBL_SUDAH
                     End If
                    
                    If OutData(CariLagi) = &H67 Then    ' 67 64 FF 36 00 00 -> PUSH FS: [0000H]
                        If OutData(CariLagi + 1) = &H64 Then
                          If OutData(CariLagi + 2) = &HFF Then
                            If OutData(CariLagi + 3) = &H36 Then
                              If OutData(CariLagi + 4) = &H0 Then
                                If OutData(CariLagi + 5) = &H0 Then
                                  
                                     CekEntryPoint = "W32.Virut.OQ"    ' YES, ketemu . .
                                  
                                '  CekEntryPointRTP = "Positif"
                                  GoTo LBL_SUDAH
                               End If
                              End If
                            End If
                         End If
                       End If
                    End If
                  Next
              End If
            End If
        Next
     End If
     End If
     
     'If heurvb <> "" Then GoTo LBL_SUDAH
     
     
     
     ' Deteksi Virut.AA -- Update 12/11/2010
       If PosisiEP = nSection - 1 Then
         If InfectChecksum(OutData, 1, 10) = "2B22E" Then
          
               CekEntryPoint = "W32.Virut.AA"    ' YES, ketemu . .
          
          ' CekEntryPointRTP = "Positif"
           GoTo LBL_SUDAH
         End If
       End If
     
     
     ' Deteksi Virut.AI -- Update 13/11/2010
       If PosisiEP = nSection - 1 Then
         If InfectChecksum(OutData, 1, 17) = "496ED" Then
           
              CekEntryPoint = "W32.Virut.AI"    ' YES, ketemu . .
           
         '  CekEntryPointRTP = "Positif"
           GoTo LBL_SUDAH
         End If
       End If
       
       ' Deteksi Virut.AJ -- Update 13/11/2010
       If PosisiEP = nSection - 1 Then
         If InfectChecksum(OutData, 1, 17) = "48C69" Then
          
               CekEntryPoint = "W32.Virut.AJ"   ' YES, ketemu . .
           
         '  CekEntryPointRTP = "Positif"
           GoTo LBL_SUDAH
         End If
       End If
       
       ' Deteksi Virut.AF -- Update 13/11/2010
       If PosisiEP = nSection - 1 Then
         If InfectChecksum(OutData, 8, 18) = "2EE39" Then
           
               CekEntryPoint = "W32.Virut.AF"    ' YES, ketemu . .
           
          ' CekEntryPointRTP = "Positif"
           GoTo LBL_SUDAH
         End If
       End If

       ' Deteksi Virut.BA -- Update 27/11/2010
       If PosisiEP = nSection - 1 Then
         If InfectChecksum(OutData, 1, 17) = "4C3BE" Then
           
              CekEntryPoint = "W32.Virut.BA"    ' YES, ketemu . .
           
          ' CekEntryPointRTP = "Positif"
           GoTo LBL_SUDAH
         End If
       End If


     
     ' Deteksi Virut.C -- Update 6/11/2010
       If PosisiEP = nSection - 1 Then
         If InfectChecksum(OutData, 1, 10) = "2C5D4" Then
           
               CekEntryPoint = "W32.Virut.C"    ' YES, ketemu . .
           
           'CekEntryPointRTP = "Positif"
           GoTo LBL_SUDAH
         End If
       End If
            
       ' Deteksi Virut.Z -- Update 6/11/2010
       If PosisiEP = nSection - 1 Then
         If InfectChecksum(OutData, 1, 5) = "141C6" Then
          
              CekEntryPoint = "W32.Virut.Z"    ' YES, ketemu . .
          
           'CekEntryPointRTP = "Positif"
           GoTo LBL_SUDAH
         End If
       End If
       
       ' Deteksi W32.Sality.E -- 17/12/2010
       If PosisiEP = nSection - 1 Then
         If InfectChecksum(OutData, 1, 18) = "51C8B" Then
           
               CekEntryPoint = "W32.Sality.E"   ' YES, ketemu . .
           
          ' CekEntryPointRTP = "Positif"
           GoTo LBL_SUDAH
         End If
       End If

       ' Deteksi W32.Slugin.A
         If InfectChecksum(OutData, 1, 10) = "2C9DF" Then
           
               CekEntryPoint = "W32.Slugin.A"    ' YES, ketemu . .
           
         '  CekEntryPointRTP = "Positif"
           GoTo LBL_SUDAH
         End If
               
      '
     ' W32.Lamechi.A merubah section .text menjadi memiliki atribut Writeable (0x80000000)
     If CS = "E" Or CS = "A" Then
        If OutData(1) = &H53 Then ' 53 - Push EBX
           If OutData(2) = &H60 Then ' 60 - PUSHAD
                If OutData(3) = &H83 Then       ' 83 -
                  If OutData(4) = &HEC Then     ' EC | } SUB ESP, 00000050H
                    If OutData(5) = &H50 Then   ' 50 -
                        
                            CekEntryPoint = "W32.Lamechi.A"
                        
                       ' CekEntryPointRTP = "Positif"
                        GoTo LBL_SUDAH
                    End If
                  End If
                End If
           End If
        End If
     End If
     
          
     If OutData(1) = &H60 Then
       Call ReadFileADV(hFilePE, ISECH(Posdata).PointerToRawData + 1, ISECH(Posdata).SizeOfRawData, Saldata)
       SalityBack = StrConv(Saldata, vbUnicode)
       If InStr(SalityBack, "KUKU") > 0 Then
         
           CekEntryPoint = "W32.Sality.P"
         
         'CekEntryPointRTP = "Positif"
         GoTo LBL_SUDAH
        End If
        Erase Saldata
     End If
     
    'Heur Sality
    'Ciri - ciri :
    ' - Membuat section baru yang hampir sama dengan section ke 2 (jika dihitung dari 1) hanya berbeda 2 huruf didepanya saja
    ' - SizeOfRawData dari Section baru sama dengan besarnya Section
    NS = TrimNull0(Mid$(StrConv(ISECH(nSection - 1).SectionName(), vbUnicode), 3))
    If nSection > 1 Then NS2 = TrimNull0(Mid$(StrConv(ISECH(1).SectionName(), vbUnicode), 2))
    If NS = NS2 Then
      If OutData(1) = &H60 Then ' PUSHAD
       If ISECH(nSection - 1).VirtualSize = ISECH(nSection - 1).SizeOfRawData Then
          Call ReadFileADV(hFilePE, ISECH(nSection - 1).PointerToRawData + 1, ISECH(nSection - 1).SizeOfRawData, Saldata)
          For CariString = 0 To (ISECH(nSection - 1).SizeOfRawData - 1)
            If Saldata(CariString) = &H8A Then '---8A 44 05 00 = MOV AL,BYTE PTR SS:[EBP+EAX]
              If Saldata(CariString + 1) = &H44 Then
                If Saldata(CariString + 2) = &H5 Then
                   If Saldata(CariString + 3) = &H0 Then
                     For CariLagi = (CariString + 4) To (ISECH(nSection - 1).SizeOfRawData - 1)
                        If Saldata(CariLagi) = &H30 Then '---30 07 = XOR BYTE PTR DS:[EDI],AL
                          If Saldata(CariLagi + 1) = &H7 Then
                            
                                CekEntryPoint = "W32.Sality.X"
                            
                            'CekEntryPointRTP = "Positif"
                            GoTo LBL_SUDAH
                          End If
                        End If
                      Next
                    End If
                End If
              End If
             End If
            Next
    Else
       
      End If
      End If
      End If
   ' End If

    
   
    GoTo LBL_GAGAL
Exit Function
LBL_SUDAH:
    ' Bersih - bersih variabel ahh...
    Call KosongkanMemory(StrPtr(CariString), Len(CariString))
    Call KosongkanMemory(StrPtr(CariLagi), Len(CariLagi))
    Call KosongkanMemory(VarPtr(PPhysicEP), Len(PPhysicEP))
    Call KosongkanMemory(VarPtr(TotalUkuranSection), Len(TotalUkuranSection))
    Call KosongkanMemory(VarPtr(UkuranAsli), Len(UkuranAsli))
    Call KosongkanMemory(VarPtr(RetFunct), Len(RetFunct))
    
    Erase OutData
    Exit Function

LBL_GAGAL:
    Call KosongkanMemory(StrPtr(CariString), Len(CariString))
    Call KosongkanMemory(StrPtr(CariLagi), Len(CariLagi))
    Call KosongkanMemory(VarPtr(PPhysicEP), Len(PPhysicEP))
    Call KosongkanMemory(VarPtr(TotalUkuranSection), Len(TotalUkuranSection))
    Call KosongkanMemory(VarPtr(UkuranAsli), Len(UkuranAsli))
    Call KosongkanMemory(VarPtr(RetFunct), Len(RetFunct))
    CekEntryPoint = ""
    Erase OutData
     
   
End Function



Private Function TrimNull0(sKar As String) As String
   If InStr(sKar, Chr$(0)) > 0 Then TrimNull0 = Left$(sKar, InStr(sKar, Chr$(0)) - 1)
End Function

Public Function TerjemahBhsMesin(teks As String) As String
  Dim convert, hasil As String
  Dim i              As Integer
  
  For i = 1 To Len(teks)
    convert = Mid$(teks, i, 1)
    If hasil = "" Then
      hasil = Asc(convert)
    Else
      hasil = hasil & Asc(Chr$(0)) & Asc(convert)
    End If
    DoEvents
  Next i
 ' Debug.Print hasil
  TerjemahBhsMesin = hasil
End Function

Public Function ConvertBhsMesin(teks As String) As String
  Dim convert, hasil As String
  Dim i              As Integer
  
  For i = 1 To Len(teks)
    convert = Mid$(teks, i, 1)
    If hasil = "" Then
      hasil = Asc(convert)
    Else
      hasil = hasil & Asc(convert)
    End If
    DoEvents
  Next i
  ConvertBhsMesin = hasil
End Function

Private Function PindahkeVar(inpt() As Byte) As String
  Dim i As Integer
  For i = 1 To UBound(inpt) + 1
    DoEvents
    PindahkeVar = PindahkeVar & inpt(i - 1)
  Next i
End Function

Private Function getHeur(inputs As String, str As String) As Boolean
  If InStr(ConvertBhsMesin(inputs), TerjemahBhsMesin(str)) > 0 Then
    getHeur = True
  Else
    getHeur = False
  End If
End Function


