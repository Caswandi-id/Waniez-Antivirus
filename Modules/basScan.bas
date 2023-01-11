Attribute VB_Name = "basScan"
' Module untuk penanganan pencarian file

Option Explicit
Private Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long 'Tanpa fungsi LockWindowUpdate

Private Const MAX_PATH  As Long = 260
Private Const MAX_BUF   As Long = 512

Private Const FILE_ATTRIBUTE_READONLY = &H1     '
Private Const FILE_ATTRIBUTE_HIDDEN = &H2     '
Private Const FILE_ATTRIBUTE_SYSTEM = &H4     '
Private Const FILE_ATTRIBUTE_DIRECTORY = &H10     'folder.
Private Const FILE_ATTRIBUTE_ARCHIVE = &H20     '
Private Const FILE_ATTRIBUTE_DEVICE = &H40     '
Private Const FILE_ATTRIBUTE_NORMAL = &H80     '
Private Const FILE_ATTRIBUTE_TEMPORARY = &H100     '
Private Const FILE_ATTRIBUTE_SPARSE_FILE = &H200     '
Private Const FILE_ATTRIBUTE_REPARSE_POINT = &H400     '
Private Const FILE_ATTRIBUTE_COMPRESSED = &H800     'terkompres ntfs.
Private Const FILE_ATTRIBUTE_OFFLINE = &H1000     '
Private Const FILE_ATTRIBUTE_NOT_CONTENT_INDEXED = &H2000     'tidak masuk dalam index pencarian file.
Private Const FILE_ATTRIBUTE_ENCRYPTED = &H4000     'enkripsi ntfs.
Private Const FILE_ATTRIBUTE_VIRTUAL = &H10000     'device virtual;

Private Type FILETIME
    dwLowDateTime       As Long
    dwHighDateTime      As Long
End Type
Private Type WIN32_FIND_DATA
    dwFileAttributes    As Long 'FILE_ATTRIBUTES
    ftCreationTime      As FILETIME
    ftLastAccessTime    As FILETIME
    ftLastWriteTime     As FILETIME
    nFileSizeHigh       As Long
    nFileSizeLow        As Long
    dwReserved0         As Long
    dwReserved1         As Long
    cFileName           As String * MAX_PATH '<>MAX_BUF
    cAlternate          As String * 14
End Type

Private Declare Function FindFirstFileW Lib "kernel32" (ByVal lpFileName As Long, ByVal lpFindFileData As Long) As Long
Private Declare Function FindNextFileW Lib "kernel32" (ByVal hFindFile As Long, ByVal lpFindFileData As Long) As Long
Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Boolean
Public total_size As Double
Private bLagiJalan  As Boolean
Public BufferFile As Integer
Public BufferFolder As Integer
Dim filekah As Boolean
Public Hidden As Boolean

'Mendapatkan directory induk dari path
Public Function GetDir(PathFile As String) As String
Dim i As Long
Dim CutDirString As Long
    
    For i = 1 To Len(PathFile)
        If Mid$(PathFile, i, 1) = "\" Then CutDirString = i
    Next i
    GetDir = Left$(PathFile, CutDirString)
End Function

Public Sub KumpulkanFile(ByVal szNamaTarget As String, lbFile As Label, bInfo As Boolean, Optional ByVal YangPertama As Boolean = False)
On Error Resume Next
Const MAXDWORD = 2 ^ 32
Dim file_size As Double
Dim WFD             As WIN32_FIND_DATA, Cont As Integer
Dim hFind           As Long
Dim NextStack       As Long
Dim zSlash          As String
Dim szFullPath      As String
Dim szFileName      As String
Dim bIsFolder       As Boolean
Dim DOT1        As String
Dim DOT2        As String
    DOT1 = ChrW$(46)    '"." dos dir$ 1.
    DOT2 = DOT1 & DOT1  '".." dos dir$ 2.
    zSlash = ChrW$(42) '"\"
    
    If YangPertama = True Then
        bLagiJalan = True
    End If
        
    If bLagiJalan = False Then GoTo ERRHD
    szNamaTarget = AddSlashW(szNamaTarget)
    hFind = FindFirstFileW(StrPtr(szNamaTarget & zSlash), VarPtr(WFD))
    If hFind < 1 Then
        GoTo ERRHD
    End If
   '    LockWindowUpdate frmmain.hwnd

    Do
        If bLagiJalan = False Then Exit Do
        szFileName = TrimNullW(WFD.cFileName) 'hilangkan NullChar paling kanan.
        szFullPath = szNamaTarget & szFileName
        bIsFolder = ((WFD.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) = FILE_ATTRIBUTE_DIRECTORY)
     file_size = (WFD.nFileSizeHigh * MAXDWORD) + WFD.nFileSizeLow
        If szFileName <> DOT1 And szFileName <> DOT2 Then
           If bIsFolder = False Then
                 
                FileFound = FileFound + 1
                frmMain.lbFileFound = Right$(FileFound, 8)
                If FrmConfig.ck1.Value = 1 Then
                    If isProperFile(szFullPath, "TMP CPL SYS LNK VBE HTT EXE DLL VBS VMX TML .DB COM SCR BAT INF TML CMD TXT PIF MSI HTM HTML") = True Then
                        FileCheck = FileCheck + 1
                        frmMain.lbFileCheck = Right$(FileCheck, 8)
                        'frmScanWith.lbFileCheck.Caption = ": " & Right$("00000000" & FileCheck, 8)
                        VirStatus = False
                        CocokanDataBase szFullPath ' cek virus atau bukan
                      Else
                        FileNotCheck = FileNotCheck + 1
                        frmMain.lbBypass = Right$(FileNotCheck, 8)
                    End If
                Else
                    FileCheck = FileCheck + 1
                    frmMain.lbFileCheck = Right$(FileCheck, 8)
                    'frmScanWith.lbFileCheck.Caption = ": " & Right$("00000000" & FileCheck, 8)
                    CocokanDataBase szFullPath ' cek virus atau bukan
                End If
                If FrmConfig.ck4.Value = 1 Then HiddenFolder szNamaTarget, szFileName, True
            Else
               If FrmConfig.ck4.Value = 1 Then HiddenFolder szNamaTarget, szFileName, False
            
            End If
            
            If bInfo = True And WithBuffer = True Then
                frmMain.PB1.Value = FileFound
                'frmScanWith.PB1.Value = FileFound
            End If
   
        Else
            bIsFolder = False
        End If
        
        If file_size > 0 Then
             total_size = total_size + file_size
              '  End If
            Else
'                    lst.AddItem start_dir & fname
         End If
        If bIsFolder = True Then
        FolToScan = FolToScan + 1
            If bLagiJalan = False Then Exit Do
            '====
            If Len(GetDir(szFullPath) & GetFileName(szFullPath)) > 70 Then 'jika panjang nama file > 50
                  If Len(GetFileName(szFullPath)) < 15 Then
                    lbFile.Caption = Mid$(GetDir(szFullPath), 1, InStr(4, GetDir(szFullPath), "\")) & "..." & "\" & GetFileName(szFullPath) & "\*.*"
                  Else
                    lbFile.Caption = Mid$(GetDir(szFullPath), 1, InStr(4, GetDir(szFullPath), "\")) & "..." & "\" & "..." & Right(GetFileName(szFullPath) & "\*.*", 15)
                  End If
                 Else 'jika tidak
                 lbFile.Caption = GetDir(szFullPath) & GetFileName(szFullPath)
                End If ' akhir jika panjangfile > 50
            Call KumpulkanFile(szFullPath, lbFile, bInfo, False) 'enumerasi lagi...
        End If
        
        NextStack = FindNextFileW(hFind, VarPtr(WFD)) 'unicode;
    'Sleep 100
    DoEvents
    Loop While NextStack
    frmMain.lbStatus22.Caption = FolToScan & " " & j_bahasa(39) & ", " & FileFound & " " & j_bahasa(38) & ", Size: " & _
            FormatSize(total_size) ' / 1024, "0.000") & " KB, " & _
            Format$(total_size / 1024 / 1024, "0.0") & " MB, " & _
            Format$(total_size / 1024 / 1024 / 1024, "0.00") & " GB"
    
 '   LockWindowUpdate 0
    Call FindClose(hFind)


ERRHD:
    If err.Number > 0 Then
        err.Clear
    End If
End Sub
Private Sub AnalisisGetHidden(path As String, SFILE As String, filekah As Boolean)
    'DrawIco path & file, frmMain.picBufferw, ricnSmall

    AddList frmMain.lvHidden, path & SFILE, frmMain.picBufferw, 2, SFILE, Array(IIf(filekah = True, "Hidden File", "Hidden Folder"), path & SFILE)
   ' AutoLst frmMain.lvHidden ': lblHid = frmMain.lvHidden.ListItems.Count
    nHiddenObj = nHiddenObj + 1
    frmMain.lvHidden.ListItems.Item(nHiddenObj).Cut = True
    frmMain.lbHidden.Caption = Right$(nHiddenObj, 6) & " "

End Sub
Private Sub HiddenFolder(path As String, SFILE As String, filekah As Boolean)
  'If ValidFile(sfile) = True Then
  'filekah = True
    If filekah = True Then
    If IsHidden(path & SFILE) Then Call AnalisisGetHidden(path, SFILE, True)
    Else
    If IsHidden(path & SFILE) Then Call AnalisisGetHidden(path, SFILE, False)
    End If
  '  End If
End Sub
Public Function IsHidden(path As String) As Boolean
If GetFileName(path) = "Thumbs.db" Or GetFileName(path) = "Desktop.ini" Then GoTo errr
If GetFileAttributes(StrPtr(path)) And FILE_ATTRIBUTE_HIDDEN Then IsHidden = True Else IsHidden = False
errr:
End Function

Public Function NormalHidden(path As String) As Boolean
NormalHidden = SetFileAttributes(StrPtr(path), FILE_ATTRIBUTE_NORMAL)
End Function
Public Function FormatSize(ByVal SIZE As Currency) As String
    Const Kilobyte As Currency = 1024@
    Const HundredK As Currency = 102400@
    Const ThousandK As Currency = 1024000@
    Const Megabyte As Currency = 1048576@
    Const HundredMeg As Currency = 104857600@
    Const ThousandMeg As Currency = 1048576000@
    Const Gigabyte As Currency = 1073741824@
    Const Terabyte As Currency = 1099511627776@
    
    If SIZE < Kilobyte Then
        FormatSize = Int(SIZE) & " bytes"
    ElseIf SIZE < HundredK Then
        FormatSize = Format(SIZE / Kilobyte, "#.0") & " KB [Kilo Byte]"
    ElseIf SIZE < ThousandK Then
        FormatSize = Int(SIZE / Kilobyte) & " KB [Kilo Byte]"
    ElseIf SIZE < HundredMeg Then
        FormatSize = Format(SIZE / Megabyte, "#.0") & " MB [Mega Byte]"
    ElseIf SIZE < ThousandMeg Then
        FormatSize = Int(SIZE / Megabyte) & " MB [Mega Byte]"
    ElseIf SIZE < Terabyte Then
        FormatSize = Format(SIZE / Gigabyte, "#.00") & " GB [Giga Byte]"
    Else
        FormatSize = Format(SIZE / Terabyte, "#.00") & " TB [Tera Byte]"
    End If
End Function


Public Function BufferPath(szNamaTarget As String, Optional ByVal YangPertama As Boolean = False) As Boolean
Dim WFD             As WIN32_FIND_DATA
Dim hFind           As Long
Dim NextStack       As Long
Dim zSlash          As String
Dim szFullPath      As String

Dim szFileName      As String
Dim bIsFolder       As Boolean
Dim DOT1        As String
Dim DOT2        As String
DoEvents
    DOT1 = ChrW$(46)    '"." dos dir$ 1.
    DOT2 = DOT1 & DOT1  '".." dos dir$ 2.
    zSlash = ChrW$(42) '"\"
    If YangPertama = True Then
        bLagiJalan = True
    End If
    
    If bLagiJalan = False Then GoTo ERRHD
    
    szNamaTarget = AddSlashW(szNamaTarget)
    hFind = FindFirstFileW(StrPtr(szNamaTarget & zSlash), VarPtr(WFD))
    If hFind < 1 Then
       GoTo ERRHD
    End If
    Do
    DoEvents
        If bLagiJalan = False Then GoTo keluar
        szFileName = TrimNullW(WFD.cFileName) 'hilangkan NullChar paling kanan.
        szFullPath = szNamaTarget & szFileName
        bIsFolder = ((WFD.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) = FILE_ATTRIBUTE_DIRECTORY)
       
        If szFileName <> DOT1 And szFileName <> DOT2 Then
       
       'szFullPath
       'eventa disini
        Else
            bIsFolder = False
        End If
        
        If bIsFolder = True Then
           If bLagiJalan = False Then GoTo keluar
           FolToScan = FolToScan + 1
           Call BufferPath(szFullPath, False)  'enumerasi lagi...
        Else
           If bLagiJalan = False Then GoTo keluar
           If PathIsDirectory(StrPtr(szFullPath)) = 0 Then FileToScan = FileToScan + 1
        End If
        NextStack = FindNextFileW(hFind, VarPtr(WFD)) 'unicode;
    DoEvents
    Loop While NextStack
    
    frmMain.lbStatus.Caption = j_bahasa(37) & " ! [ " & FolToScan & " " & j_bahasa(39) & ", " & FileToScan & " " & j_bahasa(38) & " ]"
    'frmScanWith.Labelfile.Caption = j_bahasa(37) & " ! [ " & FolToScan & " " & j_bahasa(39) & ", " & FileToScan & " " & j_bahasa(38) & " ]"
    Call FindClose(hFind)
ERRHD:
    If err.Number > 0 Then
        err.Clear
    End If
Exit Function
keluar:
WithBuffer = False
End Function



' Pengganti GetFile yang pake FSO
Public Function GetFile(sPath As String, ArrFile() As String) As Long
Dim WFD             As WIN32_FIND_DATA
Dim hFind           As Long
Dim NextStack       As Long
Dim cNumber         As Long
Dim zSlash          As String
Dim szFullPath      As String
Dim szFileName      As String
Dim bIsFolder       As Boolean
Dim DOT1        As String
Dim DOT2        As String

ReDim ArrFile(1000) As String ' max 1001 file

On Error GoTo ERRHD

    DOT1 = ChrW$(46)    '"." dos dir$ 1.
    DOT2 = DOT1 & DOT1  '".." dos dir$ 2.
    zSlash = ChrW$(42) '"\"
    sPath = AddSlashW(sPath)
    hFind = FindFirstFileW(StrPtr(sPath & zSlash), VarPtr(WFD))
    If hFind < 1 Then
       GoTo ERRHD
    End If
    Do
        szFileName = TrimNullW(WFD.cFileName) 'hilangkan NullChar paling kanan.
        szFullPath = sPath & szFileName
        bIsFolder = ((WFD.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) = FILE_ATTRIBUTE_DIRECTORY)
    
        If szFileName <> DOT1 And szFileName <> DOT2 Then
           If ValidFile(szFullPath) = True Then
              ArrFile(cNumber) = szFullPath
              'MsgBox ArrFile(cNumber)
              cNumber = cNumber + 1
           End If
        End If
        NextStack = FindNextFileW(hFind, VarPtr(WFD)) 'unicode;
    
    DoEvents
    Loop While NextStack
        
    GetFile = cNumber
    Call FindClose(hFind)
ERRHD:
    If err.Number > 0 Then
        err.Clear
    End If
End Function

Public Sub ScanRTP(ByRef sPath As String)
Dim WFD             As WIN32_FIND_DATA
Dim hFind           As Long
Dim NextStack       As Long
Dim zSlash          As String
Dim szFullPath      As String
Dim szFileName      As String
Dim bIsFolder       As Boolean
Dim DOT1        As String
Dim DOT2        As String

On Error GoTo ERRHD

    DOT1 = ChrW$(46)    '"." dos dir$ 1.
    DOT2 = DOT1 & DOT1  '".." dos dir$ 2.
    zSlash = ChrW$(42) '"\"
    sPath = AddSlashW(sPath)
    hFind = FindFirstFileW(StrPtr(sPath & zSlash), VarPtr(WFD))
    If hFind < 1 Then
       GoTo ERRHD
    End If
    Do
        szFileName = TrimNullW(WFD.cFileName) 'hilangkan NullChar paling kanan.
        szFullPath = sPath & szFileName
        bIsFolder = ((WFD.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) = FILE_ATTRIBUTE_DIRECTORY)
    
        If szFileName <> DOT1 And szFileName <> DOT2 Then
           If bIsFolder = False Then
           
                If FrmConfig.ck1.Value = 1 Then
                    If isProperFile(szFullPath, "TMP MDB CPL SYS LNK EXE DLL VBS VMX TML .DB COM SCR INF TML PIF MSI HTM HTML") = True Then
                         CocokanDataBaseRTP szFullPath ' cek virus atau bukan
                      Else
                        'FileNotCheck = FileNotCheck + 1
                        'frmMain.lbBypass = ": " & Right$("00000000" & FileNotCheck, 8)
                    End If
                Else
                    'FileCheck = FileCheck + 1
                    'frmMain.lbFileCheck = ": " & Right$("00000000" & FileCheck, 8)
                    'frmScanWith.lbFileCheck.Caption = ": " & Right$("00000000" & FileCheck, 8)
                    CocokanDataBaseRTP szFullPath ' cek virus atau bukan
                End If
           'CocokanDataBaseRTP szFullPath
           End If
        End If
        NextStack = FindNextFileW(hFind, VarPtr(WFD)) 'unicode;
    
    'DoEvents
    Loop While NextStack
     
    Call FindClose(hFind)
ERRHD:
    If err.Number > 0 Then
        err.Clear
    End If

End Sub


' Scan Semua file yang ada di root drive 2 dan 3
Public Function ScanRootDrive(lblFile As Label)
On Error Resume Next
Dim lstDrive() As String
Dim lstFile()  As String
Dim nDrive  As Long
Dim nFileX  As Long
Dim nTurn   As Long
Dim nTurn2  As Long

nDrive = GetDrive(lstDrive())

For nTurn = 1 To nDrive
    nFileX = GetFile(lstDrive(nTurn), lstFile)
    For nTurn2 = 1 To nFileX
        If BERHENTI = True Then Exit Function
        If ValidFile(lstFile(nTurn2 - 1)) = True Then
           FileFound = FileFound + 1
           FileCheck = FileCheck + 1
           lblFile.Caption = lstFile(nTurn2 - 1)
           If CekAutorun(lstFile(nTurn2 - 1)) = False Then
              ' cek autotun dulu sebelumnya
              If isProperFile(lstFile(nTurn2 - 1), "TMP CPL SYS LNK EXE DLL VBS VMX TML .DB COM SCR INF TML PIF MSI HTM HTML") = True Then
                     CocokanDataBase lstFile(nTurn2 - 1)
                      Else
                        'FileNotCheck = FileNotCheck + 1
                       ' frmMain.lbBypass = Right$(FileNotCheck, 8)
                    End If
              'CocokanDataBase lstFile(nTurn2 - 1)
           End If
        End If
    DoEvents
    Next
    nTurn2 = 1
Next

Erase lstDrive
Erase lstFile
End Function
Public Sub ScanRTPHook(ByRef sPath As String)
Dim WFD             As WIN32_FIND_DATA
Dim hFind           As Long
Dim NextStack       As Long
Dim zSlash          As String
Dim szFullPath      As String
Dim szFileName      As String
Dim bIsFolder       As Boolean
Dim DOT1        As String
Dim DOT2        As String

On Error GoTo ERRHD

    DOT1 = ChrW$(46)    '"." dos dir$ 1.
    DOT2 = DOT1 & DOT1  '".." dos dir$ 2.
    zSlash = ChrW$(42) '"\"
    sPath = AddSlashW(sPath)
    hFind = FindFirstFileW(StrPtr(sPath & zSlash), VarPtr(WFD))
    If hFind < 1 Then
       GoTo ERRHD
    End If
    Do
        szFileName = TrimNullW(WFD.cFileName) 'hilangkan NullChar paling kanan.
        szFullPath = sPath & szFileName
        bIsFolder = ((WFD.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) = FILE_ATTRIBUTE_DIRECTORY)
    
        If szFileName <> DOT1 And szFileName <> DOT2 Then
           If bIsFolder = False Then
             If isProperFile(szFullPath, "TMP EXE CPL DLL .DM MDB VBS INF VMX COM SCR PIF HTM HTML") = True Then
                 CocokanDataBaseHook szFullPath
             End If
           End If
        End If
        NextStack = FindNextFileW(hFind, VarPtr(WFD)) 'unicode;
    
    DoEvents
    Loop While NextStack
     
    Call FindClose(hFind)
ERRHD:
    If err.Number > 0 Then
        err.Clear
    End If

End Sub
Private Function AddSlashW(ByVal StrInW As String) As String 'OK
On Error Resume Next    'tambah "\" di sebelah kanan string unicode.
    If Right$(StrInW, 1) <> ChrW$(92) Then
        AddSlashW = StrInW & ChrW$(92) 'unicode string;
    Else
        AddSlashW = StrInW
    End If
    err.Clear
End Function

Private Function TrimNullW(ByVal StInpW As String) As String 'OK
On Error Resume Next
Dim AlignW As Long: AlignW = InStr(StInpW, ChrW$(0))
    If AlignW > 0 Then
        TrimNullW = Left$(StInpW, AlignW - 1) 'unicode string;
    Else
        TrimNullW = StInpW
    End If
End Function

Private Function PotongTampilanKar(sKar As String, nLimit As Byte) As String
If Len(sKar) >= nLimit Then PotongTampilanKar = Left$(sKar, nLimit - 30) & "...\" & GetFileName(sKar) Else PotongTampilanKar = sKar
End Function


Public Sub StopKumpulkan()
On Error Resume Next
    bLagiJalan = False
End Sub

