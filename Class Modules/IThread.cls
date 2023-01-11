VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IThread"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Event TaskComplete(ByVal FunctionName As String, Data As Variant)
Event BufferComplete(FileToScan As Long, FolderToScan As Long)
Event ScanComplete(FileterScan As Long)
Event ScanRTPComplete()
Event ScanHOOKComplete()

Event BufferCompleteRTP(FileToScan As Long, FolderToScan As Long)
Event ThisPath(PathApa As String, FileterScan As Long)
Private PathScan As String 'local copy
Private YangPertama As Boolean 'local copy
'Private path As String
Private WithEvents Synchronizer As CLiteTimer
Attribute Synchronizer.VB_VarHelpID = -1

'this is a list to hold tasks to be performed
Private Tasks As Collection
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
Private aLagiJalan  As Boolean
Private Fol_toScan As Long
Private File_toScan As Long
Private File_terScan As Long
Private Fol_toScanRTP As Long
Private File_toScanRTP As Long
Private File_terScanRTP As Long
Public Sub StartScan()
File_terScan = 0
total_size = 0
FolToScan = 0
Call KumpulkanFile(PathScan, YangPertama)
RaiseEvent ScanComplete(File_terScan)
End Sub
Public Sub StartScanRTP()
'File_terScan = 0
Call Scan_fast(PathScan)
'Call sCANFile(PathScan, YangPertama)
RaiseEvent ScanRTPComplete
End Sub
Public Sub BufferScanRTP()
'Fol_toScan = 0
'File_toScan = 0
Call BufferPathRTP(PathScan, YangPertama)
RaiseEvent BufferCompleteRTP(File_toScan, Fol_toScan)
End Sub
Public Sub BufferScan()
File_terScan = 0
'total_size = 0
FolToScan = 0
Call BufferPath(PathScan, YangPertama)
RaiseEvent BufferComplete(File_toScan, Fol_toScan)
End Sub
Public Function StopSearch()
StopSearch = aLagiJalan = False
End Function
Private Sub KumpulkanFile(PathApaIni As String, Optional TheOne As Boolean = False)
On Error Resume Next
Const MAXDWORD = 2 ^ 32
Dim file_size As Double
Dim WFD             As WIN32_FIND_DATA
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
    If TheOne = True Then
        bLagiJalan = True
    End If
    If bLagiJalan = False Then GoTo ERRHD
    PathApaIni = AddSlashW(PathApaIni)
    Print #1, PathApaIni & zSlash
    hFind = FindFirstFileW(StrPtr(PathApaIni & zSlash), VarPtr(WFD))
    If hFind < 1 Then
        GoTo ERRHD
    End If
    
    Do
        If bLagiJalan = False Then Exit Do
        szFileName = TrimNullW(WFD.cFileName) 'hilangkan NullChar paling kanan.
        szFullPath = PathApaIni & szFileName
        bIsFolder = ((WFD.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) = FILE_ATTRIBUTE_DIRECTORY)
        file_size = (WFD.nFileSizeHigh * MAXDWORD) + WFD.nFileSizeLow
        If szFileName <> DOT1 And szFileName <> DOT2 Then
           If bIsFolder = False Then
                FileFound = FileFound + 1
                File_terScan = File_terScan + 1
                'frmMain.PB1.Max = File_terScan
                 frmMain.PB1.Value = File_terScan
                frmMain.lbFileFound = ": " & Right$("00000000" & FileFound, 8)
                If FrmConfig.ck1.Value = 1 Then
                    If isProperFile(szFullPath, "TMP CPL INI SYS LNK VBE HTT EXE DLL VBS VMX TML .DB COM SCR BAT INF TML CMD TXT PIF MSI HTM HTML") = True Then
                        FileCheck = FileCheck + 1
                        'File_terScan = File_terScan + 1
                       ' frmMain.PB1.Value = File_terScan

                        frmMain.lbFileCheck = ": " & Right$("00000000" & FileCheck, 8)
                        'frmScanWith.lbFileCheck.Caption = ": " & Right$("00000000" & FileCheck, 8)
                        VirStatus = False
                        CocokanDataBase szFullPath ' cek virus atau bukan
                      Else
                        FileNotCheck = FileNotCheck + 1
                        frmMain.lbBypass = ": " & Right$("00000000" & FileNotCheck, 8)
                    End If
                Else
                    FileCheck = FileCheck + 1
                   ' File_terScan = File_terScan + 1
                    'frmMain.PB1.Value = File_terScan

                    frmMain.lbFileCheck = ": " & Right$("00000000" & FileCheck, 8)
                    'frmScanWith.lbFileCheck.Caption = ": " & Right$("00000000" & FileCheck, 8)
                    CocokanDataBase szFullPath ' cek virus atau bukan
                End If
                If FrmConfig.ck4.Value = 1 And VirStatus = False Then CheckAttrib szFullPath, False
            Else
             FolToScan = FolToScan + 1
               If FrmConfig.ck4.Value = 1 Then CheckAttrib szFullPath, True
            End If

         '   frmMain.lbStatus.Caption = d_bahasa(15) ' & " [ " & FileFound & " " & j_bahasa(39) & ", " & FileCheck & " " & j_bahasa(38) & " ]"
           ' If bInfo = True And WithBuffer = True Then
              '  frmmain.PB1.value = File_terScan
               ' frmScanWith.PB1.Value = FileFound
           ' End If
   
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
            If bLagiJalan = False Then Exit Do
            RaiseEvent ThisPath(szFullPath & "\*.*", File_terScan)
            Call KumpulkanFile(szFullPath, False)  'enumerasi lagi...
        End If
        
        NextStack = FindNextFileW(hFind, VarPtr(WFD)) 'unicode;
    DoEvents
    frmMain.lbStatus22.Caption = d_bahasa(15) & ": " & FolToScan & " " & j_bahasa(39) & ", " & FileFound & " " & j_bahasa(38) & ", Size: " & _
            Format$(total_size / 1024, "0.000") & " KB, " & _
            Format$(total_size / 1024 / 1024, "0.0") & " MB, " & _
            Format$(total_size / 1024 / 1024 / 1024, "0.00") & " GB"
    Loop While NextStack
    
    Call FindClose(hFind)
ERRHD:
    If err.Number > 0 Then
        err.Clear
    End If
End Sub
Private Function BufferPath(szNamaTarget As String, Optional ByVal YangPertama As Boolean = False) As Boolean
Dim WFD             As WIN32_FIND_DATA
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
        aLagiJalan = True
    End If
    If aLagiJalan = False Then GoTo ERRHD
    szNamaTarget = AddSlashW(szNamaTarget)
    hFind = FindFirstFileW(StrPtr(szNamaTarget & zSlash), VarPtr(WFD))
    If hFind < 1 Then
       GoTo ERRHD
    End If
    Do
        If aLagiJalan = False Then GoTo keluar
        szFileName = TrimNullW(WFD.cFileName) 'hilangkan NullChar paling kanan.
        szFullPath = szNamaTarget & szFileName
        bIsFolder = ((WFD.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) = FILE_ATTRIBUTE_DIRECTORY)
        If szFileName <> DOT1 And szFileName <> DOT2 Then
        Else
            bIsFolder = False
        End If
        
        If bIsFolder = True Then
           If aLagiJalan = False Then GoTo keluar
           Fol_toScan = Fol_toScan + 1
           Call BufferPath(szFullPath, False)  'enumerasi lagi...
        Else
           If aLagiJalan = False Then GoTo keluar
           If PathIsDirectory(StrPtr(szFullPath)) = 0 Then File_toScan = File_toScan + 1
        End If
        NextStack = FindNextFileW(hFind, VarPtr(WFD)) 'unicode;
    DoEvents
    Loop While NextStack
    frmMain.lbStatus.Caption = j_bahasa(37) & " [ " & Fol_toScan & " " & j_bahasa(39) & ", " & File_toScan & " " & j_bahasa(38) & " ]"

    Call FindClose(hFind)
ERRHD:
    If err.Number > 0 Then
        err.Clear
    End If
Exit Function
keluar:
WithBuffer = False
End Function
Private Function Scan_fast(PathScan As String)
Dim FileNow As String
Dim FSO As Object
Dim SFILE As Object
Set FSO = Nothing
On Error GoTo keluar:
    'lvMalware.ListItems.Clear
    Set FSO = CreateObject("Scripting.FileSystemObject")
    For Each SFILE In FSO.GetFolder(PathScan).Files
        DoEvents
        FileNow = SFILE
    If IsFileX(SFILE) = True Then
     'If isProperFile(FileNow, 1000000) = True Then
     CocokanDataBaseRTP (FileNow)
             
    'End If
     End If
    Next
keluar:
End Function

Private Function BufferPathRTP(szNamaTarget As String, Optional ByVal YangPertama As Boolean = False) As Boolean
Dim WFD             As WIN32_FIND_DATA
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
        aLagiJalan = True
    End If
    If aLagiJalan = False Then GoTo ERRHD
    szNamaTarget = AddSlashW(szNamaTarget)
    hFind = FindFirstFileW(StrPtr(szNamaTarget & zSlash), VarPtr(WFD))
    If hFind < 1 Then
       GoTo ERRHD
    End If
    File_toScanRTP = 0
    Do
        If aLagiJalan = False Then GoTo keluar
        szFileName = TrimNullW(WFD.cFileName) 'hilangkan NullChar paling kanan.
        szFullPath = szNamaTarget & szFileName
        bIsFolder = ((WFD.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) = FILE_ATTRIBUTE_DIRECTORY)
        If szFileName <> DOT1 And szFileName <> DOT2 Then
        Else
            bIsFolder = False
        End If
        
        If bIsFolder = True Then
           If aLagiJalan = False Then GoTo keluar
           Fol_toScanRTP = Fol_toScanRTP + 1
           'Call BufferPath(szFullPath, False)  'enumerasi lagi...
        Else
           If aLagiJalan = False Then GoTo keluar
           If PathIsDirectory(StrPtr(szFullPath)) = 0 Then File_toScanRTP = File_toScanRTP + 1
        End If
        NextStack = FindNextFileW(hFind, VarPtr(WFD)) 'unicode;
    DoEvents
    Loop While NextStack
    FrmRTP.lbFileCheck.Caption = j_bahasa(43) & ", " & File_toScanRTP & " " & j_bahasa(38)

    Call FindClose(hFind)
ERRHD:
    If err.Number > 0 Then
        err.Clear
    End If
Exit Function
keluar:
WithBuffer = False
End Function
Private Sub sCANFile(PathApaIni As String, Optional TheOne As Boolean = False)
On Error Resume Next
Dim WFD             As WIN32_FIND_DATA
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
    If TheOne = True Then
        bLagiJalan = True
    End If
    If bLagiJalan = False Then GoTo ERRHD
    PathApaIni = AddSlashW(PathApaIni)
    Print #1, PathApaIni & zSlash
    hFind = FindFirstFileW(StrPtr(PathApaIni & zSlash), VarPtr(WFD))
    If hFind < 1 Then
        GoTo ERRHD
    End If
    'FileCheck = 0
    Do
        If bLagiJalan = False Then Exit Do
        szFileName = TrimNullW(WFD.cFileName) 'hilangkan NullChar paling kanan.
        szFullPath = PathApaIni & szFileName
        bIsFolder = ((WFD.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) = FILE_ATTRIBUTE_DIRECTORY)
        If szFileName <> DOT1 And szFileName <> DOT2 Then
           If bIsFolder = False Then
             
                If FrmConfig.ck1.Value = 1 Then
                    If isProperFile(szFullPath, "TMP CPL SYS INI LNK VBE HTT EXE DLL VBS VMX TML .DB COM SCR BAT INF TML CMD TXT PIF MSI HTM HTML") = True Then
                   ' FileCheck = FileCheck + 1
                        CocokanDataBaseRTP szFullPath ' cek virus atau bukan
                     End If
                Else
                   'FileCheck = FileCheck + 1
                    CocokanDataBaseRTP szFullPath ' cek virus atau bukan
                End If
                'If frmMain.ck4.Value = 1 And VirStatus = False Then CheckAttrib szFullPath, False
      ' frmRTP.lbFileCheck = "Checked: " & right$("00000000" & FileCheck, 8) & " File(s)"

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

Function StartTask(ByVal strFunctionName As String, ByVal blnAsync As Boolean, ByVal enmCallType As VbCallType, ParamArray varData() As Variant) As Variant

    If blnAsync Then
        
        'add it to the task list
        Tasks.Add strFunctionName
        'check to make sure timer is running
        If Not Synchronizer.Enabled Then
            'set timer to respond/tick within 1 millisecond
            Synchronizer.Interval = 1
            Synchronizer.Enabled = True
        End If
    Else
        'if not asynchronous then we act like a normal function call
        If enmCallType = VbMethod Then
            StartTask = CallByName(Me, strFunctionName, enmCallType, varData(0))
        Else
            CallByName Me, strFunctionName, enmCallType, varData(0)
        End If
    End If
End Function

Private Sub Class_Initialize()

 'initialize the tasklist
 Set Tasks = New Collection
 
 'initialize the timer
 Set Synchronizer = New CLiteTimer
 
End Sub

Private Sub Synchronizer_Timer()
   'the timer has fired so we must check the tasklist to see if any tasks
   'are waiting to be done
    If Tasks.Count > 0 Then
        Synchronizer.Enabled = False
        'do the first added task in the list
        CallByName Me, Tasks(1), VbMethod
        'remove it
        Tasks.Remove 1
    End If
    'if no tasks left then disable timer
   ' If BERHENTI = False Then
    If Tasks.Count > 0 Then
        Synchronizer.Enabled = True
   ' End If
   'Else
   End If
End Sub

Public Property Let path(ByVal vData As String)
    PathScan = vData
End Property
Public Property Let TheFirst(ByVal vData As Boolean)
    YangPertama = vData
End Property
Public Property Let LagiJalan(ByVal vData As Boolean)
    bLagiJalan = vData
End Property
Public Property Let bufJalan(ByVal vData As Boolean)
    aLagiJalan = vData
End Property
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


