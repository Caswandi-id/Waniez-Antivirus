Attribute VB_Name = "basFile"
' ########################################################
' Module untuk penanganan akses file dan folder
'
'

Declare Function GetVolumeInformationW Lib "kernel32" (ByVal pv_lpRootPathName As Long, ByVal pv_lpVolumeNameBuffer As Long, ByVal nVolumeNameSize As Long, lpVolumeSerialNumber As Long, lpMaximumComponentLength As Long, lpFileSystemFlags As Long, ByVal pv_lpFileSystemNameBuffer As Long, ByVal nFileSystemNameSize As Long) As Long

Private Declare Function DeleteFile Lib "kernel32" Alias "DeleteFileW" (ByVal lpFileName As Long) As Long
Declare Function SetFileAttributes Lib "kernel32" Alias "SetFileAttributesW" (ByVal lpFileName As Long, ByVal dwFileAttributes As Long) As Long
Declare Function GetFileAttributes Lib "kernel32" Alias "GetFileAttributesW" (ByVal lpFileName As Long) As Long

Declare Function PathIsDirectory Lib "shlwapi.dll" Alias "PathIsDirectoryW" (ByVal pszPath As Long) As Long
Declare Function PathFileExists Lib "shlwapi.dll" Alias "PathFileExistsW" (ByVal pszPath As Long) As Long

Private Declare Function CopyFile Lib "kernel32" Alias "CopyFileW" (ByVal lpExistingFileName As Long, ByVal lpNewFileName As Long, ByVal bFailIfExists As Long) As Long
Private Declare Function CreateDirectory Lib "kernel32" Alias "CreateDirectoryW" (ByVal lpPathName As Long, lpSecurityAttributes As Long) As Long

Private Type SECURITY_ATTRIBUTES
        nLength As Long
        lpSecurityDescriptor As Long
        bInheritHandle As Long
End Type

Dim hKunci(300)  As Long ' max 301 file yang dikunci
Dim lKunci       As Long


Dim RDF As New classFile


' Untuk membaca file penuh dari ofset 00-akhir
Public Function ReadUnicodeFile(ByRef sFilePath As String) As String
On Error GoTo TERAKHIR
Dim zFileName   As String
Dim hFile       As Long 'nomor file handle, valid jika > 0;
Dim nFileLen    As Long
Dim nOperation  As Long

    zFileName = sFilePath
     hFile = RDF.VbOpenFile(zFileName, FOR_BINARY_ACCESS_READ, LOCK_NONE)
    'selanjutnya:
    If hFile > 0 Then 'jika berhasil membuka file hFile/Handel file > 0;
        'cari tahu ukuran filenya:
        nFileLen = RDF.VbFileLen(hFile)
        If nFileLen > 60000000 Then Exit Function ' nyerah aja klo file-nya lebih dari 130.000.000 B
        Dim bufdata()   As Byte
            nOperation = RDF.VbReadFileB(hFile, 1, nFileLen, bufdata)
            ReadUnicodeFile = StrConv(bufdata, vbUnicode) ' Ralat pada buku tadinya Cstr(buffdata)
            RDF.VbCloseFile hFile 'harus tutup handle ke file setelah mengaksesnya !!!
        Erase bufdata()
    Else 'jika gagal membuka file;
            GoTo TERAKHIR
    End If
Exit Function

TERAKHIR:
End Function
Public Function ReadFileADV(hFile As Long, nStart As Long, nLenght As Long, bufdata() As Byte) As String
Dim nOperation  As Long
 nOperation = RDF.VbReadFileB(hFile, nStart, nLenght, bufdata)
End Function
' Kusus baca file dengan kondisi tertentu (untuk optimalisasi pemindaian)
Public Function ReadUnicodeFile2(hFile As Long, nStart As Long, nLenght As Long, ByRef DataOut() As Byte) As String
Dim nOperation  As Long
  
  nOperation = RDF.VbReadFileB(hFile, nStart, nLenght, DataOut)
  'ReadUnicodeFile2 = StrConv(DataOut, vbUnicode)
  
End Function


Public Function WriteUnicodeFile(spath As String, nStart As Long, bWriteData() As Byte) As Boolean
Dim hFile       As Long 'nomor file handle, valid jika > 0;
Dim nFileLen    As Long
Dim nOperation  As Long


    hFile = RDF.VbOpenFile(spath, FOR_BINARY_ACCESS_READ_WRITE, LOCK_NONE)
    nFileLen = UBound(bWriteData) + 1
    
    If hFile > 0 Then
       RDF.VbWriteFileB hFile, nStart, nFileLen, bWriteData
       WriteUnicodeFile = True
       RDF.VbCloseFile (hFile)
    Else
       WriteUnicodeFile = False
    End If

End Function

' Menulis File ANSI (untuk backup ajah)
Private Sub WriteAnsiFile(spath As String, sContent As String)
Open spath For Binary As #1
    Put #1, , sContent
Close #1
End Sub

' He9x... padahal pke ANSI, tapi coba nembus UNI -> Fungsi Write File Unicode Sudah Ada :))
Public Sub WriteFileUniSim(sPathUni As String, sContent As String)
Dim TMP As String
    TMP = GetSpecFolder(USER_DOC) & "\TMP.TMP.$$" ' tmp-nya di My DOC aj
    WriteAnsiFile TMP, sContent
    CopiFile TMP, sPathUni, True
End Sub

Public Function CopiFile(sTarget As String, sDest As String, bCut As Boolean)
On Error Resume Next
CopyFile StrPtr(sTarget), StrPtr(sDest), 0 ' copi mode overwrite
If bCut = True Then
   HapusFile sTarget
End If
End Function

Public Function BuatFolder(sFolder As String)
    CreateDirectory StrPtr(sFolder), VarPtr(SECURITY_ATTRIBUTES) ' kayanya salah
End Function


Public Function KunciFile(spath As String) As Boolean ' Untuk mengunci file yang pasif [file yang aktif harus diterminate dulu] kok gagal ngunci kalo udah dicompile
Dim hFile       As Long
'Sleep 200 ' tunda 0.2 detik dulu
On Error GoTo LBLFALSE
hFile = RDF.VbOpenFile(spath, FOR_BINARY_ACCESS_READ, LOCK_READ) '
If hFile > 0 Then ' File Bisa diKunci
    hKunci(lKunci) = hFile
    lKunci = lKunci + 1
    KunciFile = True
End If

LBLFALSE:
End Function

Public Function LepasSemuaKunci() ' fungsi pendamping kuncifile
Dim iNum As Long
For iNum = 0 To UBound(hKunci)
    If hKunci(iNum) = 0 Then Exit Function
    TutupFile (hKunci(iNum))
Next
End Function

Public Function NormalizeAttribute(spath As String) ' Menormalkan Atribute
      SetFileAttributes StrPtr(spath), 0
End Function

Public Function HapusFile(spath As String) As Boolean
On Error GoTo Falsex

NormalizeAttribute spath

If DeleteFile(StrPtr(spath)) = 1 Then
   HapusFile = True
Else
   If DeleteFile(StrPtr("\\.\" & spath)) = 1 Then
      HapusFile = True
   End If
End If

If ValidFile(spath) = True Then GoTo Falsex

Exit Function
Falsex:
HapusFile = False
End Function

Public Function ValidFile(ByRef sFile As String) As Boolean ' Memvalidasi file
If PathIsDirectory(StrPtr(sFile)) = 0 And PathFileExists(StrPtr(sFile)) = 1 Then
    ValidFile = True
Else
    ValidFile = False
End If
End Function

Public Function ValidFile2(sFileV As String) As Long ' Memvalidasi file (kusus yang bisa dibuka aj)
Dim MyHnd   As Long
MyHnd = GetHandleFile(sFileV)
If MyHnd > 0 Then
    ValidFile2 = MyHnd
Else
    ValidFile2 = 0
End If
End Function
Public Function ValidFile3(spath As String) As Boolean
If ValidFile(spath) = True Then
    ValidFile3 = True
Else
    If PathIsDirectory(StrPtr(spath)) = 0 Then
        ValidFile3 = False
    Else
        ValidFile3 = True
    End If
End If
End Function
Public Function isProperFile(spath As String, sExt As String) As Boolean ' file yang tepat atau bukan
On Error Resume Next

If InStr(1, UCase$(sExt), UCase$(Right$(spath, 3))) > 0 Then
   isProperFile = True
Else
   isProperFile = False
End If

End Function

Public Function GetFileName(sFile As String) As String ' Mendapatkan nama file+extensi secara normal
On Error Resume Next
Dim TMP As String
Dim nTmp  As Long

    TMP = StrReverse(sFile)
    nTmp = InStr(TMP, "\")
    TMP = Left(TMP, nTmp - 1)

GetFileName = StrReverse(TMP)

End Function

Public Function GetFilePath(sFile As String) As String ' Mendapatkan path file secara normal
Dim sTemp()   As String
Dim lngFile As Long
    
    sTemp = Split(sFile, "\")
    lngFile = Len(sTemp(UBound(sTemp)))
    
GetFilePath = Left$(sFile, Len(sFile) - lngFile - 1)

End Function

Public Function DapatkanUkuranFile(Where As String) As Long
Dim nFileLen     As Long
Dim hFile        As Long

On Error GoTo keluar
    
    hFile = RDF.VbOpenFile(Where, FOR_BINARY_ACCESS_READ, LOCK_NONE)
    nFileLen = RDF.VbFileLen(hFile)
    RDF.VbCloseFile hFile ' Hayoo jangan lupa tutup

    DapatkanUkuranFile = nFileLen
keluar:
End Function

' Untuk Membuka File Awal kali lalu Handle dan ukuran akan di lempar-lempar
Public Sub OpenFileNow(spath As String)
    hGlobal = RDF.VbOpenFile(spath, FOR_BINARY_ACCESS_READ, LOCK_NONE) ' Public
    nSizeGlobal = RDF.VbFileLen(hGlobal) ' Public
End Sub

' Dapatkan handlenya saja
Public Function GetHandleFile(PathFileTarget As String) As Long
     GetHandleFile = RDF.VbOpenFile(PathFileTarget, FOR_BINARY_ACCESS_READ, LOCK_NONE)
End Function

' Dapatkan ukuranya saja
Public Function GetSizeFile(FileHandle As Long) As Long
     GetSizeFile = RDF.VbFileLen(FileHandle)
End Function

Public Sub TutupFile(hFile As Long)
    RDF.VbCloseFile hFile
End Sub
