Attribute VB_Name = "basException"


' ini juga public
Public PathExcep() As String ' penampung path yang di kecualikan (dari indek 1)
Public FileExcep() As String
Public RegExcep()  As String

Public JumPathExcep As Long ' jumlah path yang di kecualikan
Public JumFileExcep As Long ' file
Public JumRegExcep  As Long ' reg

Public Enum ExceptionType
    REG_EXC = 0
    FILE_EXC = 1
    PATH_EXC = 2
End Enum


Public Function ReadExceptPath(sPath As String, LstOut As ListBox) As Long
Dim IsiFile     As String
Dim SplitFile() As String
Dim Counter     As Long
Dim UpperIndek  As Long

If ValidFile(sPath) = False Then GoTo LBL_AKHIR

LstOut.Clear
IsiFile = ReadUnicodeFile(sPath)
IsiFile = Replace(IsiFile, Chr$(0), "") ' Null dihilangkan
SplitFile = Split(IsiFile, Chr$(13))

UpperIndek = UBound(SplitFile)

ReDim PathExcep(UpperIndek) As String

For Counter = 1 To UpperIndek
    PathExcep(Counter) = BuangBadKarPertama(SplitFile(Counter))
    If Len(PathExcep(Counter)) < 4 Then
       UpperIndek = UpperIndek - 1
       GoTo LBL_BROAD
    End If
    LstOut.AddItem Right$("00" & CStr(Counter), 3) & "-" & PathExcep(Counter)
Next

LBL_BROAD:
ReadExceptPath = UpperIndek

Exit Function
LBL_AKHIR:
    ReadExceptPath = 0
    Erase PathExcep()
    LstOut.Clear


End Function

Public Function ReadExceptFile(sPath As String, LstOut As ListBox) As Long
Dim IsiFile     As String
Dim SplitFile() As String
Dim Counter     As Long
Dim UpperIndek  As Long

If ValidFile(sPath) = False Then GoTo LBL_AKHIR

LstOut.Clear

IsiFile = ReadUnicodeFile(sPath)
IsiFile = Replace(IsiFile, Chr$(0), "") ' Null dihilangkan
SplitFile = Split(IsiFile, Chr$(13))

UpperIndek = UBound(SplitFile)

ReDim FileExcep(UpperIndek) As String

For Counter = 1 To UpperIndek
    FileExcep(Counter) = BuangBadKarPertama(SplitFile(Counter))
    If Len(FileExcep(Counter)) < 4 Then
       UpperIndek = UpperIndek - 1
       GoTo LBL_BROAD
    End If
    LstOut.AddItem Right$("00" & CStr(Counter), 3) & "-" & FileExcep(Counter)
Next

LBL_BROAD:
ReadExceptFile = UpperIndek

Exit Function
LBL_AKHIR:
    ReadExceptFile = 0
    Erase FileExcep()
    LstOut.Clear
End Function

Public Function ReadExceptReg(sPath As String, LstOut As ListBox) As Long
Dim IsiFile     As String
Dim SplitFile() As String
Dim Counter     As Long
Dim UpperIndek  As Long

If ValidFile(sPath) = False Then GoTo LBL_AKHIR

LstOut.Clear

IsiFile = ReadUnicodeFile(sPath)
IsiFile = Replace(IsiFile, Chr$(0), "") ' Null dihilangkan
SplitFile = Split(IsiFile, Chr$(13))

UpperIndek = UBound(SplitFile)

ReDim RegExcep(UpperIndek) As String

For Counter = 1 To UpperIndek
    RegExcep(Counter) = BuangBadKarPertama(SplitFile(Counter))
    If Len(RegExcep(Counter)) < 4 Then
       UpperIndek = UpperIndek - 1
       GoTo LBL_BROAD
    End If
    LstOut.AddItem Right$("00" & CStr(Counter), 3) & "-" & RegExcep(Counter)
Next


LBL_BROAD:
ReadExceptReg = UpperIndek

Exit Function
LBL_AKHIR:
    ReadExceptReg = 0
    Erase RegExcep()
    LstOut.Clear
End Function



Public Function ReBuildFileException(sNewFilePath As String, sPathSave As String, LstOut As ListBox)
Dim IsiFile     As String

On Error Resume Next

IsiFile = ReadUnicodeFile(sPathSave)
If IsiFile = "" Then IsiFile = ChrW$(&H6DE) & "-[WAN'IEZ FILE EXCEPTION]-" & ChrW$(&H6DE)  ' klo masih kosong
IsiFile = Replace(IsiFile, Chr$(0), "") ' Null dihilangkan
sNewFilePath = Replace(sNewFilePath, Chr$(0), "") ' Null dihilangkan [bffer ajh]
IsiFile = IsiFile & Chr$(13) & Chr$(10) & sNewFilePath

IsiFile = BuangBadKarKanan(IsiFile)
IsiFile = StrConv(IsiFile, vbUnicode)

WriteFileUniSim sPathSave, IsiFile

    JumFileExcep = ReadExceptFile(sPathSave, LstOut)
MsgBox i_bahasa(6), vbExclamation
End Function

' Cuma nyampe value aja lho, data enggak
Public Function ReBuildPathException(sNewPath As String, sPathSave As String, LstOut As ListBox)
Dim IsiFile     As String

On Error Resume Next

IsiFile = ReadUnicodeFile(sPathSave)
If IsiFile = "" Then IsiFile = ChrW$(&H6DE) & "-[WAN'IEZ PATH EXCEPTION]-" & ChrW$(&H6DE)  ' klo masih kosong
IsiFile = Replace(IsiFile, Chr$(0), "") ' Null dihilangkan
'sNewPath = Replace(sNewPath, Chr$(0), "") ' Null dihilangkan [bffer ajh] - gperlu buffer
IsiFile = IsiFile & Chr$(13) & Chr$(10) & sNewPath

IsiFile = BuangBadKarKanan(IsiFile)
IsiFile = StrConv(IsiFile, vbUnicode)

WriteFileUniSim sPathSave, IsiFile

    JumPathExcep = ReadExceptFile(sPathSave, LstOut)
MsgBox i_bahasa(6), vbExclamation
End Function

Public Function ReBuildRegException(sNewRegPath As String, sPathSave As String, LstOut As ListBox)
Dim IsiFile     As String

On Error Resume Next

IsiFile = ReadUnicodeFile(sPathSave)
If IsiFile = "" Then IsiFile = ChrW$(&H6DE) & "-[WAN'IEZ REG EXCEPTION]-" & ChrW$(&H6DE)  ' klo masih kosong
IsiFile = Replace(IsiFile, Chr$(0), "") ' Null dihilangkan
sNewRegPath = Replace(sNewRegPath, Chr$(0), "") ' Null dihilangkan [bffer ajh]
IsiFile = IsiFile & Chr$(13) & Chr$(10) & sNewRegPath

IsiFile = BuangBadKarKanan(IsiFile)
IsiFile = StrConv(IsiFile, vbUnicode)

WriteFileUniSim sPathSave, IsiFile

    JumRegExcep = ReadExceptReg(sPathSave, LstOut)

MsgBox i_bahasa(6), vbExclamation
End Function


Public Sub RemoveExceptionByIndek(nIndek As Long, ExcType As ExceptionType)
Dim HeaderFileExc As String
Dim UpIndek       As Long
Dim iCounter      As Long
Dim TmpIsiFile    As String
Dim FileNameX     As String

Dim TmpWhile      As String
Dim MyPath        As String

On Error Resume Next

Select Case ExcType
    Case 0: HeaderFileExc = ChrW$(&H6DE) & "-[WAN'IEZ REG EXCEPTION]-" & ChrW$(&H6DE): UpIndek = UBound(RegExcep()): FileNameX = "Reg.lst"
    Case 1: HeaderFileExc = ChrW$(&H6DE) & "-[WAN'IEZ FILE EXCEPTION]-" & ChrW$(&H6DE): UpIndek = UBound(FileExcep()): FileNameX = "File.lst"
    Case 2: HeaderFileExc = ChrW$(&H6DE) & "-[WAN'IEZ PATH EXCEPTION]-" & ChrW$(&H6DE): UpIndek = UBound(PathExcep()): FileNameX = "Path.lst"
End Select


If nIndek >= 0 Then ' indek dimulai 0, tapi list di file dimulai 1
    Select Case ExcType
       Case 0: RegExcep(nIndek + 1) = ""
       Case 1: FileExcep(nIndek + 1) = ""
       Case 2: PathExcep(nIndek + 1) = ""
    End Select
    
    TmpIsiFile = HeaderFileExc & Chr$(13) & Chr$(10)
    
    For iCounter = 1 To UpIndek ' bangun isi list
        
      Select Case ExcType
         Case 0
           TmpWhile = RegExcep(iCounter) & Chr$(13) & Chr$(10)
           If Len(TmpWhile) < 4 Then GoTo LBL_LANJUT
           TmpIsiFile = TmpIsiFile & TmpWhile
         Case 1
           TmpWhile = FileExcep(iCounter) & Chr$(13) & Chr$(10)
           If Len(TmpWhile) < 4 Then GoTo LBL_LANJUT
           TmpIsiFile = TmpIsiFile & TmpWhile
         Case 2
           TmpWhile = PathExcep(iCounter) & Chr$(13) & Chr$(10)
           If Len(TmpWhile) < 4 Then GoTo LBL_LANJUT
           TmpIsiFile = TmpIsiFile & TmpWhile
      End Select
LBL_LANJUT:
    Next
    
    MyPath = GetFilePath(App_FullPathW(False)) & "\" & FileNameX
    
    TmpIsiFile = BuangBadKarKanan(TmpIsiFile)

    TmpIsiFile = StrConv(TmpIsiFile, vbUnicode)

    HapusFile MyPath
    WriteFileUniSim MyPath, TmpIsiFile

End If



End Sub

Private Function BuangBadKarPertama(sKar As String) As String
If Asc(Left$(sKar, 1)) <= 11 Then
   BuangBadKarPertama = Mid$(sKar, 2)
Else
   BuangBadKarPertama = sKar
End If
End Function


Private Function BuangBadKarKanan(sKar As String) As String
Dim Stopper    As Boolean
Dim JumsKar    As Long

Do
 JumsKar = Len(sKar)
 If Asc(Right$(sKar, 1)) <= 14 Then
    sKar = Left$(sKar, JumsKar - 1)
 Else
    sKar = sKar
    Stopper = True
 End If

Loop While Stopper = False

BuangBadKarKanan = sKar
End Function
