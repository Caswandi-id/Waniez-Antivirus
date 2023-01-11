Attribute VB_Name = "BasInstal"
Dim wadahAv As String

Public Sub BukaFile(BukaApa As String, Optional parameter As String, Optional hwnd As Long)
    On Error Resume Next
    Call ShellExecute(hwnd, "open", BukaApa, parameter, CurDir$(), vbNormalFocus)
End Sub

Private Function KopiKeSistem() As Boolean
Dim sPath As String
Dim plug As String
Dim iTurn       As Byte
wadahAv = GetFilePath(App_FullPathW(False))
On Error Resume Next ' Redimensi dulu
ReDim JumlahVirus(15) As Long
ReDim JumlahVirusNonPE(15) As Long
ReDim JumlahIconMenu(6) As Long


  '  CopiFile wadahAv & "\" & App.EXEName & ".exe", PathAv, False
    If ValidFile(PathRME) = False Then CopiFile wadahAv & "\Read me.txt", PathRME, False

    CopiFile wadahAv & "\ChangeLog.txt", PathLOG, False

 SignBack = "_a.upx"
    For iTurn = 0 To UBound(JumlahVirus) ' 0-15
    iCount = 0
    sPath = GetFilePath(App_FullPathW(False)) & "\upx\" & Hex$(iTurn) & SignBack
    CopiFile sPath, FolderAv & "\upx\" & Hex(iTurn) & SignBack, False
    Next
  SignBack = "_x.upx"
    For iTurn = 0 To UBound(JumlahVirus) ' 0-15
    iCount = 0
    sPath = GetFilePath(App_FullPathW(False)) & "\upx\" & Hex$(iTurn) & SignBack
    CopiFile sPath, FolderAv & "\upx\" & Hex(iTurn) & SignBack, False
    Next

End Function
Public Function VersiIni() As Boolean
    Dim Versi      As String, VersiKu As String
    VersiKu = CreateObject("Scripting.FileSystemObject").GetFileVersion(App.path & "\" & App.EXEName & ".exe")
    Versi = CreateObject("Scripting.FileSystemObject").GetFileVersion(PathAv)
    VersiIni = (VersiKu > Versi)
End Function
Public Sub TanyaSistem(Optional VersiKu As Boolean = False)
   'If ValidFile(PathAv) = False Then
    If VersiKu = False Then
        If MsgBox(i_bahasa(45) & vbCrLf & _
            i_bahasa(46), vbExclamation Or vbYesNo) = vbYes Then
            PasangAv True
        End If
    Else
    'If VersiIni Then
            If MsgBox(i_bahasa(47) & vbCrLf & _
                i_bahasa(48), vbExclamation Or vbYesNo) = vbYes Then
                PasangAv False
                FormSplash.Tunggu 5
                PasangAv True
            End If
   '  End If
             
    End If
 ' END IF
End Sub
Public Sub PasangAv(Optional instal As Boolean = True)

If instal = True Then
Call BuatPathAv
Call KopiKeSistem

    BukaFile PathAv, "-A", frmMain.hwnd
     SetDwordValue &H80000001, "Software\KEM" & ChrW$(&H394) & "\", "Rtp", 1
      Call loadRTP
Else
KillProByPath PathAv
HapusFile PathAv
DoEvents
End If
End Sub
Public Function PathLOG() As String
PathLOG = GetSpecFolder(PROGRAM_FILE) & "\Canvas Software\Wan'iez Antivirus\ChangeLog.txt"
End Function
Public Function PathRME() As String
PathRME = GetSpecFolder(PROGRAM_FILE) & "\Canvas Software\Wan'iez Antivirus\Read me.txt"
End Function
Public Function PathHelp() As String
PathHelp = GetSpecFolder(PROGRAM_FILE) & "\Canvas Software\Wan'iez Antivirus\waniez.exe"
End Function
Public Sub BuatPathAv()
FolderAv = GetSpecFolder(PROGRAM_FILE) & "\Canvas Software\Wan'iez Antivirus"
FolderHelp = GetSpecFolder(PROGRAM_FILE) & "\Canvas Software\Wan'iez Antivirus\help"
FolderPlugin = GetSpecFolder(PROGRAM_FILE) & "\Canvas Software\Wan'iez Antivirus\plugin"
FolderSign = GetSpecFolder(PROGRAM_FILE) & "\Canvas Software\Wan'iez Antivirus\upx"
FolderSignx = GetSpecFolder(PROGRAM_FILE) & "\Canvas Software\Wan'iez Antivirus\upx"
If PathIsDirectory(StrPtr(FolderAv)) = 0 Then
   BuatFolder FolderAv
   Else
   SetFileAttributes StrPtr(FolderAv), vbNormal
End If
If PathIsDirectory(StrPtr(FolderHelp)) = 0 Then
   BuatFolder FolderHelp
   Else
   SetFileAttributes StrPtr(FolderHelp), vbNormal
End If
If PathIsDirectory(StrPtr(FolderPlugin)) = 0 Then
   BuatFolder FolderPlugin
   Else
   SetFileAttributes StrPtr(FolderPlugin), vbNormal
End If
If PathIsDirectory(StrPtr(FolderSign)) = 0 Then
   BuatFolder FolderSign
   Else
   SetFileAttributes StrPtr(FolderSign), vbNormal
End If
If PathIsDirectory(StrPtr(FolderSignx)) = 0 Then
   BuatFolder FolderSignx
   Else
   SetFileAttributes StrPtr(FolderSignx), vbNormal
End If
End Sub

Public Sub PaintTile(hdc As Long, X1 As Long, Y1 As Long, X2 As Long, Y2 As Long, Optional pakepic As Boolean = False, Optional pic As PictureBox)
Dim ctile As New cDIBTile
'ctile.SetPattern frmMain.picTile.Picture
If pakepic = False Then
ctile.Tile hdc, X1, Y1, X2, Y2
Else
pic.AutoRedraw = True
ctile.Tile pic.hdc, X1, Y1, X2, Y2
End If
End Sub


