Attribute VB_Name = "basUpdate"

Option Explicit

Public HentikanUpdate As Boolean
Public bUpdateCompon  As Boolean

Private Declare Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" (ByVal pCaller As Long, ByVal szURL As String, ByVal szFileName As String, ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long
Private Declare Function InternetOpen Lib "wininet" Alias "InternetOpenA" (ByVal sAgent As String, ByVal lAccessType As Long, ByVal sProxyName As String, ByVal sProxyBypass As String, ByVal lFlags As Long) As Long
Private Declare Function InternetCloseHandle Lib "wininet" (ByVal hInet As Long) As Integer

Private Const UpdateInfo = "[UpdateInfo]"

Dim PenampungInfoUpdate As String
Dim PenampungBerita     As String

Function InternetGetFile(sURLFileName As String, sSaveToFile As String, Optional bOverwriteExisting As Boolean = False) As Boolean
    Dim lRet As Long
    
    On Error Resume Next
        
    If bOverwriteExisting Then
        If ValidFile(sSaveToFile) = True Then
            HapusFile sSaveToFile
        End If
    End If
    
    
    
    If Len(Dir$(sSaveToFile)) = 0 Then
        
      If IsConnectedToInternet = True Then
         frmMain.Downloader1.BeginDownload sURLFileName, sSaveToFile, lRet
         InternetGetFile = True
      Else
         frmMain.Downloader1.BeginDownload sURLFileName, sSaveToFile, lRet
         InternetGetFile = False
      End If
    
    End If
    
End Function

Public Function AmbilUpdateInfo(UrlUpdateInfo As String, sFileSave As String)
       InternetGetFile UrlUpdateInfo, sFileSave, True
End Function
Public Function CheckUpdate(TmpFile As String, lblOut As Label) As String
    TampilkanBalon FrmSysTray, j_bahasa(26), i_bahasa(26), NIIF_INFO
    lblOut.Caption = j_bahasa(33)
    If ValidFile(TmpFile) = True Then
       DoEvents
       If AdakahUpdateUntukVersiDipakai(TmpFile) = True Then
          CheckUpdate = PenampungInfoUpdate & " dan Wan'iez Antivirus " & ". - [Kabar Baru Web Wan'iez] " & PenampungBerita
          bUpdateCompon = True
          TampilkanBalon FrmSysTray, PenampungInfoUpdate & " dan Wan'iez Antivirus ", i_bahasa(26), NIIF_INFO
       Else
          bUpdateCompon = False
          CheckUpdate = j_bahasa(25) & ". - [Kabar Terbaru Wan'iez Antivirus]  " & PenampungBerita
          TampilkanBalon FrmSysTray, j_bahasa(25), i_bahasa(26), NIIF_INFO
       End If
    Else
       CheckUpdate = j_bahasa(24)
       bUpdateCompon = False
       TampilkanBalon FrmSysTray, j_bahasa(24), i_bahasa(27), NIIF_WARNING
    End If
    lblOut.Caption = j_bahasa(31)
    HapusFile TmpFile ' hapus setelah dibaca
End Function

Public Function AdakahUpdateUntukVersiDipakai(SFileInfo As String) As Boolean
Dim sTmpIsi      As String
Dim SplitterA()  As String
Dim SoftInfo     As String
Dim SoftVers(1)  As Long
Dim SoftBuild(1) As Long
Dim DbWorm(1)    As Long
Dim DbVirus(1)   As Long

On Error GoTo LBL_FALSE

sTmpIsi = ReadUnicodeFile(SFileInfo)
sTmpIsi = Mid$(sTmpIsi, InStr(sTmpIsi, UpdateInfo) + 12)

SplitterA = Split(sTmpIsi, "~")
SoftVers(0) = CLng(SplitterA(0)) ' PH di info update
SoftBuild(0) = CLng(SplitterA(1)) ' build di info update
DbWorm(0) = CLng(SplitterA(2)) ' jumlah worm di info update
PenampungBerita = SplitterA(3) ' berita dari web cmc
SoftInfo = SplitterA(4) ' informasi di update

PenampungInfoUpdate = SoftInfo ' dikenali di module ini ja

With FrmAbout
    SoftVers(1) = Mid$(.lbEngine.Caption, 4) ' PH yang dipakai
    SoftBuild(1) = Mid$(.LbBuildNumber.Caption, 3) ' Build yang dipakai
    DbWorm(1) = Mid$(.lbWorm.Caption, 3) ' jumlah womr yang dipakai
End With

If SoftVers(1) <= SoftVers(0) Then ' jika versi yang dipakai </= update
   ' cek build
   If SoftBuild(1) <= SoftBuild(0) Then  ' jika build yang dipakai < update
      If DbWorm(1) < DbWorm(0) Then ' ada worm tambahan nih dari update
         AdakahUpdateUntukVersiDipakai = True ' ada update
      Else
         AdakahUpdateUntukVersiDipakai = False
      End If
   Else
      AdakahUpdateUntukVersiDipakai = False
   End If
Else
  AdakahUpdateUntukVersiDipakai = False
End If

Exit Function
LBL_FALSE:
      AdakahUpdateUntukVersiDipakai = False

End Function
' Update Component
Public Sub UpdateKomponen(PB As ucProgressBar, lblOut As Label, KompNumber As Long)
Dim URLToUpdateAV As String
Dim iCounter           As Long
Dim TmpSimpan          As String
Dim FolderSign         As String

TmpSimpan = GetSpecFolder(USER_DOC) & "\wztmp.txt"
FolderSign = GetFilePath(App_FullPathW(False)) & "\upx\"

PB.Value = KompNumber
PB.Max = 31

URLToUpdateAV = "http://waniez.p.ht/download/upx/" & Hex$(KompNumber) & "_a.zip"


    If HentikanUpdate = True Then GoTo LBL_BERHENTI
    lblOut.Caption = j_bahasa(30) & " : " & URLToUpdateAV
    
    InternetGetFile URLToUpdateAV, TmpSimpan, True
    
    PB.Value = KompNumber
    
Exit Sub
LBL_BERHENTI:
    lblOut.Caption = j_bahasa(32)
End Sub

' Update Component
Public Sub UpdateKomponenNonPE(PB As ucProgressBar, lblOut As Label, KompNumber As Long)
Dim URLToUpdateAV     As String
Dim iCounter           As Long
Dim TmpSimpan          As String
Dim FolderSign         As String

TmpSimpan = GetSpecFolder(USER_DOC) & "\wztmp.txt"
FolderSign = GetFilePath(App_FullPathW(False)) & "\upx\"

PB.Value = KompNumber + 15
PB.Max = 31
' Update http://waniez.p.ht/download/upx/0z.upx (0z.upx adalah file)

URLToUpdateAV = "http://waniez.p.ht/download/upx/" & Hex$(KompNumber) & "_x.zip"


    If HentikanUpdate = True Then GoTo LBL_BERHENTI
    lblOut.Caption = j_bahasa(30) & " : " & URLToUpdateAV
    
    InternetGetFile URLToUpdateAV, TmpSimpan, True
    
   PB.Value = KompNumber + 16
    
Exit Sub
LBL_BERHENTI:
    lblOut.Caption = j_bahasa(32)
End Sub

Public Sub MoveIfValidComp(sFileComp As String, TargetMove As String, lblOut As Label)
    If IsValidComponenDB(sFileComp) = True Then
       HapusFile TargetMove
       CopiFile sFileComp, TargetMove, True
       lblOut.Caption = j_bahasa(29) & " : " & TargetMove
    Else
       HapusFile sFileComp ' hapus disini
    End If
End Sub

Public Function IsValidComponenDB(sFileComp As String) As Boolean
Dim TmpDataFile  As String
Dim UkuranKotor  As Long
Dim UkuranBersih As Long

On Error GoTo LBL_FALSE

TmpDataFile = ReadUnicodeFile(sFileComp)
UkuranKotor = Len(TmpDataFile)

If CLng(Mid$(TmpDataFile, 3, 6)) = UkuranKotor - 10 Then
   IsValidComponenDB = True
Else
   IsValidComponenDB = False
End If

Exit Function

LBL_FALSE:
    IsValidComponenDB = False
End Function

