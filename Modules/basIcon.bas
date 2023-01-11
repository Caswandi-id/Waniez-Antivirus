Attribute VB_Name = "basIcon"
' Module Untuk Mendapatkan Ceksum icon dan Cek Icon

Declare Function ExtractIconEx Lib "shell32.dll" Alias "ExtractIconExW" (ByVal lpszFile As Long, ByVal nIconIndex As Long, ByRef phiconLarge As Long, ByRef phiconSmall As Long, ByVal nIcons As Long) As Long
Declare Function DrawIconEx Lib "user32" (ByVal hdc As Long, ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As Long, ByVal cxWidth As Long, ByVal cyWidth As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Boolean

Dim CFLFL        As New classFile



' ----------------------------------------------      CEK ICON      -------------------------------------

Public Function DRAW_ICO(PathToDraw As String, PicBox As PictureBox) As Boolean ' Yang dipanggil untuk Cek Icon
Dim hIcon       As Long
Dim IconExist   As Long

Dim HashIco As String
Dim SaveTmp As String
Dim Ukuran  As String
wadahAv = GetFilePath(App_FullPathW(False))

On Error GoTo keluar
DoEvents
DRAW_ICO = False ' init nilainya ke False
PicBox.Cls

SaveTmp = wadahAv & "\icon.tmp"
IconExist = ExtractIconEx(StrPtr(PathToDraw), 0, ByVal 0&, hIcon, 1)

If IconExist <= 0 Then
    IconExist = ExtractIconEx(StrPtr(PathToDraw), 0, hIcon, ByVal 0&, 1)
    If IconExist <= 0 Then Exit Function
End If

DrawIconEx PicBox.hdc, 0, 0, hIcon, 0, 0, 0, 0, &H3

SavePicture PicBox.Image, SaveTmp ' Simpan Dulu Gambarnya
HashIco = CALC_BYTE_ICON(SaveTmp) ' Calculasikan Byte Simpanan

If CEK_ICON(HashIco, PathToDraw) = True Then
    DRAW_ICO = True
Else
    DRAW_ICO = False
End If

'HapusFile SaveTmp ' Hapus Simpanan Icon
keluar:
End Function

' sementara
Private Function CALC_BYTE_ICON(path As String) As String ' Kalkulasikan Byte Icon
On Error Resume Next
    Dim hFileIcon  As Long
    Dim iTurn      As Long
    
    hFileIcon = CFLFL.VbOpenFile(path, FOR_BINARY_ACCESS_READ, LOCK_NONE)
    
    If hFileIcon > 0 Then
       CALC_BYTE_ICON = MYCeksumCadangan(path, hFileIcon)
    Else
       CALC_BYTE_ICON = "00"
    End If
    
    FrmRTP.txtpath.Text = CALC_BYTE_ICON
    CFLFL.VbCloseFile hFileIcon
End Function



