Attribute VB_Name = "basContol"
' Module untuk penanganan Control thdp aplikasi
'
'
Public Const MerahB = 0
Public Const Kuning = 1
Public Const Hijau = 2
Public Const Merah = 3

' API Untuk menunda Eksekusi
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

' API Untuk mengatur peletakan Form
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cX As Long, ByVal cY As Long, ByVal wFlags As Long) As Long

'File Prop
Public Declare Function ShellExecuteEx Lib "shell32.dll" (SEI As SHELLEXECUTEINFO) As Long

Private Declare Function _
SetLayeredWindowAttributes Lib "user32.dll" _
(ByVal hwnd As Long, ByVal crKey As Long, _
ByVal bAlpha As Byte, _
ByVal dwFlags As Long) As Long

Private Declare Function GetWindowLong Lib _
"user32" Alias "GetWindowLongA" _
(ByVal hwnd As Long, _
ByVal nIndex As Long) As Long

Private Declare Function SetWindowLong Lib _
"user32" Alias "SetWindowLongA" _
(ByVal hwnd As Long, ByVal nIndex As Long, _
ByVal dwNewLong As Long) As Long


' Konstanta peletakan form
Public Const SWP_NOMOVE                  As Long = &H2
Public Const SWP_NOSIZE                  As Long = &H1
Public Const HWND_NOTOPMOST              As Long = -2
Public Const HWND_TOPMOST                As Long = -1
Public Const Flags                       As Long = SWP_NOMOVE Or SWP_NOSIZE
Public dirIcon(5)   As Long

Private Const SEE_MASK_INVOKEIDLIST = &HC
Private Const SEE_MASK_NOCLOSEPROCESS = &H40
Private Const SEE_MASK_FLAG_NO_UI = &H400

Private Type SHELLEXECUTEINFO
    cbSize As Long
    fMask As Long
    hwnd As Long
    lpVerb As String
    lpFile As String
    lpParameters As String
    lpDirectory As String
    nShow As Long
    hInstApp As Long
    lpIDList As Long
    lpClass As String
    hkeyClass As Long
    dwHotKey As Long
    hIcon As Long
    hProcess As Long
End Type
'Global Const AlterExtn = ".{f39a0dc0-9cc8-11d0-a599-00c04fd64433}" 'Class ID for Channel File

Public Sub ShowProperties(FileName As String, OwnerhWnd As Long)
On Error Resume Next
    Dim SEI As SHELLEXECUTEINFO
    With SEI
        .cbSize = Len(SEI)
        .fMask = SEE_MASK_NOCLOSEPROCESS Or SEE_MASK_INVOKEIDLIST Or SEE_MASK_FLAG_NO_UI
        .hwnd = OwnerhWnd
        .lpVerb = "properties"
        .lpFile = FileName
        .lpParameters = vbNullChar
        .lpDirectory = vbNullChar
        .nShow = 0
        .hInstApp = App.hInstance
        .lpIDList = 0
    End With
    ShellExecuteEx SEI
End Sub


' Fungsi untuk meletakan Form
Public Function LetakanForm(frm As Form, Top As Boolean)
If Top = True Then
    Call SetWindowPos(frm.hwnd, HWND_TOPMOST, 0&, 0&, 0&, 0&, Flags)
Else
    Call SetWindowPos(frm.hwnd, HWND_NOTOPMOST, 0&, 0&, 0&, 0&, Flags)
End If
End Function
'coba pakai limav
Public Sub AddList(lstView As ucListView, SFILE As String, picTemp As PictureBox, imglist As Integer, Text As String, Optional SubItem, Optional sKey As String)
    On Error Resume Next
    'kali ini versi unicodenya ayo berjuang :D
    With picTemp
        .Cls: .AutoRedraw = True
        If ValidFile(SFILE) Or PathIsDirectory(StrPtr(SFILE)) Then RetrieveIcon SFILE, picTemp, ricnSmall Else GoTo TakAdaIkon
    End With
    lstView.ImageList.AddFromDc picTemp.hdc, 16, 16
    lstView.ListItems.Add sKey, Text, , dirIcon(imglist), , , , , SubItem
    dirIcon(imglist) = Val(dirIcon(imglist)) + Val(1)
    Exit Sub
TakAdaIkon:
    lstView.ListItems.Add sKey, Text, , , , , , , SubItem
End Sub

' Sub untuk menambahkan informasi ke Listview sampai 3 sub item
Public Sub AddInfoToList(lv As ucListView, sItem As String, sSub1 As String, sSub2 As String, sSub3 As String, iIcon As Long, nScroll As Long)
Dim lstLV As cListItem
Set lstLV = lv.ListItems.Add(, sItem, , iIcon, , , , , "")
    lstLV.SubItem(2).Text = sSub1
    lstLV.SubItem(3).Text = sSub2
    lstLV.SubItem(4).Text = sSub3
If lv.ListItems.Count > nScroll Then lv.Scroll 0, 25

Set lstLV = Nothing
End Sub
Public Sub AutoLst(ListApa As ucListView)
    Dim i      As Integer, ClmAuto As eListViewColumnAutoSize
    If ListApa.ListItems.Count = 0 Then ClmAuto = lvwColumnSizeToColumnText Else ClmAuto = lvwColumnSizeToItemText
    For i = 1 To ListApa.Columns.Count
        ListApa.Columns.Item(i).AutoSize ClmAuto
    Next
End Sub


' Sub untuk menambahkan informasi ke Listview sampai 2 sub item
Public Sub AddInfoToListDua(lv As ucListView, sItem As String, sSub1 As String, sSub2 As String, iIcon As Long, nScroll As Long)
Dim lstLV As cListItem

Set lstLV = lv.ListItems.Add(, sItem, , iIcon, , , , , "")
    lstLV.SubItem(2).Text = sSub1
    lstLV.SubItem(3).Text = sSub2
    'lstLV.SubItem(4).Text = sSub3
If lv.ListItems.Count > nScroll Then lv.Scroll 0, 25

Set lstLV = Nothing
End Sub
Public Sub AddInfoToListSatu(lv As ucListView, sItem As String, iIcon As Long, nScroll As Long)
Dim lstLV As cListItem

Set lstLV = lv.ListItems.Add(, sItem, , iIcon, , , , , "")
    'lstLV.SubItem(2).Text = sSub1
    'lstLV.SubItem(3).Text = sSub2
    'lstLV.SubItem(4).Text = sSub3
If lv.ListItems.Count > nScroll Then lv.Scroll 0, 25

Set lstLV = Nothing
End Sub
Public Sub GaweTransparan(lHwnd As Long, ByVal bLevel As Byte)
On Error GoTo salah
    Dim lOldStyle As Long
    
    If (lHwnd <> 0) Then
        lOldStyle = GetWindowLong(lHwnd, (-20))
        SetWindowLong lHwnd, (-20), lOldStyle Or &H80000
        SetLayeredWindowAttributes lHwnd, 0, bLevel, &H2&
    End If
salah:
End Sub

Public Function GetFileTitle(ByVal sFilename As String) As String
    'Returns FileTitle Without Path
    Dim lPos As Long
    lPos = InStrRev(sFilename, "\")


    If lPos > 0 Then


        If lPos < Len(sFilename) Then
            GetFileTitle = Mid$(sFilename, lPos + 1)
        Else
            GetFileTitle = ""
        End If
    Else
        GetFileTitle = sFilename
    End If
    
End Function
