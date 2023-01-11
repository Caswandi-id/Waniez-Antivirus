Attribute VB_Name = "ModGui"
Option Explicit

Private Type GUID
     Data1 As Long
     Data2 As Integer
     Data3 As Integer
     Data4(7) As Byte
End Type

Private Type PicBmp
     SIZE As Long
     Type As Long
     hBmp As Long
     hCur As Long
     hPal As Long
     Reserved As Long
End Type

Private Declare Function OleCreatePictureIndirect Lib "olepro32.dll" (PicDesc As PicBmp, RefIID As GUID, ByVal fPictureOwnHandle As Long, IPic As IPicture) As Long
Private Declare Function LoadBitmap Lib "user32.dll" Alias "LoadBitmapA" (ByVal hInstance As Long, ByVal lpBitmapID As Long) As Long
Private Declare Function LoadCursor Lib "user32.dll" Alias "LoadCursorA" (ByVal hInstance As Long, ByVal lpBitmapID As Long) As Long
Private Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long
Private Declare Function LoadLibrary Lib "kernel32.dll" Alias "LoadLibraryW" (ByVal lpLibFileName As Long) As Long
Private Declare Function FreeLibrary Lib "kernel32.dll" (ByVal hLibModule As Long) As Long
Public Function LoadPictureDLL(ByVal lResourceId As Long) As Picture
On Error GoTo err

Dim hInst As Long
Dim hBmp  As Long
Dim pic As PicBmp
Dim IPic As IPicture
Dim IID_IDispatch As GUID
Dim lRC As Long

hInst = LoadLibrary(StrPtr(App.path & "\WanUI.dll"))
If hInst <> 0 Then
    hBmp = LoadBitmap(hInst, lResourceId)
    If hBmp <> 0 Then
        IID_IDispatch.Data1 = &H20400
        IID_IDispatch.Data4(0) = &HC0
        IID_IDispatch.Data4(7) = &H46
        pic.SIZE = Len(pic)
        pic.Type = vbPicTypeBitmap
        pic.hBmp = hBmp
        pic.hPal = 0
        lRC = OleCreatePictureIndirect(pic, IID_IDispatch, 1, IPic)
        If lRC = 0 Then
            Set LoadPictureDLL = IPic
            Set IPic = Nothing
        Else
            DeleteObject hBmp
        End If
    End If
    FreeLibrary hInst
    hInst = 0
End If
Exit Function

err:
End Function

Public Function LoadGUI()

With frmMain
'pic Click
'.PicMainSummary.Picture = LoadPictureDLL(900)
.PicMainScan.Picture = LoadPictureDLL(902)
.Pictolls.Picture = LoadPictureDLL(902)
.PicQuar.Picture = LoadPictureDLL(902)
.PicUpd.Picture = LoadPictureDLL(902)






End With
End Function


