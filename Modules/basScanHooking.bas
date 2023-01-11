Attribute VB_Name = "basScanHooking"
Option Explicit


Type COPYDATASTRUCT
    dwData As Long
    cbData As Long
    lpData As Long
End Type

Public Const GWL_WNDPROC = (-4)
Public Const WM_COPYDATA = &H4A
Public GlobalRespon As Boolean
Public hMalware     As Long
Public IPCScan      As Boolean
Public hFile As Long
Global lpPrevWndProc As Long
Global Aku As Long

'Copies a block of memory from one location to another.
Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lngParam As Long) As Long
Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Public Sub Hook()
    lpPrevWndProc = SetWindowLong(Aku, GWL_WNDPROC, AddressOf WindowProc)
End Sub

Public Sub Unhook()
    Dim Temp As Long
    Temp = SetWindowLong(Aku, GWL_WNDPROC, lpPrevWndProc)
End Sub

Function WindowProc(ByVal hw As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lngParam As Long) As Long
    If uMsg = WM_COPYDATA Then
        Call InterProcessComms(lngParam)
    End If
    WindowProc = CallWindowProc(lpPrevWndProc, hw, uMsg, wParam, lngParam)
End Function
Public Function Scan_IPC(singpath As String)
Dim FileNow As String
Dim FSO As Object
Dim SFILE As Object
Set FSO = Nothing
On Error GoTo keluar:
    'lvMalware.ListItems.Clear
    Set FSO = CreateObject("Scripting.FileSystemObject")
    For Each SFILE In FSO.GetFolder(singpath).Files
        DoEvents
        FileNow = SFILE
    If IsFileX(SFILE) = True Then
    If FrmConfig.ck1.Value = 1 Then
     If isProperFile(FileNow, 1000000) = True Then
     CocokanDataBaseRTP (FileNow)
     End If
    Else
         CocokanDataBaseRTP (FileNow)

    End If
     End If
    Next
keluar:
End Function
Private Function isProperFile(Where As String, UkuranByte As Long) As Boolean
On Error GoTo FixE
Dim DaftarEktensi As String
DaftarEktensi = "EXE PIF CPL COM SCR VBS HTM HTML VMX INF DLL" ' Tambahkan ndiri Exkstensi
    If FileLen(Where) <= UkuranByte Then
        isProperFile = True
    Else
        isProperFile = False
    End If
Exit Function

FixE:
isProperFile = False
End Function
Sub InterProcessComms(lngParam As Long)
          Dim cdCopyData As COPYDATASTRUCT
          Dim byteBuffer(1 To 255) As Byte
          Dim strTemp As String
          Dim sebelumnya As String
          Call CopyMemory(cdCopyData, ByVal lngParam, Len(cdCopyData))
          'Debug.Print cdCopyData.dwData
          If (cdCopyData.dwData = 0) Or (cdCopyData.dwData = 3) Then
            Call CopyMemory(byteBuffer(1), ByVal cdCopyData.lpData, cdCopyData.cbData)
            strTemp = StrConv(byteBuffer, vbUnicode)
            strTemp = Left$(strTemp, InStr(1, strTemp, Chr$(0)) - 1)
            Debug.Print strTemp
          '  If InStr(strTemp, "$DR") > 0 Then
               'rmMain.tmrDetek.Enabled = False
          ' On Error Resume Next
          '   If StatusRTP = True And frmMain.Ck14.value = 1 Then ScanPatWithRTPHook Mid(strTemp, 5) & "\"
           
          ' End If
           If InStr(strTemp, "***") > 0 Then
           ' Dim sebelumnya As String
           If StatusRTP = True And FrmConfig.Ck14.Value = 1 And (Mid$(strTemp, 5)) <> sebelumnya And (Mid$(strTemp, 5)) <> "" Then CekViri (Mid$(strTemp, 5))
           sebelumnya = (Mid$(strTemp, 5))
            End If
          End If
         ' End If
      
                    
End Sub
Public Function CekViri(sPath As String) As Boolean
On Error GoTo metu
     If IsFileX(sPath) = True Then
     If isProperFile(sPath, 1000000) = True Then
     If frmMain.rtpkerja = False Then CocokanDataBaseHook sPath
   End If
   End If
metu:
End Function


Public Function isExeFile(Where As String) As Boolean
    Dim exeText As String
    
    exeText = OpenTxtFile(Where, 1)
    
    If Left$(exeText, 2) = "PK" Or Left$(exeText, 3) = "Rar" Then GoSub Warning Else isExeFile = False
    
    Exit Function
    
Warning:
    isExeFile = True
    
End Function

Private Function OpenTxtFile(SFILE As String, PosStart As Long) As String
    Dim Bin As String
    
    Open SFILE For Binary As 1
        Bin = Space$(LOF(1))
        Get #1, PosStart, Bin
    Close #1
    
    OpenTxtFile = Bin
End Function


