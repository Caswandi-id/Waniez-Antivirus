Attribute VB_Name = "basProses"
' Untuk Enum proses dan akses proses lain dari module ProsesAkses

Dim stFileStart() As String ' untuk penampung file-file startup
Dim JumStart      As Long ' penampung jumlahnya

Dim cImgList  As gComCtl

Public Sub ENUM_PROSES(ByRef lv As ucListView, pcBuffer As PictureBox)
On Error Resume Next
Dim path As String

Dim LEAX            As Long
Dim CTurn           As Long
Dim ENPC()          As ENUMERATE_PROCESSES_OUTPUT
Dim sTmp(8)         As String
Dim BufPathUni      As String
Set cImgList = New gComCtl
frmMain.lvProses.ListItems.Clear
    LEAX = PamzEnumerateProcesses(ENPC())
        If LEAX <= 0 Then
        GoTo LBL_TERAKHIR
    End If
    
    Set lv.ImageList = cImgList.NewImageList(16, 16, imlColor32)
   
    EnumStatStart stFileStart() ' di enum startupnya dulu
    
    For CTurn = 0 To (LEAX - 1)
        sTmp(0) = ENPC(CTurn).szNtExecutableNameW
        sTmp(1) = GetStatStart(PamzNtPathToUserFriendlyPathW(ENPC(CTurn).szNtExecutablePathW)) & " value(s)"
        sTmp(2) = ENPC(CTurn).nProcessID
        sTmp(3) = ENPC(CTurn).nParentProcessID
        sTmp(4) = ENPC(CTurn).bIsHiddenProcess
        sTmp(5) = ENPC(CTurn).bIsBeingDebugged
        sTmp(6) = ENPC(CTurn).bIsLockedProcess
        sTmp(7) = Format$(ENPC(CTurn).nSizeOfExecutableOpInMemory, "#,#")
        sTmp(8) = PamzNtPathToUserFriendlyPathW(ENPC(CTurn).szNtExecutablePathW)
        
        
        BufPathUni = sTmp(8)
        If ValidFile(sTmp(8)) = False Then
           MakeExeBuffer GetSpecFolder(USER_DOC) & "\00.exe"
           DrawIco GetSpecFolder(USER_DOC) & "\00.exe", frmMain.picBuffer, ricnSmall
           HapusFile GetSpecFolder(USER_DOC) & "\00.exe"
        Else
           DrawIco BufPathUni, frmMain.picBuffer, ricnSmall
        End If
        
        lv.ImageList.AddFromDc frmMain.picBuffer.hdc, 16, 16
        
        lv.ListItems.Add , sTmp(0), , (lv.ImageList.IconCount - 1), , , , , Array(sTmp(1), sTmp(2), sTmp(3), sTmp(4), sTmp(5), sTmp(6), sTmp(7), sTmp(8))

        If ENPC(CTurn).bIsHiddenProcess = True Or IsHiddenFilePros(sTmp(8)) = True Then
           lv.ListItems.Item(CTurn + 1).Cut = True
           
        End If
    Next
    
    Erase ENPC()
    frmMain.lstModule.Clear
LBL_TERAKHIR:
    '---tambahan ajah:
           AutoLst frmMain.lvProses

    frmMain.frProses.Caption = c_bahasa(2) & " ( " & frmMain.lvProses.ListItems.Count & " )"
    Set cImgList = Nothing
End Sub


Public Sub ENUM_MODULE(pId As Long, lstMod As ListBox)
On Error Resume Next
Dim LEAX            As Long
Dim CTurn           As Long
Dim pAddress        As String
Dim ENMC()          As ENUMERATE_MODULES_OUTPUT
    LEAX = PamzEnumerateModules(pId, ENMC)
    lstMod.Clear '---hapus isi yg lama.
    
    If LEAX <= 0 Then
        lstMod.AddItem j_bahasa(22) & " !"
        GoTo LBL_TERAKHIR
    End If
    For CTurn = 0 To (LEAX - 1)
        pAddress = Hex$(CLng(PamzNtPathToUserFriendlyPathW(CStr(ENMC(CTurn).pBaseAddress))))
        lstMod.AddItem "0x" & Right$("000000000" & pAddress, 8) & " - " & PamzNtPathToUserFriendlyPathW(ENMC(CTurn).szNtModulePathW)
    Next
    Erase ENMC()
    lbPID = TargetPID
LBL_TERAKHIR:
    '---tambahan ajah:
    frmMain.frModule.Caption = c_bahasa(3) & " (" & LEAX & ")"
    '-----------------;

End Sub

' Kill, Kunci, dan retstart
Public Function KillProses(pId As Long, sPath As String, bRestart As Boolean, bKunci As Boolean) As Boolean
    If PamzTerminateProcess(pId) > 0 Then
       KillProses = True
       If bRestart = True Then ' mau direstart
          Shell sPath, vbNormalNoFocus
           
       End If
       If bKunci = True Then
          KunciFile sPath
       End If
    Else
       KillProses = False
    End If
End Function

Public Function SuspendProses(pId As Long, bPause As Boolean) As String
If bPause = True Then ' mau pause
   If PamzSuspendResumeProcessThreads(pId, False) > 0 Then
      SuspendProses = "Paused"
   Else
      SuspendProses = "Ps-Failed"
   End If
Else
   If PamzSuspendResumeProcessThreads(pId, True) > 0 Then
      SuspendProses = "Resumed"
   Else
      SuspendProses = "Rs-Failed"
   End If
End If
End Function

Public Function UnloadModuleForce(sModulePathinLstBox As String, lstMod As ListBox, proPID As Long)
Dim pAddr   As Long
pAdd = GetHexAdressToLng(sModulePathinLstBox)
If PamzForceUnLoadProcessModule32(proPID, pAdd) > 0 Then
   ' refresh
    Call ENUM_MODULE(CLng(proPID), lstMod)
    MsgBox i_bahasa(4) & " ! ( 0x" & Hex$(pAdd) & " )", vbInformation
Else
    MsgBox i_bahasa(5) & " ! ( 0x" & Hex$(pAdd) & " )", vbExclamation
End If
End Function

' ----------------------------- Fungsi-Funsgi Buffer

' Melakukan enumerisasi Startup pada titik-titik yang dimasukan saat enum RegStartUp
Private Sub EnumStatStart(ByRef strFile() As String) 'output pada strFile
Dim iNum     As Long
Dim iCount   As Long
Dim stFile() As String
    iNum = EnumRegStartup(stFile, True)
    ReDim strFile(iNum) As String ' indeknya ga kepakai satu gpp
    
    For iCount = 1 To iNum
        strFile(iCount - 1) = stFile(iCount - 1)
    Next
    JumStart = iNum
    'MsgBox iNum
End Sub

' Untuk mencocokan berapa nilai startup suatu alamat file
Private Function GetStatStart(SFILE As String) As Long
Dim iNum As Long
Dim nJum As Long
For iNum = 1 To JumStart
    If UCase$(SFILE) = UCase$(stFileStart(iNum - 1)) Then
       nJum = nJum + 1
    End If
Next
GetStatStart = nJum
End Function

' Menggambar icon ke picture box untuk tampilan listview proses
Public Sub DrawIco(sPath As String, oPic As PictureBox, nDimension As IconRetrieve)
    With oPic
        .Cls: .AutoRedraw = True
        RetrieveIcon sPath, oPic, nDimension
    End With
End Sub

' Cadangan path file yang tak mampu di enum
Private Sub MakeExeBuffer(sPath As String)
Dim sTmp(1) As Byte
WriteUnicodeFile sPath, 1, sTmp
End Sub


'--- dari 0x000000F menjadi F
Private Function GetHexAdressToLng(sModulePath As String) As Long
Dim sTmp() As String
On Error GoTo AKHIR
    sTmp = Split(sModulePath, "-")
    GetHexAdressToLng = CLng("&H" & Mid$(sTmp(0), 3))
Exit Function
AKHIR:
GetHexAdressToLng = 0
End Function


Private Function IsHiddenFilePros(sFilePro As String) As Boolean
Dim NAT As Long

NAT = GetFileAttributes(StrPtr(sFilePro))

If (NAT = 2 Or NAT = 34 Or NAT = 3 Or NAT = 6 Or NAT = 22 Or NAT = 18 Or NAT = 50 Or NAT = 19 Or NAT = 35) Then
    IsHiddenFilePros = True
Else
    IsHiddenFilePros = False
End If

End Function
Public Function KillProByName(sProName As String) As Long ' membunuh proses dari path penuh dan nama
On Error Resume Next
Dim LEAX            As Long
Dim CTurn           As Long
Dim pId             As Long
Dim nSize           As Long

Dim ENPC()          As ENUMERATE_PROCESSES_OUTPUT
Dim sProsesPath     As String
    
    LEAX = PamzEnumerateProcesses(ENPC())
    MYID = GetCurrentProcessId()
    If LEAX <= 0 Then
        GoTo LBL_TERAKHIR
    End If
    
    For CTurn = 0 To (LEAX - 1)
        sProsesPath = PamzNtPathToUserFriendlyPathW(ENPC(CTurn).szNtExecutableNameW)
        pId = ENPC(CTurn).nProcessID
        If InStr(UCase$(sProsesPath), UCase$(sProName)) > 0 Then
           
           If PamzTerminateProcess(pId) > 0 Then
              KillProByName = 1
           Else
              KillProByName = 0
           End If
        End If
    Next
       
    Erase ENPC()

LBL_TERAKHIR:
End Function

