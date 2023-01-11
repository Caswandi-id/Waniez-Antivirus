Attribute VB_Name = "basRTP"

' Public Varnya disini sebagian
Public StatusRTP   As Boolean ' Klo False berarti RTP mati
Public Sub ScanPatWithRTP(PathWillScan As String)

If ApaPengecualian(PathWillScan, JumPathExcep) = True Then GoTo BERAKHIR
If StatusRTP = False Then GoTo BERAKHIR
 'If PathWillScan = GetSpecFolder(RECENT_DOC) Then GoTo BERAKHIR
  ' DoEvents
   
  'frmMain.ScanFileRTP PathWillScan
  'ScanRTP PathWillScan
  Scan_IPC PathWillScan

BERAKHIR:
End Sub
'Public Sub ScanPatWithRTPHook(PathWillScan As String)

'If ApaPengecualian(PathWillScan, JumPathExcep) = True Then GoTo BERAKHIR
'If StatusRTP = False Then GoTo BERAKHIR
   
  ' DoEvents
   
 '  frmMain.ScanFileHOOK PathWillScan

'BERAKHIR:
'End Sub

Public Function tampilkanRTP()
If FrmRTP.lvRTP.ListItems.Count > 0 Then
With FrmRTP
     .Show
    .SetFocus
  ' .Image1.Picture = .imgAman(1)
    KunciFileYangDiRTP .lvRTP
End With
Call LetakanForm(FrmRTP, True)

FrmRTP.Caption = d_bahasa(61) & "[ " & FrmRTP.lvRTP.ListItems.Count & " ]" & " Virus!!!" ' in opened folder, Please select the action..."
FrmRTP.LbDetec.Caption = "Detected [ " & FrmRTP.lvRTP.ListItems.Count & " ] Virus!!!" ' dalam folder yang sedang anda buka. Segera ambil tindakan..."
End If
End Function

' Private karena untuk module ini saja
Private Function ApaPengecualian(sPath As String, JumlahPath As Long) As Boolean
Dim iCount As Integer
On Error GoTo LBL_AKHIR
For iCount = 1 To JumlahPath
    If InStr(UCase$(sPath), UCase$(PathExcep(iCount))) > 0 Then
       ApaPengecualian = True
       Exit Function
    End If
Next
LBL_AKHIR:
End Function


' Publik karena mau diakses dimana-mana
Public Function ApaPengecualianFile(sFile As String, JumFileExc As Long) As Boolean
Dim iCount As Integer

If sFile = "" Then Exit Function

For iCount = 1 To JumFileExc
    If UCase$(FileExcep(iCount)) = UCase$(sFile) Then
       ApaPengecualianFile = True
       Exit Function
    End If
Next

End Function

' Publik karena mau diakses dimana-mana (senearnya cuma di modReg dan DbReg aj)
Public Function ApaPengecualianReg(sRegPathAndValue As String, JumRegExc As Long) As Boolean
Dim iCount As Integer

If sRegPathAndValue = "" Then Exit Function

For iCount = 1 To JumRegExc
    If UCase$(RegExcep(iCount)) = UCase$(sRegPathAndValue) Then
       ApaPengecualianReg = True
       Exit Function
    End If
Next


End Function

' Yang masuk di RTP coba kunci semua
Public Sub KunciFileYangDiRTP(listRTP As ucListView)
Dim CountToBeLock   As Long
Dim PthKunci        As String
On Error Resume Next

For CountToBeLock = 1 To listRTP.ListItems.Count
    PthKunci = listRTP.ListItems.Item(CountToBeLock).SubItem(2).Text
    KunciFile PthKunci
Next
End Sub

