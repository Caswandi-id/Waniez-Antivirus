Attribute VB_Name = "infecceksum"
' Penampung Isi Hex suatu File secara Globaal
Dim isihexfile As String
Public panjangfile As Long



'Public fileee As New clsFile


Public Function InfectChecksum(IsiFile() As Byte, awal As Long, panjang As Long) As String ' Default panjang = 10
On Error Resume Next
Dim sTmpFile As String
Dim iCount   As Integer
Dim iCek1    As Long
Dim iCek2    As Long

iCek1 = 0
'DoEvents

If awal <= 0 Then awal = 1

For iCount = awal To panjang
    'DoEvents
    iCek1 = iCek1 + Asc(IsiFile(iCount)) ^ 2.5
Next

If iCek1 = "0" Then
    InfectChecksum = "0"
Else
    InfectChecksum = Hex$(iCek1)
End If

Exit Function

LBL_AKHIR:
InfectChecksum = "0"
End Function




