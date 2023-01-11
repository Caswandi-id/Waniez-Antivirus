Attribute VB_Name = "basDBext"
Dim sPath As String
Public InternalDb(999999) As String ' Untuk menampung ceksum dan nama virus Internal Database + User Database
Public BanyakUDB As Integer

Public Sub Init_UserDatabase()
Dim InternalCount, x As Integer
Dim strUserDb As String
Dim splUser() As String
sPath = GetFilePath(App_FullPathW(False))
If IsFile(sPath & "\WanUDB.dll") = False Then Exit Sub
    InternalCount = 2 ' Banyaknya Internal Database ( Hitung dari 0 )
    strUserDb = OpenFileInTeks(sPath & "\WanUDB.dll")
    splUser = Split(strUserDb, ";")
    BanyakUDB = UBound(splUser())
    For x = 1 To BanyakUDB ' Banyaknya Vrs di User DB
       InternalDb(InternalCount + x) = splUser(x)
    Next
    JumlahVirusM31 = BanyakUDB
    'MsgBox JumlahVirusM31
End Sub
Public Function OpenFileInTeks(Where As String) As String
Dim BinTeks, Temp As String
If IsFile(Where) = True Then
Open Where For Input As #2
    On Error Resume Next
    Do While Not (EOF(2))
       Input #2, Temp
       BinTeks = BinTeks & Temp
    Loop
Close #2
OpenFileInTeks = BinTeks
End If
End Function
Public Sub Init_Dtabase()
InternalDb(0) = "A6377B3722371A372637F13737376637943796370000000000|Fake.Antikill":         InternalDb(1) = "05201052010665A66ED665566AE663B66C1660000000000|Worm.Wukil.A":                     InternalDb(2) = "27DFF7A9203FD212|Agent Trojan"
InternalDb(3) = "5937E83776377B370C379D3726379B37263701370000000000|Worm.rmnitdo_induk"
Init_UserDatabase
End Sub

