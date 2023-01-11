Attribute VB_Name = "basCompiler"
Function stringcheck(SFILE As String, hexstring As String, namavirus As String)
'Fungsi untuk mencocokkan string sampel dan string pada file
    stringcheck = ""
    Dim filedata As String
    Dim A As Integer
    Open SFILE For Binary As #1
        filedata = Space$(LOF(1))
        Get #1, , filedata
        If InStr(1, filedata, hexstring) > 0 Then
            stringcheck = namavirus
        Else
            stringcheck = ""
        End If
    'akhir dari fungsi
    Close #1
End Function
Function hex2ascii(ByVal hextext As String) As String
'Fungsi untuk menterjemahkan dari hexadecimal ke dalam string biasa
    On Error Resume Next
    Dim y As Integer
    Dim Num As String
    Dim Value As String
    For y = 1 To Len(hextext)
        Num = Mid(hextext, y, 2)
        Value = Value & Chr(Val("&h" & Num))
        y = y + 1
    Next y
    hex2ascii = Value
End Function
'Fungsi yang berisi sampel dari packernya.
Function ambil_sampel_packer(i As Integer)
    Dim sampel(100) As String
    sampel(1) = "0000004d4557:MEW"
    sampel(2) = "555058210c09:UPX"
    sampel(3) = "c02e61737061636b00:Aspack"
    sampel(4) = "89085045436f6d70616374:PECompact"
    sampel(5) = "Selesai:Selesai"
    ambil_sampel_packer = sampel(i) 'hasil yang diberikan
End Function
'Akhir dari Fungsi
'Fungsi yang berisi sampel dari compiler
Function ambil_sampel_compiler(i As Integer)
    Dim sampel(100) As String
    sampel(1) = "0000004d535642564d36302e444c4c000000:MS Visual Basic 6.0"
    sampel(2) = "5700650064000300540068007500030046007200690003005300610074:Borland Delphi 7"
    sampel(3) = "000000004d6963726f736f66742056697375616c20432b2b2052756e74696d65:MS Visual C++"
    sampel(4) = "Selesai:Selesai"
    ambil_sampel_compiler = sampel(i) 'hasil yang diberikan
End Function 'Ahir dari fungsi
'Akhir dari fungsi
'Fungsi untuk mendapatkan informasi tentang packer
Function get_Packer(SFILE As String) As String
Dim sampel(100) As String
Dim signa(100) As String
Dim PackerName(100) As String
Dim i As Integer
i = 1
Do 'Jika sampel i sebelumnya adalah Selesai:Selesai maka berhenti looping
    sampel(i) = ambil_sampel_packer(i) 'sampel i adalah hasil dari fungsi ambil sampel packer
    signa(i) = Mid(sampel(i), 1, InStr(1, sampel(i), ":") - 1)
    PackerName(i) = Mid(sampel(i), InStr(1, sampel(i), ":") + 1, Len(sampel(i)) - InStr(1, sampel(i), ":") + 1)
    hasil = stringcheck(SFILE, hex2ascii(signa(i)), PackerName(i))
    If hasil <> "" And hasil <> "Selesai" Then 'Jika hasil tidak = "" atau tidak = "Selesai"
        get_Packer = hasil 'Kembalikan Hasilnya
        Exit Do 'Berhenti Looping
    End If
    get_Packer = "Tiada"
i = i + 1
Loop Until sampel(i - 1) = "Selesai:Selesai" ' akhir dari looping
End Function

Function get_Compiler(SFILE As String) As String
Dim sampel(100) As String
Dim signa(100) As String
Dim CompilerName(100) As String
Dim i As Integer
i = 1
Do 'Jika sampel i sebelumnya adalah Selesai:Selesai maka berhenti looping
    sampel(i) = ambil_sampel_compiler(i) 'sampel i adalah hasil dari fungsi ambil sampel packer
    signa(i) = Mid(sampel(i), 1, InStr(1, sampel(i), ":") - 1)
    CompilerName(i) = Mid(sampel(i), InStr(1, sampel(i), ":") + 1, Len(sampel(i)) - InStr(1, sampel(i), ":") + 1)
    hasil = stringcheck(SFILE, hex2ascii(signa(i)), CompilerName(i))
    If hasil <> "" And hasil <> "Selesai" Then 'Jika hasil tidak = "" atau tidak = "Selesai"
        get_Compiler = hasil 'Kembalikan Hasilnya
        Exit Do 'Berhenti Looping
    End If
    get_Compiler = "Tak Diketahui"
i = i + 1
Loop Until sampel(i - 1) = "Selesai:Selesai" ' akhir dari looping
End Function

Public Function isVBa(AlamaT As String) As Boolean
On Error Resume Next
Dim AA As String
    Open AlamaT For Binary As #1
        AA = Space(LOF(1))
        Get #1, , AA
    Close #1
If InStr(AA, "MSVB") > 0 And UCase(App.path & "\" & App.EXEName & ".exe") <> UCase(AlamaT) Then isVBa = True Else isVBa = False
End Function

Public Function isUPX(AlamaT As String) As Boolean
On Error Resume Next
Dim AA As String
    Open AlamaT For Binary As #1
        AA = Space(LOF(1))
        Get #1, , AA
    Close #1
If InStr(AA, "UPX1") - InStr(AA, "UPX0") = 40 Then isUPX = True Else isUPX = False
End Function

