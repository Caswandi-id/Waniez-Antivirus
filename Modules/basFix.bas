Attribute VB_Name = "basFix"
' Segala Sesuatu Tentang FIX objek
Private Const TheDocXls As String = "ÐÏà¡±á"
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)


Public Enum FIX_MALWARE_TYPE
    BY_CHECKED = 0
    BY_SELECT = 1
    BY_ALL = 2
End Enum
' FIX untuk MALWARE
Public Function Quar_Malware(lv As ucListView, TypeFix As FIX_MALWARE_TYPE, nScrolBatas As Integer) As String ' kembali ke status
' semua kepenjara aja ya
Dim MalwareName     As String
Dim MalwarePath     As String
Dim StatusNow       As String
Dim iCount          As Long
Dim HasBeenFix      As Long
Dim ToBeFix         As Long
Dim RetSuOrNot      As Long
Dim RetJail         As Byte
Dim Cawang          As String
Dim StrTypeFix      As String


Cawang = ChrW$(&H221A)
Call LepasSemuaKunci ' buat jaga-jaga

Select Case TypeFix
    Case 0: StrTypeFix = j_bahasa(13)
    Case 1: StrTypeFix = j_bahasa(14)
    Case Else: StrTypeFix = j_bahasa(15)
End Select

LBL_INIT: 'periksa yang mau di fix
For iCount = 1 To lv.ListItems.Count

    MalwareName = lv.ListItems.Item(iCount).Text

    If TypeFix = 0 Then ' by checked
       If lv.ListItems.Item(iCount).Checked = False Then GoTo LBL_LEWAT
    ElseIf TypeFix = 1 Then ' by select
       If lv.ListItems.Item(iCount).Selected = False Then GoTo LBL_LEWAT
    End If
    
    If Left$(MalwareName, 1) = Cawang Or Left$(MalwareName, 1) = "!" Then
       HasBeenFix = HasBeenFix + 1 ' yang dimaksud udah di fix
    Else
       ToBeFix = ToBeFix + 1
    End If
    
LBL_LEWAT:
Next

If ToBeFix > 0 Or HasBeenFix > 0 Then

Else
'frmRTP.Timer5.Enabled = True
    Exit Function ' gak ada yang mau di fix
End If
For iCount = 1 To lv.ListItems.Count
    If iCount > nScrolBatas Then lv.Scroll 0, 10

    MalwareName = lv.ListItems.Item(iCount).Text
    MalwarePath = lv.ListItems.Item(iCount).SubItem(2).Text
    StatusNow = lv.ListItems.Item(iCount).SubItem(4).Text
    
    If TypeFix = 0 Then ' by checked
       If lv.ListItems.Item(iCount).Checked = False Then GoTo LBL_LANJUT
    ElseIf TypeFix = 1 Then ' by select
       If lv.ListItems.Item(iCount).Selected = False Then GoTo LBL_LANJUT
    End If
    
    
    If Left$(MalwareName, 1) = Cawang Or Left$(MalwareName, 1) = "!" Then GoTo LBL_LANJUT ' yang sudah di FIX kasih cawang/!(gagal) soalnya
    If IsFileProtectedBySystem(MalwarePath) = False Then ' yakinkan bukan file milik system
       
          If LayakDirestore(MalwareName) = False Then
             If DeteksiMuatanPE32(MalwarePath) > 5000 Then
                Call PenjarakanMalware(FolderJail, MalwarePath, MalwareName, lv, iCount, f_bahasa(21)) ' ada status tambahan
             Else
                Call PenjarakanMalware(FolderJail, MalwarePath, MalwareName, lv, iCount, "")
             End If
          Else ' berarti di restore gak dipenjara
             lv.ListItems.Item(iCount).SubItem(4).Text = RestoreInfection(MalwareName, MalwarePath, RetSuOrNot)
             lv.ListItems.Item(iCount).Text = Cawang & " - " & MalwareName
             If RetSuOrNot = 1 Then lv.ListItems.Item(iCount).IconIndex = 5
          End If
       
    
    Else ' waduh file milik sistem
        lv.ListItems.Item(iCount).IconIndex = 7
        lv.ListItems.Item(iCount).Text = f_bahasa(19)
        lv.ListItems.Item(iCount).SubItem(4).Text = f_bahasa(20)
    End If
    
    DoEvents
    
LBL_LANJUT:
Next

End Function

' FIX untuk MALWARE
Public Function FiX_Malware(lv As ucListView, TypeFix As FIX_MALWARE_TYPE, nScrolBatas As Integer) As String ' kembali ke status
' Jika statusnya virus dan heuristic icon/vbs maka ke Penjara aj
Dim MalwareName     As String
Dim MalwarePath     As String
Dim StatusNow       As String
Dim iCount          As Long
Dim HasBeenFix      As Long
Dim ToBeFix         As Long
Dim RetSuOrNot      As Long
Dim RetJail         As Byte
Dim Cawang          As String
Dim StrTypeFix      As String


Cawang = ChrW$(&H221A)

Call LepasSemuaKunci ' buat jaga-jaga

Select Case TypeFix
    Case 0: StrTypeFix = j_bahasa(13)
    Case 1: StrTypeFix = j_bahasa(14)
    Case Else: StrTypeFix = j_bahasa(15)
End Select

LBL_INIT: 'periksa yang mau di fix
For iCount = 1 To lv.ListItems.Count

    MalwareName = lv.ListItems.Item(iCount).Text

    If TypeFix = 0 Then ' by checked
       If lv.ListItems.Item(iCount).Checked = False Then GoTo LBL_LEWAT
    ElseIf TypeFix = 1 Then ' by select
       If lv.ListItems.Item(iCount).Selected = False Then GoTo LBL_LEWAT
    End If
    
    If Left$(MalwareName, 1) = Cawang Or Left$(MalwareName, 1) = "!" Then
       HasBeenFix = HasBeenFix + 1 ' yang dimaksud udah di fix
    Else
       ToBeFix = ToBeFix + 1
    End If
    
LBL_LEWAT:
Next

If ToBeFix > 0 Or HasBeenFix > 0 Then

Else
    Exit Function ' gak ada yang mau di fix
End If


For iCount = 1 To lv.ListItems.Count
    If iCount > nScrolBatas Then lv.Scroll 0, 10

    MalwareName = lv.ListItems.Item(iCount).Text
    MalwarePath = lv.ListItems.Item(iCount).SubItem(2).Text
    StatusNow = lv.ListItems.Item(iCount).SubItem(4).Text
    
    If TypeFix = 0 Then ' by checked
       If lv.ListItems.Item(iCount).Checked = False Then GoTo LBL_LANJUT
    ElseIf TypeFix = 1 Then ' by select
       If lv.ListItems.Item(iCount).Selected = False Then GoTo LBL_LANJUT
    End If
    
    
    If Left$(MalwareName, 1) = Cawang Or Left$(MalwareName, 1) = "!" Then GoTo LBL_LANJUT ' yang sudah di FIX kasih cawang/!(gagal) soalnya
    
    If IsFileProtectedBySystem(MalwarePath) = False Then ' yakinkan bukan file milik system
       If LayakDihapus(MalwareName, MalwarePath, StatusNow) = True Then ' yang layak dihapus, harus dihapus
          'HapusFile MalwarePath
          If HapusFile(MalwarePath) = False Then ' wah gagal hapus kasih tanda
             lv.ListItems.Item(iCount).Text = "! - " & MalwareName
             lv.ListItems.Item(iCount).SubItem(4).Text = "Cannot be deleted !!"
          Else
             lv.ListItems.Item(iCount).IconIndex = 5
             lv.ListItems.Item(iCount).Text = Cawang & " - " & MalwareName
             lv.ListItems.Item(iCount).SubItem(4).Text = "Has been deleted !"
          End If
       Else ' gak layak kepenjara ajah atau restore klo bisa
          If LayakDirestore(MalwareName) = False Then
             If DeteksiMuatanPE32(MalwarePath) > 5000 Then
                Call PenjarakanMalware(FolderJail, MalwarePath, MalwareName, lv, iCount, f_bahasa(21)) ' ada status tambahan
             Else
                Call PenjarakanMalware(FolderJail, MalwarePath, MalwareName, lv, iCount, "")
             End If
          Else ' berarti di restore gak dipenjara
             lv.ListItems.Item(iCount).SubItem(4).Text = RestoreInfection(MalwareName, MalwarePath, RetSuOrNot)
             lv.ListItems.Item(iCount).Text = Cawang & " - " & MalwareName
             If RetSuOrNot = 1 Then lv.ListItems.Item(iCount).IconIndex = 5
          End If
       End If
    
    Else ' waduh file milik sistem
        lv.ListItems.Item(iCount).IconIndex = 7
        lv.ListItems.Item(iCount).Text = f_bahasa(19)
        lv.ListItems.Item(iCount).SubItem(4).Text = f_bahasa(20)
    End If
    DoEvents
    
LBL_LANJUT:
Next

End Function
' FIX untuk MALWARE
Public Function HAPUS_Malware(lv As ucListView, TypeFix As FIX_MALWARE_TYPE, nScrolBatas As Integer) As String ' kembali ke status
' Jika statusnya virus dan heuristic icon/vbs maka ke Penjara aj
Dim MalwareName     As String
Dim MalwarePath     As String
Dim StatusNow       As String
Dim iCount          As Long
Dim HasBeenFix      As Long
Dim ToBeFix         As Long
Dim RetSuOrNot      As Long
Dim RetJail         As Byte
Dim Cawang          As String
Dim StrTypeFix      As String


Cawang = ChrW$(&H221A)

Call LepasSemuaKunci ' buat jaga-jaga

Select Case TypeFix
    Case 0: StrTypeFix = j_bahasa(13)
    Case 1: StrTypeFix = j_bahasa(14)
    Case Else: StrTypeFix = j_bahasa(15)
End Select

LBL_INIT: 'periksa yang mau di fix
For iCount = 1 To lv.ListItems.Count

    MalwareName = lv.ListItems.Item(iCount).Text

    If TypeFix = 0 Then ' by checked
       If lv.ListItems.Item(iCount).Checked = False Then GoTo LBL_LEWAT
    ElseIf TypeFix = 1 Then ' by select
       If lv.ListItems.Item(iCount).Selected = False Then GoTo LBL_LEWAT
    End If
    
    If Left$(MalwareName, 1) = Cawang Or Left$(MalwareName, 1) = "!" Then
       HasBeenFix = HasBeenFix + 1 ' yang dimaksud udah di fix
    Else
       ToBeFix = ToBeFix + 1
    End If
    
LBL_LEWAT:
Next

If ToBeFix > 0 Or HasBeenFix > 0 Then

Else
    Exit Function ' gak ada yang mau di fix
End If


For iCount = 1 To lv.ListItems.Count
    If iCount > nScrolBatas Then lv.Scroll 0, 10

    MalwareName = lv.ListItems.Item(iCount).Text
    MalwarePath = lv.ListItems.Item(iCount).SubItem(2).Text
    StatusNow = lv.ListItems.Item(iCount).SubItem(4).Text
    
    If TypeFix = 0 Then ' by checked
       If lv.ListItems.Item(iCount).Checked = False Then GoTo LBL_LANJUT
    ElseIf TypeFix = 1 Then ' by select
       If lv.ListItems.Item(iCount).Selected = False Then GoTo LBL_LANJUT
    End If
    
    
    If Left$(MalwareName, 1) = Cawang Then GoTo LBL_LANJUT  ' yang sudah di FIX kasih cawang/!(gagal) soalnya
    
    If IsFileProtectedBySystem(MalwarePath) = False Then ' yakinkan bukan file milik system
       If LayakDirestore(MalwareName) = True Then
          ' berarti di restore gak dipenjara
             lv.ListItems.Item(iCount).SubItem(4).Text = RestoreInfection(MalwareName, MalwarePath, RetSuOrNot)
             lv.ListItems.Item(iCount).Text = Cawang & " - " & MalwareName
             If RetSuOrNot = 1 Then lv.ListItems.Item(iCount).IconIndex = 5
          Else
          
          If HapusFile(MalwarePath) = False Then ' wah gagal hapus kasih tanda
             lv.ListItems.Item(iCount).Text = "! - " & MalwareName
             lv.ListItems.Item(iCount).SubItem(4).Text = "Cannot be deleted !!"
          Else
             lv.ListItems.Item(iCount).IconIndex = 5
             lv.ListItems.Item(iCount).Text = Cawang & " - " & MalwareName
             lv.ListItems.Item(iCount).SubItem(4).Text = "Has been deleted !"
          End If
        End If
     Else
     lv.ListItems.Item(iCount).IconIndex = 7
        lv.ListItems.Item(iCount).Text = f_bahasa(19)
        lv.ListItems.Item(iCount).SubItem(4).Text = f_bahasa(20)
    End If
    'DoEvents
    
LBL_LANJUT:
Next

End Function

' FIX untuk RECENT
Public Function FiX_RECENT(lv As ucListView, TypeFix As FIX_MALWARE_TYPE, nScrolBatas As Integer) As String ' kembali ke status
' Jika statusnya virus dan heuristic icon/vbs maka ke Penjara aj
Dim MalwareName     As String
Dim MalwarePath     As String
Dim StatusNow       As String
Dim iCount          As Long
Dim HasBeenFix      As Long
Dim ToBeFix         As Long
Dim RetSuOrNot      As Long
Dim RetJail         As Byte
Dim Cawang          As String
Dim StrTypeFix      As String


Cawang = ChrW$(&H221A)

'Call LepasSemuaKunci ' buat jaga-jaga

Select Case TypeFix
    Case 0: StrTypeFix = j_bahasa(13)
    Case 1: StrTypeFix = j_bahasa(14)
    Case Else: StrTypeFix = j_bahasa(15)
End Select

LBL_INIT: 'periksa yang mau di fix
For iCount = 1 To lv.ListItems.Count

    MalwareName = lv.ListItems.Item(iCount).Text

    If TypeFix = 0 Then ' by checked
       If lv.ListItems.Item(iCount).Checked = False Then GoTo LBL_LEWAT
    ElseIf TypeFix = 1 Then ' by select
       If lv.ListItems.Item(iCount).Selected = False Then GoTo LBL_LEWAT
    End If
    
    If Left$(MalwareName, 1) = Cawang Or Left$(MalwareName, 1) = "!" Then
       HasBeenFix = HasBeenFix + 1 ' yang dimaksud udah di fix
    Else
       ToBeFix = ToBeFix + 1
    End If
    
LBL_LEWAT:
Next

If ToBeFix > 0 Or HasBeenFix > 0 Then


Else
    Exit Function ' gak ada yang mau di fix
End If


For iCount = 1 To lv.ListItems.Count
    If iCount > nScrolBatas Then lv.Scroll 0, 10

    MalwareName = lv.ListItems.Item(iCount).Text
    MalwarePath = lv.ListItems.Item(iCount).SubItem(2).Text
    StatusNow = lv.ListItems.Item(iCount).SubItem(3).Text
    
    If TypeFix = 0 Then ' by checked
       If lv.ListItems.Item(iCount).Checked = False Then GoTo LBL_LANJUT
    ElseIf TypeFix = 1 Then ' by select
       If lv.ListItems.Item(iCount).Selected = False Then GoTo LBL_LANJUT
    End If
    
    
    If Left$(MalwareName, 1) = Cawang Or Left$(MalwareName, 1) = "!" Then GoTo LBL_LANJUT ' yang sudah di FIX kasih cawang/!(gagal) soalnya
    
    If IsFileProtectedBySystem(MalwarePath) = False Then ' yakinkan bukan file milik system
        If LayakDihapus(MalwareName, MalwarePath, StatusNow) = True Then ' yang layak dihapus, harus dihapus
          'HapusFile MalwarePath
          If HapusFile(MalwarePath) = False Then ' wah gagal hapus kasih tanda
             lv.ListItems.Item(iCount).Text = "! - " & MalwareName
             lv.ListItems.Item(iCount).SubItem(3).Text = "Cannot be deleted !!"
          Else
             lv.ListItems.Item(iCount).IconIndex = 5
             lv.ListItems.Item(iCount).Text = Cawang & " - " & MalwareName
             lv.ListItems.Item(iCount).SubItem(3).Text = "Has been deleted !"
          End If
     Else
    End If
    End If
    'DoEvents
    
LBL_LANJUT:
Next

End Function

' FIX untuk Registry
Public Function FiX_REGISTRY(lv As ucListView, TypeFix As FIX_MALWARE_TYPE) As String  ' kembali ke status
' Jika statusnya virus dan heuristic icon/vbs maka ke Penjara aj
Dim InfoBadReg      As String
Dim KeyPath         As String
Dim Cawang          As String
Dim KeyName         As String
Dim PathReg         As String
Dim TrueData        As String
Dim NewStatus       As String
Dim iCount          As Integer
Dim MainKeyReg      As Long
Dim HasBeenFix      As Long
Dim ToBeFix         As Long

Dim IconIndek       As Byte


Cawang = ChrW$(&H221A)

Select Case TypeFix
    Case 0: StrTypeFix = j_bahasa(13)
    Case 1: StrTypeFix = j_bahasa(14)
    Case Else: StrTypeFix = j_bahasa(15)
End Select

LBL_INIT: 'periksa yang mau di fix
For iCount = 1 To lv.ListItems.Count

    MalwareName = lv.ListItems.Item(iCount).Text

    If TypeFix = 0 Then ' by checked
       If lv.ListItems.Item(iCount).Checked = False Then GoTo LBL_LEWAT
    ElseIf TypeFix = 1 Then ' by select
       If lv.ListItems.Item(iCount).Selected = False Then GoTo LBL_LEWAT
    End If
    
    If Left$(MalwareName, 1) = Cawang Or Left$(MalwareName, 1) = "!" Then
       HasBeenFix = HasBeenFix + 1 ' yang dimaksud udah di fix
    Else
       ToBeFix = ToBeFix + 1
    End If
    
LBL_LEWAT:
Next

If ToBeFix > 0 Or HasBeenFix > 0 Then

Else
    Exit Function ' gak ada yang mau di fix
End If


For iCount = 1 To lv.ListItems.Count
    If iCount > 17 Then lv.Scroll 0, 10
    
    KeyName = lv.ListItems.Item(iCount).Text ' harusnya namanya value bukan key, gpp udah terlanjur
    InfoBadReg = lv.ListItems.Item(iCount).SubItem(4).Text
    KeyPath = lv.ListItems.Item(iCount).SubItem(2).Text ' belum dibuffer
    
    If TypeFix = 0 Then ' by checked
       If lv.ListItems.Item(iCount).Checked = False Then GoTo LBL_LANJUT
    ElseIf TypeFix = 1 Then ' by select
       If lv.ListItems.Item(iCount).Selected = False Then GoTo LBL_LANJUT
    End If
    
    
    If Left$(KeyName, 1) = Cawang Then GoTo LBL_LANJUT ' yang sudah di FIX kasih cawang soalnya

    MainKeyReg = GetKeyInLong(KeyPath)
    PathReg = GetKeyPathClean(KeyPath)
    TrueData = GetTrueDataReg(InfoBadReg)
    Select Case Left$(InfoBadReg, 15) 'ambil 15 ajah
           Case Left$(f_bahasa(8), 15)
                IconIndek = 3
                DeleteValue MainKeyReg, PathReg, KeyName
                NewStatus = f_bahasa(10)
           Case "Startup Virus (" ' bahasa tetap
                IconIndek = 3
                DeleteValue MainKeyReg, PathReg, KeyName
                NewStatus = f_bahasa(10)
           Case Left$("Bad String Value, Should : ", 15) 'bahasa tetap
                IconIndek = 3
                SetStringValue MainKeyReg, PathReg, ClearDefaultStr(KeyName), TrueData
                NewStatus = f_bahasa(11)
           Case Left$("Bad DWORD Value, Should :", 15) 'bahasa tetap
                IconIndek = 1
                SetDwordValue MainKeyReg, PathReg, KeyName, CLng(TrueData) 'bahasa tetap
                NewStatus = f_bahasa(12)
           Case Left$(f_bahasa(9), 15)
                IconIndek = 3
                DeleteValue MainKeyReg, PathReg, KeyName
                NewStatus = f_bahasa(10)
           Case Left$("Kunci startup tak lazim", 15)
                IconIndek = 3
                DeleteValue MainKeyReg, PathReg, KeyName
                NewStatus = f_bahasa(10)
    End Select
  
    lv.ListItems.Item(iCount).IconIndex = IconIndek
    lv.ListItems.Item(iCount).Text = Cawang & " - " & KeyName
    lv.ListItems.Item(iCount).SubItem(4).Text = NewStatus
    

    'DoEvents
LBL_LANJUT:
Next

End Function


Public Function FIX_HIDDEN(lv As ucListView, TypeFix As FIX_MALWARE_TYPE)
Dim FileName    As String
Dim ObjPath     As String
Dim StatusObj   As String
Dim Cawang      As String
Dim IconIndek   As Long

Cawang = ChrW$(&H221A)


For iCount = 1 To lv.ListItems.Count
    If iCount > 17 Then lv.Scroll 0, 10
    FileName = lv.ListItems.Item(iCount).Text
    ObjPath = lv.ListItems.Item(iCount).SubItem(3).Text
    StatusObj = lv.ListItems.Item(iCount).SubItem(2).Text
    
    If TypeFix = 0 Then ' by checked
       If lv.ListItems.Item(iCount).Checked = False Then GoTo LBL_LANJUT
    ElseIf TypeFix = 1 Then ' by select
       If lv.ListItems.Item(iCount).Selected = False Then GoTo LBL_LANJUT
    End If
    
    If Left$(FileName, 1) = Cawang Then GoTo LBL_LANJUT ' yang sudah di FIX kasih cawang soalnya
    
    
    If IsFileProtectedBySystem(ObjPath) = False Then ' yakinkan bukan file milik system
    
       SetFileAttributes StrPtr(ObjPath), 0
    
       If StatusObj = "Hidden File" Then
          'lv.ListItems.Item(iCount).IconIndex = 0
          lv.ListItems.Item(iCount).Cut = False
          lv.ListItems.Item(iCount).SubItem(2).Text = f_bahasa(13)
          lv.ListItems.Item(iCount).Text = Cawang & " - " & FileName
       Else
         'lv.ListItems.Item(iCount).IconIndex = 1
          lv.ListItems.Item(iCount).Cut = False
          lv.ListItems.Item(iCount).SubItem(2).Text = f_bahasa(14)
          lv.ListItems.Item(iCount).Text = Cawang & " - " & FileName
       End If
      'DoEvents
    Else
       lv.ListItems.Item(iCount).IconIndex = 4
       lv.ListItems.Item(iCount).Text = f_bahasa(19)
       lv.ListItems.Item(iCount).SubItem(2).Text = f_bahasa(20)
    End If
LBL_LANJUT:
Next
End Function


' Klo bisa restore panggil ini
Private Function RestoreInfection(MalwareName As String, sFilePath As String, ByRef SuccesOrNot As Long) As String
Dim nCount  As Long
Select Case MalwareName
       Case "W32/Srigala.A"
            nCount = HealInfeksiSrigala(sFilePath)
            RestoreInfection = nCount & " " & f_bahasa(15)
            If nCount = 0 Then SuccesOrNot = 0 Else SuccesOrNot = 1 ' sukses atau gagal
       Case "Win32/Gaelicum.A"
            nCount = DisInfectTenga(sFilePath)
            If nCount = 1 Then
               SuccesOrNot = 1
               RestoreInfection = "Success Disinfected"
            Else
               SuccesOrNot = 0 ' sukses
               RestoreInfection = "Fail Disinfected"
            End If
       Case "Chirb@mm"
            nCount = DisInfectRunouce(sFilePath)
            If nCount = 1 Then
               SuccesOrNot = 1
               RestoreInfection = "Success Disinfected"
            Else
               SuccesOrNot = 0 ' sukses
               RestoreInfection = "Fail Disinfected"
            End If
        Case "Win32/Ramnit.A"
            nCount = FixRamnit(sFilePath)
            If nCount = 1 Then
               SuccesOrNot = 1
               RestoreInfection = "Success Disinfected"
            Else
               SuccesOrNot = 0 ' sukses
               RestoreInfection = "Success Disinfected"
            End If
End Select ' baru bisa heal vir2 di atas ajah

End Function


'Klo mau penjara pake ini
Private Sub PenjarakanMalware(sFolderJail As String, MalwarePath As String, MalwareName As String, lv As ucListView, ItemIndek As Long, SpecialStatus As String)
Dim Cawang          As String

Cawang = ChrW$(&H221A)

RetJail = JailFile(MalwarePath, sFolderJail, MalwareName)
  If RetJail = 0 Then ' gagal memenjarakan dan hapus sumber
     lv.ListItems.Item(ItemIndek).Text = "! - " & MalwareName
     lv.ListItems.Item(ItemIndek).SubItem(4).Text = SpecialStatus & f_bahasa(18)
  ElseIf RetJail = 1 Then ' sukses ke penjara tapi gagal remove sumber
     lv.ListItems.Item(ItemIndek).Text = "! - " & MalwareName
     lv.ListItems.Item(ItemIndek).SubItem(4).Text = SpecialStatus & f_bahasa(17)
  ElseIf RetJail = 3 Then ' sukses ke penjara tapi gagal remove sumber
     lv.ListItems.Item(ItemIndek).Text = "! - " & MalwareName
     lv.ListItems.Item(ItemIndek).SubItem(4).Text = SpecialStatus & "File terlalu besar tidak dikarantina"
  Else ' berhasil semua
     lv.ListItems.Item(ItemIndek).IconIndex = 6
     lv.ListItems.Item(ItemIndek).Text = Cawang & " - " & MalwareName
     lv.ListItems.Item(ItemIndek).SubItem(4).Text = SpecialStatus & f_bahasa(16)
  End If
End Sub


'buat cek LvMalware ajah (Reg ga Usah)
Public Function AdakahYangBelumDiFix(lv As ucListView) As Boolean
Dim iCount  As Long
Dim StrM    As String
Dim Caeang  As String

Cawang = ChrW$(&H221A)

For iCount = 1 To lv.ListItems.Count
    StrM = lv.ListItems.Item(iCount).Text
    If Left$(StrM, 1) = Cawang Or Left$(StrM, 1) = "!" Then
    Else
        GoTo LBL_BROAD
    End If
Next

AdakahYangBelumDiFix = False

Exit Function

LBL_BROAD: ' berakhir
AdakahYangBelumDiFix = True
End Function

'''' Fungsi-Fungsi Buffer
Private Function GetKeyInLong(sPathReg As String) As Long
Dim MainKeyReg    As String
Dim SplitPath()    As String

    SplitPath = Split(sPathReg, "\")
    MainKeyReg = SplitPath(0) ' yang ke 0

GetKeyInLong = StringToMain(MainKeyReg)
End Function

Private Function GetKeyPathClean(sPathDirty As String) As String
Dim SplitTmp()  As String
Dim StrTmp      As String
' kalo ada string (  =>) -- sebagai saparator
If InStr(sPathDirty, " =>") > 0 Then
   SplitTmp = Split(sPathDirty, " =>")
   StrTmp = SplitTmp(0)
   StrTmp = Mid$(StrTmp, InStr(StrTmp, "\") + 1) ' potong main key
   StrTmp = GetFilePath(StrTmp)
   GetKeyPathClean = StrTmp
Else
   StrTmp = Mid$(sPathDirty, InStr(sPathDirty, "\") + 1) ' potong main key
   StrTmp = GetFilePath(StrTmp)
   GetKeyPathClean = StrTmp
End If
End Function

' dipakai diluar module (tapi dikit)
Public Function GetKeyPathAndValueClean(sPathDirty As String) As String
Dim SplitTmp()  As String
Dim StrTmp      As String
' kalo ada string (  =>) -- sebagai saparator
If InStr(sPathDirty, " =>") > 0 Then
   SplitTmp = Split(sPathDirty, " =>")
   StrTmp = SplitTmp(0)
   GetKeyPathAndValueClean = StrTmp
Else
   GetKeyPathAndValueClean = sPathDirty
End If
End Function


Private Function GetTrueDataReg(sInformasi As String) As String ' klo untuk long di konversi ajh
Dim SplitTmp()    As String
Dim StrTmp        As String
Dim nPos          As Long
' should : -> ambil ":" untuk split
nPos = InStr(sInformasi, ":")
If nPos > 0 Then
   StrTmp = Mid$(sInformasi, nPos + 1)
   StrTmp = BuangSpaceAwal(StrTmp)
Else
   StrTmp = ""
End If
GetTrueDataReg = StrTmp
End Function

Private Function LayakDihapus(MalwareName As String, MalwarePath As String, StatusVirs As String) As Boolean
Dim nLong   As Long

LBL_SELEKSI1:
Select Case MalwareName
       Case "Chirb@mm": LayakDihapus = False
       Case "Win32/Sality.A": LayakDihapus = False
       Case "Win32/Alman.A": LayakDihapus = False
       Case "Win32/Alman.B": LayakDihapus = False
       Case "Win32/Tanatos.M": LayakDihapus = False
       Case "Suspect/Tanatos.M-1": LayakDihapus = False
       Case "Suspect/Tanatos.M-2": LayakDihapus = False
       Case "Suspect/Tanatos.M-3": LayakDihapus = False
       Case "Win32/Ramnit.H": LayakDihapus = False
       Case "Suspect/Sality.A": LayakDihapus = False
       Case "Win32/Mabezat": LayakDihapus = False
       Case "Win32/Gaelicum.A": LayakDihapus = False
       Case "Win32/Ramnit.H": LayakDihapus = False
       Case "70% Suspect Tanatos": LayakDihapus = False
       Case "Virus [ArrS Method]": LayakDihapus = False
       Case "Win32/Downloader.NAE": LayakDihapus = False
       Case "Win32/Expiro": LayakDihapus = False
       Case "Suspect.PEHeur.1": LayakDihapus = False
       Case "Suspect.PEHeur.2": LayakDihapus = False
       Case "Win32/Ramnit.A": LayakDihapus = False
       Case "Win32/Ramnit.H": LayakDihapus = False
       Case "Win32/Virut.S": LayakDihapus = False
       Case "Win32/Virut.BT": LayakDihapus = False
       Case "Win32/Virut.NBH": LayakDihapus = False
       Case "Win32/Ramnit.H": LayakDihapus = False
       Case "Win32/Virut.NBP": LayakDihapus = False
       Case "Win32/Virut.NBC": LayakDihapus = False
       Case "W32.Lethic.AA": LayakDihapus = False
       Case "W32/Srigala.A": LayakDihapus = False ' W32/Srigala, coba restore
       Case "W32.Slugin.A": LayakDihapus = False
       Case "W32.Sality.E": LayakDihapus = False
       Case "W32.Sality.NBA": LayakDihapus = False
       Case "W32.Sality.X": LayakDihapus = False
       Case "W32.Sality.N": LayakDihapus = False
       Case "W32.Sality.P": LayakDihapus = False
       Case "W32.Sality.NAO": LayakDihapus = False
       Case "W32.Virut.NBP": LayakDihapus = False
       Case "W32.Virut.NBU": LayakDihapus = False
       Case "W32.Virut.AA": LayakDihapus = False
       Case "W32.Virut.AF": LayakDihapus = False
       Case "W32.Virut.AJ": LayakDihapus = False
       Case "W32.Virut.AI": LayakDihapus = False
       Case "W32.Virut.BA": LayakDihapus = False
       Case "W32.Virut.OQ": LayakDihapus = False
       Case "W32.Lamechi.A": LayakDihapus = False
       Case "Win32/CnsMin": LayakDihapus = False
       Case "Suspect.PEHeur.2": LayakDihapus = False
       Case "Suspect.PEHeur.1": LayakDihapus = False
       Case "Win32/NgrBot(one)": LayakDihapus = False
       Case Else: GoTo LBL_SELEKSI2
End Select

Exit Function

LBL_SELEKSI2:
Select Case StatusVirs
       Case f_bahasa(2): LayakDihapus = False ' suspected file
       Case Else: GoTo LBL_SELEKSI3
End Select

Exit Function

LBL_SELEKSI3:
nLong = DeteksiMuatanPE32(MalwarePath)
Select Case nLong
       Case Is >= 1000: LayakDihapus = False: Exit Function ' klo memuat lebih dari 1.000 bytes tambahan stataus layak=false (siapa tahu ada dokumen atau lainya)
       Case Else: LayakDihapus = True 'true TAPI cek lagi
End Select

LBL_SELEKSI4: ' seleksi jika file doc/xls (jaga-jaga false detek)
If YakinkanFileTakPunyaPola(MalwarePath, TheDocXls) = True Then
   LayakDihapus = True
Else
   LayakDihapus = False
End If

End Function
 
 ' Kalo gak layak dihapus, aksinya restore atau goto jail
Private Function LayakDirestore(MalwareName As String) As Boolean
Select Case MalwareName
       Case "W32/Srigala.A": LayakDirestore = True
       Case "Win32/Gaelicum.A": LayakDirestore = True
       Case "Chirb@mm": LayakDirestore = True
       Case "Win32/Ramnit.A": LayakDirestore = True
       Case Else: LayakDirestore = False
End Select
End Function

Private Function ClearDefaultStr(strKeyName As String) As String
If strKeyName = "(Default)" Then ClearDefaultStr = "" Else ClearDefaultStr = strKeyName
End Function

Private Function YakinkanFileTakPunyaPola(ByRef PathFileNya As String, ByRef sPolaDimksd As String) As Boolean
Dim TheHand  As Long
Dim OData()  As Byte
Dim TheStrng As String
    TheHand = GetHandleFile(PathFileNya)
If TheHand > 0 Then
   Call ReadUnicodeFile2(TheHand, 1, Len(sPolaDimksd), OData)
   TheStrng = StrConv(OData, vbUnicode)
   If sPolaDimksd <> TheStrng Then
      YakinkanFileTakPunyaPola = True
   Else
      YakinkanFileTakPunyaPola = False
   End If
Else
   YakinkanFileTakPunyaPola = True
End If
TutupFile TheHand
End Function
Public Function KillFolder(ByVal Fullpath As String) _
   As Boolean
 
On Error Resume Next
Dim oFso As New Scripting.FileSystemObject

'deletefolder method does not like the "\"
'at end of fullpath

If Right$(Fullpath, 1) = "\" Then Fullpath = _
    Left$(Fullpath, Len(Fullpath) - 1)

If oFso.FolderExists(Fullpath) Then
    
    'Setting the 2nd parameter to true
    'forces deletion of read-only files
    oFso.DeleteFolder Fullpath, True
    
    KillFolder = err.Number = 0 And _
      oFso.FolderExists(Fullpath) = False
End If

End Function
