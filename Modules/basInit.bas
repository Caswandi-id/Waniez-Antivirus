Attribute VB_Name = "basInit"

Dim cImgMal  As New gComCtl
Public Sub InitAplikasi()

    ' init ektensi penjara
    JailExt = "." & ChrW$(&H634) & ChrW$(&H630) & ChrW$(&H628) & ChrW$(&H632) & ChrW$(&H6BE)
    ' ini folder jail
    FolderJail = Left(GetSpecFolder(WINDOWS_DIR), 3) & "Wan'iez Antivirus"
    'Init Var BERHENTI
    BERHENTI = True
With frmMain
    Call LoadConfig '(Environ$("windir") & "\MO.LOG") ' load config
    Call InitLanguange(LangUsed): Call loadRTP: Call loadPassword: Call loadLNK: Call loadCONTEKMENU:
    Call loadSTARTUP: Call SystemEditor: Call BuildListView: Call BacaDatabase: Call InitImageList:
    Call LoadDataIcon: Call InitPHPattern: Call ENUM_PROSES(.lvProses, frmMain.picBuffer): Call BuatPenjara
    Call READ_DATA_JAIL(FolderJail): Call ListVirus(.lstListWorm): Call BuilDirTree
    Call EnumPlugin(GetFilePath(App_FullPathW(False)) & "\plugin", .lstPlugin)
    Call EnumLangAvalaible(GetFilePath(App_FullPathW(False)) & "\lang", FrmConfig.lstLanguage)
    Call Init_Dtabase 'Init_UserDatabase
    JumPathExcep = ReadExceptPath(GetFilePath(App_FullPathW(False)) & "\Path.lst", .lstExceptFolder)
    JumFileExcep = ReadExceptFile(GetFilePath(App_FullPathW(False)) & "\File.lst", .lstExceptFile)
    JumRegExcep = ReadExceptReg(GetFilePath(App_FullPathW(False)) & "\Reg.lst", .lstExceptReg)
If UCase$(Left$(Command, 2)) <> "-K" Then
    Shell_NotifyIcon NIM_DELETE, nID
    'rtp
    If FrmConfig.ck8.Value = 1 Then
    Call UpdateIcon(FrmSysTray.Icon, "Wan'iez Antivirus™ - Your System Is Secured", FrmSysTray)
    Else
   Call UpdateIconRTPmati(frmMain.Icon, "Wan'iez Antivirus™ - Your System Is Not Secured", FrmSysTray)
    End If
End If
'Call Regis_EXT_Waniez
Call GetOS

End With

End Sub
Public Sub BuilDirTree()
If Left$(Command, 2) = "-K" Then
   frmMain.DirTree.LoadTreeDir False, False
   RegNode = False
   StartUpNode = False
   ProsesNode = False

Else
   frmMain.DirTree.LoadTreeDir True, False
   RegNode = True
   StartUpNode = True
   ProsesNode = True
End If
End Sub

' Init sesudah tampilan muncul
Public Sub InitAplikasi2()
Dim nFS As Long
If UCase$(Left$(Command, 2)) <> "-K" Then
    frmMain.terapkanIcon
End If
    nFS = EnumFileSystem
    
    If nFS = 0 Then
        ' TampilPesan frmmain.Popup, i_bahasa(23) & " !", Kuning

       TampilkanBalon frmMain, i_bahasa(23) & " !", i_bahasa(27), NIIF_WARNING
      ' Sleep 2000
    End If
    
    If IsWinXPOS = False Then
       If FrmConfig.ck3.Value = 1 Then TampilkanBalon frmMain, j_bahasa(40) & " !", NIIF_WARNING
'TampilkanBalon frmMain, j_bahasa(40) & " : " & h_bahasa(2) & " !", i_bahasa(27), NIIF_WARNING: frmMain.ck3.Value = 0: frmMain.ck3.Enabled = False
    End If


End Sub
Private Sub BuildListView()

With frmMain
     With .lvMalware '=================================================== Listview Malware
          .Font.FaceName = "Tahoma"
          .CheckBoxes = True
          '.Font.
          .Columns.Add , e_bahasa(0), , , lvwAlignCenter, 2200
          .Columns.Add , e_bahasa(1), , , lvwAlignLeft, 5000
          .Columns.Add , e_bahasa(2), , , lvwAlignRight, 1300
          .Columns.Add , e_bahasa(3), , , lvwAlignLeft, 2200
                     
          ' Init image list
          Set .ImageList = cImgMal.NewImageList(16, 16, imlColor32)

     End With '===========================================================================
      
      With .ucListVirus2 '====================================================daftar virus
          .Font.FaceName = "Tahoma"
          .Columns.Add , e_bahasa(0), , , lvwAlignLeft, 1500
          .Columns.Add , e_bahasa(22), , , lvwAlignCenter, 700
          .Columns.Add , e_bahasa(21), , , lvwAlignLeft, 1000
          'img
     Set .ImageList = cImgMal.NewImageList(16, 16, imlColor32)
     .ImageList.AddFromDc frmMain.pic1.hdc, 16, 16: .ImageList.AddFromDc frmMain.pic2.hdc, 16, 16
     .ImageList.AddFromDc frmMain.pic3.hdc, 16, 16: .ImageList.AddFromDc frmMain.pic4.hdc, 16, 16
     .ImageList.AddFromDc frmMain.pic5.hdc, 16, 16: .ImageList.AddFromDc frmMain.pic6.hdc, 16, 16
     .ImageList.AddFromDc frmMain.pic13.hdc, 16, 16
         'list virus
AddInfoToListDua frmMain.ucListVirus2, "Alman.A", "Virus", "Detector", 0, 10
AddInfoToListDua frmMain.ucListVirus2, "Alman.B", "Virus", "Detector", 0, 10
AddInfoToListDua frmMain.ucListVirus2, "W32.Chir.B@mm.A", "Virus", "DisInfected", 3, 10
AddInfoToListDua frmMain.ucListVirus2, "W32.Downloader.NAE", "Virus", "Detector", 0, 10
AddInfoToListDua frmMain.ucListVirus2, "Win32/CnsMin", "Virus", "Detector", 0, 10
AddInfoToListDua frmMain.ucListVirus2, "W32.Expiro", "Virus", "Detector", 0, 10
AddInfoToListDua frmMain.ucListVirus2, "W32.Gaelicum.A", "Virus", "DisInfected", 3, 10
AddInfoToListDua frmMain.ucListVirus2, "W32.Lamechi.A", "Virus", "DisInfected", 3, 10
AddInfoToListDua frmMain.ucListVirus2, "W32.lethic.AA", "Virus", "Detector", 0, 10
AddInfoToListDua frmMain.ucListVirus2, "W32.Mabezat", "Virus", "Detector", 0, 10
AddInfoToListDua frmMain.ucListVirus2, "W32.Polyene", "Virus", "Detector", 0, 10
AddInfoToListDua frmMain.ucListVirus2, "W32.Recure", "Virus", "Detector", 0, 10
AddInfoToListDua frmMain.ucListVirus2, "W32.Sality.A", "Virus", "Detector", 0, 10
AddInfoToListDua frmMain.ucListVirus2, "W32.Sality.X", "Virus", "Detector", 0, 10
AddInfoToListDua frmMain.ucListVirus2, "W32.Sality.NAO", "Virus", "Detector", 0, 10
AddInfoToListDua frmMain.ucListVirus2, "W32.Sality.NBA", "Virus", "Detector", 0, 10
AddInfoToListDua frmMain.ucListVirus2, "W32.Sality.P", "Virus", "Detector", 0, 10
AddInfoToListDua frmMain.ucListVirus2, "W32.Slugin.A", "Virus", "Detector", 0, 10
AddInfoToListDua frmMain.ucListVirus2, "W32.Srigala.A", "Virus", "DisInfected", 3, 10
AddInfoToListDua frmMain.ucListVirus2, "W32.Tanatos.M", "Virus", "Detector", 0, 10
AddInfoToListDua frmMain.ucListVirus2, "W32.Virut.NBP", "Virus", "Detector", 0, 10
AddInfoToListDua frmMain.ucListVirus2, "W32.Virut.NBP.Variant", "Virus", "Detector", 0, 10
AddInfoToListDua frmMain.ucListVirus2, "W32.Virut.NBH", "Virus", "Detector", 0, 10
AddInfoToListDua frmMain.ucListVirus2, "W32.Virut.NBU", "Virus", "Detector", 0, 10
AddInfoToListDua frmMain.ucListVirus2, "W32.Virut.AA", "Virus", "Detector", 0, 10
AddInfoToListDua frmMain.ucListVirus2, "W32.Virut.AF", "Virus", "Detector", 0, 10
AddInfoToListDua frmMain.ucListVirus2, "W32.Virut.AJ", "Virus", "Detector", 0, 10
AddInfoToListDua frmMain.ucListVirus2, "W32.Virut.AI", "Virus", "Detector", 0, 10
AddInfoToListDua frmMain.ucListVirus2, "W32.Virut.AN", "Virus", "Detector", 0, 10
AddInfoToListDua frmMain.ucListVirus2, "W32.Virut.BA", "Virus", "Detector", 0, 10
AddInfoToListDua frmMain.ucListVirus2, "W32.Virut.OQ", "Virus", "Detector", 0, 10
AddInfoToListDua frmMain.ucListVirus2, "W32.Virut.C", "Virus", "Detector", 0, 10
AddInfoToListDua frmMain.ucListVirus2, "W32.Virut.O", "Virus", "Detector", 0, 10
AddInfoToListDua frmMain.ucListVirus2, "W32.Virut.S", "Virus", "Detector", 0, 10
AddInfoToListDua frmMain.ucListVirus2, "W32.Virut.BT", "Virus", "Detector", 0, 10
AddInfoToListDua frmMain.ucListVirus2, "W32.Virut.Z", "Virus", "Detector", 0, 10 'Win32/CnsMin
AddInfoToListDua frmMain.ucListVirus2, "Win32.Ramnit.H", "Virus", "Detector", 0, 10
AddInfoToListDua frmMain.ucListVirus2, "Win32.Ramnit.A", "Virus", "Disinfected", 3, 10
AddInfoToListDua frmMain.ucListVirus2, "Win32/NgrBot(one)", "Virus", "Detector", 3, 10
AddInfoToListDua frmMain.ucListVirus2, "W32/Ramnit.H", "Virus", "Detector", 3, 10
AddInfoToListDua frmMain.ucListVirus2, "Win32/Ramnit.CPL", "Virus", "Detector", 3, 10
AddInfoToListDua frmMain.ucListVirus2, "Win32/Ramnit.Ax", "Virus", "Detector", 3, 10
AddInfoToListDua frmMain.ucListVirus2, "Win32/Ramnit.G", "Virus", "Detector", 3, 10
AddInfoToListDua frmMain.ucListVirus2, "Win32/Vitro", "Virus", "Detector", 3, 10
AddInfoToListDua frmMain.ucListVirus2, "Lyzapo", "Virus", "Detector", 3, 10
AddInfoToListDua frmMain.ucListVirus2, "Conficker", "Virus", "Detector", 3, 10
AddInfoToListDua frmMain.ucListVirus2, "Win32/Sality", "Virus", "Detector", 3, 10
AddInfoToListDua frmMain.ucListVirus2, "Win32/WaterMark", "Virus", "Detector", 3, 10
AddInfoToListDua frmMain.ucListVirus2, "Win32/Ramnit", "Virus", "Detector", 3, 10
AddInfoToListDua frmMain.ucListVirus2, "RontokBro", "Virus", "Detector", 3, 10
AddInfoToListDua frmMain.ucListVirus2, "Conficker.G", "Virus", "Detector", 3, 10
AddInfoToListDua frmMain.ucListVirus2, "AMG", "Virus", "Detector", 3, 10
AddInfoToListDua frmMain.ucListVirus2, "Fanny", "Virus", "Detector", 3, 10
               
                JumlahVirusINT = .ListItems.Count
                frmMain.LbInDB.Caption = "Internal Virus : " & .ListItems.Count & ""
                AutoLst frmMain.ucListVirus2
     End With '=================================================================================
     
     With .lvm31 '  ================================================================ bd external
          .Font.FaceName = "Tahoma"
          Set .ImageList = cImgMal.NewImageList(16, 16, imlColor32)
     .CheckBoxes = True
      .ImageList.AddFromDc frmMain.pic2.hdc, 16, 16
      .ImageList.AddFromDc frmMain.pic3.hdc, 16, 16
      .ImageList.AddFromDc frmMain.pic4.hdc, 16, 16
      .ImageList.AddFromDc frmMain.pic5.hdc, 16, 16
      .ImageList.AddFromDc frmMain.pic1.hdc, 16, 16
          .Columns.Add , "No", , , lvwAlignLeft, 700
          .Columns.Add , e_bahasa(0), , , lvwAlignLeft, 3000
          .Columns.Add , e_bahasa(19), , , lvwAlignLeft, 4000
     End With '=================================================================================
    
     With .lvRegistry ' ====================================================== Listview Registry
          .Font.FaceName = "Tahoma"
          .Columns.Add , e_bahasa(4), , , lvwAlignLeft, 2000
          .Columns.Add , e_bahasa(5), , , lvwAlignLeft, 7000
          .Columns.Add , e_bahasa(2), , , lvwAlignRight, 1000
          .Columns.Add , e_bahasa(3), , , lvwAlignLeft, 3000
          
           'Init image list
          Set .ImageList = cImgMal.NewImageList(16, 16, imlColor32)
     End With '==================================================================================

     With .lvHidden ' Listview Hidden
          .Font.FaceName = "Tahoma"
          .CheckBoxes = True
          .Columns.Add , e_bahasa(6), , , lvwAlignLeft, 1500
          .Columns.Add , e_bahasa(18), , , lvwAlignLeft, 1000
          .Columns.Add , e_bahasa(1), , , lvwAlignLeft, 3000
         ' .Columns.Add , e_bahasa(3), , , lvwAlignLeft, 2000
          
          ' Init image list
          Set .ImageList = cImgMal.NewImageList(16, 16, imlColor32)
     End With '==================================================================================
     
     With .lvInfo '========================================================= Listview Information
          .Font.FaceName = "Tahoma"
          .ColorFore = vbBlue
          .Columns.Add , e_bahasa(7), , , lvwAlignLeft, 2000
          .Columns.Add , e_bahasa(1), , , lvwAlignLeft, 4000
          .Columns.Add , e_bahasa(2), , , lvwAlignRight, 1300
          .Columns.Add , e_bahasa(3), , , lvwAlignLeft, 3000
          
           'InitImageList
          Set .ImageList = cImgMal.NewImageList(16, 16, imlColor32)
     End With '=================================================================================
     
     With .lvProses ' =========================================================== Listview Proses
          .Font.FaceName = "Tahoma"
          
          .Columns.Add , e_bahasa(8), , , lvwAlignLeft, 2000
          .Columns.Add , e_bahasa(9), , , lvwAlignCenter, 1200
          .Columns.Add , "PID", , , lvwAlignCenter, 1200
          .Columns.Add , e_bahasa(10), , , lvwAlignCenter, 1300
          .Columns.Add , e_bahasa(12), , , lvwAlignLeft, 1200
          .Columns.Add , e_bahasa(13), , , lvwAlignLeft, 1300
          .Columns.Add , e_bahasa(14), , , lvwAlignLeft, 1300
          .Columns.Add , e_bahasa(2), , , lvwAlignRight, 1300
          .Columns.Add , e_bahasa(11), , , lvwAlignLeft, 5000
     End With '==================================================================================
     
     With .lvJail '=================================================================== Quarantine
          .Font.FaceName = "Tahoma"
          
          .Columns.Add , e_bahasa(15), , , lvwAlignLeft, 2000
          .Columns.Add , e_bahasa(16), , , lvwAlignLeft, 4000
          .Columns.Add , e_bahasa(2), , , lvwAlignRight, 1300
          .Columns.Add , e_bahasa(17), , , lvwAlignRight, 1600
          
          Set .ImageList = cImgMal.NewImageList(16, 16, imlColor32)
     End With '==================================================================================
     
       With .lstKunci ' ============================================================== Driver Lock
          .Font.FaceName = "Tahoma"
          .Columns.Add , e_bahasa(20), , , lvwAlignLeft, 2000
          .Columns.Add , e_bahasa(21), , , lvwAlignLeft, 7000
          
           'Init image list
          Set .ImageList = cImgMal.NewImageList(16, 16, imlColor32)
          End With '==============================================================================
          
End With
    
          
With FrmRTP.lvRTP '================================================================= LIst Virus RTP
     .CheckBoxes = True
     .Font.FaceName = "Tahoma"
     .Columns.Add , e_bahasa(0), , , lvwAlignCenter, 1800
     .Columns.Add , e_bahasa(1), , , lvwAlignLeft, 4000
     .Columns.Add , e_bahasa(2), , , lvwAlignRight, 1200
     .Columns.Add , e_bahasa(3), , , lvwAlignLeft, 2000
     
     Set .ImageList = cImgMal.NewImageList(16, 16, imlColor32)
End With '=========================================================================================

End Sub
Private Sub InitImageList()

With frmMain

     With .lvMalware
          .ImageList.AddFromDc frmMain.pic1.hdc, 16, 16
          .ImageList.AddFromDc frmMain.pic2.hdc, 16, 16
          .ImageList.AddFromDc frmMain.pic3.hdc, 16, 16
          .ImageList.AddFromDc frmMain.pic4.hdc, 16, 16
          .ImageList.AddFromDc frmMain.pic5.hdc, 16, 16
          .ImageList.AddFromDc frmMain.pic6.hdc, 16, 16
          .ImageList.AddFromDc frmMain.pic13.hdc, 16, 16
          .ImageList.AddFromDc frmMain.picCaution.hdc, 16, 16
    End With
    With .lvRegistry
         .ImageList.AddFromDc frmMain.pic7.hdc, 16, 16
         .ImageList.AddFromDc frmMain.pic8.hdc, 16, 16
         .ImageList.AddFromDc frmMain.pic9.hdc, 16, 16
         .ImageList.AddFromDc frmMain.pic10.hdc, 16, 16
    End With
    With .lvHidden
    End With
    With .lvInfo
         .ImageList.AddFromDc frmMain.pic14.hdc, 16, 16
         .ImageList.AddFromDc frmMain.pic3.hdc, 16, 16
    End With
    With .lvJail
         .ImageList.AddFromDc frmMain.pic13.hdc, 16, 16
    End With
    
        With .lstKunci
        .ImageList.AddFromDc frmMain.Pic15.hdc, 16, 16
         .ImageList.AddFromDc frmMain.Pic16.hdc, 16, 16
    End With
    
End With

With FrmRTP.lvRTP
     .ImageList.AddFromDc frmMain.pic1.hdc, 16, 16
     .ImageList.AddFromDc frmMain.pic2.hdc, 16, 16
     .ImageList.AddFromDc frmMain.pic3.hdc, 16, 16
     .ImageList.AddFromDc frmMain.pic4.hdc, 16, 16
     .ImageList.AddFromDc frmMain.pic5.hdc, 16, 16
     .ImageList.AddFromDc frmMain.pic6.hdc, 16, 16
     .ImageList.AddFromDc frmMain.pic13.hdc, 16, 16
     .ImageList.AddFromDc frmMain.picCaution.hdc, 16, 16
End With
End Sub

Public Sub InitAplikasi3()
With frmMain
    Call InitLanguange(LangUsed)
    Call EnumLangAvalaible(GetFilePath(App_FullPathW(False)) & "\lang", FrmConfig.lstLanguage)
End With
End Sub
