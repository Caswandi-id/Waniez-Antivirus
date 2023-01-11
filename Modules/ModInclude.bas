Attribute VB_Name = "ModInclude"
Public Function LoadDll()
If Dir(App.path & "\WanUI.dll") = "" Then
    MsgBox j_bahasa(16) & vbCrLf & j_bahasa(12), vbCritical
End
End If
If Dir(App.path & "\WanSM.dll") = "" Then
    MsgBox j_bahasa(17) & vbCrLf & j_bahasa(12), vbCritical
End
End If
If Dir(App.path & "\WanUDB.dll") = "" Then
    MsgBox j_bahasa(18) & vbCrLf & j_bahasa(12), vbCritical
End
End If
End Function
'Main Menu
Public Function summary() 'Home
With frmMain
.PicMainSummary.Visible = True: .PicMainScan.Visible = False: .Pictolls.Visible = False: .PicQuar.Visible = False: .PicUpd.Visible = False
.Picture = LoadPictureDLL(100)
End With
End Function
Public Function scancomputer() 'Scan
With frmMain
.PicMainSummary.Visible = False: .PicMainScan.Visible = True: .Pictolls.Visible = False: .PicQuar.Visible = False: .PicUpd.Visible = False
.Picture = LoadPictureDLL(101)
End With
End Function
Public Function additionalprotec()
With frmMain
.PicMainSummary.Visible = False: .PicMainScan.Visible = False: .Pictolls.Visible = True: .PicQuar.Visible = False: .PicUpd.Visible = False
.Picture = LoadPictureDLL(102)
End With
End Function
Public Function update()
With frmMain
.PicMainSummary.Visible = False: .PicMainScan.Visible = False: .Pictolls.Visible = False: .PicQuar.Visible = False: .PicUpd.Visible = True
.Picture = LoadPictureDLL(104)
End With
End Function
Public Function quaranntine()
With frmMain
.PicMainSummary.Visible = False: .PicMainScan.Visible = False: .Pictolls.Visible = False: .PicQuar.Visible = True: .PicUpd.Visible = False
.Picture = LoadPictureDLL(103)
End With
End Function
'Akhiri Main Menu
'Cleaner Virus
Public Function CkFixHidden()
With frmMain
.cmdFixHiddenall.Enabled = False: .cmdFixHidden.Enabled = False
    Call FIX_HIDDEN(.lvHidden, BY_CHECKED)
.cmdFixHiddenall.Enabled = True: .cmdFixHidden.Enabled = True
   AutoLst .lvHidden
   End With
End Function
Public Function AllFixHidden()
With frmMain
.cmdFixHiddenall.Enabled = False: .cmdFixHidden.Enabled = False
    Call FIX_HIDDEN(.lvHidden, BY_ALL)
.cmdFixHiddenall.Enabled = True: .cmdFixHidden.Enabled = True
   AutoLst .lvHidden
       End With
End Function
Public Function CkFixMalware()
With frmMain
.cmdFixMalware.Enabled = False: .cmdFixMalwareAll.Enabled = False: .CmdQuarAll.Enabled = False
    FiX_Malware .lvMalware, BY_CHECKED, 16
.cmdFixMalwareAll.Enabled = True: .cmdFixMalware.Enabled = True: .CmdQuarAll.Enabled = True
    AutoLst frmMain.lvMalware
    Call READ_DATA_JAIL(FolderJail)
        End With
End Function
Public Function AllFixMalware()
With frmMain
.cmdFixMalware.Enabled = False: .cmdFixMalwareAll.Enabled = False: .CmdQuarAll.Enabled = False
    Call FiX_Malware(.lvMalware, BY_ALL, 16)
.cmdFixMalwareAll.Enabled = True: .cmdFixMalware.Enabled = True: .CmdQuarAll.Enabled = True
    Call READ_DATA_JAIL(FolderJail)
    AutoLst .lvMalware
        End With
End Function
Public Function CkFixReg()
With frmMain
.cmdFixRegAll.Enabled = False: .cmdFixReg.Enabled = False
    Call FiX_REGISTRY(.lvRegistry, BY_CHECKED)
.cmdFixRegAll.Enabled = True: .cmdFixReg.Enabled = True
    AutoLst .lvRegistry
        End With
End Function
Public Function AllFixReg()
With frmMain
.cmdFixRegAll.Enabled = False: .cmdFixReg.Enabled = False
    Call FiX_REGISTRY(.lvRegistry, BY_ALL)
.cmdFixRegAll.Enabled = True: .cmdFixReg.Enabled = True
    AutoLst .lvRegistry
    End With
End Function
