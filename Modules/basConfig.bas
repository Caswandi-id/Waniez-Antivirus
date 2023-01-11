Attribute VB_Name = "basCOnfig"
Public Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long 'Tanpa fungsi LockWindowUpdate
Public Function loadPassWordAktif()
With FrmConfig
'Text Color
.TextConfPass.Enabled = True: .TextNewPass.Enabled = True: .TextOlPass.Enabled = True
.TextConfPass.BackColor = &HFFFFFF: .TextNewPass.BackColor = &HFFFFFF: .TextOlPass.BackColor = &HFFFFFF
'Pass App
.CkPotection.Enabled = True: .CkQuar.Enabled = True: .CkUpdate.Enabled = True: .CkAdditional.Enabled = True
.CkDellUs.Enabled = True: .CkExit.Enabled = True:
End With
End Function
Public Function loadPassWordNonAktif()
With FrmConfig
'Text Color
.TextConfPass.Enabled = False: .TextNewPass.Enabled = False: .TextOlPass.Enabled = False
.TextConfPass.BackColor = &HE0E0E0: .TextNewPass.BackColor = &HE0E0E0: .TextOlPass.BackColor = &HE0E0E0
'Pass App
.CkExit.Enabled = False:
.CkDellUs.Enabled = False:
.CkPotection.Enabled = False: .CkQuar.Enabled = False: .CkUpdate.Enabled = False: .CkAdditional.Enabled = False
End With
End Function
Public Function loadPassword()
With FrmConfig
Dim A As String
A = GetSTRINGValue(&H80000001, "Software\Wan'iez\", "password")
.CkPass.Value = GetDWORDValue(&H80000001, "Software\Wan'iez\", "EnablePassword")
If .CkPass.Value = 1 And A = "" Then
Call loadPassWordAktif
ElseIf .CkPass.Value = 1 And A <> "" Then
'Text Color
.TextConfPass.Enabled = False: .TextNewPass.Enabled = False: .TextOlPass.Enabled = False
.TextConfPass.BackColor = &HFFFFFF: .TextNewPass.BackColor = &HFFFFFF: .TextOlPass.BackColor = &HFFFFFF
'Pass App
.CkDellUs.Enabled = True: .CkExit.Enabled = True
.CkPotection.Enabled = True: .CkQuar.Enabled = True: .CkUpdate.Enabled = True: .CkAdditional.Enabled = True
Else
Call loadPassWordNonAktif
End If
End With
End Function
Public Function loadSTARTUP()
With FrmConfig
.ck7.Value = GetDWORDValue(&H80000001, "Software\Wan'iez\", "Startup")  ' autorun
If .ck7.Value = 1 Then
   Call InstalInReg(App_FullPathW(False), " -A")
   FrmSysTray.mnRun.Checked = True
Else
   Call UnInstalInReg("Wan'iez Antivirus")
   FrmSysTray.mnRun.Checked = False
End If

End With
End Function
Public Function loadCONTEKMENU()
wadahAV = GetFilePath(App_FullPathW(False))

With FrmConfig
.ck12.Value = GetDWORDValue(&H80000001, "Software\Wan'iez\", "ConteckMenu") ' Contek Menu
If .ck12.Value = 1 Then
frmMain.aktifContek

Else
frmMain.UnaktifContek
End If
End With
End Function

Public Function loadLNK()
With FrmConfig
.ck15.Value = GetDWORDValue(&H80000001, "Software\Wan'iez\", "Shortcut") ' shortcut

If .ck15.Value = 1 Then
Call LayOnDekstop
Else
HapusFile GetSpecFolder(DEKSTOP_PATH) & "\Wan'iez Antivirus.lnk"
End If
End With
End Function
Public Function loadRTP()
FrmConfig.ck8.Value = GetDWORDValue(&H80000001, "Software\Wan'iez\", "Rtp")
If FrmConfig.ck8.Value = 1 Then
   StatusRTP = True
   If UCase$(Left$(Command, 2)) <> "-K" Then frmMain.terapkanIcon
   FrmSysTray.mnEPro.Checked = True
   frmMain.PicMainSummary.Picture = LoadPictureDLL(900) 'on
   frmMain.LbstatusRTP(0).ForeColor = &HC000&
   frmMain.LbstatusRTP(0).Caption = "SECURED"
   frmMain.LbstatusRTP(1).Caption = "Your System is fully protected"
   frmMain.RTPdet.Visible = False
   frmMain.BtnAtur.Visible = falue
   frmMain.cekstatus
Else
   StatusRTP = False
  If UCase$(Left$(Command, 2)) <> "-K" Then frmMain.terapkanIcon
   FrmSysTray.mnEPro.Checked = False
  frmMain.PicMainSummary.Picture = LoadPictureDLL(901) 'off
  frmMain.LbstatusRTP(0).ForeColor = &HC0&
  frmMain.LbstatusRTP(0).Caption = "NOT SECURED"
  frmMain.LbstatusRTP(1).Caption = "Your System is not fully protected"
  frmMain.RTPdet.Visible = True: frmMain.RTPdet.Caption = ""
  frmMain.BtnAtur.Visible = True
  frmMain.cekstatus
  'TampilkanBalon frmMain, i_bahasa(25) & " !", i_bahasa(27), NIIF_ERROR


End If
End Function
Public Function Save()
SetDwordValue &H80000001, "Software\Wan'iez\", "Rtp", FrmConfig.ck8.Value 'rtp
SetDwordValue &H80000001, "Software\Wan'iez\", "ScanFD", FrmConfig.ck10.Value 'scanfd
SetDwordValue &H80000001, "Software\Wan'iez\", "Filter", FrmConfig.ck1.Value 'filter
SetDwordValue &H80000001, "Software\Wan'iez\", "Splash", FrmConfig.ck11.Value 'splas
SetDwordValue &H80000001, "Software\Wan'iez\", "ConteckMenu", FrmConfig.ck12.Value 'conteckmenu
SetDwordValue &H80000001, "Software\Wan'iez\", "Shortcut", FrmConfig.ck15.Value 'shortcut
SetDwordValue &H80000001, "Software\Wan'iez\", "Heuristic", FrmConfig.ck2.Value 'heuristic
SetDwordValue &H80000001, "Software\Wan'iez\", "Registry", FrmConfig.ck3.Value 'registry
SetDwordValue &H80000001, "Software\Wan'iez\", "DetecHidden", FrmConfig.ck4.Value 'detekhiden
SetDwordValue &H80000001, "Software\Wan'iez\", "Information", FrmConfig.ck5.Value 'information
SetDwordValue &H80000001, "Software\Wan'iez\", "Startup", FrmConfig.ck7.Value 'starup
SetDwordValue &H80000001, "Software\Wan'iez\", "Update", FrmConfig.ck9.Value 'update
SetDwordValue &H80000001, "Software\Wan'iez\", "EnablePassword", FrmConfig.CkPass.Value 'pasword

   'setting Password Protection
   SetDwordValue &H80000001, "Software\Wan'iez\", "RtpSet", FrmConfig.CkPotection.Value 'Realtime Protection Setting
   SetDwordValue &H80000001, "Software\Wan'iez\", "QuarSet", FrmConfig.CkQuar.Value 'Quarantine Pass
   SetDwordValue &H80000001, "Software\Wan'iez\", "AddProSet", FrmConfig.CkAdditional.Value 'Additional Pass
   SetDwordValue &H80000001, "Software\Wan'iez\", "UpdtSet", FrmConfig.CkUpdate.Value 'Update Pass
   SetDwordValue &H80000001, "Software\Wan'iez\", "DelUSet", FrmConfig.CkDellUs.Value 'Dell Virus By User
   SetDwordValue &H80000001, "Software\Wan'iez\", "ExitPassword", FrmConfig.CkExit.Value 'Exit Allpication

Call loadPassword: Call LoadConfig: Call loadLNK: Call loadSTARTUP: Call loadCONTEKMENU: Call LoadConfig: Call loadRTP
  FrmConfig.Lbapply.Caption = "Configuration Apply"
  LockWindowUpdate 0
End Function
Public Function LoadConfig()

On Error Resume Next

With FrmConfig

.ck1.Value = GetDWORDValue(&H80000001, "Software\Wan'iez\", "Filter")
.ck2.Value = GetDWORDValue(&H80000001, "Software\Wan'iez\", "Heuristic")
.ck3.Value = GetDWORDValue(&H80000001, "Software\Wan'iez\", "Registry")
.ck4.Value = GetDWORDValue(&H80000001, "Software\Wan'iez\", "DetecHidden")
.ck5.Value = GetDWORDValue(&H80000001, "Software\Wan'iez\", "Information")

  .CkPotection.Value = GetDWORDValue(&H80000001, "Software\Wan'iez\", "RtpSet") 'Realtime Protection Setting
  .CkQuar.Value = GetDWORDValue(&H80000001, "Software\Wan'iez\", "QuarSet") 'Quarantine Pass
  .CkAdditional.Value = GetDWORDValue(&H80000001, "Software\Wan'iez\", "AddProSet") 'Additional Pass
  .CkUpdate.Value = GetDWORDValue(&H80000001, "Software\Wan'iez\", "UpdtSet") 'Update Pass
  .CkDellUs.Value = GetDWORDValue(&H80000001, "Software\Wan'iez\", "DelUSet")  'Delet virus by user
  .CkExit.Value = GetDWORDValue(&H80000001, "Software\Wan'iez\", "ExitPassword")  'Exit app
If .Check12.Value = 1 Then
Install_CMenuDB ("Add To Wan'iez user database")
Else
UnInstall_CMenuDB ("Add To Wa'iez user database")
End If

.ck9.Value = GetDWORDValue(&H80000001, "Software\Wan'iez\", "Update") ' Online Update

.ck10.Value = GetDWORDValue(&H80000001, "Software\Wan'iez\", "ScanFD") 'USB Detect
If .ck10.Value = 1 Then
   Call GetLasFDVolume
   frmMain.tmFlash.Enabled = True
Else
   frmMain.tmFlash.Enabled = False
End If

.ck11.Value = GetDWORDValue(&H80000001, "Software\Wan'iez\", "Splash") ' Form Ontop

LangUsed = GetSTRINGValue(&H80000001, "Software\Wan'iez\", "Language")
.lblLangSel1.Caption = ": " & GetSTRINGValue(&H80000001, "Software\Wan'iez\", "Language")
.lblLangAut1.Caption = GetSTRINGValue(&H80000001, "Software\Wan'iez\", "LangAUTOR")
.lblLangID1.Caption = GetSTRINGValue(&H80000001, "Software\Wan'iez\", "LangID")

.lblLangUsed1.Caption = ": " & getNameLangFromFile(GetFilePath(App_FullPathW(False)) & "\lang\" & LangUsed)

End With
End Function
