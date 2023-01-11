VERSION 5.00
Begin VB.Form FrmSysTray 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   1335
   ClientLeft      =   150
   ClientTop       =   840
   ClientWidth     =   2865
   Icon            =   "FormSysTray.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1335
   ScaleWidth      =   2865
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.Timer Timer1 
      Left            =   960
      Top             =   120
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   0
      TabIndex        =   3
      Top             =   360
      Width           =   855
   End
   Begin VB.TextBox txtpath 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   0
      TabIndex        =   2
      Top             =   600
      Width           =   855
   End
   Begin VB.ListBox lstStatus 
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   120
      Width           =   855
   End
   Begin VB.PictureBox Picture1 
      Height          =   375
      Left            =   1440
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   0
      Top             =   120
      Width           =   375
   End
   Begin VB.Timer TimRTP 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   960
      Top             =   480
   End
   Begin VB.Image Image1 
      Height          =   255
      Left            =   1440
      Top             =   960
      Width           =   375
   End
   Begin VB.Image img4 
      Height          =   195
      Left            =   1920
      Picture         =   "FormSysTray.frx":52C2
      Top             =   1080
      Width           =   195
   End
   Begin VB.Image img3 
      Height          =   195
      Left            =   1920
      Picture         =   "FormSysTray.frx":550C
      Top             =   840
      Width           =   195
   End
   Begin VB.Image img2 
      Height          =   195
      Left            =   1920
      Picture         =   "FormSysTray.frx":5758
      Top             =   600
      Width           =   195
   End
   Begin VB.Image img1 
      Height          =   195
      Left            =   1920
      Picture         =   "FormSysTray.frx":59A4
      Top             =   360
      Width           =   195
   End
   Begin VB.Image img0 
      Height          =   195
      Left            =   1920
      Picture         =   "FormSysTray.frx":5BF0
      Top             =   120
      Width           =   195
   End
   Begin VB.Image picmti 
      Height          =   345
      Left            =   1440
      Top             =   480
      Width           =   345
   End
   Begin VB.Menu mnSystray 
      Caption         =   "Systray"
      Begin VB.Menu mnCScan 
         Caption         =   "Hide Scanner coy"
      End
      Begin VB.Menu bts1 
         Caption         =   "-"
      End
      Begin VB.Menu mnEPro2 
         Caption         =   "Real Time Protection"
         Begin VB.Menu mnEPro 
            Caption         =   "Enable Protection"
         End
         Begin VB.Menu bts2 
            Caption         =   "-"
         End
         Begin VB.Menu mndisable 
            Caption         =   "Disable for 1 minutes"
         End
         Begin VB.Menu mndisable2 
            Caption         =   "Disable for 10 minutes"
         End
      End
      Begin VB.Menu bts3 
         Caption         =   "-"
      End
      Begin VB.Menu mnRun 
         Caption         =   "Run On Startup"
      End
      Begin VB.Menu bts4 
         Caption         =   "-"
      End
      Begin VB.Menu mnseting 
         Caption         =   "Settings"
      End
      Begin VB.Menu mnAbout 
         Caption         =   "About"
      End
      Begin VB.Menu mnExit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnProses 
      Caption         =   "Proses"
      Begin VB.Menu mnRefresh 
         Caption         =   "Refresh Processes"
      End
      Begin VB.Menu mnKillPro 
         Caption         =   "Kill Process"
      End
      Begin VB.Menu mnRestartPro 
         Caption         =   "Restart Process"
      End
      Begin VB.Menu mnPausePro 
         Caption         =   "Pause Process"
      End
      Begin VB.Menu mnResumePro 
         Caption         =   "Resume Process"
      End
      Begin VB.Menu btx 
         Caption         =   "-"
      End
      Begin VB.Menu mnProProperties 
         Caption         =   "Properties"
      End
   End
   Begin VB.Menu mnjail 
      Caption         =   "jail"
      Begin VB.Menu mncleanjail 
         Caption         =   "Clear Jail"
      End
      Begin VB.Menu mnkiljail 
         Caption         =   "Kill Prisoner"
      End
      Begin VB.Menu btc 
         Caption         =   "-"
      End
      Begin VB.Menu mnsubmit 
         Caption         =   "Submit For Analysis"
      End
   End
   Begin VB.Menu mnDriveLock 
      Caption         =   "Drive Lock"
      Begin VB.Menu mnLock 
         Caption         =   "Dirve Lock"
      End
      Begin VB.Menu mnunLock 
         Caption         =   "Drive Un Lock"
      End
   End
End
Attribute VB_Name = "FrmSysTray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Icom Menu Editor :
Private Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Private Declare Function SetMenuItemBitmaps Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal hBitmapUnchecked As Long, ByVal hBitmapChecked As Long) As Long
Private Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long        ':( Missing Scope
Private Const MF_BYPOSITION = &H400&
Private mHandle As Long
Private lRet As Long
Private sHandle As Long
'
Dim second As Integer
Dim SudahJalan As Boolean
Public rtpkerja As Boolean
Dim Counter As Integer
'Paswword
Dim A As String: Dim B As String: Dim spWd As String
Dim encrypt As New clsEncryption 'Enkrip Passwoed
Private Sub PaintMenuBitmaps()  'function to set bitmaps to menus
    On Error Resume Next
    AssignMenuBitmaps Me, img0, 0, 0: AssignMenuBitmaps Me, img1, 0, 2: AssignMenuBitmaps Me, img2, 0, 6: AssignMenuBitmaps Me, img3, 0, 7: AssignMenuBitmaps Me, img4, 0, 8
    'Open,RTP,SEtting,help,exit
End Sub
Private Sub Form_Load()
Counter = 0
  Timer1.Interval = 40  'Atur kecepatannya di sini
  PaintMenuBitmaps    'calls the function to set menu bitmaps
  Call LoadDll
End Sub
Private Sub AssignMenuBitmaps(ByRef frm As Form, ByRef IMG As Image, ByVal Menu_Position As Integer, ByVal Sub_Menu_Position As Integer)
   mHandle = GetMenu(frm.hwnd)
   sHandle = GetSubMenu(mHandle, Menu_Position)
   lRet = SetMenuItemBitmaps(sHandle, Sub_Menu_Position, MF_BYPOSITION, IMG.Picture, IMG.Picture)
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
Dim lHasil  As Long
Dim HorX    As Long

    If Me.ScaleMode = vbPixels Then
        HorX = X
    Else
        HorX = X / Screen.TwipsPerPixelX
    End If
    
    Select Case HorX
        Case WM_LBUTTONDBLCLK
        'Me.Caption = ""
            lHasil = SetForegroundWindow(Me.hwnd)
            mnCScan.Caption = g_bahasa(0)
            frmMain.Show: frmMain.WindowState = vbNormal
        Case WM_RBUTTONUP 'Tampilkan menu Popup saat klik kanan.
            lHasil = SetForegroundWindow(Me.hwnd)
            Me.PopupMenu Me.mnSystray, 0, , , mnCScan
    End Select

End Sub
Private Sub Form_Unload(Cancel As Integer)
Cancel = 1
    Me.Hide
End Sub
Public Sub mnCScan_Click() 'mulai menu
If mnCScan.Caption = g_bahasa(1) Then
    frmMain.Show: frmMain.WindowState = vbNormal
    mnCScan.Caption = g_bahasa(0)
Else
    mnCScan.Caption = g_bahasa(1)
    frmMain.Hide: frmMain.WindowState = vbMinimized
End If
Call LoadDll
End Sub
Private Sub mncleanjail_Click()
If MsgBox(i_bahasa(19), vbExclamation + vbYesNo) = vbYes Then
    ClearJail frmMain.lvJail
End If
End Sub
Private Sub mnkiljail_Click()
Dim PrisName        As String
Dim IndekTerpilih   As Long
Dim Counter         As Long

If MsgBox(i_bahasa(21) & " ?", vbExclamation + vbYesNo) = vbYes Then
   For Counter = 1 To frmMain.lvJail.ListItems.Count
       If frmMain.lvJail.ListItems.Item(Counter).Selected = False Then GoTo LBL_LANJUT
       PrisName = frmMain.lvJail.ListItems.Item(Counter).SubItem(4).Text
       KillPrisonner PrisName, frmMain.lvJail
LBL_LANJUT:
   Next
Call READ_DATA_JAIL(FolderJail)
End If
End Sub
Private Sub mnKillPro_Click()
Dim nIndek As Long
Dim pId    As Long
Dim sPath  As String

nIndek = CariIndekItemTerpilih(frmMain.lvProses)

If nIndek > 0 Then
   pId = CLng(frmMain.lvProses.ListItems.Item(nIndek).SubItem(3).Text)
   sPath = frmMain.lvProses.ListItems.Item(nIndek).SubItem(9).Text
   If KillProses(pId, sPath, False, True) = True Then
      MsgBox i_bahasa(10) & ": [" & pId & "] " & i_bahasa(11)
      Call ENUM_PROSES(frmMain.lvProses, frmMain.picBuffer)
   Else
      MsgBox i_bahasa(10) & ": [" & pId & "] " & i_bahasa(12)
   End If
End If
End Sub

Private Sub mnLock_Click()
Dim DriveLet As String
    DriveLet = Left$(Right$(frmMain.lstKunci.ListItems.Item(GetSelect(frmMain.lstKunci)).Text, 3), 1)
    KunciFD DriveLet
    Call LoadKunci
End Sub
Public Sub AutoLst(ListApa As ucListView)
    Dim i      As Integer, ClmAuto As eListViewColumnAutoSize
    If ListApa.ListItems.Count = 0 Then ClmAuto = lvwColumnSizeToColumnText Else ClmAuto = lvwColumnSizeToItemText
    For i = 1 To ListApa.Columns.Count
        ListApa.Columns.Item(i).AutoSize ClmAuto
    Next
End Sub
Private Sub LoadKunci()
    Dim DriveKu As Collection
    Dim CekKunci As Boolean, i As Integer, LabelKu As String, Terkunci As Boolean
    DriveList "02346", DriveKu:
    frmMain.lstKunci.ListItems.Clear
    For i = 1 To DriveKu.Count
        LabelKu = DriveLabel(DriveKu.Item(i))
        Terkunci = PathIsDirectory(StrPtr(DriveKu.Item(i) & ":\autorun.inf\con\aux\nul\"))
        AddList frmMain.lstKunci, DriveKu(i) & ":\", frmMain.picBuffer, 4, LabelKu & " (" & DriveKu.Item(i) & ":)", _
        Array(IIf(Terkunci, g_bahasa(19), g_bahasa(20)))
    Next
   AutoLst frmMain.lstKunci
End Sub
Private Sub mnPausePro_Click()
Dim nIndek As Long
Dim pId    As Long

nIndek = CariIndekItemTerpilih(frmMain.lvProses)
If nIndek > 0 Then
   pId = CLng(frmMain.lvProses.ListItems.Item(nIndek).SubItem(3).Text)
   frmMain.lvProses.ListItems.Item(nIndek).SubItem(5).Text = SuspendProses(pId, True)
End If
End Sub
Private Sub mnProProperties_Click()
Dim sPath  As String
Dim nIndek As Long

nIndek = CariIndekItemTerpilih(frmMain.lvProses)
If nIndek > 0 Then
   sPath = frmMain.lvProses.ListItems.Item(nIndek).SubItem(9).Text
   If ValidFile(sPath) = True Then ShowProperties sPath, Me.hwnd
End If
End Sub
Private Sub mnRefresh_Click()
    Call ENUM_PROSES(frmMain.lvProses, frmMain.picBuffer) ' Refresh
End Sub
Private Sub mnRestartPro_Click()
Dim nIndek As Long
Dim pId    As Long
Dim sPath  As String

nIndek = CariIndekItemTerpilih(frmMain.lvProses)

If nIndek > 0 Then
   pId = CLng(frmMain.lvProses.ListItems.Item(nIndek).SubItem(3).Text)
   sPath = frmMain.lvProses.ListItems.Item(nIndek).SubItem(9).Text
   If KillProses(pId, sPath, True, False) = True Then
      MsgBox i_bahasa(10) & ": [" & pId & "] " & i_bahasa(13)
      Call ENUM_PROSES(frmMain.lvProses, frmMain.picBuffer)
   Else
      MsgBox i_bahasa(10) & ": [" & pId & "] " & i_bahasa(14)
   End If
End If
End Sub
Private Sub mnResumePro_Click()
Dim nIndek As Long
Dim pId    As Long

nIndek = CariIndekItemTerpilih(frmMain.lvProses)
If nIndek > 0 Then
   pId = CLng(frmMain.lvProses.ListItems.Item(nIndek).SubItem(3).Text)
   frmMain.lvProses.ListItems.Item(nIndek).SubItem(5).Text = SuspendProses(pId, False)
End If
End Sub
Private Sub mnseting_Click()
    FrmConfig.Show: FrmConfig.WindowState = vbNormal: Call LoadDll
End Sub
Private Sub mnsubmit_Click()
ShellExecute Me.hwnd, vbNullString, "mailto:caswandi14@gmail.com", vbNullString, "C:\", 1
End Sub
Private Sub mnunLock_Click()
Dim DriveLet As String
    DriveLet = Left$(Right$(frmMain.lstKunci.ListItems.Item(GetSelect(frmMain.lstKunci)).Text, 3), 1)
    BukaFD DriveLet
    Call LoadKunci
End Sub

Private Sub TimRTP_Timer()
If (second = 0) Then Call mnEPro_Click:  StatusRTP = True: mndisable.Checked = False: mndisable2.Checked = False: TimRTP.Enabled = False
    second = second - 1
    frmMain.RTPdet = second + 1
End Sub
Private Sub mndisable2_Click() 'On/Off RTP 10 minutest / 600 det
If mnEPro.Checked = True Then '

A = GetSTRINGValue(&H80000001, "Software\Wan'iez\", "password")
If A <> "" And FrmConfig.CkPass.Value <> 0 And FrmConfig.CkPotection.Value <> 0 Then 'MsgBox "ada paswor"
B = encrypt.Cryption(A, "WANPASS", False)
FrmPassword.ShowMe spWd
If spWd <> "" Then
If spWd = B Then

   TimRTP.Enabled = False
   second = 600
       mnEPro.Checked = False
       FrmConfig.ck8.Value = 0
   frmMain.PicMainSummary.Picture = LoadPictureDLL(901) 'off
 frmMain.LbstatusRTP(0).ForeColor = &HC0&
 frmMain.LbstatusRTP(0).Caption = "NOT SECURED"
 frmMain.LbstatusRTP(1).Caption = "Your System is not fully protected"
    frmMain.RTPdet.Visible = True
       mndisable.Checked = False: mndisable.Enabled = False: mndisable2.Enabled = False: mndisable2.Checked = True
       TimRTP.Enabled = True
       StatusRTP = False

    SetDwordValue &H80000001, "Software\Wan'iez\", "Rtp", 0
    Call loadRTP 'Environ$("windir") & "\MO.LOG"
    Shell_NotifyIcon NIM_DELETE, nID
    Call frmMain.cekstatus
    frmMain.terapkanIcon     'UpdateIconRTPmati
    
Exit Sub
Else
      MsgBox "password you entered is not corect!", vbCritical, i_bahasa(26)
End If
Exit Sub
End If
Else

   TimRTP.Enabled = False
   second = 600
       mnEPro.Checked = False
       FrmConfig.ck8.Value = 0
   frmMain.PicMainSummary.Picture = LoadPictureDLL(901) 'off
 frmMain.LbstatusRTP(0).ForeColor = &HC0&
 frmMain.LbstatusRTP(0).Caption = "NOT SECURED"
 frmMain.LbstatusRTP(1).Caption = "Your System is not fully protected"
    frmMain.RTPdet.Visible = True
       mndisable.Checked = False: mndisable.Enabled = False: mndisable2.Enabled = False: mndisable2.Checked = True
       TimRTP.Enabled = True
       StatusRTP = False

    SetDwordValue &H80000001, "Software\Wan'iez\", "Rtp", 0
    Call loadRTP 'Environ$("windir") & "\MO.LOG"
    Shell_NotifyIcon NIM_DELETE, nID
 
    Call frmMain.cekstatus
    frmMain.terapkanIcon     'UpdateIconRTPmati
    
End If
Else '

 second = 600
 TimRTP.Enabled = True
 
       mnEPro.Checked = False
       FrmConfig.ck8.Value = 0
       TimRTP.Enabled = False
   frmMain.PicMainSummary.Picture = LoadPictureDLL(900) 'on
 frmMain.LbstatusRTP(0).ForeColor = &HC000&
frmMain.LbstatusRTP(0).Caption = "SECURED"
frmMain.LbstatusRTP(1).Caption = "Your System is fully protected"
    frmMain.RTPdet.Visible = False
       mndisable2.Checked = True: mndisable2.Enabled = False: mndisable.Enabled = False: mndisable.Checked = False
       StatusRTP = False
       
      SetDwordValue &H80000001, "Software\Wan'iez\", "Rtp", 0
      Call loadRTP 'Environ$("windir") & "\MO.LOG"
      Shell_NotifyIcon NIM_DELETE, nID
      
      Call frmMain.cekstatus
    frmMain.terapkanIcon2     'UpdateIconRTPmati
      
End If '
End Sub
Public Sub mnEPro_Click() 'On/Off RTP
If mnEPro.Checked = True Then '

A = GetSTRINGValue(&H80000001, "Software\Wan'iez\", "password")
If A <> "" And FrmConfig.CkPass.Value <> 0 And FrmConfig.CkPotection.Value <> 0 Then 'MsgBox "ada paswor"
B = encrypt.Cryption(A, "WANPASS", False)
FrmPassword.ShowMe spWd
If spWd <> "" Then
If spWd = B Then

    mnEPro.Checked = False
    FrmConfig.ck8.Value = 0
   frmMain.PicMainSummary.Picture = LoadPictureDLL(901) 'off
 frmMain.LbstatusRTP(0).ForeColor = &HC0&
 frmMain.LbstatusRTP(0).Caption = "NOT SECURED"
 frmMain.LbstatusRTP(1).Caption = "Your System is not fully protected"
 frmMain.RTPdet.Visible = True: frmMain.RTPdet.Caption = ""
 frmMain.BtnAtur.Visible = True
     mndisable2.Enabled = False: mndisable.Enabled = False
     StatusRTP = False
     
Exit Sub
Else
      MsgBox "password you entered is not corect!", vbCritical, i_bahasa(26)
End If
Exit Sub
End If
Else
    mnEPro.Checked = False
    FrmConfig.ck8.Value = 0
   frmMain.PicMainSummary.Picture = LoadPictureDLL(901) 'off
 frmMain.LbstatusRTP(0).ForeColor = &HC0&
frmMain.LbstatusRTP(0).Caption = "NOT SECURED"
frmMain.LbstatusRTP(1).Caption = "Your System is not fully protected"
 frmMain.RTPdet.Visible = True: frmMain.RTPdet.Caption = ""
 frmMain.BtnAtur.Visible = True
     mndisable2.Enabled = False: mndisable.Enabled = False
     StatusRTP = False
     
End If
Else '
   mnEPro.Checked = True
   FrmConfig.ck8.Value = 1
   frmMain.PicMainSummary.Picture = LoadPictureDLL(900) 'on
 frmMain.LbstatusRTP(0).ForeColor = &HC000&
frmMain.LbstatusRTP(0).Caption = "SECURED"
frmMain.LbstatusRTP(1).Caption = "Your System is fully protected"
 frmMain.RTPdet.Visible = False: frmMain.RTPdet.Caption = ""
 frmMain.BtnAtur.Visible = False
       mndisable2.Checked = False: mndisable.Checked = False: mndisable2.Enabled = True: mndisable.Enabled = True
       StatusRTP = True
       TimRTP.Enabled = False
 FormSplash.buka_splas: FormSplash.Show 'vbModal
End If '
    SetDwordValue &H80000001, "Software\Wan'iez\", "Rtp", FrmConfig.ck8.Value
    Call loadRTP
    Shell_NotifyIcon NIM_DELETE, nID
      If UCase$(Left$(Command, 2)) <> "-K" Then frmMain.terapkanIcon
    frmMain.cekstatus

End Sub
Private Sub mndisable_Click() 'on/Off RTP 1 Minutes/60 det
If mnEPro.Checked = True Then '

A = GetSTRINGValue(&H80000001, "Software\Wan'iez\", "password")
If A <> "" And FrmConfig.CkPass.Value <> 0 And FrmConfig.CkPotection.Value <> 0 Then 'MsgBox "ada paswor"
B = encrypt.Cryption(A, "WANPASS", False)
FrmPassword.ShowMe spWd
If spWd <> "" Then
If spWd = B Then

  second = 60
  TimRTP.Enabled = False 'dat off
       TimRTP.Enabled = True 'det on
       mnEPro.Checked = False
       FrmConfig.ck8.Value = 0
   frmMain.PicMainSummary.Picture = LoadPictureDLL(901) 'off
  frmMain.LbstatusRTP(0).ForeColor = &HC0&
frmMain.LbstatusRTP(0).Caption = "NOT SECURED"
frmMain.LbstatusRTP(1).Caption = "Your System is not fully protected Disable 1 Minutes"
  frmMain.RTPdet.Visible = True
  frmMain.BtnAtur.Visible = True
       mndisable2.Checked = False: mndisable2.Enabled = False: mndisable.Enabled = False: mndisable.Checked = True
       StatusRTP = False

    SetDwordValue &H80000001, "Software\Wan'iez\", "Rtp", 0
    Call loadRTP 'Environ$("windir") & "\MO.LOG"
    Shell_NotifyIcon NIM_DELETE, nID
 
      Call frmMain.cekstatus
frmMain.terapkanIcon     'UpdateIconRTPmati
Exit Sub
Else
      MsgBox "password you entered is not corect!", vbCritical, i_bahasa(26)
End If
Exit Sub
End If
Else

  second = 60
  TimRTP.Enabled = False 'dat off
      TimRTP.Enabled = True 'det on
  
       mnEPro.Checked = False
       FrmConfig.ck8.Value = 0
   frmMain.PicMainSummary.Picture = LoadPictureDLL(901) 'off
  frmMain.LbstatusRTP(0).ForeColor = &HC0&
frmMain.LbstatusRTP(0).Caption = "NOT SECURED"
frmMain.LbstatusRTP(1).Caption = "Your System is not fully protected Disable 1 Minutes"
  frmMain.RTPdet.Visible = True
  frmMain.BtnAtur.Visible = True
       mndisable2.Checked = False: mndisable2.Enabled = False: mndisable.Enabled = False: mndisable.Checked = True
       StatusRTP = False

    SetDwordValue &H80000001, "Software\Wan'iez\", "Rtp", 0
    Call loadRTP 'Environ$("windir") & "\MO.LOG"
    Shell_NotifyIcon NIM_DELETE, nID
 
      Call frmMain.cekstatus
      frmMain.terapkanIcon     'UpdateIconRTPmati

End If
Else '
 second = 60
 TimRTP.Enabled = True
      
       mnEPro.Checked = False
       TimRTP.Enabled = False
   frmMain.PicMainSummary.Picture = LoadPictureDLL(900) 'on
 frmMain.LbstatusRTP(1).ForeColor = &HC000&
 frmMain.LbstatusRTP(0).Caption = "SECURED"
 frmMain.LbstatusRTP(1).Caption = "Your System is fully protected"
 frmMain.RTPdet.Visible = False
 frmMain.BtnAtur.Visible = False
 frmMain.terapkanIcon2      'UpdateIconRTPhidup
       mndisable2.Checked = False: mndisable2.Enabled = False: mndisable.Enabled = False: mndisable.Checked = True
       FrmConfig.ck8.Value = 0
       StatusRTP = False
       
    SetDwordValue &H80000001, "Software\Wan'iez\", "Rtp", 0
    Call loadRTP 'Environ$("windir") & "\MO.LOG"
    Shell_NotifyIcon NIM_DELETE, nID
    
    Call frmMain.cekstatus
End If '
End Sub
Public Sub mnRun_Click()
    If mnRun.Checked = True Then
       mnRun.Checked = False: FrmConfig.ck7.Value = 0
    Else
       mnRun.Checked = True: FrmConfig.ck7.Value = 1
    End If
   SetDwordValue &H80000001, "Software\Wan'iez\", "Startup", FrmConfig.ck7.Value
    Call loadSTARTUP: Call frmMain.cekstatus
End Sub
Private Sub mnAbout_Click()
FrmAbout.Show: FrmAbout.WindowState = vbNormal: Call LoadDll
End Sub
Private Sub mnExit_Click()
A = GetSTRINGValue(&H80000001, "Software\Wan'iez\", "password")
If A <> "" And FrmConfig.CkPass.Value <> 0 And FrmConfig.CkExit.Value <> 0 Then 'MsgBox "ada paswor"
B = encrypt.Cryption(A, "WANPASS", False)
FrmPassword.ShowMe spWd
If spWd <> "" Then
If spWd = B Then
    Shell_NotifyIcon NIM_DELETE, nID
    frmMain.mnkeluar
Exit Sub
Else
      MsgBox "password you entered is not corect!", vbCritical, i_bahasa(26)
End If
Exit Sub
End If
Else
    Shell_NotifyIcon NIM_DELETE, nID
    frmMain.mnkeluar
End If
End Sub

