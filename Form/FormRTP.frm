VERSION 5.00
Begin VB.Form FrmRTP 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Wan'iez Antivirus Real Time Protection"
   ClientHeight    =   4230
   ClientLeft      =   6630
   ClientTop       =   3840
   ClientWidth     =   8325
   Icon            =   "FormRTP.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   282
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   555
   ShowInTaskbar   =   0   'False
   Begin WANIEZ.rButton cmdIgnore 
      Height          =   375
      Left            =   4440
      TabIndex        =   12
      Top             =   3720
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      ButtonStyle     =   7
      BackColor       =   14211288
      BackColorPressed=   15715986
      BackColorHover  =   16243621
      BorderColor     =   9408398
      BorderColorPressed=   6045981
      BorderColorHover=   11632444
      Caption         =   "Ignore"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin WANIEZ.rButton mnQuar 
      Height          =   375
      Left            =   3000
      TabIndex        =   11
      Top             =   3720
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      ButtonStyle     =   7
      BackColor       =   14211288
      BackColorPressed=   15715986
      BackColorHover  =   16243621
      BorderColor     =   9408398
      BorderColorPressed=   6045981
      BorderColorHover=   11632444
      Caption         =   "Quarantine all"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin WANIEZ.rButton cmdFixRtp2 
      Height          =   375
      Left            =   1680
      TabIndex        =   10
      Top             =   3720
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      ButtonStyle     =   7
      BackColor       =   14211288
      BackColorPressed=   15715986
      BackColorHover  =   16243621
      BorderColor     =   9408398
      BorderColorPressed=   6045981
      BorderColorHover=   11632444
      Caption         =   "Fix Selected"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin WANIEZ.rButton cmdFixAllRtp 
      Height          =   375
      Left            =   360
      TabIndex        =   9
      Top             =   3720
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      ButtonStyle     =   7
      BackColor       =   14211288
      BackColorPressed=   15715986
      BackColorHover  =   16243621
      BorderColor     =   9408398
      BorderColorPressed=   6045981
      BorderColorHover=   11632444
      Caption         =   "Fix All"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox TxtRamnit 
      BackColor       =   &H00004000&
      ForeColor       =   &H00004000&
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Text            =   "<SCRIPT Language=VBScript><!--"
      Top             =   6120
      Width           =   6735
   End
   Begin VB.TextBox Text3 
      Height          =   315
      Left            =   120
      TabIndex        =   6
      Top             =   4440
      Width           =   6735
   End
   Begin VB.PictureBox picIconInfo 
      BackColor       =   &H000000FF&
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   2040
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   5
      Top             =   5880
      Width           =   255
   End
   Begin VB.PictureBox picInfoHidden 
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2280
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   4
      Top             =   5880
      Width           =   255
   End
   Begin VB.CheckBox CkAll 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   "Check All"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   7080
      TabIndex        =   3
      Top             =   3240
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   5400
      Width           =   6735
   End
   Begin VB.TextBox txtpath 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   4800
      Width           =   6735
   End
   Begin WANIEZ.ucListView lvRTP 
      Height          =   2055
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   8055
      _extentx        =   14843
      _extenty        =   3836
      styleex         =   33
      showsort        =   -1  'True
   End
   Begin VB.Label TitleRTP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Realtime Wan'iez Antivirus"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   360
      Left            =   240
      TabIndex        =   15
      Top             =   120
      Width           =   4035
   End
   Begin VB.Label LbDetec 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   240
      TabIndex        =   14
      Top             =   480
      Width           =   60
   End
   Begin VB.Label LbX 
      Caption         =   "X"
      Height          =   255
      Left            =   9120
      TabIndex        =   13
      Top             =   2040
      Width           =   255
   End
   Begin VB.Label lbFileCheck 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   240
      TabIndex        =   8
      Top             =   2760
      Width           =   45
   End
End
Attribute VB_Name = "FrmRTP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public DiIgnore As Boolean
Dim LastPathRClick    As String
Dim sPath As String
Dim IndekPluginTerplih As Long
Private WithEvents Thread As IThread
Attribute Thread.VB_VarHelpID = -1
Private WithEvents Thread2 As IThread
Attribute Thread2.VB_VarHelpID = -1

Private ctile As New cDIBTile
Private Declare Function Beep Lib "kernel32" (ByVal dwFreq As Long, _
ByVal dwDuration As Long) As Long
Dim second As Integer
Private Sub CkAll_Click()
Dim i As Long
If lvRTP.ListItems.Count <> 0 Then
        If CkAll.Value = Checked Then
            For i = 1 To lvRTP.ListItems.Count
                lvRTP.ListItems.Item(i).Checked = True
            Next
        Else
            For i = 1 To lvRTP.ListItems.Count
                lvRTP.ListItems.Item(i).Checked = False
            Next
        End If
    End If
End Sub
Private Sub cmdFixAllRtp_Click()
cmdIgnore.Enabled = False: cmdFixRtp2.Enabled = False: cmdFixAllRtp.Enabled = False
    Call FiX_Malware(lvRTP, BY_ALL, 10)
cmdIgnore.Enabled = True: cmdFixRtp2.Enabled = True: cmdFixAllRtp.Enabled = True
      AutoLst lvRTP
End Sub
Private Sub Form_Unload(Cancel As Integer)
Cancel = 1
Call LepasSemuaKunci
    Me.Hide
    lvRTP.ListItems.Clear
 LbDetec.Caption = ""
End Sub
Private Sub mnFixall_Click()
cmdIgnore.Enabled = False: cmdFixRtp2.Enabled = False: cmdFixAllRtp.Enabled = False
   Call FiX_Malware(lvRTP, BY_ALL, 10)
cmdIgnore.Enabled = True: cmdFixRtp2.Enabled = True: cmdFixAllRtp.Enabled = True
      AutoLst lvRTP
End Sub
Private Sub cmdFixRtp2_Click()
cmdIgnore.Enabled = False: cmdFixRtp2.Enabled = False: cmdFixAllRtp.Enabled = False
   Call FiX_Malware(lvRTP, BY_SELECT, 10)
cmdIgnore.Enabled = True: cmdFixRtp2.Enabled = True: cmdFixAllRtp.Enabled = True
      AutoLst lvRTP
End Sub

Private Sub cmdIgnore_Click()
If FrmSysTray.rtpkerja = True Then GoTo metE
   Call LepasSemuaKunci
    Me.Hide
    lvRTP.ListItems.Clear: FrmRTP.Caption = "Wan'iez Real Time Protection": LbDetec.Caption = ""
metE:
End Sub
Private Sub mnFixS_Click()
cmdIgnore.Enabled = False: cmdFixRtp2.Enabled = False: cmdFixAllRtp.Enabled = False
    Call FiX_Malware(lvRTP, BY_CHECKED, 10)
cmdIgnore.Enabled = True: cmdFixRtp2.Enabled = True: cmdFixAllRtp.Enabled = True
      AutoLst lvRTP
End Sub

Private Sub LbX_Click()
If FrmSysTray.rtpkerja = True Then GoTo metE
   Call LepasSemuaKunci
    Me.Hide
    lvRTP.ListItems.Clear: FrmRTP.Caption = "Wan'iez Real Time Protection": LbDetec.Caption = ""
metE:
End Sub
Private Sub mnQuar_Click()
cmdIgnore.Enabled = False: cmdFixRtp2.Enabled = False: cmdFixAllRtp.Enabled = False
  Call Quar_Malware(lvRTP, BY_ALL, 10)
cmdIgnore.Enabled = True: cmdFixRtp2.Enabled = True: cmdFixAllRtp.Enabled = True
   AutoLst lvRTP
End Sub
Private Sub mnExcL_Click()
Dim nIndek As Long
'LastPathRClick = lvRTP.Checked
nIndek = CariIndekItemTerpilih(lvRTP)
If ValidFile(LastPathRClick) = True Then
   ReBuildFileException LastPathRClick, GetFilePath(App_FullPathW(False)) & "\File.lst", frmMain.lstExceptFile
   If nIndek > 0 Then lvRTP.ListItems.Remove nIndek
Else ' berarti exception untuk registry
End If
End Sub
Private Function CariIndekItemTerpilih(LvInput As ucListView) 'base1
Dim CNTA As Long
On Error Resume Next
For CNTA = 1 To LvInput.ListItems.Count
    If LvInput.ListItems.Item(CNTA).Selected = True Then
       CariIndekItemTerpilih = CNTA
       Exit For
    End If
Next
End Function
Private Sub Form_Initialize()
Me.Left = (Screen.Width - Me.Width - 200) / 1:  Me.Top = (Screen.Height - Me.Height - 500) / 1
End Sub
Private Sub Form_Load()
Call LetakanForm(Me, True)
 FrmRTP.Picture = LoadPictureDLL(500) 'load image
End Sub
Private Sub lvRTP_ColumnClick(ByVal oColumn As cColumn)
    oColumn.Sort
End Sub

Private Sub Title_Click()

End Sub
