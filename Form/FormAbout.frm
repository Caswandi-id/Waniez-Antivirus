VERSION 5.00
Begin VB.Form FrmAbout 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Wan'iez Antivirus - About"
   ClientHeight    =   4260
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6240
   DrawMode        =   14  'Copy Pen
   Icon            =   "FormAbout.frx":0000
   LinkTopic       =   "Wan'iez Antivirus - About"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   284
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   416
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.PictureBox PicAbout 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2295
      Index           =   0
      Left            =   120
      ScaleHeight     =   2295
      ScaleWidth      =   6015
      TabIndex        =   4
      Top             =   1440
      Width           =   6015
      Begin VB.Frame frSoftInformation 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   2175
         Left            =   120
         TabIndex        =   5
         Top             =   120
         Width           =   5775
         Begin VB.Label lbMesin 
            BackStyle       =   0  'Transparent
            Caption         =   ": 32 bit - Windows XP | Vista | 7 | 8"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1440
            TabIndex        =   18
            Top             =   1440
            Width           =   3975
         End
         Begin VB.Label lbVirus 
            BackStyle       =   0  'Transparent
            Caption         =   ": 0037 + Heuristic"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1440
            TabIndex        =   17
            Top             =   1200
            Width           =   3375
         End
         Begin VB.Label lbWorm 
            BackStyle       =   0  'Transparent
            Caption         =   ": -"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1440
            TabIndex        =   16
            Top             =   1680
            Visible         =   0   'False
            Width           =   3735
         End
         Begin VB.Label lbRegDataBase 
            BackStyle       =   0  'Transparent
            Caption         =   ": 106 value(s)"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1440
            TabIndex        =   15
            Top             =   960
            Width           =   3975
         End
         Begin VB.Label lbBuildDate 
            BackStyle       =   0  'Transparent
            Caption         =   ": 12 September 2013"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1440
            TabIndex        =   14
            Top             =   720
            Width           =   3735
         End
         Begin VB.Label lbBuildNumber 
            BackStyle       =   0  'Transparent
            Caption         =   ": 06"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1440
            TabIndex        =   13
            Top             =   480
            Width           =   3975
         End
         Begin VB.Label lbEngine 
            BackStyle       =   0  'Transparent
            Caption         =   ": 1.5.0.06"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1440
            TabIndex        =   12
            Top             =   240
            Width           =   3855
         End
         Begin VB.Label lbInfo1 
            BackStyle       =   0  'Transparent
            Caption         =   "Machine"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   5
            Left            =   120
            TabIndex        =   11
            Top             =   1440
            Width           =   1215
         End
         Begin VB.Label lbInfo1 
            BackStyle       =   0  'Transparent
            Caption         =   "Virus Signature"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   4
            Left            =   120
            TabIndex        =   10
            Top             =   1200
            Width           =   1215
         End
         Begin VB.Label lbInfo1 
            BackStyle       =   0  'Transparent
            Caption         =   "Reg Database"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   9
            Top             =   960
            Width           =   1215
         End
         Begin VB.Label lbInfo1 
            BackStyle       =   0  'Transparent
            Caption         =   "Build Date"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   8
            Top             =   720
            Width           =   1215
         End
         Begin VB.Label lbInfo1 
            BackStyle       =   0  'Transparent
            Caption         =   "Build Number"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   7
            Top             =   480
            Width           =   1215
         End
         Begin VB.Label lbInfo1 
            BackStyle       =   0  'Transparent
            Caption         =   "Engine Version "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   6
            Top             =   240
            Width           =   1215
         End
      End
   End
   Begin VB.PictureBox PicAbout 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2295
      Index           =   1
      Left            =   120
      ScaleHeight     =   2295
      ScaleWidth      =   6015
      TabIndex        =   1
      Top             =   1440
      Width           =   6015
      Begin VB.ListBox List 
         Appearance      =   0  'Flat
         Height          =   1590
         ItemData        =   "FormAbout.frx":52C2
         Left            =   0
         List            =   "FormAbout.frx":52F9
         TabIndex        =   2
         Top             =   480
         Width           =   6015
      End
      Begin VB.Label LbThanks 
         BackStyle       =   0  'Transparent
         Caption         =   "Thanks To"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   120
         Width           =   3735
      End
   End
   Begin WANIEZ.Tab TabInformation 
      Height          =   2775
      Left            =   0
      TabIndex        =   0
      Top             =   1080
      Width           =   6255
      _extentx        =   11033
      _extenty        =   4895
      backcolor       =   16777215
      closebutton     =   0   'False
      blurforecolor   =   0
      activeforecolor =   0
      picture         =   "FormAbout.frx":540F
      alltabsforecolor=   -2147483630
      fontname        =   "Tahoma"
   End
   Begin VB.Label Copyright 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright © 2011 - 2013 Canvas Software."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   210
      Left            =   120
      TabIndex        =   21
      Top             =   3960
      Width           =   3405
   End
   Begin VB.Label Title 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Wan'iez Antivirus"
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
      TabIndex        =   20
      Top             =   240
      Width           =   2640
   End
   Begin VB.Label LbInfoAv 
      BackStyle       =   0  'Transparent
      Caption         =   "Information about your Wan'iez security application"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   19
      Top             =   600
      Width           =   4935
   End
End
Attribute VB_Name = "FrmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'menggerakan form tanpa border
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Long) As Long
Private Sub Form_Load()
    TabInformation.AddTabs "System Information", "Credits": TabInformation.ActiveTab = 1: TabInformation_Click (1)
    Call LetakanForm(Me, True) 'frmdi depan
    FrmAbout.Picture = LoadPictureDLL(600) 'load image
End Sub
Private Sub Form_Unload(Cancel As Integer)
Cancel = 1
Me.Hide
End Sub

Private Sub TabInformation_Click(tIndex As Integer)
    On Error Resume Next
    PicAbout(tIndex - 1).ZOrder 0
End Sub
