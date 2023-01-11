VERSION 5.00
Begin VB.Form FrmConfig 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Wan'iez Antivirus | Setting"
   ClientHeight    =   6240
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8640
   Icon            =   "FormConfig.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   416
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   576
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox frm 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      FillColor       =   &H00808080&
      ForeColor       =   &H80000008&
      Height          =   4215
      Index           =   0
      Left            =   120
      ScaleHeight     =   4215
      ScaleWidth      =   8415
      TabIndex        =   36
      Top             =   1320
      Width           =   8415
      Begin VB.CheckBox CkUpdate 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Update"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   47
         Top             =   3720
         Width           =   3015
      End
      Begin VB.CheckBox CkAdditional 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Addtional Protection"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   46
         Top             =   3000
         Width           =   2895
      End
      Begin VB.CheckBox CkPotection 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Real Time Protection"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3600
         TabIndex        =   45
         Top             =   3000
         Width           =   2895
      End
      Begin VB.CheckBox CkQuar 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Virus Chest"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   44
         Top             =   3360
         Width           =   3015
      End
      Begin VB.CheckBox CkDellUs 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Delete Virus By User"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3600
         TabIndex        =   43
         Top             =   3360
         Width           =   2895
      End
      Begin VB.CheckBox CkExit 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Exit Application"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3600
         TabIndex        =   42
         Top             =   3720
         Width           =   2895
      End
      Begin VB.TextBox TextConfPass 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   2520
         PasswordChar    =   "*"
         TabIndex        =   41
         Top             =   2160
         Width           =   2535
      End
      Begin VB.TextBox TextNewPass 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   2520
         PasswordChar    =   "*"
         TabIndex        =   40
         Top             =   1800
         Width           =   2535
      End
      Begin VB.TextBox TextOlPass 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2520
         TabIndex        =   39
         ToolTipText     =   "Change password and enter "
         Top             =   1440
         Width           =   2535
      End
      Begin VB.CheckBox CkPass 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Enable Password Protection"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   38
         Top             =   960
         Width           =   7815
      End
      Begin VB.CheckBox ck7 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Launch AntiVirus at computer startup"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   37
         Top             =   120
         Width           =   7815
      End
      Begin WANIEZ.rButton CmdEdit 
         Height          =   375
         Left            =   5400
         TabIndex        =   48
         Top             =   1800
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
         Caption         =   "Edit"
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
      Begin WANIEZ.rButton CmdSave 
         Height          =   375
         Left            =   5400
         TabIndex        =   49
         Top             =   1440
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
         Caption         =   "Save"
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
      Begin VB.Label LbInfo 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   2520
         TabIndex        =   55
         Top             =   2520
         Width           =   4335
      End
      Begin VB.Label LbGen 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Protect Wan'iez with a Password"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   0
         Left            =   360
         TabIndex        =   54
         Top             =   600
         Width           =   2820
      End
      Begin VB.Label LbGen 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Protected Areas"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   1
         Left            =   480
         TabIndex        =   53
         Top             =   2640
         Width           =   1395
      End
      Begin VB.Label ConfPass 
         BackStyle       =   0  'Transparent
         Caption         =   "&Confirm Password"
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
         Left            =   240
         TabIndex        =   52
         Top             =   2160
         Width           =   1575
      End
      Begin VB.Label NewPass 
         BackStyle       =   0  'Transparent
         Caption         =   "&New Password"
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
         Left            =   240
         TabIndex        =   51
         Top             =   1800
         Width           =   1335
      End
      Begin VB.Label OlPass 
         BackStyle       =   0  'Transparent
         Caption         =   "&Old Password"
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
         Left            =   240
         TabIndex        =   50
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Line LineSetting 
         BorderColor     =   &H00E0E0E0&
         Index           =   0
         X1              =   120
         X2              =   8160
         Y1              =   720
         Y2              =   720
      End
      Begin VB.Line LineSetting 
         BorderColor     =   &H00E0E0E0&
         Index           =   1
         X1              =   120
         X2              =   8160
         Y1              =   2760
         Y2              =   2760
      End
   End
   Begin VB.PictureBox frm 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      FillColor       =   &H00808080&
      ForeColor       =   &H80000008&
      Height          =   4215
      Index           =   1
      Left            =   120
      ScaleHeight     =   4215
      ScaleWidth      =   8415
      TabIndex        =   28
      Top             =   1320
      Width           =   8415
      Begin VB.CheckBox ck1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Enable filter file (by pass file with certain extensions) "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   240
         TabIndex        =   34
         ToolTipText     =   "Make slower but total scanning if unchecked"
         Top             =   600
         Value           =   1  'Checked
         Width           =   7335
      End
      Begin VB.CheckBox ck2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Enable use Heuristic to suspect malware"
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
         Left            =   240
         MouseIcon       =   "FormConfig.frx":52C2
         MousePointer    =   99  'Custom
         TabIndex        =   33
         Top             =   960
         Value           =   1  'Checked
         Width           =   7335
      End
      Begin VB.CheckBox ck4 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Enable detect hidden object (file and folder)"
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
         Left            =   240
         MouseIcon       =   "FormConfig.frx":5414
         MousePointer    =   99  'Custom
         TabIndex        =   32
         ToolTipText     =   "Make speed slower but detect hidden file and folder !"
         Top             =   2040
         Width           =   7335
      End
      Begin VB.CheckBox ck5 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Enable give strange  information while scanning"
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
         Left            =   240
         MouseIcon       =   "FormConfig.frx":5566
         MousePointer    =   99  'Custom
         TabIndex        =   31
         Top             =   1680
         Width           =   7335
      End
      Begin VB.CheckBox ck3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Enable detect useless registry value (XP only)"
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
         Left            =   240
         MouseIcon       =   "FormConfig.frx":56B8
         MousePointer    =   99  'Custom
         TabIndex        =   30
         ToolTipText     =   "Only for XP OS"
         Top             =   1320
         Value           =   1  'Checked
         Width           =   7335
      End
      Begin VB.CheckBox Ck14 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Scan File at (copy and exstract) not recommended"
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
         Left            =   240
         TabIndex        =   29
         Top             =   2400
         Visible         =   0   'False
         Width           =   7335
      End
      Begin VB.Label LbSetScan 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Scanning Setting"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   240
         TabIndex        =   35
         Top             =   120
         Width           =   1440
      End
      Begin VB.Line LineSetting 
         BorderColor     =   &H00E0E0E0&
         Index           =   4
         X1              =   120
         X2              =   8160
         Y1              =   240
         Y2              =   240
      End
   End
   Begin VB.PictureBox frm 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      FillColor       =   &H00808080&
      Height          =   4215
      Index           =   2
      Left            =   120
      ScaleHeight     =   4215
      ScaleWidth      =   8415
      TabIndex        =   17
      Top             =   1320
      Width           =   8415
      Begin VB.CheckBox ck6 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Place Application on Top"
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
         Left            =   240
         TabIndex        =   25
         Top             =   2280
         Visible         =   0   'False
         Width           =   7575
      End
      Begin VB.CheckBox ck15 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Create Shortcut to desktop"
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
         Left            =   240
         MaskColor       =   &H00FFFFFF&
         MouseIcon       =   "FormConfig.frx":580A
         MousePointer    =   99  'Custom
         TabIndex        =   24
         Top             =   1200
         UseMaskColor    =   -1  'True
         Width           =   7575
      End
      Begin VB.CheckBox ck12 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Install Context Menu - Scan with Wan'iez Antivirus"
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
         Left            =   240
         TabIndex        =   23
         Top             =   3240
         Width           =   7575
      End
      Begin VB.CheckBox ck9 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Auto Check Online Update"
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
         Left            =   240
         MouseIcon       =   "FormConfig.frx":595C
         MousePointer    =   99  'Custom
         TabIndex        =   22
         Top             =   1560
         Width           =   7575
      End
      Begin VB.CheckBox ck11 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Splash screen"
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
         Left            =   240
         TabIndex        =   21
         Top             =   1920
         Width           =   7575
      End
      Begin VB.CheckBox ck8 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Enable Wan'iez Real Time Protection"
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
         Left            =   240
         TabIndex        =   20
         Top             =   480
         UseMaskColor    =   -1  'True
         Width           =   7575
      End
      Begin VB.CheckBox ck10 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Auto Scan Flashdisk inserted"
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
         Left            =   240
         TabIndex        =   19
         Top             =   840
         UseMaskColor    =   -1  'True
         Width           =   7575
      End
      Begin VB.CheckBox Check12 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Install Context Menu - Add  Wan'iez user database"
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
         Left            =   240
         TabIndex        =   18
         Top             =   3600
         Visible         =   0   'False
         Width           =   7575
      End
      Begin VB.Label LbAppSet 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Context Menu Setting"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   1
         Left            =   240
         TabIndex        =   27
         Top             =   2760
         Width           =   1815
      End
      Begin VB.Label LbAppSet 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Application Setting"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   0
         Left            =   240
         TabIndex        =   26
         Top             =   120
         Width           =   1590
      End
      Begin VB.Line LineSetting 
         BorderColor     =   &H00E0E0E0&
         Index           =   3
         X1              =   120
         X2              =   8160
         Y1              =   2880
         Y2              =   2880
      End
      Begin VB.Line LineSetting 
         BorderColor     =   &H00E0E0E0&
         Index           =   2
         X1              =   120
         X2              =   8160
         Y1              =   240
         Y2              =   240
      End
   End
   Begin VB.PictureBox frm 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      FillColor       =   &H00808080&
      ForeColor       =   &H80000008&
      Height          =   4215
      Index           =   3
      Left            =   120
      ScaleHeight     =   4215
      ScaleWidth      =   8415
      TabIndex        =   6
      Top             =   1320
      Width           =   8415
      Begin VB.ListBox lstLanguage 
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
         Height          =   1620
         Left            =   120
         TabIndex        =   7
         Top             =   480
         Width           =   8055
      End
      Begin VB.Label lblAvalaibleLang 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Avalaible Language"
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
         Height          =   195
         Left            =   120
         TabIndex        =   16
         Top             =   120
         Width           =   1395
      End
      Begin VB.Label lblLangSel 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Language Selected"
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
         Height          =   195
         Left            =   240
         TabIndex        =   15
         Top             =   2400
         Width           =   1365
      End
      Begin VB.Label lblLangID 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Language ID"
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
         Height          =   195
         Left            =   240
         TabIndex        =   14
         Top             =   2640
         Width           =   915
      End
      Begin VB.Label lblLangAut 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Language Author "
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
         Height          =   195
         Left            =   240
         TabIndex        =   13
         Top             =   2880
         Width           =   1290
      End
      Begin VB.Label lblLangUsed 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Language Used"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   240
         TabIndex        =   12
         Top             =   3840
         Width           =   1290
      End
      Begin VB.Label lblLangUsed1 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   2400
         TabIndex        =   11
         Top             =   3840
         Width           =   4155
      End
      Begin VB.Label lblLangSel1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "x"
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
         Height          =   195
         Left            =   2400
         TabIndex        =   10
         Top             =   2400
         Width           =   90
      End
      Begin VB.Label lblLangID1 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "x"
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
         Height          =   195
         Left            =   2400
         TabIndex        =   9
         Top             =   2640
         Width           =   1515
      End
      Begin VB.Label lblLangAut1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "x"
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
         Height          =   195
         Left            =   2400
         TabIndex        =   8
         Top             =   2880
         Width           =   90
      End
   End
   Begin WANIEZ.Tab TabSetting 
      Height          =   4695
      Left            =   0
      TabIndex        =   4
      Top             =   960
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   8281
      BackColor       =   16777215
      CloseButton     =   0   'False
      BlurForeColor   =   0
      ActiveForeColor =   0
      picture         =   "FormConfig.frx":5AAE
      AllTabsForeColor=   -2147483630
      FontName        =   "Tahoma"
   End
   Begin VB.PictureBox frm 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4215
      Index           =   4
      Left            =   120
      ScaleHeight     =   4215
      ScaleWidth      =   8415
      TabIndex        =   2
      Top             =   1320
      Width           =   8415
   End
   Begin VB.PictureBox frm 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4215
      Index           =   5
      Left            =   120
      ScaleHeight     =   4215
      ScaleWidth      =   8415
      TabIndex        =   1
      Top             =   1320
      Width           =   8415
   End
   Begin VB.PictureBox frm 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4215
      Index           =   6
      Left            =   120
      ScaleHeight     =   4215
      ScaleWidth      =   8415
      TabIndex        =   0
      Top             =   1320
      Width           =   8415
   End
   Begin WANIEZ.rButton cmdok 
      Height          =   375
      Left            =   7320
      TabIndex        =   5
      Top             =   5760
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      ButtonStyle     =   7
      ButtonStyleColors=   3
      BackColor       =   14211288
      BackColorPressed=   15715986
      BackColorHover  =   16243621
      BorderColor     =   9408398
      BorderColorPressed=   6045981
      BorderColorHover=   11632444
      Caption         =   "Apply"
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
   Begin VB.Label TitleMini 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Setting Your Aplication Wan'iez Antivirus"
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
      Left            =   120
      TabIndex        =   57
      Top             =   600
      Width           =   3345
   End
   Begin VB.Label Title 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Wan'iez Antivirus | Setting"
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
      Left            =   120
      TabIndex        =   56
      Top             =   240
      Width           =   3975
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   3720
      Top             =   2880
      Width           =   1215
   End
   Begin VB.Label Lbapply 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "scscscs"
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
      Left            =   5280
      TabIndex        =   3
      Top             =   2880
      Width           =   690
   End
End
Attribute VB_Name = "FrmConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim encrypt As New clsEncryption 'Enkrip Passwoed
Dim A As String
Dim B As String
Dim spWd As String
Private Sub CkPass_Click()
If CkPass.Value = 1 Then '
    Call loadPassWordAktif
    CmdSave.Enabled = True: CmdEdit.Enabled = True
Else '
    Call loadPassWordNonAktif
    CmdSave.Enabled = False: CmdEdit.Enabled = False
    LbInfo.Caption = ""
End If '
End Sub
Private Sub CmdEdit_Click()
TextConfPass.Enabled = True: TextNewPass.Enabled = True: TextOlPass.Enabled = True
TextConfPass.BackColor = &HFFFFFF: TextNewPass.BackColor = &HFFFFFF: TextOlPass.BackColor = &HFFFFFF
TextOlPass.PasswordChar = ""
CmdSave.Enabled = True
End Sub
Private Sub CmdSave_Click()
Call Password
End Sub
Private Sub Form_Load()
Call LetakanForm(Me, True) 'frmdi depan
FrmConfig.Picture = LoadPictureDLL(200) 'load image
Call loadPassword: Call LoadConfig: Call loadLNK: Call loadSTARTUP: Call loadCONTEKMENU: Call LoadConfig: Call loadRTP
TabSetting.AddTabs "General", "Scanning", "Application", "Language": TabSetting.ActiveTab = 1: TabSetting_Click (1)
End Sub
Private Sub Form_Unload(Cancel As Integer)
Cancel = 1
Me.Hide
End Sub
Private Sub TabSetting_Click(tIndex As Integer)
    On Error Resume Next
    frm(tIndex - 1).ZOrder 0
End Sub
Private Sub CmdOk_Click()
Call Save
MsgBox "Setting Application for apply all change", vbInformation
End Sub
Function Password()
 'all loadPassword 'Aktif
On Error Resume Next
    If Check1.Value = Checked Then
        If TextNewPass.Text <> TextConfPass.Text Then
            LbInfo.Caption = "Invalid password combination... please try again..." ', vbInformation  ', App.Title
            TextNewPass.Text = "": TextConfPass.Text = "": TextOlPass.Text = ""
            CmdSave.Enabled = True
            TextNewPass.SetFocus
            Exit Function

        ElseIf Len(TextNewPass.Text) < 6 Then
            LbInfo.Caption = "A valid password must be 6 characters long..." ', vbInformation  ', App.Title
            TextNewPass.Text = "": TextConfPass.Text = "": TextOlPass.Text = ""
            CmdSave.Enabled = True
            TextNewPass.SetFocus
            Exit Function
        Else
            SetStringValue &H80000001, "Software\Wan'iez\", "Password", encrypt.Cryption(TextNewPass.Text, "WANPASS", True)
            LbInfo.Caption = "Administrator password successfully applied..." ', vbInformation ', 'App.Title
            LbInfo.ForeColor = &H80000012
            CmdSave.Enabled = False
            TextConfPass.Enabled = False: TextNewPass.Enabled = False: TextOlPass.Enabled = False
            TextOlPass.PasswordChar = "*"
        End If
    Else
         LbInfo.Caption = "You must aktif pasword..." ', vbInformation ', App.Title
         LbInfo.ForeColor = &H80000012
         CmdSave.Enabled = False
    End If
End Function

Private Sub lstLanguage_Click()
WriteLngInfoToLabel lstLanguage.List(lstLanguage.ListIndex), lblLangID1, lblLangSel1, lblLangAut1
Dim BhasaDipakai As String

    BhasaDipakai = lstLanguage.List(lstLanguage.ListIndex)
    BhasaDipakai = Mid(BhasaDipakai, InStr(BhasaDipakai, "| ") + 2)
    LangUsed = BhasaDipakai
    
    InitLanguange LangUsed
    SetStringValue &H80000001, "Software\Wan'iez\", "Language", LangUsed
    SetStringValue &H80000001, "Software\Wan'iez\", "LangID", lblLangID1
    SetStringValue &H80000001, "Software\Wan'iez\", "LangAUTOR", lblLangAut1

   Call LoadConfig '

End Sub
Private Sub TextOlPass_KeyPress(KeyAscii As Integer)
'Call loadPassword 'Aktif
If KeyAscii = 13 Then 'enter
A = GetSTRINGValue(&H80000001, "Software\Wan'iez\", "password")
B = encrypt.Cryption(A, "WANPASS", False)
If TextOlPass.Text = B Then
'text color
TextConfPass.Enabled = True
TextNewPass.Enabled = True
TextOlPass.Enabled = False
'Pass Setting
CkGenSet.Enabled = True: CkScanSet.Enabled = True: CkAppSet.Enabled = True: CkLanguage.Enabled = True
'CkPassProtect.Enabled = True '.Check3.Enabled = True:
'Pass App
CkPotection.Enabled = True: CkQuar.Enabled = True: CkUpdate.Enabled = True: CkAdditional.Enabled = True
CkDellUs.Enabled = True: CkExit.Enabled = True:

Else
LbInfo.Caption = "Invalid administrator password... please try again..." ', vbCritical, App.Title
TextOlPass.Text = ""
End If
ElseIf KeyAscii = 27 Then ' esc
TextOlPass.Text = ""
End If
End Sub
