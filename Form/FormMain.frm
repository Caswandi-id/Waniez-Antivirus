VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Wan'iez Antivirus"
   ClientHeight    =   8025
   ClientLeft      =   3360
   ClientTop       =   1590
   ClientWidth     =   11805
   Icon            =   "FormMain.frx":0000
   LinkTopic       =   "Wan'iez Antivirus"
   MaxButton       =   0   'False
   ScaleHeight     =   8025
   ScaleWidth      =   11805
   Begin VB.PictureBox PicMainSummary 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6735
      Left            =   3360
      ScaleHeight     =   6735
      ScaleWidth      =   8385
      TabIndex        =   204
      Top             =   720
      Width           =   8385
      Begin WANIEZ.rButton BtnAtur 
         Height          =   495
         Left            =   6720
         TabIndex        =   205
         Top             =   960
         Width           =   1095
         _extentx        =   1931
         _extenty        =   873
         buttonstyle     =   7
         buttontheme     =   2
         backcolor       =   14211288
         backcolorpressed=   15715986
         backcolorhover  =   16243621
         bordercolor     =   9408398
         bordercolorpressed=   6045981
         bordercolorhover=   11632444
         caption         =   "Enable"
         font            =   "FormMain.frx":52C2
      End
      Begin VB.Line MainLine 
         BorderColor     =   &H00C0C0C0&
         Index           =   0
         X1              =   240
         X2              =   8160
         Y1              =   4320
         Y2              =   4320
      End
      Begin VB.Label LbVirusVersion 
         BackStyle       =   0  'Transparent
         Caption         =   ":  -"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3600
         TabIndex        =   219
         Top             =   3240
         Width           =   4215
      End
      Begin VB.Label LbBuilidDate 
         BackStyle       =   0  'Transparent
         Caption         =   ": 12 September 2013"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3600
         TabIndex        =   218
         Top             =   3000
         Width           =   2655
      End
      Begin VB.Label LbBuildNumber 
         BackStyle       =   0  'Transparent
         Caption         =   ": 06"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3600
         TabIndex        =   217
         Top             =   2760
         Width           =   1575
      End
      Begin VB.Label LbEnggine 
         BackStyle       =   0  'Transparent
         Caption         =   ": 1.5.0.06"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3600
         TabIndex        =   216
         Top             =   2520
         Width           =   1575
      End
      Begin VB.Label LbInfoSummary 
         BackStyle       =   0  'Transparent
         Caption         =   "Virus Signature"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   960
         TabIndex        =   215
         Top             =   3240
         Width           =   2295
      End
      Begin VB.Label LbInfoSummary 
         BackStyle       =   0  'Transparent
         Caption         =   "Build Date"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   960
         TabIndex        =   214
         Top             =   3000
         Width           =   2295
      End
      Begin VB.Label LbInfoSummary 
         BackStyle       =   0  'Transparent
         Caption         =   "Build Number"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   960
         TabIndex        =   213
         Top             =   2760
         Width           =   2295
      End
      Begin VB.Label LbInfoSummary 
         BackStyle       =   0  'Transparent
         Caption         =   "Engine Version"
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
         Height          =   255
         Index           =   0
         Left            =   960
         TabIndex        =   212
         Top             =   2520
         Width           =   2295
      End
      Begin VB.Label LbWindows 
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   2
         Left            =   840
         TabIndex        =   211
         Top             =   5640
         Width           =   7215
      End
      Begin VB.Label LbWindows 
         BackStyle       =   0  'Transparent
         Caption         =   "-"
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
         Index           =   1
         Left            =   840
         TabIndex        =   210
         Top             =   5400
         Width           =   3975
      End
      Begin VB.Label LbWindows 
         BackStyle       =   0  'Transparent
         Caption         =   "-"
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
         Index           =   0
         Left            =   840
         TabIndex        =   209
         Top             =   5160
         Width           =   3975
      End
      Begin VB.Label LbstatusRTP 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Your System is fully protected"
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
         Height          =   375
         Index           =   1
         Left            =   1920
         TabIndex        =   208
         Top             =   1080
         Width           =   4815
      End
      Begin VB.Label LbstatusRTP 
         BackStyle       =   0  'Transparent
         Caption         =   "SECURED"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   0
         Left            =   1920
         TabIndex        =   207
         Top             =   600
         Width           =   3015
      End
      Begin VB.Label RTPdet 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   6360
         TabIndex        =   206
         Top             =   480
         Width           =   1815
      End
   End
   Begin VB.PictureBox Pictolls 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6735
      Left            =   3360
      ScaleHeight     =   6735
      ScaleWidth      =   8415
      TabIndex        =   112
      Top             =   720
      Width           =   8415
      Begin VB.PictureBox Pictools 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   5295
         Index           =   0
         Left            =   240
         ScaleHeight     =   5295
         ScaleWidth      =   7935
         TabIndex        =   198
         Top             =   1200
         Width           =   7935
         Begin VB.ListBox lstModule 
            ForeColor       =   &H00000000&
            Height          =   1230
            Left            =   120
            TabIndex        =   199
            Top             =   3360
            Visible         =   0   'False
            Width           =   8535
         End
         Begin WANIEZ.ucListView lvProses 
            Height          =   5175
            Left            =   0
            TabIndex        =   200
            Top             =   0
            Width           =   7935
            _extentx        =   15055
            _extenty        =   7858
            styleex         =   33
         End
         Begin VB.Label frModule 
            BackStyle       =   0  'Transparent
            Caption         =   "Modul List"
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
            Left            =   120
            TabIndex        =   203
            Top             =   3000
            Width           =   1215
         End
         Begin VB.Label lblSelectedPID 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            Height          =   210
            Left            =   1440
            TabIndex        =   202
            Top             =   720
            Width           =   90
         End
         Begin VB.Label frProses 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Proces "
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
            TabIndex        =   201
            Top             =   600
            Width           =   525
         End
      End
      Begin VB.PictureBox Pictools 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   5295
         Index           =   1
         Left            =   240
         ScaleHeight     =   5295
         ScaleWidth      =   7935
         TabIndex        =   173
         Top             =   1200
         Width           =   7935
         Begin VB.TextBox txtVirusName 
            Appearance      =   0  'Flat
            ForeColor       =   &H000000FF&
            Height          =   285
            Left            =   1680
            TabIndex        =   186
            Top             =   480
            Width           =   2655
         End
         Begin VB.TextBox txtVirusPath 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   1680
            Locked          =   -1  'True
            OLEDropMode     =   1  'Manual
            TabIndex        =   185
            Top             =   120
            Width           =   5295
         End
         Begin VB.CheckBox ckuserdb 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            Caption         =   "Check All"
            Height          =   255
            Left            =   6840
            TabIndex        =   184
            Top             =   3480
            Width           =   1095
         End
         Begin VB.PictureBox PicAnalisis 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            FillColor       =   &H00FFFFFF&
            FillStyle       =   0  'Solid
            ForeColor       =   &H00FFFFFF&
            Height          =   495
            Left            =   7200
            ScaleHeight     =   495
            ScaleWidth      =   615
            TabIndex        =   183
            Top             =   720
            Width           =   615
         End
         Begin VB.TextBox TxtNama 
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
            Left            =   1680
            Locked          =   -1  'True
            TabIndex        =   182
            Top             =   1320
            Width           =   6135
         End
         Begin VB.TextBox TxtSize 
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
            Left            =   1680
            Locked          =   -1  'True
            TabIndex        =   181
            Top             =   1680
            Width           =   6135
         End
         Begin VB.TextBox txtType 
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
            Left            =   1680
            Locked          =   -1  'True
            TabIndex        =   180
            Top             =   2040
            Width           =   6135
         End
         Begin VB.TextBox txtCeck 
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
            Height          =   525
            Left            =   1680
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   179
            Top             =   2880
            Width           =   6135
         End
         Begin VB.TextBox txtCecksum 
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
            Height          =   405
            Left            =   1680
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   178
            Top             =   2400
            Width           =   6135
         End
         Begin WANIEZ.rButton cmdRemmoveCk 
            Height          =   375
            Left            =   4440
            TabIndex        =   174
            Top             =   840
            Width           =   1335
            _extentx        =   2355
            _extenty        =   661
            buttonstyle     =   7
            backcolor       =   14211288
            backcolorpressed=   15715986
            backcolorhover  =   16243621
            bordercolor     =   11907757
            bordercolorpressed=   6045981
            bordercolorhover=   11632444
            caption         =   "Delete"
            enabled         =   0   'False
            font            =   "FormMain.frx":52EA
         End
         Begin WANIEZ.rButton cmdCancel 
            Height          =   375
            Left            =   3000
            TabIndex        =   175
            Top             =   840
            Width           =   1335
            _extentx        =   2355
            _extenty        =   661
            buttonstyle     =   7
            backcolor       =   14211288
            backcolorpressed=   15715986
            backcolorhover  =   16243621
            bordercolor     =   11907757
            bordercolorpressed=   6045981
            bordercolorhover=   11632444
            caption         =   "Cancle"
            enabled         =   0   'False
            font            =   "FormMain.frx":5312
         End
         Begin WANIEZ.rButton cmdAddVirus 
            Height          =   375
            Left            =   1680
            TabIndex        =   176
            Top             =   840
            Width           =   1335
            _extentx        =   2355
            _extenty        =   661
            buttonstyle     =   7
            backcolor       =   14211288
            backcolorpressed=   15715986
            backcolorhover  =   16243621
            bordercolor     =   11907757
            bordercolorpressed=   6045981
            bordercolorhover=   11632444
            caption         =   "Add"
            enabled         =   0   'False
            font            =   "FormMain.frx":533A
         End
         Begin WANIEZ.rButton cmdBrowse 
            Height          =   375
            Left            =   7080
            TabIndex        =   177
            Top             =   120
            Width           =   735
            _extentx        =   1296
            _extenty        =   661
            buttonstyle     =   7
            backcolor       =   14211288
            backcolorpressed=   15715986
            backcolorhover  =   16243621
            bordercolor     =   9408398
            bordercolorpressed=   6045981
            bordercolorhover=   11632444
            caption         =   "..."
            font            =   "FormMain.frx":5362
         End
         Begin WANIEZ.ucListView lvm31 
            Height          =   1335
            Left            =   0
            TabIndex        =   187
            Top             =   3840
            Width           =   7935
            _extentx        =   15055
            _extenty        =   4895
            style           =   4
            styleex         =   33
            showsort        =   -1  'True
         End
         Begin VB.Label frTemp 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "-"
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
            Left            =   0
            TabIndex        =   195
            Top             =   3600
            Width           =   60
         End
         Begin VB.Label lblMalwarePath 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Malware Path"
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
            Left            =   0
            TabIndex        =   194
            Top             =   120
            Width           =   975
         End
         Begin VB.Label lblMalwareName 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Malware Name"
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
            Left            =   0
            TabIndex        =   193
            Top             =   480
            Width           =   1050
         End
         Begin VB.Label LbAnalisis 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Type File "
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
            Index           =   3
            Left            =   0
            TabIndex        =   192
            Top             =   2040
            Width           =   690
         End
         Begin VB.Label LbAnalisis 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "File Name"
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
            Index           =   0
            Left            =   0
            TabIndex        =   191
            Top             =   1440
            Width           =   690
         End
         Begin VB.Label LbAnalisis 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "File Size"
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
            Index           =   1
            Left            =   0
            TabIndex        =   190
            Top             =   1680
            Width           =   570
         End
         Begin VB.Label LbAnalisis 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Checksum"
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
            Index           =   2
            Left            =   0
            TabIndex        =   189
            Top             =   2400
            Width           =   720
         End
         Begin VB.Label LbAnalisis 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Check Additional Byte"
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
            Height          =   570
            Index           =   4
            Left            =   0
            TabIndex        =   188
            Top             =   2880
            Width           =   1440
         End
      End
      Begin VB.PictureBox Pictools 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   5295
         Index           =   2
         Left            =   240
         ScaleHeight     =   5295
         ScaleWidth      =   7935
         TabIndex        =   158
         Top             =   1200
         Width           =   7935
         Begin VB.ListBox lstExceptReg 
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
            Height          =   1035
            Left            =   120
            TabIndex        =   164
            Top             =   2160
            Width           =   5895
         End
         Begin VB.ListBox lstExceptFolder 
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
            Height          =   1035
            Left            =   120
            TabIndex        =   163
            Top             =   3840
            Width           =   5880
         End
         Begin VB.ListBox lstExceptFile 
            ForeColor       =   &H00000000&
            Height          =   1035
            Left            =   120
            TabIndex        =   162
            Top             =   480
            Width           =   5895
         End
         Begin WANIEZ.rButton cmdRemExcFile 
            Height          =   375
            Left            =   6240
            TabIndex        =   159
            Top             =   1200
            Width           =   1575
            _extentx        =   2778
            _extenty        =   661
            buttonstyle     =   7
            backcolor       =   14211288
            backcolorpressed=   15715986
            backcolorhover  =   16243621
            bordercolor     =   9408398
            bordercolorpressed=   6045981
            bordercolorhover=   11632444
            caption         =   "Remove All"
            font            =   "FormMain.frx":538A
         End
         Begin WANIEZ.rButton cmdRemExcFile1 
            Height          =   375
            Left            =   6240
            TabIndex        =   160
            Top             =   840
            Width           =   1575
            _extentx        =   2778
            _extenty        =   661
            buttonstyle     =   7
            backcolor       =   14211288
            backcolorpressed=   15715986
            backcolorhover  =   16243621
            bordercolor     =   9408398
            bordercolorpressed=   6045981
            bordercolorhover=   11632444
            caption         =   "Remove Selected"
            font            =   "FormMain.frx":53B2
         End
         Begin WANIEZ.rButton cmdAddExcFile 
            Height          =   375
            Left            =   6240
            TabIndex        =   161
            Top             =   480
            Width           =   1575
            _extentx        =   2778
            _extenty        =   661
            buttonstyle     =   7
            backcolor       =   14211288
            backcolorpressed=   15715986
            backcolorhover  =   16243621
            bordercolor     =   9408398
            bordercolorpressed=   6045981
            bordercolorhover=   11632444
            caption         =   "Add File"
            font            =   "FormMain.frx":53DA
         End
         Begin WANIEZ.rButton cmdRemExcReg1 
            Height          =   375
            Left            =   6240
            TabIndex        =   165
            Top             =   2160
            Width           =   1575
            _extentx        =   2778
            _extenty        =   661
            buttonstyle     =   7
            backcolor       =   14211288
            backcolorpressed=   15715986
            backcolorhover  =   16243621
            bordercolor     =   9408398
            bordercolorpressed=   6045981
            bordercolorhover=   11632444
            caption         =   "Remove Selected"
            font            =   "FormMain.frx":5402
         End
         Begin WANIEZ.rButton cmdRemExcReg 
            Height          =   375
            Left            =   6240
            TabIndex        =   166
            Top             =   2520
            Width           =   1575
            _extentx        =   2778
            _extenty        =   661
            buttonstyle     =   7
            backcolor       =   14211288
            backcolorpressed=   15715986
            backcolorhover  =   16243621
            bordercolor     =   9408398
            bordercolorpressed=   6045981
            bordercolorhover=   11632444
            caption         =   "Remove All"
            font            =   "FormMain.frx":542A
         End
         Begin WANIEZ.rButton cmdAddExcFolder 
            Height          =   375
            Left            =   6240
            TabIndex        =   167
            Top             =   3840
            Width           =   1575
            _extentx        =   2778
            _extenty        =   661
            buttonstyle     =   7
            backcolor       =   14211288
            backcolorpressed=   15715986
            backcolorhover  =   16243621
            bordercolor     =   9408398
            bordercolorpressed=   6045981
            bordercolorhover=   11632444
            caption         =   "Add File"
            font            =   "FormMain.frx":5452
         End
         Begin WANIEZ.rButton cmdRemovePath1 
            Height          =   375
            Left            =   6240
            TabIndex        =   168
            Top             =   4200
            Width           =   1575
            _extentx        =   2778
            _extenty        =   661
            buttonstyle     =   7
            backcolor       =   14211288
            backcolorpressed=   15715986
            backcolorhover  =   16243621
            bordercolor     =   9408398
            bordercolorpressed=   6045981
            bordercolorhover=   11632444
            caption         =   "Remove Selected"
            font            =   "FormMain.frx":547A
         End
         Begin WANIEZ.rButton cmdRemovePath 
            Height          =   375
            Left            =   6240
            TabIndex        =   169
            Top             =   4560
            Width           =   1575
            _extentx        =   2778
            _extenty        =   661
            buttonstyle     =   7
            backcolor       =   14211288
            backcolorpressed=   15715986
            backcolorhover  =   16243621
            bordercolor     =   9408398
            bordercolorpressed=   6045981
            bordercolorhover=   11632444
            caption         =   "Remove All"
            font            =   "FormMain.frx":54A2
         End
         Begin VB.Label lblExceptReg 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Registry Exception - I'am sure this is normal value. Don't catch as a bad value, value(s) below"
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
            TabIndex        =   172
            Top             =   1800
            Width           =   8535
         End
         Begin VB.Label lblExceptFolder 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "RTP Exception - Dont Give me Warning about Threat in this Path"
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
            Left            =   120
            TabIndex        =   171
            Top             =   3480
            Width           =   8535
         End
         Begin VB.Label lblExceptFile 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "File Exception - I'am sure this is normal file, dont catch as a malware file(s) below"
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
            TabIndex        =   170
            Top             =   120
            Width           =   8565
         End
      End
      Begin VB.PictureBox Pictools 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   5295
         Index           =   3
         Left            =   240
         ScaleHeight     =   5295
         ScaleWidth      =   7935
         TabIndex        =   142
         Top             =   1200
         Width           =   7935
         Begin VB.ListBox lstPlugin 
            Appearance      =   0  'Flat
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
            Height          =   1980
            Left            =   120
            TabIndex        =   144
            Top             =   480
            Width           =   7575
         End
         Begin WANIEZ.rButton cmdExecutePlug 
            Height          =   495
            Left            =   5760
            TabIndex        =   143
            Top             =   2640
            Width           =   1935
            _extentx        =   3413
            _extenty        =   873
            buttonstyle     =   7
            backcolor       =   14211288
            backcolorpressed=   15715986
            backcolorhover  =   16243621
            bordercolor     =   9408398
            bordercolorpressed=   6045981
            bordercolorhover=   11632444
            caption         =   "Apply"
            font            =   "FormMain.frx":54CA
         End
         Begin VB.Label lblAvalaiblePlug 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Avalaible Plugin(s)"
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
            TabIndex        =   157
            Top             =   120
            Width           =   1305
         End
         Begin VB.Label lblPlugVer1 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   ":                    "
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
            Left            =   1560
            TabIndex        =   156
            Top             =   4080
            Width           =   2280
         End
         Begin VB.Label lblPlugVer 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Verification Code"
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
            TabIndex        =   155
            Top             =   4080
            Width           =   1215
         End
         Begin VB.Label lblPlugAutSite1 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   ":                    "
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
            Left            =   1560
            TabIndex        =   154
            Top             =   3840
            Width           =   2280
         End
         Begin VB.Label lblPlugAutSite 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Author Site"
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
            TabIndex        =   153
            Top             =   3840
            Width           =   810
         End
         Begin VB.Label lblPlugAutEmail1 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   ":                    "
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
            Left            =   1560
            TabIndex        =   152
            Top             =   3600
            Width           =   2280
         End
         Begin VB.Label lblPlugAutEmail 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Author Email"
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
            TabIndex        =   151
            Top             =   3600
            Width           =   900
         End
         Begin VB.Label lblPlugSelect 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Plugin Selected"
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
            TabIndex        =   150
            Top             =   3120
            Width           =   1080
         End
         Begin VB.Label lblPlugAut 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Plugin Author"
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
            TabIndex        =   149
            Top             =   3360
            Width           =   960
         End
         Begin VB.Label lblPlugDesc 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Plugin Description"
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
            TabIndex        =   148
            Top             =   4320
            Width           =   1260
         End
         Begin VB.Label lblPlugSelect1 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   ":                    "
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
            Left            =   1560
            TabIndex        =   147
            Top             =   3120
            Width           =   960
         End
         Begin VB.Label lblPlugAut1 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   ":                    "
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
            Left            =   1560
            TabIndex        =   146
            Top             =   3360
            Width           =   960
         End
         Begin VB.Label lblPlugDesc1 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   ":                                  "
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
            Height          =   435
            Left            =   1560
            TabIndex        =   145
            Top             =   4320
            Width           =   7095
         End
      End
      Begin VB.PictureBox Pictools 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   5295
         Index           =   4
         Left            =   240
         ScaleHeight     =   5295
         ScaleWidth      =   7935
         TabIndex        =   115
         Top             =   1200
         Width           =   7935
         Begin VB.Frame FmSystem 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Caption         =   "Windows Explorer"
            ForeColor       =   &H80000008&
            Height          =   4575
            Index           =   2
            Left            =   3720
            TabIndex        =   129
            Top             =   360
            Width           =   2895
            Begin VB.CheckBox Check28 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Full Path TitleBar"
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
               TabIndex        =   141
               Top             =   2400
               Width           =   1530
            End
            Begin VB.CheckBox Check27 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Full Path AddBars"
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
               TabIndex        =   140
               Top             =   2760
               Width           =   1650
            End
            Begin VB.CheckBox Check26 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Extension File"
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
               TabIndex        =   139
               Top             =   960
               Width           =   1410
            End
            Begin VB.CheckBox Check25 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Hidden File"
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
               TabIndex        =   138
               Top             =   240
               Width           =   1170
            End
            Begin VB.CheckBox Check29 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Turn Off Autoplay"
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
               TabIndex        =   137
               Top             =   3120
               Width           =   1800
            End
            Begin VB.CheckBox Check18 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Folder Options"
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
               TabIndex        =   136
               Top             =   1320
               Width           =   1410
            End
            Begin VB.CheckBox Check19 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Control Panel"
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
               TabIndex        =   135
               Top             =   2040
               Width           =   1410
            End
            Begin VB.CheckBox Check20 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Explorer's Context Menu"
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
               TabIndex        =   134
               Top             =   4200
               Width           =   2175
            End
            Begin VB.CheckBox Check21 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Taskbar Context Menu"
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
               TabIndex        =   133
               Top             =   3840
               Width           =   2175
            End
            Begin VB.CheckBox Check22 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Taskbar Settings"
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
               TabIndex        =   132
               Top             =   3480
               Width           =   1605
            End
            Begin VB.CheckBox Check23 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Desktop Item"
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
               TabIndex        =   131
               Top             =   1680
               Width           =   1410
            End
            Begin VB.CheckBox Check24 
               BackColor       =   &H00FFFFFF&
               Caption         =   "System File"
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
               TabIndex        =   130
               Top             =   600
               Width           =   1170
            End
         End
         Begin VB.Frame FmSystem 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Caption         =   "System Aplication"
            ForeColor       =   &H80000008&
            Height          =   2175
            Index           =   1
            Left            =   1320
            TabIndex        =   123
            Top             =   2760
            Width           =   2175
            Begin VB.CheckBox Check9 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Recent Doc"
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
               TabIndex        =   128
               Top             =   360
               Width           =   1215
            End
            Begin VB.CheckBox Check13 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Display Setting"
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
               TabIndex        =   127
               Top             =   1800
               Width           =   1365
            End
            Begin VB.CheckBox Check15 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Task Manager"
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
               TabIndex        =   126
               Top             =   1080
               Width           =   1365
            End
            Begin VB.CheckBox Check14 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Registry Editor"
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
               TabIndex        =   125
               Top             =   1440
               Width           =   1365
            End
            Begin VB.CheckBox Check17 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Win Hotkeys"
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
               TabIndex        =   124
               Top             =   720
               Width           =   1290
            End
         End
         Begin VB.Frame FmSystem 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Caption         =   "Start Menu"
            ForeColor       =   &H80000008&
            Height          =   2295
            Index           =   0
            Left            =   1320
            TabIndex        =   116
            Top             =   360
            Width           =   2175
            Begin VB.CheckBox Check12 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Help"
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
               TabIndex        =   122
               Top             =   600
               Width           =   735
            End
            Begin VB.CheckBox Check11 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Shutdown"
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
               TabIndex        =   121
               Top             =   2040
               Width           =   1095
            End
            Begin VB.CheckBox Check10 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Run"
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
               TabIndex        =   120
               Top             =   960
               Width           =   855
            End
            Begin VB.CheckBox Check8 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Log Off"
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
               TabIndex        =   119
               Top             =   1680
               Width           =   975
            End
            Begin VB.CheckBox Check7 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Find"
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
               TabIndex        =   118
               Top             =   240
               Width           =   735
            End
            Begin VB.CheckBox Check16 
               BackColor       =   &H00FFFFFF&
               Caption         =   "CMD"
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
               TabIndex        =   117
               Top             =   1320
               Width           =   690
            End
         End
      End
      Begin VB.PictureBox Pictools 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   5295
         Index           =   5
         Left            =   240
         ScaleHeight     =   5295
         ScaleWidth      =   7935
         TabIndex        =   113
         Top             =   1200
         Width           =   7935
         Begin WANIEZ.ucListView lstKunci 
            Height          =   3375
            Left            =   120
            TabIndex        =   114
            Top             =   0
            Width           =   7935
            _extentx        =   15055
            _extenty        =   8281
            styleex         =   33
         End
      End
      Begin WANIEZ.Tab TabTools 
         Height          =   5895
         Left            =   120
         TabIndex        =   196
         Top             =   720
         Width           =   8175
         _extentx        =   14420
         _extenty        =   10398
         backcolor       =   16777215
         closebutton     =   0   'False
         blurforecolor   =   0
         activeforecolor =   0
         picture         =   "FormMain.frx":54F2
         alltabsforecolor=   -2147483630
         fontname        =   "Tahoma"
      End
      Begin VB.Label Lbtitle 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Additional Protection"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   0
         TabIndex        =   197
         Top             =   120
         Width           =   8415
      End
   End
   Begin VB.PictureBox PicQuar 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6735
      Left            =   3360
      ScaleHeight     =   6735
      ScaleWidth      =   8415
      TabIndex        =   106
      Top             =   720
      Width           =   8415
      Begin WANIEZ.rButton CmdRestoreto 
         Height          =   375
         Left            =   3840
         TabIndex        =   107
         Top             =   6120
         Width           =   1575
         _extentx        =   2778
         _extenty        =   661
         buttonstyle     =   7
         backcolor       =   14211288
         backcolorpressed=   15715986
         backcolorhover  =   16243621
         bordercolor     =   9408398
         bordercolorpressed=   6045981
         bordercolorhover=   11632444
         caption         =   "Restore to ..."
         font            =   "FormMain.frx":5510
      End
      Begin WANIEZ.rButton CmdRestore 
         Height          =   375
         Left            =   2280
         TabIndex        =   108
         Top             =   6120
         Width           =   1575
         _extentx        =   2778
         _extenty        =   661
         buttonstyle     =   7
         backcolor       =   14211288
         backcolorpressed=   15715986
         backcolorhover  =   16243621
         bordercolor     =   9408398
         bordercolorpressed=   6045981
         bordercolorhover=   11632444
         caption         =   "Restore"
         font            =   "FormMain.frx":5538
      End
      Begin WANIEZ.rButton CmnKarantine 
         Height          =   375
         Left            =   120
         TabIndex        =   109
         Top             =   6120
         Width           =   1935
         _extentx        =   3413
         _extenty        =   661
         buttonstyle     =   7
         backcolor       =   14211288
         backcolorpressed=   15715986
         backcolorhover  =   16243621
         bordercolor     =   9408398
         bordercolorpressed=   6045981
         bordercolorhover=   11632444
         caption         =   "Quarantine File"
         font            =   "FormMain.frx":5560
      End
      Begin WANIEZ.ucListView lvJail 
         Height          =   5055
         Left            =   120
         TabIndex        =   110
         Top             =   840
         Width           =   8115
         _extentx        =   22463
         _extenty        =   10186
         styleex         =   33
      End
      Begin VB.Label LbQuar 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Virus Chest - 0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   0
         TabIndex        =   111
         Top             =   120
         Width           =   8415
      End
   End
   Begin VB.PictureBox PicUpd 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   6735
      Left            =   3360
      ScaleHeight     =   6735
      ScaleWidth      =   8415
      TabIndex        =   90
      Top             =   720
      Width           =   8415
      Begin VB.PictureBox PicUpdate 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   5895
         Left            =   0
         ScaleHeight     =   5895
         ScaleWidth      =   8415
         TabIndex        =   92
         Top             =   720
         Width           =   8415
         Begin VB.TextBox txtRetriveInfo 
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1455
            Left            =   1200
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   94
            Text            =   "FormMain.frx":5588
            Top             =   1440
            Width           =   6855
         End
         Begin VB.ListBox lstListWorm 
            ForeColor       =   &H00000000&
            Height          =   1230
            Left            =   120
            Sorted          =   -1  'True
            TabIndex        =   93
            Top             =   4320
            Visible         =   0   'False
            Width           =   3855
         End
         Begin WANIEZ.rButton cmdCheckUpdate 
            Height          =   495
            Left            =   120
            TabIndex        =   95
            Top             =   3480
            Width           =   1695
            _extentx        =   2990
            _extenty        =   873
            buttonstyle     =   7
            backcolor       =   14211288
            backcolorpressed=   15715986
            backcolorhover  =   16243621
            bordercolor     =   9408398
            bordercolorpressed=   6045981
            bordercolorhover=   11632444
            caption         =   "Update Now"
            font            =   "FormMain.frx":558C
         End
         Begin WANIEZ.ucProgressBar PBC 
            Height          =   255
            Left            =   120
            Top             =   600
            Width           =   4215
            _extentx        =   7223
            _extenty        =   450
         End
         Begin WANIEZ.ucProgressBar PB_UPD 
            Height          =   255
            Left            =   120
            Top             =   240
            Width           =   8175
            _extentx        =   14843
            _extenty        =   450
         End
         Begin WANIEZ.ucListView ucListVirus2 
            Height          =   1335
            Left            =   4080
            TabIndex        =   96
            Top             =   4320
            Visible         =   0   'False
            Width           =   4095
            _extentx        =   7435
            _extenty        =   4895
            style           =   4
            styleex         =   33
            showsort        =   -1  'True
         End
         Begin WANIEZ.rButton cmdupload 
            Height          =   495
            Left            =   1920
            TabIndex        =   97
            Top             =   3480
            Visible         =   0   'False
            Width           =   1695
            _extentx        =   2990
            _extenty        =   873
            buttonstyle     =   7
            backcolor       =   14211288
            backcolorpressed=   15715986
            backcolorhover  =   16243621
            bordercolor     =   9408398
            bordercolorpressed=   6045981
            bordercolorhover=   11632444
            caption         =   "Upload Virus"
            font            =   "FormMain.frx":55B4
         End
         Begin VB.Label Wormupdate 
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            Height          =   255
            Left            =   3240
            TabIndex        =   105
            Top             =   2160
            Width           =   735
         End
         Begin VB.Label lblStatusUpdate 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Update Now "
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   1200
            TabIndex        =   104
            Top             =   1080
            Width           =   1065
         End
         Begin VB.Label LbupdateMain 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Status"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   0
            Left            =   120
            TabIndex        =   103
            Top             =   1080
            Width           =   525
         End
         Begin VB.Label LbupdateMain 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Info"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   1
            Left            =   120
            TabIndex        =   102
            Top             =   1440
            Width           =   300
         End
         Begin VB.Label LbBts 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   ":"
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
            Index           =   10
            Left            =   960
            TabIndex        =   101
            Top             =   1440
            Width           =   45
         End
         Begin VB.Label LbBts 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   ":"
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
            Index           =   11
            Left            =   960
            TabIndex        =   100
            Top             =   1080
            Width           =   45
         End
         Begin VB.Line MainLine 
            BorderColor     =   &H00E0E0E0&
            Index           =   2
            X1              =   120
            X2              =   8280
            Y1              =   3240
            Y2              =   3240
         End
         Begin VB.Label LbInDB 
            BackStyle       =   0  'Transparent
            Caption         =   "Exsternal Database - 0"
            Height          =   255
            Left            =   4080
            TabIndex        =   99
            Top             =   4080
            Visible         =   0   'False
            Width           =   2175
         End
         Begin VB.Label LbExDB 
            BackStyle       =   0  'Transparent
            Caption         =   "Exsternal Database - 0"
            Height          =   255
            Left            =   120
            TabIndex        =   98
            Top             =   4080
            Visible         =   0   'False
            Width           =   2175
         End
      End
      Begin VB.Label Lbtitle 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Update"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   0
         TabIndex        =   91
         Top             =   120
         Width           =   8415
      End
   End
   Begin VB.PictureBox PicMainScan 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6735
      Left            =   3360
      ScaleHeight     =   6735
      ScaleWidth      =   8415
      TabIndex        =   22
      Top             =   720
      Width           =   8415
      Begin VB.PictureBox Picscan 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   5295
         Index           =   0
         Left            =   240
         ScaleHeight     =   5295
         ScaleWidth      =   7935
         TabIndex        =   41
         Top             =   1200
         Width           =   7935
         Begin WANIEZ.rButton BtnRefres 
            Height          =   375
            Left            =   4920
            TabIndex        =   60
            Top             =   4440
            Width           =   855
            _extentx        =   1508
            _extenty        =   661
            buttonstyle     =   7
            buttonstylecolors=   3
            buttontheme     =   3
            backcolor       =   14211288
            backcolorpressed=   15715986
            backcolorhover  =   16243621
            bordercolor     =   9408398
            bordercolorpressed=   6045981
            bordercolorhover=   11632444
            caption         =   "....."
            font            =   "FormMain.frx":55DC
         End
         Begin WANIEZ.rButton cmdStartScan 
            Height          =   615
            Left            =   5880
            TabIndex        =   59
            Top             =   4320
            Width           =   1575
            _extentx        =   2778
            _extenty        =   1085
            buttonstyle     =   7
            buttonstylecolors=   3
            buttontheme     =   2
            backcolor       =   14211288
            backcolorpressed=   15715986
            backcolorhover  =   16243621
            bordercolor     =   9408398
            bordercolorpressed=   6045981
            bordercolorhover=   11632444
            caption         =   "Start Scan"
            font            =   "FormMain.frx":5604
         End
         Begin WANIEZ.DirTree DirTree 
            Height          =   5055
            Left            =   120
            TabIndex        =   42
            Top             =   120
            Width           =   7695
            _extentx        =   15055
            _extenty        =   7858
         End
      End
      Begin VB.PictureBox Picscan 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   5295
         Index           =   1
         Left            =   240
         ScaleHeight     =   5295
         ScaleWidth      =   7935
         TabIndex        =   35
         Top             =   1200
         Width           =   7935
         Begin WANIEZ.ucProgressBar PB1 
            Height          =   255
            Left            =   120
            Top             =   840
            Width           =   7815
            _extentx        =   13785
            _extenty        =   450
         End
         Begin WANIEZ.rButton BtnFix 
            Height          =   375
            Left            =   5280
            TabIndex        =   63
            Top             =   240
            Width           =   1335
            _extentx        =   2355
            _extenty        =   661
            buttonstyle     =   7
            buttonstylecolors=   3
            backcolor       =   14211288
            backcolorpressed=   15715986
            backcolorhover  =   16243621
            bordercolor     =   11907757
            bordercolorpressed=   6045981
            bordercolorhover=   11632444
            caption         =   "Fix All"
            enabled         =   0   'False
            font            =   "FormMain.frx":562C
         End
         Begin WANIEZ.rButton BTNlOG 
            Height          =   375
            Left            =   4080
            TabIndex        =   62
            Top             =   240
            Width           =   1215
            _extentx        =   2143
            _extenty        =   661
            buttonstyle     =   7
            buttonstylecolors=   3
            backcolor       =   14211288
            backcolorpressed=   15715986
            backcolorhover  =   16243621
            bordercolor     =   11907757
            bordercolorpressed=   6045981
            bordercolorhover=   11632444
            caption         =   "Log Scan"
            enabled         =   0   'False
            font            =   "FormMain.frx":5654
         End
         Begin WANIEZ.rButton BtnResult 
            Height          =   375
            Left            =   6600
            TabIndex        =   61
            Top             =   240
            Width           =   1335
            _extentx        =   2355
            _extenty        =   661
            buttonstyle     =   7
            buttonstylecolors=   3
            backcolor       =   14211288
            backcolorpressed=   15715986
            backcolorhover  =   16243621
            bordercolor     =   11907757
            bordercolorpressed=   6045981
            bordercolorhover=   11632444
            caption         =   "Result"
            enabled         =   0   'False
            font            =   "FormMain.frx":567C
         End
         Begin VB.Label lblbanter 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Speed Scanning"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   225
            Left            =   2640
            TabIndex        =   89
            Top             =   1560
            Width           =   1365
         End
         Begin VB.Label LbBts 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   ":"
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
            Index           =   9
            Left            =   1560
            TabIndex        =   88
            Top             =   3720
            Width           =   45
         End
         Begin VB.Label lblProcessed 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Processed "
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   225
            Left            =   120
            TabIndex        =   83
            Top             =   3720
            Width           =   945
         End
         Begin VB.Label LbBts 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   ":"
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
            Index           =   8
            Left            =   4440
            TabIndex        =   82
            Top             =   3360
            Width           =   45
         End
         Begin VB.Label LbBts 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   ":"
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
            Index           =   7
            Left            =   4440
            TabIndex        =   81
            Top             =   3000
            Width           =   45
         End
         Begin VB.Label LbBts 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   ":"
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
            Index           =   6
            Left            =   1560
            TabIndex        =   80
            Top             =   3000
            Width           =   45
         End
         Begin VB.Label LbBts 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   ":"
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
            Index           =   5
            Left            =   1560
            TabIndex        =   79
            Top             =   3360
            Width           =   45
         End
         Begin VB.Label LbBts 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   ":"
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
            Index           =   4
            Left            =   1560
            TabIndex        =   78
            Top             =   2640
            Width           =   45
         End
         Begin VB.Label LbBts 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   ":"
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
            Index           =   3
            Left            =   1560
            TabIndex        =   77
            Top             =   2280
            Width           =   45
         End
         Begin VB.Label LbBts 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   ":"
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
            Index           =   2
            Left            =   1560
            TabIndex        =   76
            Top             =   1920
            Width           =   45
         End
         Begin VB.Label LbBts 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   ":"
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
            Left            =   1560
            TabIndex        =   75
            Top             =   1560
            Width           =   45
         End
         Begin VB.Label LbBts 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   ":"
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
            Left            =   1560
            TabIndex        =   74
            Top             =   1200
            Width           =   45
         End
         Begin VB.Label lbMalware1 
            BackStyle       =   0  'Transparent
            Caption         =   "Malware "
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   120
            TabIndex        =   58
            Top             =   3000
            Width           =   975
         End
         Begin VB.Label lbRegistry1 
            BackStyle       =   0  'Transparent
            Caption         =   "Registry"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   120
            TabIndex        =   57
            Top             =   3360
            Width           =   975
         End
         Begin VB.Label lbHidden1 
            BackStyle       =   0  'Transparent
            Caption         =   "Hidden"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   3000
            TabIndex        =   56
            Top             =   3000
            Width           =   975
         End
         Begin VB.Label lbMalware 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   225
            Left            =   1680
            TabIndex        =   55
            Top             =   3000
            Width           =   105
         End
         Begin VB.Label lbReg 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   225
            Left            =   1680
            TabIndex        =   54
            Top             =   3360
            Width           =   105
         End
         Begin VB.Label lbHidden 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00008080&
            Height          =   225
            Left            =   4560
            TabIndex        =   53
            Top             =   3000
            Width           =   105
         End
         Begin VB.Label lbInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00008000&
            Height          =   225
            Left            =   4560
            TabIndex        =   52
            Top             =   3360
            Width           =   105
         End
         Begin VB.Label lblInfor 
            BackStyle       =   0  'Transparent
            Caption         =   "Information"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   3000
            TabIndex        =   51
            Top             =   3360
            Width           =   975
         End
         Begin VB.Label lbTime1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Time"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   225
            Left            =   120
            TabIndex        =   50
            Top             =   1560
            Width           =   420
         End
         Begin VB.Label lbFileFound1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Founded"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   225
            Left            =   120
            TabIndex        =   49
            Top             =   1920
            Width           =   735
         End
         Begin VB.Label lbFileCheck1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Checked"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   225
            Left            =   120
            TabIndex        =   48
            Top             =   2280
            Width           =   735
         End
         Begin VB.Label lbBypass1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ByPassed"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   225
            Left            =   120
            TabIndex        =   47
            Top             =   2640
            Width           =   840
         End
         Begin VB.Label lbTime 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "0 :0 :0"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   225
            Left            =   1680
            TabIndex        =   46
            Top             =   1560
            Width           =   495
         End
         Begin VB.Label lbFileFound 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   225
            Left            =   1680
            TabIndex        =   45
            Top             =   1920
            Width           =   105
         End
         Begin VB.Label lbFileCheck 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   225
            Left            =   1680
            TabIndex        =   44
            Top             =   2280
            Width           =   105
         End
         Begin VB.Label lbBypass 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   225
            Left            =   1680
            TabIndex        =   43
            Top             =   2640
            Width           =   105
         End
         Begin VB.Label lbStatus22 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "0 folder(s), 0 file(s), Size:0 KB [Kilo Byte]"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   1680
            TabIndex        =   39
            Top             =   1200
            Width           =   6285
         End
         Begin VB.Label lbStatus1 
            BackStyle       =   0  'Transparent
            Caption         =   "Status       "
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   120
            TabIndex        =   38
            Top             =   1200
            Width           =   975
         End
         Begin VB.Label lbObject 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "-"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   1575
            Left            =   1680
            TabIndex        =   37
            Top             =   3720
            Width           =   6255
         End
         Begin VB.Label lbStatus 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Ready"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   120
            TabIndex        =   36
            Top             =   360
            Width           =   525
         End
      End
      Begin VB.PictureBox Picscan 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   5295
         Index           =   2
         Left            =   240
         ScaleHeight     =   5295
         ScaleWidth      =   7935
         TabIndex        =   31
         Top             =   1200
         Width           =   7935
         Begin WANIEZ.rButton cmdFixMalware 
            Height          =   375
            Left            =   1680
            TabIndex        =   65
            Top             =   4800
            Width           =   1695
            _extentx        =   2990
            _extenty        =   661
            buttonstyle     =   7
            buttonstylecolors=   3
            backcolor       =   14211288
            backcolorpressed=   15715986
            backcolorhover  =   16243621
            bordercolor     =   11907757
            bordercolorpressed=   6045981
            bordercolorhover=   11632444
            caption         =   "Fix check"
            enabled         =   0   'False
            font            =   "FormMain.frx":56A4
         End
         Begin WANIEZ.rButton cmdFixMalwareAll 
            Height          =   375
            Left            =   0
            TabIndex        =   64
            Top             =   4800
            Width           =   1695
            _extentx        =   2990
            _extenty        =   661
            buttonstyle     =   7
            buttonstylecolors=   3
            backcolor       =   14211288
            backcolorpressed=   15715986
            backcolorhover  =   16243621
            bordercolor     =   11907757
            bordercolorpressed=   6045981
            bordercolorhover=   11632444
            caption         =   "Fix All Object"
            enabled         =   0   'False
            font            =   "FormMain.frx":56CC
         End
         Begin VB.PictureBox picInfoMalware 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   7920
            ScaleHeight     =   615
            ScaleWidth      =   735
            TabIndex        =   33
            Top             =   4080
            Width           =   735
         End
         Begin VB.CheckBox ckScan 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            Caption         =   "Check All"
            Enabled         =   0   'False
            Height          =   255
            Index           =   0
            Left            =   6960
            TabIndex        =   32
            Top             =   4800
            Width           =   975
         End
         Begin WANIEZ.ucListView lvMalware 
            Height          =   4335
            Left            =   0
            TabIndex        =   34
            Top             =   360
            Width           =   7935
            _extentx        =   15055
            _extenty        =   6588
            styleex         =   37
            showsort        =   -1  'True
         End
         Begin WANIEZ.rButton CmdQuarAll 
            Height          =   375
            Left            =   3360
            TabIndex        =   72
            Top             =   4800
            Width           =   1695
            _extentx        =   2990
            _extenty        =   661
            buttonstyle     =   7
            buttonstylecolors=   3
            backcolor       =   14211288
            backcolorpressed=   15715986
            backcolorhover  =   16243621
            bordercolor     =   11907757
            bordercolorpressed=   6045981
            bordercolorhover=   11632444
            caption         =   "Qurantine All"
            enabled         =   0   'False
            font            =   "FormMain.frx":56F4
         End
      End
      Begin VB.PictureBox Picscan 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   5295
         Index           =   3
         Left            =   240
         ScaleHeight     =   5295
         ScaleWidth      =   7935
         TabIndex        =   28
         Top             =   1200
         Width           =   7935
         Begin VB.CheckBox ckScan 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            Caption         =   "Check All"
            Enabled         =   0   'False
            Height          =   255
            Index           =   1
            Left            =   6960
            TabIndex        =   29
            Top             =   4800
            Width           =   975
         End
         Begin WANIEZ.ucListView lvHidden 
            Height          =   4335
            Left            =   0
            TabIndex        =   30
            Top             =   360
            Width           =   7935
            _extentx        =   15055
            _extenty        =   6588
            styleex         =   37
            showsort        =   -1  'True
         End
         Begin WANIEZ.rButton cmdFixHiddenall 
            Height          =   375
            Left            =   0
            TabIndex        =   66
            Top             =   4800
            Width           =   1695
            _extentx        =   2990
            _extenty        =   661
            buttonstyle     =   7
            buttonstylecolors=   3
            backcolor       =   14211288
            backcolorpressed=   15715986
            backcolorhover  =   16243621
            bordercolor     =   11907757
            bordercolorpressed=   6045981
            bordercolorhover=   11632444
            caption         =   "Fix All Object"
            enabled         =   0   'False
            font            =   "FormMain.frx":571C
         End
         Begin WANIEZ.rButton cmdFixHidden 
            Height          =   375
            Left            =   1680
            TabIndex        =   67
            Top             =   4800
            Width           =   1695
            _extentx        =   2990
            _extenty        =   661
            buttonstyle     =   7
            buttonstylecolors=   3
            backcolor       =   14211288
            backcolorpressed=   15715986
            backcolorhover  =   16243621
            bordercolor     =   11907757
            bordercolorpressed=   6045981
            bordercolorhover=   11632444
            caption         =   "Fix check"
            enabled         =   0   'False
            font            =   "FormMain.frx":5744
         End
      End
      Begin VB.PictureBox Picscan 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   5295
         Index           =   4
         Left            =   240
         ScaleHeight     =   5295
         ScaleWidth      =   7935
         TabIndex        =   25
         Top             =   1200
         Width           =   7935
         Begin VB.CheckBox ckScan 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            Caption         =   "Check All"
            Enabled         =   0   'False
            Height          =   255
            Index           =   2
            Left            =   6960
            TabIndex        =   26
            Top             =   4800
            Width           =   975
         End
         Begin WANIEZ.ucListView lvRegistry 
            Height          =   4335
            Left            =   0
            TabIndex        =   27
            Top             =   360
            Width           =   7935
            _extentx        =   15055
            _extenty        =   6588
            styleex         =   37
            showsort        =   -1  'True
         End
         Begin WANIEZ.rButton cmdFixRegAll 
            Height          =   375
            Left            =   0
            TabIndex        =   68
            Top             =   4800
            Width           =   1695
            _extentx        =   2990
            _extenty        =   661
            buttonstyle     =   7
            buttonstylecolors=   3
            backcolor       =   14211288
            backcolorpressed=   15715986
            backcolorhover  =   16243621
            bordercolor     =   11907757
            bordercolorpressed=   6045981
            bordercolorhover=   11632444
            caption         =   "Fix All Object"
            enabled         =   0   'False
            font            =   "FormMain.frx":576C
         End
         Begin WANIEZ.rButton cmdFixReg 
            Height          =   375
            Left            =   1680
            TabIndex        =   69
            Top             =   4800
            Width           =   1695
            _extentx        =   2990
            _extenty        =   661
            buttonstyle     =   7
            buttonstylecolors=   3
            backcolor       =   14211288
            backcolorpressed=   15715986
            backcolorhover  =   16243621
            bordercolor     =   11907757
            bordercolorpressed=   6045981
            bordercolorhover=   11632444
            caption         =   "Fix check"
            enabled         =   0   'False
            font            =   "FormMain.frx":5794
         End
      End
      Begin VB.PictureBox Picscan 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   5295
         Index           =   5
         Left            =   240
         ScaleHeight     =   5295
         ScaleWidth      =   7935
         TabIndex        =   23
         Top             =   1200
         Width           =   7935
         Begin WANIEZ.ucListView lvInfo 
            Height          =   4335
            Left            =   0
            TabIndex        =   24
            Top             =   360
            Width           =   7905
            _extentx        =   15055
            _extenty        =   6588
            styleex         =   33
            showsort        =   -1  'True
         End
         Begin WANIEZ.rButton cmdExplore 
            Height          =   375
            Left            =   0
            TabIndex        =   70
            Top             =   4800
            Width           =   1695
            _extentx        =   2990
            _extenty        =   661
            buttonstyle     =   7
            buttonstylecolors=   3
            captionalignment=   4
            backcolor       =   14211288
            backcolorpressed=   15715986
            backcolorhover  =   16243621
            bordercolor     =   11907757
            bordercolorpressed=   6045981
            bordercolorhover=   11632444
            caption         =   "Explorer"
            enabled         =   0   'False
            font            =   "FormMain.frx":57BC
         End
         Begin WANIEZ.rButton cmdProperties 
            Height          =   375
            Left            =   1680
            TabIndex        =   71
            Top             =   4800
            Width           =   1695
            _extentx        =   2990
            _extenty        =   661
            buttonstyle     =   7
            buttonstylecolors=   3
            captionalignment=   4
            backcolor       =   14211288
            backcolorpressed=   15715986
            backcolorhover  =   16243621
            bordercolor     =   11907757
            bordercolorpressed=   6045981
            bordercolorhover=   11632444
            caption         =   "Propertis"
            enabled         =   0   'False
            font            =   "FormMain.frx":57E4
         End
      End
      Begin WANIEZ.Tab Tabscan 
         Height          =   5895
         Left            =   120
         TabIndex        =   40
         Top             =   720
         Width           =   8175
         _extentx        =   14420
         _extenty        =   10398
         backcolor       =   16777215
         closebutton     =   0   'False
         blurforecolor   =   0
         activeforecolor =   0
         picture         =   "FormMain.frx":580C
         alltabsforecolor=   -2147483630
         fontname        =   "Tahoma"
      End
      Begin VB.Label Lbtitle 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Scan Computer"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   0
         TabIndex        =   87
         Top             =   120
         Width           =   8415
      End
   End
   Begin WANIEZ.rtp_mode rtp_mode1 
      Index           =   0
      Left            =   6600
      Top             =   0
      _extentx        =   1720
      _extenty        =   661
   End
   Begin VB.PictureBox PBRTP 
      Height          =   1575
      Left            =   12240
      ScaleHeight     =   1515
      ScaleWidth      =   2475
      TabIndex        =   2
      Top             =   720
      Visible         =   0   'False
      Width           =   2535
      Begin VB.PictureBox picBufferw 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   135
         Left            =   480
         ScaleHeight     =   105
         ScaleWidth      =   225
         TabIndex        =   86
         Top             =   0
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.PictureBox picTmpIcon 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   135
         Left            =   0
         ScaleHeight     =   7
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   15
         TabIndex        =   85
         Top             =   0
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.PictureBox picBuffer 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   240
         ScaleHeight     =   165
         ScaleWidth      =   195
         TabIndex        =   84
         Top             =   0
         Visible         =   0   'False
         Width           =   225
      End
      Begin VB.Timer tmrSPEED 
         Enabled         =   0   'False
         Interval        =   600
         Left            =   0
         Top             =   1080
      End
      Begin VB.Timer TimerIcon 
         Interval        =   6000
         Left            =   720
         Top             =   1080
      End
      Begin VB.Timer tmUpdate 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   1440
         Top             =   1080
      End
      Begin VB.Timer tmFlash 
         Enabled         =   0   'False
         Interval        =   4000
         Left            =   1800
         Top             =   1080
      End
      Begin VB.Timer tmAwal 
         Enabled         =   0   'False
         Interval        =   500
         Left            =   360
         Top             =   1080
      End
      Begin VB.Timer tmTime 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   1080
         Top             =   1080
      End
      Begin VB.PictureBox Pic16 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   240
         Left            =   2160
         Picture         =   "FormMain.frx":582A
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   21
         Top             =   0
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.PictureBox Pic15 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   240
         Left            =   2160
         Picture         =   "FormMain.frx":5DB4
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   20
         Top             =   240
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.PictureBox pic9 
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         Height          =   240
         Left            =   1440
         Picture         =   "FormMain.frx":633E
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   19
         Top             =   720
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.PictureBox pic10 
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         Height          =   240
         Left            =   1440
         Picture         =   "FormMain.frx":68C8
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   18
         Top             =   480
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.PictureBox pic1 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   1200
         Picture         =   "FormMain.frx":6E52
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   17
         Top             =   240
         Width           =   255
      End
      Begin VB.PictureBox pic13 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   1920
         Picture         =   "FormMain.frx":7194
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   16
         Top             =   0
         Width           =   255
      End
      Begin VB.PictureBox pic5 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   1200
         Picture         =   "FormMain.frx":738C
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   15
         Top             =   0
         Width           =   255
      End
      Begin VB.PictureBox pic12 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   1680
         Picture         =   "FormMain.frx":76CE
         ScaleHeight     =   17
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   17
         TabIndex        =   14
         Top             =   480
         Width           =   255
      End
      Begin VB.PictureBox pic11 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   1920
         Picture         =   "FormMain.frx":7C58
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   13
         Top             =   480
         Width           =   255
      End
      Begin VB.PictureBox pic8 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   1200
         Picture         =   "FormMain.frx":81E2
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   12
         Top             =   720
         Width           =   255
      End
      Begin VB.PictureBox pic7 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   1200
         Picture         =   "FormMain.frx":876C
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   11
         Top             =   480
         Width           =   255
      End
      Begin VB.PictureBox pic6 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   1440
         Picture         =   "FormMain.frx":8CF6
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   10
         Top             =   240
         Width           =   255
      End
      Begin VB.PictureBox pic4 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   1440
         Picture         =   "FormMain.frx":903A
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   9
         Top             =   0
         Width           =   255
      End
      Begin VB.PictureBox pic3 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   1680
         Picture         =   "FormMain.frx":92B2
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   8
         Top             =   240
         Width           =   255
      End
      Begin VB.PictureBox pic2 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   1680
         Picture         =   "FormMain.frx":95F4
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   7
         Top             =   0
         Width           =   255
      End
      Begin VB.PictureBox pic14 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   2160
         Picture         =   "FormMain.frx":9936
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   6
         Top             =   480
         Width           =   255
      End
      Begin VB.PictureBox picFolHid 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   1680
         Picture         =   "FormMain.frx":9C50
         ScaleHeight     =   17
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   17
         TabIndex        =   5
         Top             =   720
         Width           =   255
      End
      Begin VB.PictureBox picFileHid 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   1920
         Picture         =   "FormMain.frx":9F92
         ScaleHeight     =   17
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   17
         TabIndex        =   4
         Top             =   720
         Width           =   255
      End
      Begin VB.PictureBox picCaution 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   1920
         Picture         =   "FormMain.frx":A2D4
         ScaleHeight     =   17
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   17
         TabIndex        =   3
         Top             =   240
         Width           =   255
      End
      Begin WANIEZ.Downloader Downloader1 
         Left            =   0
         Top             =   120
         _extentx        =   847
         _extenty        =   847
      End
      Begin WANIEZ.UniDialog UniDialog1 
         Left            =   0
         Top             =   600
         _extentx        =   847
         _extenty        =   847
         fileflags       =   2621444
         folderflags     =   323
         filecustomfilter=   "FormMain.frx":A616
         filedefaultextension=   "FormMain.frx":A636
         filefilter      =   "FormMain.frx":A656
         fileopentitle   =   "FormMain.frx":A69E
         filesavetitle   =   "FormMain.frx":A6D6
         foldermessage   =   "FormMain.frx":A70E
      End
      Begin VB.Image picmti 
         Height          =   480
         Left            =   600
         Picture         =   "FormMain.frx":A74A
         Top             =   600
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image gmb 
         Height          =   480
         Left            =   720
         Picture         =   "FormMain.frx":FA0C
         Top             =   120
         Visible         =   0   'False
         Width           =   480
      End
   End
   Begin VB.TextBox pathRTP 
      Height          =   315
      Left            =   12240
      TabIndex        =   1
      Top             =   2400
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.TextBox UniLabel1 
      Height          =   285
      Left            =   12240
      TabIndex        =   0
      Top             =   2760
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Label LbFooter 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "About"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   210
      Index           =   0
      Left            =   11160
      TabIndex        =   229
      Top             =   7680
      Width           =   435
   End
   Begin VB.Label LbFooter 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Setting"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   210
      Index           =   1
      Left            =   10440
      TabIndex        =   228
      Top             =   7680
      Width           =   495
   End
   Begin VB.Image ImgClick 
      Height          =   615
      Index           =   4
      Left            =   0
      Top             =   2610
      Width           =   3255
   End
   Begin VB.Image ImgClick 
      Height          =   615
      Index           =   3
      Left            =   0
      Top             =   3885
      Width           =   3255
   End
   Begin VB.Image ImgClick 
      Height          =   615
      Index           =   2
      Left            =   0
      Top             =   3255
      Width           =   3255
   End
   Begin VB.Image ImgClick 
      Height          =   615
      Index           =   1
      Left            =   0
      Top             =   1965
      Width           =   3255
   End
   Begin VB.Image ImgClick 
      Height          =   615
      Index           =   0
      Left            =   0
      Top             =   1320
      Width           =   3255
   End
   Begin VB.Label LAv 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Update"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   240
      Index           =   4
      Left            =   840
      TabIndex        =   227
      Top             =   2805
      Width           =   615
   End
   Begin VB.Label LAv 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Virus Chest"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   240
      Index           =   3
      Left            =   840
      TabIndex        =   226
      Top             =   4095
      Width           =   1020
   End
   Begin VB.Label LAv 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Additional Protection"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   240
      Index           =   2
      Left            =   840
      TabIndex        =   225
      Top             =   3480
      Width           =   1800
   End
   Begin VB.Label LAv 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Scan Computer"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   240
      Index           =   1
      Left            =   840
      TabIndex        =   224
      Top             =   2175
      Width           =   1350
   End
   Begin VB.Label LAv 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Summary"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   240
      Index           =   0
      Left            =   840
      TabIndex        =   223
      Top             =   1560
      Width           =   840
   End
   Begin VB.Label LbVersi 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1.5"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   3480
      TabIndex        =   222
      Top             =   120
      Width           =   330
   End
   Begin VB.Image ImgNotifasi 
      Height          =   8055
      Left            =   0
      Top             =   0
      Width           =   11895
   End
   Begin VB.Label LbSlogan 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Safeguard and Protect your Computer"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   210
      Left            =   8520
      TabIndex        =   221
      Top             =   240
      Width           =   3135
   End
   Begin VB.Line MainLine 
      BorderColor     =   &H00C0C0C0&
      Index           =   1
      X1              =   11040
      X2              =   11040
      Y1              =   7680
      Y2              =   7920
   End
   Begin VB.Label LbCopyright 
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
      ForeColor       =   &H00808080&
      Height          =   210
      Left            =   120
      TabIndex        =   220
      Top             =   7680
      Width           =   3375
   End
   Begin WANIEZ.ucAniGIF GifScan 
      Height          =   240
      Left            =   2760
      Top             =   2160
      Visible         =   0   'False
      Width           =   240
      _extentx        =   423
      _extenty        =   423
      gif             =   "FormMain.frx":14CCE
      enabled         =   0   'False
   End
   Begin VB.Label LbX 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   12840
      TabIndex        =   73
      Top             =   8400
      Width           =   165
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'multi thread
Private WithEvents Thread As IThread: Private WithEvents Thread2 As IThread
Attribute Thread.VB_VarHelpID = -1
Attribute Thread2.VB_VarHelpID = -1
'tread
Dim Working As Boolean: Dim File_terScan As Long: Dim File_toScan As Long
Dim File_terScanRTP As Long: Dim File_toScanRTP As Long: Dim Fol_toScan As Long

Dim iMax_ID2 As Integer
Dim encrypt As New clsEncryption 'Enkrip PassWord

Dim i As Long

Private aLagiJalan  As Boolean
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public wadahPath As String

'Shell Contac Menu
Dim WithEvents ContextMenu As wodShellMenu
Attribute ContextMenu.VB_VarHelpID = -1
Dim WithEvents ShellIE As SHDocVw.ShellWindows 'shell menu
Attribute ShellIE.VB_VarHelpID = -1

Dim isCompatch As Boolean
Dim LewatExit  As Boolean
Dim SudahJalan As Boolean ' buffer ajh
Dim sebelumnya As String

Dim IndekPluginTerplih As Long ' buat buffer plugin aj karena gak support unic ListBoxnya
Dim Detik        As Long
Dim Detik2       As Long
Dim Menit        As Long
Dim Jam          As Integer
Dim LastTime As Long
Dim NewTime As Long
Dim AllTime As Long
Dim TimeSpeedCount As Long
Dim SpeedRate As Long
Dim MaxSpeed As Long

Dim StatScan      As String ' status scan
'unDialog
Dim UniDialogPath As String: Dim UniDialogFile As String

Dim LastPathRClick    As String ' path terakhir yang dikil kanan dari LV2 hasil scan
Dim ListViewClicked   As String ' status nama listview yang di klik kanan
'Path
Dim PathsCANtURBO As String: Dim PathDariShellMenu As String: Dim PathDariShellMenuDB As String: Dim PathLainArr(2)   As String
'Dim PathHooking As String:
'Dim byteBuffer(1 To 255) As Byte
Dim NewFDMasuk        As String
Dim BufferUpdate      As Long ' penanda sampai mana updatenya
Dim second As Integer
Dim strPath As String
'Dim sebelumnya As String
Public rtpkerja As Boolean
Dim ThreadID As Long

'For usefull test... compile this example and open the exe in some debugger (like ADA, OLLY, etc). Debug this code before install the "Antidebugger"... then debug again after install the "Antidebugger"
Private ctile As New cDIBTile
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
Private Sub BtnAtur_Click()
Dim wadah As String
wadah = GetFilePath(App_FullPathW(False))
FrmSysTray.mnEPro_Click
    Shell_NotifyIcon NIM_DELETE, nID
    If UCase$(Left$(Command, 2)) <> "-K" Then frmMain.terapkanIcon
    GoTo simpan
simpan:
   Call LoadConfig 'Environ$("windir") & "\MO.LOG"
   Call cekstatus
End Sub
Public Function cekstatus() 'Check Status
Dim X As String
If FrmConfig.ck8.Value = 1 Then
'imgStat(1).Picture = FrmRTP.imgAman(0)
BtnAtur.Caption = a_bahasa(30)
'imgStat(1).Picture = FrmRTP.imgAman(1)
Else
BtnAtur.Caption = a_bahasa(29)
End If
End Function
Private Sub BtnFix_Click() 'FIX Btn
If lvMalware.ListItems.Count > 0 Then
Call cmdFixMalwareAll_Click
End If
If lvRegistry.ListItems.Count > 0 Then
Call cmdFixRegAll_Click
End If
If lvHidden.ListItems.Count > 0 Then
Call cmdFixHiddenall_Click
End If
BtnFix.Enabled = False
End Sub
Private Sub BTNlOG_Click() 'Log Reslut Scan
On Error Resume Next
    Shell "Notepad.exe " & App.path & "\LOG.scan", vbNormalFocus
End Sub
Private Sub BtnRefres_Click() 'Refles Dtree
Call BuilDirTree
End Sub
Private Sub BtnResult_Click() 'Btn Resluh Scan
If BtnResult.Caption = a_bahasa(25) Then
If lvMalware.ListItems.Count > 0 Then
Tabscan.ActiveTab = 3: Tabscan_Click (3)
ElseIf lvRegistry.ListItems.Count > 0 Then
Tabscan.ActiveTab = 5: Tabscan_Click (5)
ElseIf lvHidden.ListItems.Count > 0 Then
Tabscan.ActiveTab = 4: Tabscan_Click (4)
Else
BtnResult.Enabled = False
End If
Else
Call cmdStartScan_Click
End If
End Sub
Private Sub ckScan_Click(Index As Integer)
Select Case Index
Case 0 'cek malware
If lvMalware.ListItems.Count <> 0 Then
        If ckScan(0).Value = Checked Then
            For i = 1 To lvMalware.ListItems.Count
                lvMalware.ListItems.Item(i).Checked = True
            Next
        Else
            For i = 1 To lvMalware.ListItems.Count
                lvMalware.ListItems.Item(i).Checked = False
            Next
        End If
    End If

Case 1 'cek hidden
If lvHidden.ListItems.Count <> 0 Then
        If ckScan(1).Value = Checked Then
            For i = 1 To lvHidden.ListItems.Count
                lvHidden.ListItems.Item(i).Checked = True
            Next
        Else
            For i = 1 To lvHidden.ListItems.Count
                lvHidden.ListItems.Item(i).Checked = False
            Next
        End If
    End If
Case 2 'cek reg
        If lvRegistry.ListItems.Count <> 0 Then
        If ckScan(2).Value = Checked Then
            For i = 1 To lvRegistry.ListItems.Count
                lvRegistry.ListItems.Item(i).Checked = True
            Next
        Else
            For i = 1 To lvRegistry.ListItems.Count
                lvRegistry.ListItems.Item(i).Checked = False
            Next
        End If
    End If
    End Select
End Sub
Private Sub ckuserdb_Click()
If lvm31.ListItems.Count <> 0 Then
        If ckuserdb.Value = Checked Then
            For i = 1 To lvm31.ListItems.Count
                lvm31.ListItems.Item(i).Checked = True
            Next
        Else
            For i = 1 To lvm31.ListItems.Count
                lvm31.ListItems.Item(i).Checked = False
            Next
        End If
    End If
End Sub
Private Sub cmdAddExcFile_Click() 'AddExcFile
UniDialog1.ShowOpen
If ValidFile(UniDialogFile) = True Then
   ReBuildFileException UniDialogFile, GetFilePath(App_FullPathW(False)) & "\File.lst", lstExceptFile
End If
End Sub
Private Sub cmdAddExcFolder_Click() 'AddExcFolder
UniDialog1.ShowFolder
If Len(UniDialogPath) > 2 Then
   ReBuildPathException UniDialogPath, GetFilePath(App_FullPathW(False)) & "\Path.lst", lstExceptFolder
End If
End Sub
Private Sub cmdAddVirus_Click() 'Add Virus USER
On Error Resume Next
If txtVirusPath.Text = "" Then
   MsgBox i_bahasa(7), vbExclamation
   Exit Sub
End If

If AddVirusTemp(UniDialogFile, txtVirusName.Text) = True Then
    
    If Replace(txtVirusName.Text, " ", "") <> "" Then

      AddInfo txtVirusPath, txtVirusName
      txtVirusPath.Text = "": txtVirusName.Text = ""
      cmdAddVirus.Enabled = False: cmdCancel.Enabled = False

     If ValidFile(App.path & "\WanUDB.dll") = True Then ReadUDB lvm31
     Call Init_Dtabase
    Else
        MsgBox "Please Give your virus name on Text Box of Virus Name !", vbInformation, "change virus name"
    End If
   ' End If
ElseIf AddVirusTemp(txtVirusPath, txtVirusName.Text) = True Then 'Ukuran minimal dengan M31 Pattern adalah 610 an Byte
    
    If Replace(txtVirusName.Text, " ", "") <> "" Then
      AddInfo txtVirusPath, txtVirusName
      Call Init_Dtabase
      'MsgBox "Penambahan database sukses", vbInformation
      txtVirusPath.Text = "": txtVirusName.Text = ""
      TxtNama.Text = "": TxtSize.Text = "": txtType.Text = "": txtCecksum.Text = "": txtCeck.Text = ""
      PicAnalisis.Visible = False: cmdAddVirus.Enabled = False: cmdCancel.Enabled = False
       
     If ValidFile(App.path & "\WanUDB.dll") = True Then ReadUDB lvm31
    Else
        MsgBox "Please Give your virus name on Text Box of Virus Name !", vbExclamation, "change virus name"
    End If
End If
PathDariShellMenuDB = ""
End Sub
Private Sub AddInfo(txtpath As TextBox, TxtVname As TextBox) 'Add Info Virus User
Dim strFile As String
Dim M31_Hash As String
Dim V_Name As String
    If IsFile(App.path & "\WanUDB.dll") = True Then strFile = OpenFileInTeks(App.path & "\WanUDB.dll") ' Buka File User.DAT
    M31_Hash = FrmRTP.Text1 'MeiPattern(TxtPath.Text, 202) ' Ingat ya HR BOX harus 202 [ Standar disini ]
    V_Name = TxtVname.Text
Open App.path & "\WanUDB.dll" For Output As #3 ' Tulis ke User.DAT
    Write #3, strFile & ";" & M31_Hash & "|" & V_Name ' Sesuai Aturan penulisan
Close #3
End Sub
Private Sub cmdBrowse_Click() 'Brouse File
Dim fileku As String
Dim TmpHGlobal As Long
Dim RetPE        As Long
Dim Ukuran       As String
Dim sCeksum     As String
Dim nDB         As Long
Dim nJumVirus   As Long
Dim MyHandle    As Long
Dim FalsCek     As String

    UniDialog1.ShowOpen

    If ValidFile(UniDialogFile) = True Then
    txtVirusPath.Text = UniDialogFile
   wadahPath = UniDialogFile
    If IsFileProtectedBySystem(txtVirusPath.Text) = True Then GoTo metu
    UniDialogFile = "": FrmRTP.Text1.Text = ""
    AnalisisPE wadahPath
    End If
    Exit Sub

metu:
   MsgBox f_bahasa(20), vbCritical
Call cmdCancel_Click
End Sub
Private Sub cmdCancel_Click() 'Batal Tambahkan
    txtVirusName.Text = "": UniDialogFile = "": txtVirusPath = "": FrmRTP.Text1.Text = ""
    cmdAddVirus.Enabled = False: cmdCancel.Enabled = False: PicAnalisis.Visible = False
    TxtNama.Text = "": TxtSize.Text = "": txtType.Text = "": txtCecksum.Text = "": txtCeck.Text = ""
End Sub
Private Sub CmdQuarAll_Click()
    CmdQuarAll.Enabled = False: cmdFixMalware.Enabled = False: cmdFixMalwareAll.Enabled = False
    Quar_Malware lvMalware, BY_ALL, 16
    cmdFixMalwareAll.Enabled = True: cmdFixMalware.Enabled = True: CmdQuarAll.Enabled = True
    AutoLst frmMain.lvMalware
    Call READ_DATA_JAIL(FolderJail)
End Sub
Private Sub CmdRestore_Click()
Dim PrisName        As String
Dim IndekTerpilih   As Long

IndekTerpilih = CariIndekItemTerpilih(frmMain.lvJail)
If IndekTerpilih > 0 Then
   If MsgBox(i_bahasa(22) & " ?", vbExclamation + vbYesNo) = vbYes Then
      ReleasePrisoner frmMain.lvJail.ListItems.Item(IndekTerpilih).SubItem(2).Text, frmMain.lvJail.ListItems.Item(IndekTerpilih).SubItem(4).Text, frmMain.lvJail
   End If
End If
End Sub
Private Sub CmdRestoreto_Click()
Dim PrisName        As String
Dim IndekTerpilih   As Long
Dim PrisonerFName  As String

IndekTerpilih = CariIndekItemTerpilih(frmMain.lvJail)
frmMain.UniDialog1.ShowFolder
    If PathIsDirectory(StrPtr(UniDialogPath)) <> 0 Then
       If IndekTerpilih > 0 Then
          PrisonerFName = GetFileName(frmMain.lvJail.ListItems.Item(IndekTerpilih).SubItem(2).Text)
          If MsgBox(i_bahasa(22) & " here ?", vbExclamation + vbYesNo) = vbYes Then
             ReleasePrisoner UniDialogPath & "\" & PrisonerFName, frmMain.lvJail.ListItems.Item(IndekTerpilih).SubItem(4).Text, frmMain.lvJail
          End If
       End If
    End If
End Sub
Private Sub CmnKarantine_Click() 'Karantine
On Error GoTo gagal
UniDialog1.ShowOpen
FrmRTP.Text3 = UniDialogFile
    If ValidFile(UniDialogFile) = True Then
        If IsFileProtectedBySystem(FrmRTP.Text3.Text) = False Then
        JailFile FrmRTP.Text3, FolderJail, "Added by User"
         Call READ_DATA_JAIL(FolderJail)
        If ValidFile(FrmRTP.Text3.Text) = True Then GoTo gagal
        MsgBox f_bahasa(16) + "-" + FrmRTP.Text3.Text
        Else
        MsgBox f_bahasa(20) + "-" + FrmRTP.Text3.Text, vbCritical
        End If
    End If
    Exit Sub
gagal:
    MsgBox f_bahasa(18) + "-" + FrmRTP.Text3.Text, vbExclamation
End Sub

Private Sub ImgNotifasi_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
LAv(0).ForeColor = &H808080: LAv(1).ForeColor = &H808080: LAv(2).ForeColor = &H808080: LAv(3).ForeColor = &H808080: LAv(4).ForeColor = &H808080
LbFooter(0).ForeColor = &H808080: LbFooter(1).ForeColor = &H808080
End Sub
Private Sub LbFooter_Click(Index As Integer)
Select Case Index
Case 0
FrmAbout.Show
Case 1
FrmConfig.Show
End Select
End Sub
Private Sub LbFooter_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, y As Single)
 Dim i As Byte
 For i = LbFooter.LBound To LbFooter.UBound
        If i <> Index Then
            LbFooter(i).ForeColor = &H808080
        End If
    Next
    With LbFooter(Index)
      .ForeColor = &H0&
    End With
End Sub
Private Sub ImgClick_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, y As Single)
 Dim i As Byte
 For i = LAv.LBound To LAv.UBound
        If i <> Index Then
            LAv(i).ForeColor = &H808080
        End If
    Next
    With LAv(Index)
      .ForeColor = &H0&
    End With
End Sub
Private Sub ImgClick_Click(Index As Integer)
Dim A As String
Dim B As String
Dim spWd As String
A = GetSTRINGValue(&H80000001, "Software\Wan'iez\", "password")
Select Case Index
        Case 0 'sumary
        Call summary
        Case 1 'computer scan
        Call scancomputer
        Tabscan.ActiveTab = 1: Tabscan_Click (1)
        Case 2 'addition protected
If A <> "" And FrmConfig.CkPass.Value <> 0 And FrmConfig.CkAdditional.Value <> 0 Then
B = encrypt.Cryption(A, "WANPASS", False)
FrmPassword.ShowMe spWd
If spWd <> "" Then
If spWd = B Then
        Call additionalprotec
        TabTools.ActiveTab = 1: TabTools_Click (1)
Else
    MsgBox "password you entered is not corect!", vbCritical, i_bahasa(26)
    Call ImgClick_Click(0)
Exit Sub
End If
Else
    Call ImgClick_Click(0)
End If
Else
        Call additionalprotec
        TabTools.ActiveTab = 1: TabTools_Click (1)
End If
        Case 3 'virus chets
If A <> "" And FrmConfig.CkPass.Value <> 0 And FrmConfig.CkQuar.Value <> 0 Then
B = encrypt.Cryption(A, "WANPASS", False)
FrmPassword.ShowMe spWd
If spWd <> "" Then
If spWd = B Then
        Call quaranntine
Else
    MsgBox "password you entered is not corect!", vbCritical, i_bahasa(26)
    Call ImgClick_Click(0)
Exit Sub
End If
Else
    Call ImgClick_Click(0)
End If
Else
        Call quaranntine
End If
        Case 4 'update
        Call update
End Select
End Sub

Private Sub LbX_Click()
FrmSysTray.mnCScan_Click
End Sub
Private Sub lstKunci_ColumnClick(ByVal oColumn As cColumn)
oColumn.Sort
End Sub
Private Sub lstKunci_ItemClick(ByVal oItem As cListItem, ByVal iButton As evbComCtlMouseButton)
    If iButton = vbccMouseRButton Then FrmSysTray.mnsubmit.Enabled = True: FrmSysTray.PopupMenu FrmSysTray.mnDriveLock, 0, , , FrmSysTray.mnLock
End Sub
Private Sub lvM31_ColumnClick(ByVal oColumn As cColumn)
oColumn.Sort
End Sub
Private Sub cmdExecutePlug_Click()
If IndekPluginTerplih >= 0 Then
    RunPlugin IndekPluginTerplih
End If
End Sub
Private Sub cmdExplore_Click()
Dim IndekTerpilih   As Long
Dim sPathFile       As String
IndekTerpilih = CariIndekItemTerpilih(lvInfo)
If IndekTerpilih > 0 Then
   sPathFile = lvInfo.ListItems.Item(IndekTerpilih).SubItem(2).Text
   Shell "Explorer.exe /e," & GetFilePath(sPathFile), vbNormalFocus
End If
End Sub
Private Sub cmdFixHidden_Click() 'Fix Select Hiden
Call CkFixHidden
End Sub
Private Sub cmdFixHiddenall_Click() 'Fix All Hiden
Call AllFixHidden
End Sub
Private Sub cmdFixMalware_Click() 'Fix Select Virus
Call CkFixMalware
End Sub
Private Sub cmdFixMalwareAll_Click() 'FIX all Virus
Call AllFixMalware
End Sub
Private Sub cmdFixReg_Click() 'Fix Select Reg
Call CkFixReg
End Sub
Private Sub cmdFixRegAll_Click() 'Fix All Reg
Call AllFixReg
End Sub
Private Sub cmdProperties_Click() 'Propertis
Dim IndekTerpilih   As Long
IndekTerpilih = CariIndekItemTerpilih(lvInfo)
If IndekTerpilih > 0 Then
   ShowProperties lvInfo.ListItems.Item(IndekTerpilih).SubItem(2).Text, Me.hwnd
End If
End Sub
Private Sub cmdRemExcFile_Click()
    HapusFile GetFilePath(App_FullPathW(False)) & "\File.lst"
    ReadExceptFile GetFilePath(App_FullPathW(False)) & "\File.lst", lstExceptFile
End Sub
Private Sub cmdRemExcFile1_Click()
    RemoveExceptionByIndek lstExceptFile.ListIndex, FILE_EXC
    JumFileExcep = ReadExceptFile(GetFilePath(App_FullPathW(False)) & "\File.lst", lstExceptFile)
End Sub
Private Sub cmdRemovePath_Click()
    HapusFile GetFilePath(App_FullPathW(False)) & "\Path.lst"
    ReadExceptPath GetFilePath(App_FullPathW(False)) & "\Path.lst", lstExceptFolder
End Sub
Private Sub cmdRemovePath1_Click()
RemoveExceptionByIndek lstExceptFolder.ListIndex, PATH_EXC
JumPathExcep = ReadExceptPath(GetFilePath(App_FullPathW(False)) & "\Path.lst", lstExceptFolder)
End Sub
Private Sub cmdStartScan_Click()
Dim lstCek      As Collection
Dim iCount      As Long
If ValidFile(App.path & "\LOG.scan") = True Then HapusFile App.path & "\LOG.scan"
Tabscan.ActiveTab = 2: Tabscan_Click (2)

Call awal 'log

StatScan = d_bahasa(17)
Set lstCek = New Collection
DirTree.OutPutPath lstCek
If cmdStartScan.Caption = a_bahasa(5) Then

   If MaulanjutScan(lvMalware) = False Then Exit Sub
   GifScan.Enabled = True
GifScan.Visible = True
   Call ResetObjek
   If NewFDMasuk <> "" Then GoTo LBL_SCAN_FD Else GoTo END_LBL_FD
LBL_SCAN_FD:    ' scan FD masuk
  'mulai buffer path yang akan di scan
  BufferPath NewFDMasuk, True
  Call BtnAtur_Click
  cmdStartScan.Caption = a_bahasa(7): BtnResult.Caption = a_bahasa(7)
  ' init Progress Bar
  PB1.Value = 0
  PB1.Max = FileToScan
  
  lbStatus.Caption = d_bahasa(15)
  KumpulkanFile NewFDMasuk, lbObject, True, True
  GoTo LBL_LOMPATAN_FD
END_LBL_FD:

   If WinNode = True Then PathLainArr(0) = GetSpecFolder(WINDOWS_DIR) Else PathLainArr(0) = ""
   If DocNode = True Then PathLainArr(1) = GetSpecFolder(USER_DOC) Else PathLainArr(1) = ""
   If ProgNode = True Then PathLainArr(2) = GetSpecFolder(PROGRAM_FILE) Else PathLainArr(2) = ""
   'Sleep 300
   cmdStartScan.Caption = a_bahasa(6): BtnResult.Caption = a_bahasa(6)
   'mulai buffer path yang akan di scan
   For iCount = 1 To lstCek.Count
      If WithBuffer = False Then Exit For
      BufferPath lstCek(iCount), True
   Next
   iCount = 0

   For iCount = 0 To 2
      If WithBuffer = False Then Exit For
      If Len(PathLainArr(iCount)) > 0 Then BufferPath PathLainArr(iCount), True
   Next
   'buffer buat shell menu (jika folder)
   If Len(PathDariShellMenu) > 0 And ValidFile(PathDariShellMenu) = False Then BufferPath PathDariShellMenu, True

   cmdStartScan.Caption = a_bahasa(7): BtnResult.Caption = a_bahasa(7)
      
   If RegNode = True Then
      lbStatus.Caption = d_bahasa(10)
      If FrmConfig.ck3.Value = 1 Then 'dengan auto detect useles value
         ScanRegistry lbObject, False, True
      Else
         ScanRegistry lbObject, False, False
      End If
   End If
   
   If ProsesNode = True Then 'scan proses + service
      lbStatus.Caption = d_bahasa(11)
      Call ScanService(lbObject, True)
      
      lbStatus.Caption = d_bahasa(12)
      Call ScanProses(False, lbObject)
   
      lbStatus.Caption = j_bahasa(0) ' module
      Call ScanProses(True, lbObject)

   End If
   
   If StartUpNode = True Then ' scan startup
      lbStatus.Caption = d_bahasa(13)
      ScanRegStartup lbObject, True
   End If
   
   If UCase$(Left$(Command, 2)) <> "-K" Then lbStatus.Caption = d_bahasa(14) ' root drive
   If UCase$(Left$(Command, 2)) <> "-K" Then ScanRootDrive lbObject

   ' init Progress Bar
   PB1.Value = 0
   PB1.Max = FileToScan
   ' reset
   iCount = 0
   ' Mulai pindai
   lbStatus.Caption = d_bahasa(15) '& " [ " & FolToScan & " " & j_bahasa(39) & ", " & FileToScan & " " & j_bahasa(38) & " ]"
     
   'path lain dulu
   For iCount = 0 To 2
       If BERHENTI = True Then Exit For
       If Len(PathLainArr(iCount)) > 0 Then KumpulkanFile PathLainArr(iCount), lbObject, True, True
   Next
      'scan dari path shellmenu
   'Sleep 300
   If Len(PathDariShellMenu) > 0 Then
      If ValidFile(PathDariShellMenu) = False Then ' jika folder
         KumpulkanFile PathDariShellMenu, lbObject, True, True
      Else ' jika file
         CocokanDataBase PathDariShellMenu
         FileCheck = FileCheck + 1
         FileFound = FileFound + 1
         lbFileCheck.Caption = Right$(FileCheck, 8)
        ' frmScanWith.lbFileCheck.Caption = ": " & Right$("00000000" & FileCheck, 8)
         lbFileFound.Caption = Right$(FileFound, 8)
      End If
   End If
   tmrSPEED.Enabled = True
   iCount = 1
  FileToScan = 0
  FolToScan = 0
   For iCount = 1 To lstCek.Count
       If BERHENTI = True Then Exit For
       
      ' sCANFile lstCek(iCount) '
       KumpulkanFile lstCek(iCount), lbObject, True, True
   Next
LBL_LOMPATAN_FD:
   
   tmrSPEED.Enabled = False: tmTime.Enabled = False
   cmdStartScan.Caption = a_bahasa(5): BtnResult.Caption = a_bahasa(5)
      
   If StatScan = d_bahasa(16) And WithBuffer = True Then StatScan = d_bahasa(16) & " !"
   lbStatus.Caption = StatScan
         GifScan.Enabled = False
      GifScan.Visible = False
   BERHENTI = True
   Me.Show
           
   Dim jumVir As Long, Msg As String: jumVir = lvMalware.ListItems.Count + lvHidden.ListItems.Count + nErrorReg
    Msg = i_bahasa(33) & IIf(jumVir > 10, i_bahasa(35), IIf(jumVir > 0, i_bahasa(34), i_bahasa(36)))
    If jumVir > 0 Then
        MsgBox Msg & vbCrLf & i_bahasa(37) & IIf(lvMalware.ListItems.Count > 0, vbCrLf & "- " & lvMalware.ListItems.Count & " " & i_bahasa(40), "") & _
        IIf(nErrorReg <> 0, vbCrLf & "- " & nErrorReg & " " & i_bahasa(39), "") & IIf(lvHidden.ListItems.Count > 0, vbCrLf & "- " & lvHidden.ListItems.Count & " " & i_bahasa(41), ""), vbExclamation, i_bahasa(38)
      GoTo xer
    End If
    
   MsgBox i_bahasa(42), vbInformation, i_bahasa(38)
   BtnResult.Caption = a_bahasa(25): BtnResult.Enabled = False
      GifScan.Enabled = False
      GifScan.Visible = False

xer:
      If lvMalware.ListItems.Count > 0 Or lvRegistry.ListItems.Count > 0 Or lvHidden.ListItems.Count > 0 Then
      Call ImgClick_Click(1)
      Tabscan.ActiveTab = 2: Tabscan_Click (2)
      BtnResult.Caption = a_bahasa(25): BtnFix.Enabled = True
      End If
Call akir
    Call ReBack ' Aktifkan yang peru diaktifkan
      If FrmConfig.ck7.Value = 1 Then tmFlash.Enabled = True
     ' End If
ElseIf cmdStartScan.Caption = a_bahasa(6) Then
   Call StopKumpulkan
   WithBuffer = False
         GifScan.Enabled = False
      GifScan.Visible = False
Else

   Call StopKumpulkan
   BERHENTI = True
   StatScan = d_bahasa(16): cmdStartScan.Caption = a_bahasa(5)
End If
End Sub
Public Sub Form_Load()
TabTools.AddTabs "Process Manager", "Virus By User", "Exception Scan", "Plugin", "System Editor", "Drive Lock"
Tabscan.AddTabs "Path Scan", "Current Report", "Virus", "Hidden", "Registry", "Information"
Call Protect
Call LoadDll
If App.PrevInstance = True And UCase$(Left$(Command, 2)) <> "-K" Then
   MsgBox "Wan'iez AntiVirus sudah Aktif pada sistem Anda !", vbExclamation
   End
  End If
If UCase$(Left$(Command, 2)) <> "-K" Then ThreadID = InstallAntiDebugger

Set ContextMenu = New wodShellMenu
    Dim mItem As MenuItem
    Set mItem = ContextMenu.MenuItems.Add("Scan With Wan'iez")
    mItem.BitmapFile = App.path + "\icon.bmp"
    ContextMenu.Enabled = True
    Me.AutoRedraw = True
    
  If UCase$(Left$(Command, 2)) <> "-K" Then Load FrmSysTray
 ' End If
 'Khusus Call
  Call InitAplikasi: Call LoadKunci: Call ReduceMemory: Call LoadGUI: Call cekstatus: Call ImgClick_Click(0): Call ReBack
  FrmAbout.lbVirus.Caption = ": " & CStr(JumVirus) + JumlahVirusINT & " Virus , " & lvm31.ListItems.Count & " Malware User + Heuristic": DoEvents
  LbVirusVersion.Caption = ": " & CStr(JumVirus) + JumlahVirusINT & " Virus , " & lvm31.ListItems.Count & " Malware User + Heuristic": DoEvents

GoTo BAR
lanjut:
BAR:
If ValidFile(App.path & "\WanUDB.dll") = True Then ReadUDB lvm31

Select Case Left$(Command, 2)
  Case "-A" '
   Me.Hide
  End Select
   
   BERHENTI = True
    ' Call bunyi
lbObject.Caption = ""

If UCase$(Left$(Command, 2)) <> "-K" Then
Call CompactObject
End If

tmAwal.Enabled = True

SudahJalan = True ' informasikan sopware sudah jalan
IndekPluginTerplih = -1
keluar:
End Sub

Function Protect() 'Cegah Pemalsuan Antivirus
If App.Comments <> "" Then
'jika komentar diganti maka program ditutup
MsgBox j_bahasa(41) & vbCrLf & j_bahasa(42), vbOKOnly + vbExclamation
GoTo keluar
End If
'mencegah nama pembuat diganti
If App.CompanyName <> "Canvas Soft.™" Then
MsgBox j_bahasa(41) & vbCrLf & j_bahasa(42), vbOKOnly + vbExclamation
GoTo keluar
End If
'jika legalcopyright tidak sesuai
If App.LegalCopyright <> "Canvas Soft@2011-2013" Then
MsgBox j_bahasa(41) & vbCrLf & j_bahasa(42), vbOKOnly + vbExclamation
GoTo keluar
End If
'jika deskripsi tidak sesuai
If App.FileDescription <> "Wan'iez Antivirus" Then
MsgBox j_bahasa(41) & vbCrLf & j_bahasa(42), vbOKOnly + vbExclamation
GoTo keluar
End If
'jika legaltrademarks tidak sesuai
If App.LegalTrademarks <> "" Then
MsgBox j_bahasa(41) & vbCrLf & j_bahasa(42), vbOKOnly + vbExclamation
GoTo keluar
End If
Exit Function
keluar:
Unload FormSplash: Unload FrmRTP: Unload FrmAbout: Unload Me
End
End Function
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
    If Button = 1 Then DragForm Me
End Sub
Private Sub Form_Resize()
On Error Resume Next
Me.Caption = "Wan'iez Antivirus"
    If Me.WindowState = vbMinimized Then
       FrmSysTray.mnCScan.Caption = g_bahasa(1)
   End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
If LewatExit = False And Left$(Command, 2) <> "-K" Then
    Cancel = 1
    Shell_NotifyIcon NIM_DELETE, nID
  Call terapkanIcon

    Me.WindowState = vbMinimized
    Me.Hide

ElseIf BERHENTI = False Then
    Cancel = 1
    MsgBox i_bahasa(8), vbExclamation
Else

'For i = LBound(cAccess) To UBound(cAccess)
   ' cAccess(i).AllowAccess
'Next

        Me.Show
Shell_NotifyIcon NIM_DELETE, nID
Unload FrmRTP: Unload FormSplash: Unload FrmAbout: Unload FormSplash: Unload Me
    End
End If
End Sub
Private Sub Header_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
If Button = 1 Then DragForm Me
End Sub
Private Sub lstExceptFile_Click()
    lstExceptFile.ToolTipText = lstExceptFile.List(lstExceptFile.ListIndex)
End Sub
Private Sub lstExceptFolder_Click()
    lstExceptFolder.ToolTipText = lstExceptFolder.List(lstExceptFolder.ListIndex)
End Sub
Private Sub lstExceptReg_Click()
    lstExceptReg.ToolTipText = lstExceptReg.List(lstExceptReg.ListIndex)
End Sub
Private Sub lstModule_DblClick()
If Left$(lstModule.List(lstModule.ListIndex), 2) = "0x" Then
    If MsgBox(i_bahasa(9), vbExclamation + vbYesNo) = vbYes Then
        Call UnloadModuleForce(lstModule.List(lstModule.ListIndex), lstModule, lblSelectedPID.Caption)
    End If
End If
End Sub
Private Sub lstPlugin_Click()
Dim MyIndek As Long
On Error Resume Next
MyIndek = lstPlugin.ListIndex
If MyIndek >= 0 Then
   IndekPluginTerplih = MyIndek
   RetrievePlugInfo MyIndek, lstPlugin, lblPlugSelect1, lblPlugAut1, lblPlugAutEmail1, lblPlugAutSite1, lblPlugVer1, lblPlugDesc1
End If
End Sub
Private Sub lvHidden_ColumnClick(ByVal oColumn As cColumn)
    oColumn.Sort
End Sub
Private Sub lvHidden_ItemClick(ByVal oItem As cListItem, ByVal iButton As evbComCtlMouseButton)
      FrmRTP.picInfoHidden.Cls
      RetrieveIcon oItem.SubItem(2).Text, FrmRTP.picInfoHidden, ricnLarge
End Sub
Private Sub lvInfo_ColumnClick(ByVal oColumn As cColumn)
    oColumn.Sort
End Sub
Private Sub lvInfo_ItemClick(ByVal oItem As cListItem, ByVal iButton As evbComCtlMouseButton)
  If iButton = vbccMouseRButton Or iButton = vbccMouseLButton Then
     LastPathRClick = oItem.SubItem(2).Text ' berikan info klik kanan terkahir
  Else
     FrmRTP.picIconInfo.Cls
     RetrieveIcon oItem.SubItem(2).Text, FrmRTP.picIconInfo, ricnLarge
  End If
End Sub
Private Sub lvJail_ColumnClick(ByVal oColumn As cColumn)
    oColumn.Sort
End Sub
Private Sub lvJail_ItemClick(ByVal oItem As cListItem, ByVal iButton As evbComCtlMouseButton)
    LastPathRClick = oItem.SubItem(2).Text
    If iButton = vbccMouseRButton Then FrmSysTray.mnsubmit.Enabled = True: FrmSysTray.PopupMenu FrmSysTray.mnjail, 0, , , FrmSysTray.mncleanjail
    lvJail.ToolTipText = LastPathRClick
End Sub
Private Sub lvMalware_ColumnClick(ByVal oColumn As cColumn)
    oColumn.Sort
End Sub
Private Sub lvMalware_ItemClick(ByVal oItem As cListItem, ByVal iButton As evbComCtlMouseButton)
      picInfoMalware.Cls
      RetrieveIcon oItem.SubItem(2).Text, picInfoMalware, ricnLarge
      lvMalware.ToolTipText = LastPathRClick
End Sub
Private Sub lvProses_ColumnClick(ByVal oColumn As cColumn)
    oColumn.Sort
End Sub
Private Sub lvProses_ItemClick(ByVal oItem As cListItem, ByVal iButton As evbComCtlMouseButton)
    Call ENUM_MODULE(oItem.SubItem(3).Text, lstModule)
    lblSelectedPID.Caption = oItem.SubItem(3).Text
If iButton = vbccMouseRButton Or iButton = vbccMouseLButton Then FrmSysTray.PopupMenu FrmSysTray.mnProses, 0, , , FrmSysTray.mnRefresh
End Sub
Private Sub lvRegistry_ColumnClick(ByVal oColumn As cColumn)
    oColumn.Sort
End Sub
Sub segarkanDB() 'Segarkan DB User
If ValidFile(App.path & "\WanUDB.dll") = True Then ReadUDB lvm31
Call Init_Dtabase
End Sub
Private Sub cmdRemExcReg_Click()
    HapusFile GetFilePath(App_FullPathW(False)) & "\Reg.lst"
    ReadExceptReg GetFilePath(App_FullPathW(False)) & "\Reg.lst", lstExceptReg
End Sub
Private Sub cmdRemExcReg1_Click()
    RemoveExceptionByIndek lstExceptReg.ListIndex, REG_EXC
    JumRegExcep = ReadExceptReg(GetFilePath(App_FullPathW(False)) & "\Reg.lst", lstExceptReg)
End Sub
Sub mnkeluar()
    Call LepasSemuaKunci
    LewatExit = True
        Me.WindowState = vbNormal
        Me.Show
         TimerIcon.Enabled = False

    Unload FrmRTP: Unload FormSplash: Unload FrmAbout: Unload FrmFDmsk: Unload FrmSysTray: Unload FrmPassword: Unload Me
    End
End Sub
Private Sub pic_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
If Button = 1 Then DragForm Me
End Sub
Private Function PotongTampilanKar(sKar As String, nLimit As Byte) As String 'Potong File Kar
If Len(sKar) >= nLimit Then PotongTampilanKar = Left$(sKar, nLimit - 30) & "...\" & GetFileName(sKar) Else PotongTampilanKar = sKar
End Function
Private Sub PicHomeTolls_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
Call LoadGUI
End Sub


Private Sub Tabscan_Click(tIndex As Integer)
    On Error Resume Next
    Picscan(tIndex - 1).ZOrder 0
End Sub
Private Sub TabTools_Click(tIndex As Integer)
    On Error Resume Next
    Pictools(tIndex - 1).ZOrder 0
End Sub

Private Sub TimerIcon_Timer() 'iCon Timer
Dim sPath As String
sPath = GetFilePath(App_FullPathW(False)) ' & "\"
Select Case Left$(Command, 2)
  Case "-A" ' dari auto run
   Case "-K"
   Case Else
    If frmMain.WindowState <> vbNormal Then
  If Left$(Command, 2) <> "-K" Then Call terapkanIcon2
    End If

    End Select
End Sub
Public Sub bUKAFROM() 'Open Form
 If lvMalware.ListItems.Count > 0 Then
     Call ImgClick_Click(1): Tabscan.ActiveTab = 3: Tabscan_Click (3)
     CmdQuarAll.Enabled = True
 ElseIf BERHENTI = False Then
     Call ImgClick_Click(1): Tabscan.ActiveTab = 2: Tabscan_Click (2)
 Else
     Call ImgClick_Click(0)
 End If
    Call loadRTP: Me.Show
    FrmSysTray.mnCScan.Caption = g_bahasa(0)
    Call LoadDll
End Sub
Public Sub berhentiScanYa() 'Stop Scan
Call BtnResult_Click
End Sub
Private Sub tmAwal_Timer() 'Timmer Awal , Jalann ?
Dim WinPath
WinPath = Environ$("windir")
Dim nFS As Long
Detik2 = Detik2 + 1
If Detik2 = 1 Then
Select Case Left$(Command, 2)
  Case "-A" ' dari auto run

    If FrmConfig.ck11.Value = 1 Then FormSplash.Show: FormSplash.buka_splas

     FrmSysTray.mnCScan.Caption = g_bahasa(1)
      Call InitAplikasi2
      If StatusRTP = True Then

         Call LetakanForm(FormSplash, True)
        If UCase$(Left$(Command, 2)) <> "-K" Then Call terapkanIcon
      End If
  Case "-K" ' scan dari context menu

      TimerIcon.Enabled = False

     PathDariShellMenu = Mid$(Command, 4)
     If Len(PathDariShellMenu) > 0 And ValidFile(PathDariShellMenu) = False Then
     Call ImgClick_Click(1): Tabscan.ActiveTab = 2: Tabscan_Click (2)
     
 Call cmdStartScan_Click
     ElseIf Len(PathDariShellMenu) > 0 And ValidFile(PathDariShellMenu) = True Then
     Call ImgClick_Click(2): TabTools.ActiveTab = 2: TabTools_Click (2)
     txtVirusPath.Text = PathDariShellMenu
     Else
     End If
      tmAwal.Enabled = False
  Case Else ' double klik
      Call InitAplikasi2: Me.Show
End Select
End If

If Detik2 = 8 Then
   CabutBalon Me
End If

If Detik2 = 12 Then ' klo mau cek online update
   If FrmConfig.ck9.Value = 1 Then Call AmbilUpdateInfo("http://waniez.p.ht/download/updateinfo.txt", GetSpecFolder(USER_DOC) & "\updwaniez.tmp")
   HentikanUpdate = True
End If

If Detik2 = 22 Then ' bayangkan aja udah selesai ambil informasinya
   If FrmConfig.ck9.Value = 1 Then Call CheckUpdate(GetSpecFolder(USER_DOC) & "\updwaniez.tmp", lblStatusUpdate)
   tmAwal.Enabled = False
   If bUpdateCompon = True Then
   'Call bunyi

                Me.Show
              Call ImgClick_Click(4)
   TampilkanBalon frmMain, "Update New Wan'iez Antivirus" & " !", a_bahasa(26), NIIF_INFO
      If MsgBox("Wan'iez Antivirus Detected File Update to server !" & Chr$(13) & _
                "Update Now !!", vbInformation + vbYesNo) = vbYes Then

      Call cmdCheckUpdate_Click
      End If
   End If

End If

End Sub
Private Sub tmFlash_Timer() 'Timmer Falsh
Dim sDriveName          As String
Dim DriveLabel          As String
Dim nDriveNameLen       As Long
Dim hajarWMX          As String
Dim hajarWMX2          As String
Dim hajarWMX3          As String
Dim hajarWMX4          As String
Dim sPath As String
sPath = GetFilePath(App_FullPathW(False))
If BERHENTI = False Or UCase$(Left$(Command, 2)) = "-K" Then ' jika scan scan jalan
   tmFlash.Enabled = False
   Exit Sub
End If
If AdakahFDBaru(LastFlashVolume) = True Then
'Call bunyi
Call LoadDll
   nDriveNameLen = 128
   sDriveName = String$(nDriveNameLen, 0)
   If GetVolumeInformationW(StrPtr(Chr$(LastFlashVolume) & ":\"), StrPtr(sDriveName), nDriveNameLen, ByVal 0, ByVal 0, ByVal 0, 0, 0) Then
       DriveLabel = Left$(sDriveName, InStr(1, sDriveName, ChrW$(0)) - 1)
   Else
       DriveLabel = vbNullString
   End If
   Call BuilDirTree ' refresh dir tree
   'TampilkanBalon Me, j_bahasa(60) & " [ " & DriveLabel & " (" & Chr$(LastFlashVolume) & ") ] !", i_bahasa(26), NIIF_INFO
  ' If MsgBox("Wan'iez AntiVirus " & j_bahasa(61) & " Wan'iez AntiVirus" & " [ " & DriveLabel & " (" & Chr$(LastFlashVolume) & ") ]", vbYesNo + vbExclamation, "Wan'iez AntiVirus") = vbYes Then
      NewFDMasuk = Chr$(LastFlashVolume) & ":\"
     hajarWMX = Chr$(LastFlashVolume) & ":\RECYCLER"
       
    If PathIsDirectory(StrPtr(hajarWMX)) <> 0 Then 'ScanPatWithRTPUSB hajarWMX
    FrmFDmsk.Show
          TampilkanBalon frmMain, j_bahasa(63) & " [ " & DriveLabel & " (" & Chr$(LastFlashVolume) & ") ]" & j_bahasa(64), NIIF_WARNING

        If MsgBox(j_bahasa(63) & " [ " & DriveLabel & " (" & Chr$(LastFlashVolume) & ") ]" & j_bahasa(64), vbCritical + vbYesNo) = vbYes Then
           KillFolder hajarWMX
           '"Wan'iez AntiVirus found 'RECYCLER' folder in"                                                                             "this folder is recognized as the potential threat and cntains virus! you are suggested to click YES to remove and immune it"
           ExtractRes (NewFDMasuk & "RECYCLER"), 5, "REG"
          ' If ValidFile(NewFDMasuk & "RECYCLER") = True Then SetAttr (NewFDMasuk & "RECYCLER"), vbHide + vbReadOnly + vbSystem
           If ValidFile(NewFDMasuk & "RECYCLER") = True And PathIsDirectory(StrPtr(hajarWMX)) <> 0 Then
              MsgBox j_bahasa(65) & " [ " & DriveLabel & " (" & Chr$(LastFlashVolume) & ") ]" & j_bahasa(66), vbExclamation
           Else
           SetAttr (NewFDMasuk & "RECYCLER"), vbHidden + vbReadOnly + vbSystem
              MsgBox j_bahasa(67) & " [ " & DriveLabel & " (" & Chr$(LastFlashVolume) & ") ]" & j_bahasa(68), vbInformation
           End If
           GoTo lanjut
       Else
       GoTo lanjut
       End If
    Else
       GoTo lanjut
    End If
lanjut:

Call BuilDirTree ' refresh dir tree
     ' FrmFDmsk.lblFD(0).Caption = "Full Scan [ " & DriveLabel & " (" & Chr$(LastFlashVolume) & ")\*.* ]"
     ' FrmFDmsk.lblFD(1).Caption = "Scan Area [ " & DriveLabel & " (" & Chr$(LastFlashVolume) & ") ]"
     ' FrmFDmsk.Lbtitle = " Scanning " & DriveLabel & " (" & Chr$(LastFlashVolume) & ")"

     ' NewFDMasuk = Chr$(LastFlashVolume) & ":\"
     ' Load FrmFDmsk
     ' FrmFDmsk.Show
End If
End Sub
Public Sub ScanFD() 'Scan USB
With frmMain
Call LetakanForm(Me, True)
      Call cmdStartScan_Click: Call ImgClick_Click(1): Tabscan.ActiveTab = 2: Tabscan_Click (2)
      .Show
      .WindowState = vbNormal
End With
End Sub
Public Sub ScanFDArea() 'Scan Are USB
ScanPatWithRTP NewFDMasuk
If FrmRTP.lvRTP.ListItems.Count = 0 Then
MsgBox "No Virus Detected in Area Flasdhisk", vbInformation, "Information"
End If
End Sub
Public Sub BatalScanFD() 'Batal USB Scan
NewFDMasuk = ""
End Sub
Private Sub tmrSPEED_Timer() 'Timer Speed
On Error GoTo metu
Dim CurSpeed As String
If BERHENTI = False Then
    NewTime = FileFound
    CurSpeed = Int((NewTime - LastTime) * 2.5)
    AllTime = AllTime + 1
    TimeSpeedCount = TimeSpeedCount + 1
    lblbanter.Caption = "Speed " & Abs(CurSpeed) & " File(s)/sec"
    If CurSpeed > MaxSpeed Then MaxSpeed = CurSpeed
    LastTime = NewTime
Else
    SpeedRate = Abs(Int(FileFound / (AllTime * 0.4)))
    lblbanter.Caption = "Speed " & SpeedRate & " File(s)/sec"
    If SpeedRate > MaxSpeed Then MaxSpeed = SpeedRate
    tmrSPEED.Enabled = False
End If
metu:
End Sub
Private Sub tmTime_Timer() 'Timer Scan
Detik = Detik + 1
If Detik = 60 Then
   Detik = 0
   Menit = Menit + 1
End If
If Menit = 60 Then
   Menit = 0
   Jam = Jam + 1
End If
lbTime.Caption = Jam & " :" & Menit & " :" & Detik
End Sub

Private Sub txtVirusName_KeyPress(KeyAscii As Integer) 'Text Name Virus
If KeyAscii = 13 Then 'enter
Call cmdAddVirus_Click
End If
If KeyAscii = 27 Then 'enter
txtVirusName.Text = ""
End If
End Sub
Private Sub txtVirusPath_Change() 'Text Path
nVirusTmp = lvm31.ListItems.Count ' + 1
txtVirusName.Text = j_bahasa(9) & " (" & (nVirusTmp + 1) & ")"
cmdAddVirus.Enabled = True: cmdCancel.Enabled = True
Me.WindowState = vbNormal
Me.Show
Call ImgClick_Click(2): frmMain.TabTools.ActiveTab = 2: TabTools_Click (2):
AnalisisPE PathDariShellMenu
End Sub
Private Sub txtVirusPath_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, y As Single) 'text Path Virus
txtVirusPath.Text = Data.Files(vbCFText)
If ValidFile(Data.Files(vbCFText)) = True Then
    If IsFileProtectedBySystem(Data.Files(vbCFText)) = True Then GoTo metu

nVirusTmp = lvm31.ListItems.Count
txtVirusName.Text = j_bahasa(9) & " (" & (nVirusTmp + 1) & ")"
Else
MsgBox "File [ " & txtVirusPath.Text & " ]" & Chr$(13) & _
          i_bahasa(20) & " !", vbExclamation
   txtVirusPath.Text = "": txtVirusName.Text = "": FrmRTP.Text1.Text = ""
End If
Exit Sub
metu:
  MsgBox f_bahasa(20), vbCritical
Call cmdCancel_Click
End Sub
Private Sub ucListVirus2_ColumnClick(ByVal oColumn As cColumn) 'List Virus
oColumn.Sort
End Sub
Private Sub UniDialog1_FolderCancel(ByVal CancelType As UniDialogFolderCancel) 'Open Path
    UniDialogPath = ""
End Sub
Private Sub UniDialog1_FolderSelect(ByVal path As String) 'Open Path
    UniDialogPath = path
End Sub
Private Sub UniDialog1_OpenCancel(ByVal CancelType As UniDialogFileCancel) 'Open Folder
    UniDialogFile = ""
End Sub
Private Sub UniDialog1_OpenFile(ByVal FileName As String) 'Open Folder
    UniDialogFile = FileName
End Sub
Private Sub ResetObjek() 'Reset Objeck

   lbMalware.Caption = "0": lbReg.Caption = "0": lbHidden.Caption = "0": lbInfo.Caption = "0"
   lbBypass.Caption = "0": lbFileFound.Caption = "0": lbFileCheck.Caption = "0"
   lvMalware.ListItems.Clear: lvRegistry.ListItems.Clear: lvHidden.ListItems.Clear: lvInfo.ListItems.Clear
   PB1.Value = 0: VirusFound = 0: FileFound = 0: FileCheck = 0: FileNotCheck = 0: nRegVal = 0: nErrorReg = 0:
   FileToScan = 0: InfoFound = 0: total_size = 0: Detik = 0: Menit = 0: Jam = 0
   lbStatus22.Caption = "0 " & j_bahasa(39) & ", 0 " & j_bahasa(38) & ", Size:0 KB [Kilo Byte]"
   CmdQuarAll.Enabled = False: cmdFixMalware.Enabled = False: cmdFixMalwareAll.Enabled = False
   cmdFixReg.Enabled = False: cmdFixRegAll.Enabled = False
   cmdFixHidden.Enabled = False: cmdFixHiddenall.Enabled = False
   cmdProperties.Enabled = False: cmdProperties.Enabled = False: cmdExplore.Enabled = False
   BTNlOG.Enabled = False: BtnFix.Enabled = False: BtnResult.Enabled = True
   ckScan(0).Enabled = False: ckScan(1).Enabled = False: ckScan(2).Enabled = False
   WithBuffer = True ' nilai awal true
   BERHENTI = False
   tmTime.Enabled = True: TimerIcon.Enabled = False
   lbObject.Caption = ""
   Call LepasSemuaKunci
   DataAutorun = "" ' Reset
   TargetShorcutOnFD = "" ' Reset
End Sub
Private Sub ReBack() 'Reg Back Awal
TimerIcon.Enabled = True: BTNlOG.Enabled = True: lblbanter.Caption = ""
'lbStatus22.Caption = "0 " & j_bahasa(39) & ", 0 " & j_bahasa(38) & ", Size:0 KB [Kilo Byte]"

   If lvMalware.ListItems.Count > 0 Then
      cmdFixMalware.Enabled = True: cmdFixMalwareAll.Enabled = True: CmdQuarAll.Enabled = True: ckScan(0).Enabled = True
   End If
   If lvRegistry.ListItems.Count > 0 Then
      cmdFixReg.Enabled = True: cmdFixRegAll.Enabled = True: ckScan(2).Enabled = True
   End If
   If lvHidden.ListItems.Count > 0 Then
      cmdFixHidden.Enabled = True: cmdFixHiddenall.Enabled = True: ckScan(1).Enabled = True
   End If
   If lvInfo.ListItems.Count > 0 Then
      cmdExplore.Enabled = True: cmdProperties.Enabled = True
   End If
End Sub
Public Sub cmdCheckUpdate_Click() 'Check Update
Dim A As String
Dim B As String
Dim spWd As String
A = GetSTRINGValue(&H80000001, "Software\Wan'iez\", "password")
If A <> "" And FrmConfig.CkPass.Value <> 0 And FrmConfig.CkUpdate.Value <> 0 Then 'MsgBox "ada paswor"
B = encrypt.Cryption(A, "WANPASS", False)
FrmPassword.ShowMe spWd
If spWd <> "" Then
If spWd = B Then

    If cmdCheckUpdate.Caption = a_bahasa(26) Then
        bUpdateCompon = False
        HentikanUpdate = False
        BufferUpdate = -1
  
        AmbilUpdateInfo "http://waniez.p.ht/download/updateinfo.txt", GetSpecFolder(USER_DOC) & "\updwaniez.tmp"
        cmdCheckUpdate.Caption = j_bahasa(34): lblStatusUpdate.Caption = j_bahasa(33)
    Else
        HentikanUpdate = True
        cmdCheckUpdate.Caption = a_bahasa(26): lblStatusUpdate.Caption = j_bahasa(32)
    End If

Exit Sub
Else
      MsgBox "password you entered is not corect!", vbCritical, i_bahasa(26)
End If
Exit Sub
End If
Else
    If cmdCheckUpdate.Caption = a_bahasa(26) Then
        bUpdateCompon = False
        HentikanUpdate = False
        BufferUpdate = -1
  
        AmbilUpdateInfo "http://waniez.p.ht/download/updateinfo.txt", GetSpecFolder(USER_DOC) & "\updwaniez.tmp"
        cmdCheckUpdate.Caption = j_bahasa(34): lblStatusUpdate.Caption = j_bahasa(33)
    Else
        HentikanUpdate = True
        cmdCheckUpdate.Caption = a_bahasa(26): lblStatusUpdate.Caption = j_bahasa(32)
    End If
End If
End Sub
Private Sub Downloader1_DownloadComplete(MaxBytes As Long, SaveFile As String) 'Download Complite
   BufferUpdate = BufferUpdate + 1
   tmUpdate.Enabled = True
End Sub
Private Sub Downloader1_DownloadError(SaveFile As String) 'Tampilkan Download
    HentikanUpdate = True
    TampilkanBalon frmMain, "Error Download ....", i_bahasa(26), NIIF_WARNING
    lblStatusUpdate.Caption = "Error Download..."
End Sub
Private Sub Downloader1_DownloadProgress(CurBytes As Long, MaxBytes As Long, SaveFile As String) 'Click Download
    PBC.Max = MaxBytes: PBC.Value = CurBytes
End Sub
Private Sub tmUpdate_Timer() 'Update File DB
Dim TmpPath  As String
Dim TmpPath2 As String
Dim MyPath   As String

If HentikanUpdate = True Then GoTo LBL_MATI_AJ
'GifUpdate.Enabled = True
'GifUpdate.Visible = True
TmpPath = GetSpecFolder(USER_DOC) & "\updwz.tmp"
TmpPath2 = GetSpecFolder(USER_DOC) & "\wztmp.txt"
MyPath = GetFilePath(App_FullPathW(False))

Select Case BufferUpdate
    Case 0 ' baru ambil updateinfo.txt
         txtRetriveInfo.Text = CheckUpdate(TmpPath, lblStatusUpdate)
         If bUpdateCompon = True Then ' berarti ada update terbaru

            cmdCheckUpdate.Caption = j_bahasa(34)
            UpdateKomponen PB_UPD, lblStatusUpdate, BufferUpdate
         End If
    Case 1 ' selesai update komponen db-0 (0x.upx)
         MoveIfValidComp TmpPath2, MyPath & "\upx\" & Hex$(BufferUpdate - 1) & "_a.upx", lblStatusUpdate
         UpdateKomponen PB_UPD, lblStatusUpdate, BufferUpdate
    Case Is < 16 ' selesai update komponen db-1 (1x.upx) dst
         MoveIfValidComp TmpPath2, MyPath & "\upx\" & Hex$(BufferUpdate - 1) & "_a.upx", lblStatusUpdate
         UpdateKomponen PB_UPD, lblStatusUpdate, BufferUpdate
    Case 16 ' selesai update komponen db-15 (terakhir PE)
         MoveIfValidComp TmpPath2, MyPath & "\upx\" & Hex$(BufferUpdate - 1) & "_a.upx", lblStatusUpdate
         UpdateKomponenNonPE PB_UPD, lblStatusUpdate, BufferUpdate - 16
    Case Is < 32
         MoveIfValidComp TmpPath2, MyPath & "\upx\" & Hex$(BufferUpdate - 17) & "_x.upx", lblStatusUpdate
         UpdateKomponenNonPE PB_UPD, lblStatusUpdate, BufferUpdate - 16
    Case 32 ' selsai sampai akhir
         MoveIfValidComp TmpPath2, MyPath & "\upx\" & Hex$(BufferUpdate - 17) & "_x.upx", lblStatusUpdate
         cmdCheckUpdate.Caption = a_bahasa(26): lblStatusUpdate.Caption = j_bahasa(31)
         
         PBC.Value = PBC.Max
         tmUpdate.Enabled = False
         BufferUpdate = -1
         Call BacaDatabase: Call ListVirus(lstListWorm)
         Exit Sub
End Select
PBC.Value = 0
tmUpdate.Enabled = False ' matiin lagi...
Exit Sub
'GifUpdate.Enabled = False
'GifUpdate.Visible = False
LBL_MATI_AJ:
   PBC.Value = 0
   tmUpdate.Enabled = False: cmdCheckUpdate.Caption = a_bahasa(26)
End Sub
Private Sub cmdupload_Click() 'Upload Firus
'cmdupload.Enabled = False:
'cmdupload.Enabled = True
'ShellExecute Me.hwnd, vbNullString, "http://www.4shared.com/account/dir/ev0v2YST/sharing.html", vbNullString, "C:\", 1
'Exit Sub
End Sub
Private Sub Check10_Click() 'Run
Call Check10M
End Sub
Private Sub Check11_Click() 'Shutdown
Call Check11M
End Sub
Private Sub Check12_Click() 'Help
Call Check12M
End Sub
Private Sub Check13_Click() 'Display Setting
Call Check13M
End Sub
Private Sub Check14_Click() 'Registry Editor
Call Check14M
End Sub
Private Sub Check15_Click() 'task manager
Call Check15M
End Sub
Private Sub Check16_Click() 'CMD
Call Check16M
End Sub
Private Sub Check17_Click() 'Windows Hotkey
Call Check17M
End Sub
Private Sub Check18_Click() 'FolderOptions
Call Check18M
End Sub
Private Sub Check19_Click() 'ControlPanel
Call Check19M
End Sub
Private Sub Check20_Click() 'ViewContextMenu
Call Check20M
End Sub
Private Sub Check21_Click() 'TrayContextMenu
Call Check21M
End Sub
Private Sub Check22_Click() 'SetTaskbar
Call Check22M
End Sub
Private Sub Check23_Click() 'Desktop
Call Check23M
End Sub
Private Sub Check24_Click() 'Show superhidden
Call Check24M
End Sub
Private Sub Check25_Click() 'Show Hidden file
Call Check25M
End Sub
Private Sub Check26_Click() 'Show Extension file
Call Check26M
End Sub
Private Sub Check27_Click() 'Show Full Path in Address Bar
Call Check27M
End Sub
Private Sub Check28_Click() 'Show Full Path in Title Bar
Call Check28M
End Sub
Private Sub Check29_Click() 'Stop Autoplay Drive
Call Check29M
End Sub
Private Sub Check7_Click() 'Find
Call Check7M
End Sub
Private Sub Check8_Click() 'LogOff
Call Check8M
End Sub
Private Sub Check9_Click() 'RecentDocs
Call Check9M
End Sub 'cek all
Public Function ScanFileRTP(path As String) 'Scan RTP FILE
Thread.StartTask "Path", False, VbLet, path
Thread.StartTask "TheFirst", False, VbLet, True
Thread2.StartTask "Path", False, VbLet, path
Thread2.StartTask "TheFirst", False, VbLet, True
Thread.StartTask "StartScanRTP", True, VbMethod
End Function
Private Sub Thread2_BufferCompleteRTP(FileToScan As Long, FolderToScan As Long) 'Thread Buffer Compile RTP
Working = False
End Sub
Public Function sCANFile(path As String) 'Scan File
Thread.StartTask "Path", False, VbLet, path
Thread.StartTask "TheFirst", False, VbLet, True
Thread2.StartTask "Path", False, VbLet, path
Thread2.StartTask "TheFirst", False, VbLet, True
Thread.StartTask "BufferScan", True, VbMethod
Thread.StartTask "StartScan", True, VbMethod
End Function
Private Sub Thread2_BufferComplete(FileToScan As Long, FolderToScan As Long) 'Thread Buffer Complite
Working = False
Fol_toScan = FolderToScan
File_toScan = FileToScan
frmMain.lbStatus.Caption = d_bahasa(17) '& " [ " & FolderToScan & " " & j_bahasa(39) & ", " & FileToScan & " " & j_bahasa(38) & " ]"

PB1.Max = FileToScan
BtnResult.Enabled = True
tmrSPEED.Enabled = True
End Sub
Private Sub HasilScan() 'Hasil Scanning
      GifScan.Enabled = False
      GifScan.Visible = False
Call akir
PB1.Value = PB1.Max
tmTime.Enabled = False: cmdStartScan.Enabled = True: tmrSPEED.Enabled = False: BERHENTI = True
  Me.Show
           
   Dim jumVir As Long, Msg As String: jumVir = lvMalware.ListItems.Count + lvHidden.ListItems.Count + nErrorReg
    Msg = i_bahasa(33) & IIf(jumVir > 10, i_bahasa(35), IIf(jumVir > 0, i_bahasa(34), i_bahasa(36)))
    If jumVir > 0 Then
        MsgBox Msg & vbCrLf & i_bahasa(37) & IIf(lvMalware.ListItems.Count > 0, vbCrLf & "- " & lvMalware.ListItems.Count & " " & i_bahasa(40), "") & _
        IIf(nErrorReg <> 0, vbCrLf & "- " & nErrorReg & " " & i_bahasa(39), "") & IIf(lvHidden.ListItems.Count > 0, vbCrLf & "- " & lvHidden.ListItems.Count & " " & i_bahasa(41), ""), vbExclamation, i_bahasa(38)
      GoTo xer
    End If
    
   MsgBox i_bahasa(42), vbInformation, i_bahasa(38)
    BtnResult.Caption = a_bahasa(25): BtnResult.Enabled = False
xer:
      If lvMalware.ListItems.Count > 0 Or lvRegistry.ListItems.Count > 0 Or lvHidden.ListItems.Count > 0 Then
     Call ImgClick_Click(1): Tabscan.ActiveTab = 2: Tabscan_Click (2)
      BtnResult.Caption = a_bahasa(25): BtnFix.Enabled = True ':
      End If
      Call ReBack ' Aktifkan yang peru diaktifkan
      PathDariShellMenu = "" 'kosongkan lagi yang dari shell menu
      If FrmConfig.ck10.Value = 1 Then tmFlash.Enabled = True
End Sub
Private Sub Thread_ScanComplete(FileterScan As Long) 'Thread Scan Complite
Working = False
lbStatus.Caption = d_bahasa(17)
Call HasilScan
End Sub
Private Sub Thread_ScanRTPComplete() 'Thread Scan RTP
Working = False
End Sub
Private Sub Thread_TaskComplete(ByVal FunctionName As String, Data As Variant) 'this is an event from the active x exe, you can add any events you want similar to an ocx
    'the thread is done working!
    If FunctionName = "DoSomeLongWork" Then
        Working = False
        'process the finished data
        ThreadStatus "'" & FunctionName & "' has completed its task!!"
        ThreadStatus "Separate Thread Processed " & CStr(Data) & " Characters!"
    End If
    'add all the function names here that are in your "threaded" activeX exe control
End Sub
Sub ThreadStatus(ByVal StatusText As String) 'this just adds text to the listbox for display purposes
    FrmSysTray.lstStatus.AddItem StatusText
    FrmSysTray.lstStatus.ListIndex = FrmSysTray.lstStatus.NewIndex
End Sub
Private Sub Thread_ThisPath(PathApa As String, FileterScan As Long) 'Thread Path
    lbObject.Caption = PathApa
    File_terScan = FileterScan
    If File_toScan > 0 Then
    End If
End Sub
Function file_isFolder(path As String) As Long 'File Folder
On Error GoTo salah

Dim ret As VbFileAttribute
    ret = GetAttr(path) And vbDirectory
    If ret = vbDirectory Then
        file_isFolder = 1
    Else
        file_isFolder = 0
    End If
    
Exit Function

salah:
file_isFolder = -1
End Function
Public Function terapkanIcon() 'Terapkan Icon Pertama yah
If Left$(Command, 2) = "-K" Then GoTo metu
 FrmSysTray.Refresh
 If FrmConfig.ck8.Value = 1 Then
    'UpdateIconRTPmati
   Shell_NotifyIcon NIM_DELETE, nID
  Call UpdateIcon(FrmSysTray.Icon, "Wan'iez Antivirus™ - Your System Is Secured", FrmSysTray)

    Else
    Shell_NotifyIcon NIM_DELETE, nID
  Call UpdateIconRTPmati(FrmSysTray.Icon, "Wan'iez Antivirus™ - Your System Is Not Secured", FrmSysTray)

    End If
metu:
End Function
Public Function terapkanIcon2() 'Terapkan icon 2
If Left$(Command, 2) = "-K" Then GoTo metu
 FrmSysTray.Refresh
 If FrmConfig.ck8.Value = 1 Then
  Call UpdateIcon(FrmSysTray.Icon, "Wan'iez Antivirus™ - Your System Is Secured", FrmSysTray)
    Else
  Call UpdateIconRTPmati(FrmSysTray.Icon, "Wan'iez Antivirus™ - Your System Is Not Secured", FrmSysTray)
    End If
metu:
End Function
Private Sub ReadUDB(lv As ucListView) ' baca database ceksum kedua dan memasukkannya ke lisview
Dim strUDb() As String
Dim TMP() As String
Dim sNum As Integer
lv.ListItems.Clear
    strUDb = Split(OpenFileInTeks(App.path & "\WanUDB.dll"), ";")
  For sNum = 1 To UBound(strUDb)
    TMP = Split(strUDb(sNum), "|")

    
  AddInfoToListDua lvm31, "", TMP(1), TMP(0), 4, 5
 Next
JumlahVirusOUT = lvm31.ListItems.Count
frTemp.Caption = "User Database : " & lv.ListItems.Count & " worm"
LbVirusVersion.Caption = ": " & CStr(JumVirus) + JumlahVirusINT & " Virus , " & lv.ListItems.Count & " Malware User + Heuristic": DoEvents
FrmAbout.lbVirus.Caption = ": " & CStr(JumVirus) + JumlahVirusINT & " Virus , " & lvm31.ListItems.Count & " Malware User + Heuristic": DoEvents

If JumlahVirusOUT > 0 Then
cmdRemmoveCk.Enabled = True
Else
cmdRemmoveCk.Enabled = False
End If
AutoLst lvm31
End Sub
Private Sub cmdRemmoveCk_Click() 'Remove Check User DB
Dim i As Integer
Dim A As String
Dim B As String
Dim spWd As String
A = GetSTRINGValue(&H80000001, "Software\Wan'iez", "password")
If A <> "" And FrmConfig.CkPass.Value <> 0 And FrmConfig.CkDellUs.Value <> 0 Then 'MsgBox "ada paswor"
B = encrypt.Cryption(A, "WANPASS", False)
FrmPassword.ShowMe spWd
If spWd <> "" Then
If spWd = B Then
If MsgBox("Are you sure to delete the chosen database user?", vbInformation + vbYesNo) = vbYes Then

For i = lvm31.ListItems.Count To 1 Step -1
            If lvm31.ListItems.Item(i).Checked = True Then
                lvm31.ListItems.Remove i
                
            End If
        Next
End If
WriteDbAgain lvm31
                AutoLst lvm31
Exit Sub
Else
      MsgBox "password you entered is not corect!", vbCritical, i_bahasa(26)
End If
Exit Sub
End If
Else
If MsgBox("Are you sure to delete  database user?", vbInformation + vbYesNo) = vbYes Then

For i = lvm31.ListItems.Count To 1 Step -1
            If lvm31.ListItems.Item(i).Checked = True Then
                lvm31.ListItems.Remove i
            End If
        Next
End If
WriteDbAgain lvm31
                AutoLst lvm31
End If
End Sub
Private Sub WriteDbAgain(lv As ucListView) 'Database User Write Db
Dim Penampung As String
Dim X As Integer
For X = 1 To lv.ListItems.Count
    Penampung = Penampung & ";" & lv.ListItems.Item(X).SubItem(3) & "|" & lv.ListItems.Item(X).SubItem(2)
Next
Open App.path & "\WanUDB.dll" For Output As #4 ' Tulis ke User.DAT
    Write #4, Penampung
Close #4
Call segarkanDB
End Sub
Public Function GenerateRandomTitle() As String 'Enkrip Rondom

    Dim sTitle() As Variant
    
    sTitle = Array("a", "b", "c", "d", "e", "f", "g", _
        "h", "i", "j", "k", "l", "m", "n", "o", "p", "q", "r", "s", "t", "u", "v", "w", "x", "y", "z", "A", "B", "C", "D", "E", "F", "G", "I", _
        "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z")
    
    Randomize
    
    GenerateRandomTitle = sTitle(Rnd * UBound(sTitle)) & sTitle(Rnd * UBound(sTitle)) & sTitle(Rnd * _
        UBound(sTitle)) & sTitle(Rnd * UBound(sTitle)) & sTitle(Rnd * UBound(sTitle)) & sTitle(Rnd * _
        UBound(sTitle)) & sTitle(Rnd * UBound(sTitle)) & sTitle(Rnd * UBound(sTitle))
    GenerateRandomTitle = EncryptText(GenerateRandomTitle)
        
End Function
Private Function EncryptText(ByVal sText As String) As String 'Enkrip Text
    
    Dim intLen As Integer
    Dim sNewText As String
    Dim sChar As String
    Dim i As Integer
    
    sChar = ""
    intLen = Len(sText)
    For i = 1 To intLen
        sChar = Mid(sText, i, 1)
        Select Case Asc(sChar)
            Case 65 To 90: sChar = Chr(Asc(sChar) + 127)
            Case 97 To 122: sChar = Chr(Asc(sChar) + 121)
            Case 48 To 57: sChar = Chr(Asc(sChar) + 196)
            Case 32: sChar = Chr(32)
        End Select
        sNewText = sNewText + sChar
    Next i
    EncryptText = sNewText
    
    Exit Function
    
End Function
Private Sub ContextMenu_Click(ByVal Item As WODSHMENUCOMLib.IMenuItem, ByVal Name As String, ByVal Directory As String, ByVal Targets As String) 'Context Menu
  
        PathDariShellMenu = Targets
   UniLabel1 = Targets
End Sub
Private Sub ContextMenu_BeforePopup(ByVal Name As String, ByVal Targets As String) 'Context Menu Name

  If ValidFile(Targets) = False Then
  ContextMenu.MenuItems.Item(0).Caption = "Scan With Wan'iez Antivirus" ' & Targets
        ContextMenu.MenuItems(0).Enabled = True
        Else
      ContextMenu.MenuItems(0).Caption = "Add Wan'iez User Database" ' File: " & GetFileTitle(Targets)
        ContextMenu.MenuItems(0).Enabled = True
        End If
End Sub
Private Sub UniLabel1_Change()

If Len(PathDariShellMenu) > 0 And ValidFile(PathDariShellMenu) = False Then
    frmMain.DirTree.LoadTreeDir False, False
   RegNode = False: StartUpNode = False: ProsesNode = False
   
 Me.WindowState = vbNormal
 Me.Show
 Call ImgClick_Click(1): Tabscan.ActiveTab = 2: Tabscan_Click (2)
 Call cmdStartScan_Click: PathDariShellMenu = "": UniLabel1.Text = ""
 Else
 txtVirusPath.Text = PathDariShellMenu: PathDariShellMenu = ""
 End If
End Sub
Public Function aktifContek() ''Context menu aktif
     ContextMenu.Enabled = True
End Function
Public Function UnaktifContek() 'Context menu non aktif
     ContextMenu.Enabled = False
End Function
Private Sub ShellIE_WindowRegistered(ByVal lCookie As Long) ' user membuka explorer baru
    If StatusRTP = True Then Call MulaiRTP ' jika TRUE ajh
End Sub
Private Sub rtp_mode1_PathChange(Index As Integer, strPath As String) 'Mulai Scan RTP deh
If StatusRTP = True And BERHENTI = True And UCase$(Left$(Command, 2)) <> "-K" Then
       ScanPatWithRTP strPath
End If
keluar:
End Sub
Private Sub MulaiRTP() ' Mulai RTP
On Error Resume Next
If isCompatch = False Then
   Dim i As Integer, Cnt As Integer
   Cnt = ShellIE.Count - 1
   For i = 0 To Cnt
       If (rtp_mode1.Count - 1) < Cnt Then
          AddIEObj i
       End If
          If FindID(ShellIE(i).hwnd) = False Then
             rtp_mode1(i).EnabledMonitoring True
             rtp_mode1(i).AddSubClass ShellIE(i)
          End If
   Next i
End If
End Sub
Sub AddIEObj(Index As Integer) 'IE Object
On Error GoTo salah
    Load rtp_mode1(Index)
salah:
End Sub
Function FindID(id As Long) As Boolean 'FIDND ID
On Error GoTo salah
    Dim i As Integer
    For i = 0 To rtp_mode1.Count - 1
        If rtp_mode1(i).IEKey = id Then
           FindID = True
        End If
    Next i
salah:
End Function
Private Sub CompactObject() ' untuk aktifkan rtp
On Error Resume Next
isCompatch = True
   Dim i As Integer, Cnt As Integer
   For i = 0 To rtp_mode1.Count - 1
       rtp_mode1(i).SetIENothing
   Next i
       
   Set ShellIE = Nothing
   For i = 1 To rtp_mode1.Count - 1
        Unload rtp_mode1(i)
   Next i
   
   Set ShellIE = New SHDocVw.ShellWindows
   Cnt = ShellIE.Count - 1
   For i = 0 To Cnt
       If i > 0 Then
          AddIEObj i
       End If
          rtp_mode1(i).AddSubClass ShellIE(i)
   Next i
isCompatch = False
End Sub
