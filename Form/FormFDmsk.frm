VERSION 5.00
Begin VB.Form FrmFDmsk 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Scan Flsadisk"
   ClientHeight    =   1920
   ClientLeft      =   0
   ClientTop       =   -60
   ClientWidth     =   4455
   Icon            =   "FormFDmsk.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   128
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   297
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin WANIEZ.rButton CmdScanFD 
      Height          =   375
      Index           =   0
      Left            =   3120
      TabIndex        =   0
      Top             =   600
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      ButtonStyle     =   7
      BackColor       =   14211288
      BackColorPressed=   15715986
      BackColorHover  =   16243621
      BorderColor     =   9408398
      BorderColorPressed=   6045981
      BorderColorHover=   11632444
      Caption         =   "Scan"
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
   Begin VB.Timer Timer1 
      Interval        =   2
      Left            =   2040
      Top             =   600
   End
   Begin WANIEZ.rButton CmdScanFD 
      Height          =   375
      Index           =   1
      Left            =   3120
      TabIndex        =   1
      Top             =   1200
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      ButtonStyle     =   7
      BackColor       =   14211288
      BackColorPressed=   15715986
      BackColorHover  =   16243621
      BorderColor     =   9408398
      BorderColorPressed=   6045981
      BorderColorHover=   11632444
      Caption         =   "Scan"
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
   Begin WANIEZ.rButton CmdScanFD 
      Height          =   375
      Index           =   2
      Left            =   4080
      TabIndex        =   2
      Top             =   0
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      ButtonStyle     =   7
      BackColor       =   14211288
      BackColorPressed=   15715986
      BackColorHover  =   16243621
      BorderColor     =   9408398
      BorderColorPressed=   6045981
      BorderColorHover=   11632444
      Caption         =   "X"
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
   Begin VB.Label LbTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "Scanning"
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
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   2655
   End
   Begin VB.Label lblFDw 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Scan all files"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   120
      MouseIcon       =   "FormFDmsk.frx":52C2
      TabIndex        =   6
      Top             =   840
      Width           =   870
   End
   Begin VB.Label lblFDh 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Scan area drive"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   120
      MouseIcon       =   "FormFDmsk.frx":5414
      TabIndex        =   5
      Top             =   1440
      Width           =   1125
   End
   Begin VB.Label lblFD 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Scan Area"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   1
      Left            =   120
      MouseIcon       =   "FormFDmsk.frx":5566
      TabIndex        =   4
      Top             =   1200
      Width           =   885
   End
   Begin VB.Label lblFD 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Full Scan"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   0
      Left            =   120
      MouseIcon       =   "FormFDmsk.frx":56B8
      TabIndex        =   3
      Top             =   600
      Width           =   780
   End
End
Attribute VB_Name = "FrmFDmsk"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Dim Naik As Boolean
'Oval
Private Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function CreateEllipticRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Private Sub Command1_Click()
    Naik = False
    Timer1.Enabled = True
End Sub
Private Sub Form_Load()
 Call LetakanForm(Me, True)
FrmFDmsk.Picture = LoadPictureDLL(400) 'load image

    Top = ((GetSystemMetrics(17) + GetSystemMetrics(4)) * Screen.TwipsPerPixelY)
    Left = (GetSystemMetrics(16) * Screen.TwipsPerPixelX) - Width
    Naik = True
    
'membuat form oval
Dim L As Long
L = CreateRoundRectRgn(0, 0, Me.Width / Screen.TwipsPerPixelX, Me.Height / Screen.TwipsPerPixelY, 20, 20)
SetWindowRgn Me.hwnd, L, 0

End Sub
Private Sub CmdScanFD_Click(Index As Integer)
If Index = 0 Then
Call LetakanForm(Me, False)

Naik = False
    Timer1.Enabled = True
    Sleep 100
    frmMain.ScanFD
ElseIf Index = 1 Then
Call LetakanForm(Me, False)

Naik = False
    Timer1.Enabled = True
    Sleep 100
frmMain.ScanFDArea
ElseIf Index = 2 Then
Call LetakanForm(Me, False)
frmMain.BatalScanFD
Naik = False
    Timer1.Enabled = True
 Else
 
End If
End Sub

Private Sub Timer1_Timer()
    Const s = 80 'kecepatan gerak / slide
    Dim v As Single
    v = (GetSystemMetrics(17) + GetSystemMetrics(4)) * Screen.TwipsPerPixelY
    
    If Naik = True Then
        If Top - s <= v - Height Then
            Top = Top - (Top - (v - Height))
            Timer1.Enabled = False
        Else
            Top = Top - s
        End If
        
    Else
        Top = Top + s
        If Top >= v Then Unload Me
    End If
End Sub

