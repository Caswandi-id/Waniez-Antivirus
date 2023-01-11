VERSION 5.00
Begin VB.Form FormSplash 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   4110
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5715
   Icon            =   "FormSplash.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4110
   ScaleWidth      =   5715
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin WANIEZ.ucProgressBar Loading 
      Height          =   255
      Left            =   240
      Top             =   3720
      Width           =   3855
      _ExtentX        =   9551
      _ExtentY        =   450
   End
   Begin VB.Timer TmrAll 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   1200
      Top             =   3240
   End
   Begin VB.Timer tmrLoad 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   120
      Top             =   3240
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   25
      Left            =   840
      Top             =   3240
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   480
      Top             =   3240
   End
   Begin VB.Label LbInfoSummary 
      BackStyle       =   0  'Transparent
      Caption         =   "Build Number"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   5
      Top             =   2160
      Width           =   1335
   End
   Begin VB.Label LbBuildNumber 
      BackStyle       =   0  'Transparent
      Caption         =   ":  06"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1560
      TabIndex        =   4
      Top             =   2160
      Width           =   615
   End
   Begin VB.Label LbInfoSummary 
      BackStyle       =   0  'Transparent
      Caption         =   "Engine Version"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   3
      Top             =   1920
      Width           =   1335
   End
   Begin VB.Label LbEnggine 
      BackStyle       =   0  'Transparent
      Caption         =   ": 1.5.0.06"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1560
      TabIndex        =   2
      Top             =   1920
      Width           =   735
   End
   Begin VB.Label Copyright 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright © 2011 - 2013 Canvas Software."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   165
      Left            =   240
      TabIndex        =   1
      Top             =   2400
      Width           =   2835
   End
   Begin VB.Label lbktotDB 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
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
      TabIndex        =   0
      Top             =   3360
      Width           =   3615
   End
End
Attribute VB_Name = "FormSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim transparan As Integer
Dim Mulai As Boolean
'Oval
Private Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function CreateEllipticRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Private Sub Form_Load()
Call LetakanForm(Me, True): GaweTransparan Me.hwnd, 100
FormSplash.Picture = LoadPictureDLL(700) 'load image
'membuat form oval
Dim L As Long
L = CreateRoundRectRgn(0, 0, Me.Width / Screen.TwipsPerPixelX, Me.Height / Screen.TwipsPerPixelY, 20, 20)
SetWindowRgn Me.hwnd, L, 0
End Sub
Public Sub buka_splas()
Call LetakanForm(Me, True)
Explode Me, 1 'Up Animation
tmrLoad.Enabled = True
End Sub

Private Sub tmrLoad_Timer()
   With Loading

        If .Value < 100 Then
            DoEvents
           Call LetakanForm(Me, True): GaweTransparan Me.hwnd, 255
            .Value = .Value + 1
            If .Value = 20 Then
                lbktotDB.Caption = "Init Database...": DoEvents
                Call Init_Dtabase
            Tunggu 1
            End If
            If .Value = 30 Then
                lbktotDB.Caption = "Signature database: " & CStr(JumVirus) + JumlahVirusINT + JumlahVirusOUT & " Virus + Heuristic": DoEvents
                Tunggu 2
            End If
            If .Value = 40 Then
                lbktotDB.Caption = "Scanning Memori..."
                Tunggu 1
                If UCase$(Left$(Command, 2)) <> "-K" Then
                BERHENTI = False
                VirusFound = 0
                ScanProses False, lbktotDB
                BERHENTI = True
                End If
                Tunggu 1
            Sleep 100
            End If
            If .Value = 75 Then
                If VirusFound > 0 Then
                lbktotDB.Caption = "Wan'iez Antivirus Found Virus In Your Memori"
                Else
                lbktotDB.Caption = "Wan'iez Antivirus Not Found Virus In Memori"
                End If
                Tunggu 2
            End If
            If .Value = 80 Then

            With frmMain
            If FrmConfig.ck8.Value = 1 Then
                lbktotDB.Caption = "Status Enable Protection"
                Else
                lbktotDB.Caption = "Status Not Protection"
            End If
            End With
                 Tunggu 2
            End If

          If .Value = 100 Then
          
            Tunggu 1
            
            tmrLoad.Enabled = False
            If VirusFound > 0 Then
            GaweTransparan Me.hwnd, 100
             Unload Me
            MsgBox "Wan'iez Antivirus Found Virus In Your Memori  !", vbOKOnly + vbExclamation
           With frmMain
           .Show
           .WindowState = vbNormal
           .ckScan(0).Enabled = True: .cmdFixMalware.Enabled = True: .cmdFixMalwareAll.Enabled = True: .CmdQuarAll.Enabled = True
             End With
             End If
             GaweTransparan Me.hwnd, 100
            Unload Me
            End If
        End If
       End With
End Sub
Sub Tunggu(ByVal scd As Single)
Dim isStop As Boolean
On Error Resume Next
    Dim Mulai As Variant
    Mulai = Timer
    Do While Timer < Mulai + scd
      DoEvents
      If isStop = True Then Exit Do
    Loop
End Sub
Private Sub Timer3_Timer()
Mulai = False: Timer1.Enabled = True
End Sub
Private Sub TmrAll_Timer()
Timer1.Enabled = False: tmrLoad.Enabled = False: Timer3.Enabled = False
End Sub
Private Sub Timer1_Timer()
On Error Resume Next
If Mulai Then
  transparan = transparan + 5
  If transparan > 255 Then transparan = 255: Timer1.Enabled = False ': tmrLoad.Enabled = True
Else
  transparan = transparan - 3
  If transparan < 0 Then transparan = 0:   TmrAll.Enabled = True: Unload Me
End If
SetTransparan FormSplash.hwnd, transparan
End Sub
