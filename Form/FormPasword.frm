VERSION 5.00
Begin VB.Form FrmPassword 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Type your password"
   ClientHeight    =   1725
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3975
   Icon            =   "FormPasword.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   115
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   265
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin WANIEZ.rButton X 
      Height          =   495
      Left            =   3360
      TabIndex        =   3
      Top             =   1080
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
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
   Begin WANIEZ.rButton cmdok 
      Height          =   495
      Left            =   2880
      TabIndex        =   1
      Top             =   1080
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      ButtonStyle     =   7
      BackColor       =   14211288
      BackColorPressed=   15715986
      BackColorHover  =   16243621
      BorderColor     =   9408398
      BorderColorPressed=   6045981
      BorderColorHover=   11632444
      Caption         =   "Ok"
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
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
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
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   120
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   600
      Width           =   3735
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
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Top             =   1080
      Width           =   2655
   End
   Begin VB.Label lbtitle 
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
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
      TabIndex        =   4
      Top             =   60
      Width           =   3135
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
      Left            =   3960
      TabIndex        =   2
      Top             =   0
      Width           =   165
   End
End
Attribute VB_Name = "FrmPassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim s2 As String
'menggerakan form tanpa border
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Long) As Long
'akhirnya border
Private Sub Form_Load()
    Call LetakanForm(Me, True) 'frmdi depan
    lbInfo.Caption = ""
    FrmPassword.Picture = LoadPictureDLL(800) 'load image

End Sub
Public Function ShowMe(ByRef spWd As String) As Boolean
On Error GoTo ErrHandling
Me.Show vbModal
spWd = s2
ShowMe = True
Text1.SetFocus
ErrHandling:
End Function
Private Sub CmdOk_Click()
If Text1.Text = "" Then
Call LetakanForm(Me, True) 'frmdi depan
    lbInfo.Caption = "Password you entered is not Correct !"
    Text1.SetFocus
    Exit Sub
End If
s2 = Text1.Text
Unload Me
End Sub

Private Sub Image1_Click()

End Sub

Private Sub LbX_Click()
s2 = ""
Unload Me
End Sub

Private Sub X_Click()
s2 = ""
Unload Me
End Sub

Private Sub rButton2_Click()

End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then 'enter
Call CmdOk_Click
End If
If KeyAscii = 27 Then 'enter
End If
End Sub
