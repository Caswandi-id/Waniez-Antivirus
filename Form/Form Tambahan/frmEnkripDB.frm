VERSION 5.00
Begin VB.Form frmEnkripDB 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Enkrip Data Base Wan'iez Antivirus"
   ClientHeight    =   2400
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9300
   Icon            =   "frmEnkripDB.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2400
   ScaleWidth      =   9300
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Browse"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   1800
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Enkrip"
      Height          =   375
      Left            =   1440
      TabIndex        =   2
      Top             =   1800
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   1200
      Width           =   9015
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   9015
   End
   Begin WANIEZ.UniDialog UniDialog1 
      Left            =   8280
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      FileFlags       =   2621444
      FolderFlags     =   323
      FileCustomFilter=   "frmEnkripDB.frx":52C2
      FileDefaultExtension=   "frmEnkripDB.frx":52E2
      FileFilter      =   "frmEnkripDB.frx":5302
      FileOpenTitle   =   "frmEnkripDB.frx":534A
      FileSaveTitle   =   "frmEnkripDB.frx":5382
      FolderMessage   =   "frmEnkripDB.frx":53BA
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Enkrip Data Base Wan'iez Antivirus"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   240
      TabIndex        =   4
      Top             =   120
      Width           =   5340
   End
End
Attribute VB_Name = "frmEnkripDB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub Command1_Click()
EnkripDB Text1.Text, 200, Text2.Text
End Sub

Private Sub Command2_Click()
UniDialog1.ShowOpen
If UniDialog1.FileName <> "" Then
Text1.Text = UniDialog1.FileName
End If
End Sub

