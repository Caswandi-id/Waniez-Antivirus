VERSION 5.00
Begin VB.UserControl rtp_mode 
   BackColor       =   &H00808080&
   CanGetFocus     =   0   'False
   ClientHeight    =   3150
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4695
   ClipControls    =   0   'False
   HasDC           =   0   'False
   InvisibleAtRuntime=   -1  'True
   ScaleHeight     =   3150
   ScaleWidth      =   4695
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "RTP_mod"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   75
      TabIndex        =   0
      Top             =   90
      Width           =   720
   End
End
Attribute VB_Name = "rtp_mode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Event PathChange(strPath As String)
Event FileNameSeletedChange(strFilename As String, Fullpath As String)
'Event FileNameSeletedStart()
Event IEClosed()

Dim WithEvents IEObject As SHDocVw.WebBrowser
Attribute IEObject.VB_VarHelpID = -1
Dim WithEvents SpaceIE  As shell32.ShellFolderView
Attribute SpaceIE.VB_VarHelpID = -1

Dim var_Quit As Boolean
Dim var_Enabled As Boolean
Dim var_hwnd As Long
Private Declare Function IsWindow Lib "user32" (ByVal hwnd As Long) As Long

Property Get IEKey() As Long
On Error GoTo salah
    IEKey = IEObject.hwnd
salah:
End Property

Function AddSubClass(Value As Object) As Boolean
On Error GoTo salah
  'If var_hwnd = 0 Or IsWindow(var_hwnd) = False Then
     Set IEObject = Nothing
     Set SpaceIE = Nothing
     
     Set IEObject = Value
     Set SpaceIE = IEObject.Document
         var_hwnd = Value.hwnd
         Call ChangePath
         SpaceIE_SelectionChanged
         AddSubClass = True
  'End If
  Exit Function
salah:
'MsgBox Error
End Function

Private Sub ChangePath()
On Error Resume Next
Dim Buff As String
Buff = ValidatePath(CStr(IEObject.LocationURL))
If Trim$(Buff) <> "" Then
   RaiseEvent PathChange(Buff)
End If
End Sub
Private Sub IEObject_OnQuit()
var_hwnd = 0
On Error GoTo salah
Set IEObject = Nothing
Set SpaceIE = Nothing
RaiseEvent IEClosed
salah:
End Sub
Private Sub IEObject_DocumentComplete(ByVal pDisp As Object, URL As Variant)
On Error Resume Next
    RaiseEvent PathChange(CStr(URL))

End Sub
Function ValidatePath(nStr As String) As String
On Error Resume Next
Dim Buff As String
Dim i As Integer
Buff = nStr
Buff = Replace(Buff, "file:///", "")
Buff = Replace(Buff, "/", "\")

For i = 32 To 255
    Buff = Replace(Buff, "%" & Hex$(i), Chr$(i), , , vbTextCompare)
Next i
ValidatePath = Buff
End Function


Private Sub IEObject_TitleChange(ByVal Text As String)
On Error Resume Next
     Set SpaceIE = Nothing
     Set SpaceIE = IEObject.Document
End Sub

Private Sub SpaceIE_SelectionChanged()
If var_Enabled Then
    Dim FI As Object
    'RaiseEvent FileNameSeletedStart
    For Each FI In SpaceIE.SelectedItems
        RaiseEvent FileNameSeletedChange(FI.Name, FI.path)
        DoEvents
    Next FI
    Set FI = Nothing
End If
End Sub


Sub SetIENothing()
On Error GoTo salah
Set IEObject = Nothing
Set SpaceIE = Nothing
var_hwnd = 0
salah:
End Sub



Sub EnabledMonitoring(Value As Boolean)
var_Enabled = Value
End Sub

'Private Sub UserControl_Terminate()
'On Error GoTo salah
'Set IEObject = Nothing
'Set SpaceIE = Nothing
'var_hwnd = 0
'salah:
'End Sub

Private Sub UserControl_Initialize()

End Sub
