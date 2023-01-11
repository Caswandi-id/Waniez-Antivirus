Attribute VB_Name = "ModFadeForm"
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal Color As Long, ByVal x As Byte, ByVal Alpha As Long) As Boolean


Private Const LWA_COLORKEY = 1
Private Const LWA_ALPHA = 2
Private Const LWA_BOTH = 3
Private Const WS_EX_LAYERED = &H80000
Private Const GWL_EXSTYLE = -20
Public Sub Explode(ByRef frm As Form, ByRef efek As Boolean)
  With frm
    .Width = 0
    .Height = 0
    .Show
    
    If efek Then
      For x = 0 To 4110 Step 5
        .Move (Screen.Width - .Width) / 2, (Screen.Height - .Height) / 2, 5730, x
        DoEvents
      Next
    Else
      For x = 10000 To 0 Step -50
        .Move (Screen.Width - .Width) / 2, (Screen.Height - .Height) / 2, x, x
        DoEvents
      Next
      End
    End If
  End With
End Sub

Sub SetTransparan(Form As Long, nilai As Integer)
On Error GoTo err

Dim attrib As Long
attrib = GetWindowLong(Form, GWL_EXSTYLE)
SetWindowLong Form, GWL_EXSTYLE, attrib Or WS_EX_LAYERED
SetLayeredWindowAttributes Form, RGB(255, 255, 0), nilai, LWA_ALPHA
Exit Sub

err:
    MsgBox "Error!" & vbCrLf & _
            "Source :" & vbTab & err.Source & vbCrLf & _
            "Reason :" & vbTab & err.reason _
            , vbCritical + vbOKOnly, "Error!"
End Sub

