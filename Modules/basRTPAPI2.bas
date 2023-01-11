Attribute VB_Name = "basRTPAPI2"
Dim eWindow As InternetExplorer
Private cWindow As New ShellWindows
Public Function GetTangkap() As String
Dim buffer As String
Dim cLocData As String
Dim Files As Collection
Dim clocation As String
'Dim rts As New RTimeStatus
    ' p1.Value = p1.Value + 1
     On Error Resume Next
    Timerrtp.Timer1.Enabled = False
            
    For Each eWindow In cWindow
    
        DoEvents
        If eWindow.Busy Then
            GoTo winBusy
        End If
        
        clocation = eWindow.LocationURL
        cLocData = InStr(1, buffer, eWindow.LocationName & "|" & eWindow.LocationURL & "|")
        
        If cLocData = 0 Then
            If Mid$(clocation, 1, 7) = "file://" Then
                 clocation = Replace(clocation, "file:///", "")
                 clocation = Replace(clocation, "%20", " ")
                 clocation = Replace(clocation, "/", "\")
                 
                ' rts.ActiveFileMonitor clocation, Files
                 'frmMain.Label3 = clocation
                   
            End If
        End If
        
winBusy:
        
    Next
    Timerrtp.Timer1.Enabled = True
    On Error GoTo 0
    
End Function
