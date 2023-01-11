Attribute VB_Name = "MLiteTimer"
Option Explicit

'i think this timer came from vb2themax, but im not sure where i got it anymore...
'much love and thanks to whoever coded it though :)

Private Declare Function SetTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Private Declare Function KillTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (dest As Any, Source As Any, ByVal Bytes As Long)

Private Const WM_TIMER = &H113

Private mobjTimers As Collection

Private Sub TimerProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal idEvent As Long, ByVal dwTime As Long)
  Dim cTime As CLiteTimer
    
  On Error GoTo ErrorHandler
    ' Make sure that the message is WM_TIMER.
    If uMsg = WM_TIMER Then
        ' It is a timer event.
            Set cTime = mobjTimers("T" & CStr(idEvent))
            cTime.TimerCallBack idEvent
    End If

Exit Sub

ErrorHandler:
   Debug.Print "TimerProc Error " & err.Number & ": " & err.Description
   
End Sub

Public Function StartTimer(ByVal objTimer As CLiteTimer, ByVal lngInterval As Long, ByVal lngTimerID As Long) As Long

 On Error GoTo ErrorHandler
    ' Create the collection to store the timers if it hasn't been already.
    If mobjTimers Is Nothing Then
        Set mobjTimers = New Collection
    End If

    ' Check to see if the timer is already running.
    If lngTimerID = 0 Then
        If lngInterval > 0 Then

            lngTimerID = SetTimer(0, 0, lngInterval, AddressOf TimerProc)
            mobjTimers.Add objTimer, "T" & lngTimerID

        End If
    End If

    StartTimer = lngTimerID

Exit Function
    
ErrorHandler:
    Debug.Print "StartTimer Error " & err.Number & ": " & err.Description
    
End Function

Public Sub StopTimer(ByRef lngTimerID As Long)

   On Error GoTo ErrorHandler
    
    If Not (mobjTimers Is Nothing) Then
        ' Is the timer running?
        If lngTimerID > 0 Then
            ' The timer is running. Kill it.
            If KillTimer(0, lngTimerID) <> 0 Then

                mobjTimers.Remove "T" & lngTimerID
                lngTimerID = 0

                If mobjTimers.Count = 0 Then
                    Set mobjTimers = Nothing
                End If
            End If
        End If
    End If

    Exit Sub
    
ErrorHandler:
    Debug.Print "StopTimer Error " & err.Number & ": " & err.Description
    
End Sub



