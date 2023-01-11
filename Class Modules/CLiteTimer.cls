VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CLiteTimer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mlngInterval As Long ' The timer interval.
Private mlngTimerID As Long

Public Event Timer()

Public Property Let Enabled(ByVal blnEnabled As Boolean)
On Error GoTo ErrorHandler

    ' Is the timer being started?
    If blnEnabled Then
        ' Is the timer currently running?
        If Not (mlngTimerID > 0) Then
            ' Start the timer.
            mlngTimerID = StartTimer(Me, mlngInterval, mlngTimerID)
        End If
    Else
        ' Stop the timer.
        StopTimer mlngTimerID
    End If

    Exit Property
    
ErrorHandler:
    Debug.Print "Enabled Let Error " & err.Number & ": " & err.Description
End Property

Public Property Get Enabled() As Boolean
On Error GoTo ErrorHandler

    Enabled = (mlngTimerID > 0)
    
    Exit Property
    
ErrorHandler:
    Debug.Print "Enabled Get Error " & err.Number & ": " & err.Description
End Property

Public Property Let Interval(ByVal lngInterval As Long)
On Error GoTo ErrorHandler

    ' Check the interval value.
    If lngInterval > 0 Then
        ' Check to see if the timer is currently running.
        If (mlngTimerID > 0) Then
            ' Stop the timer.
            StopTimer mlngTimerID
            
            ' Restart the timer with the new interval.
            mlngTimerID = StartTimer(Me, lngInterval, mlngTimerID)
        End If

        ' Store the new interval value.
        mlngInterval = lngInterval
    Else
        mlngInterval = 0
        
        ' Stop the timer.
        MLiteTimer.StopTimer mlngTimerID
    End If

    Exit Property
    
ErrorHandler:
    Debug.Print "Interval Let Error " & err.Number & ": " & err.Description
End Property

Public Property Get Interval() As Long
On Error GoTo ErrorHandler

    Interval = mlngInterval
    
    Exit Property
    
ErrorHandler:
    Debug.Print "Interval Get Error " & err.Number & ": " & err.Description
End Property

Friend Sub TimerCallBack(ByVal lngTimerID As Long)
On Error GoTo ErrorHandler

    ' Raise the event if the timer id's match.
    If mlngTimerID = lngTimerID Then
        RaiseEvent Timer
    End If

    Exit Sub

ErrorHandler:
    Debug.Print "TimerCallBack Error " & err.Number & ": " & err.Description
End Sub

Private Sub Class_Initialize()
    
    mlngInterval = 10

End Sub

Private Sub Class_Terminate()

On Error GoTo ErrorHandler

    ' Stop the timer if it is running.
    If (mlngTimerID > 0) Then
        StopTimer mlngTimerID
    End If
     
    Exit Sub
    
ErrorHandler:
    Debug.Print "Class_Initialize Error " & err.Number & ": " & err.Description
End Sub

