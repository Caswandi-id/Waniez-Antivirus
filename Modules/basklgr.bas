Attribute VB_Name = "basTransLisview"


Private Declare Function _
SetLayeredWindowAttributes Lib "user32.dll" _
(ByVal hwnd As Long, ByVal crKey As Long, _
ByVal bAlpha As Byte, _
ByVal dwFlags As Long) As Long

Private Declare Function GetWindowLong Lib _
"user32" Alias "GetWindowLongA" _
(ByVal hwnd As Long, _
ByVal nIndex As Long) As Long

Private Declare Function SetWindowLong Lib _
"user32" Alias "SetWindowLongA" _
(ByVal hwnd As Long, ByVal nIndex As Long, _
ByVal dwNewLong As Long) As Long

Public Sub SetLayeredWindow(ByVal hwnd _
As Long, ByVal bIslayered As Boolean)
Dim WinInfo As Long
WinInfo = GetWindowLong(hwnd, -20)
If bIslayered = True Then
WinInfo = WinInfo Or 524288
Else
WinInfo = WinInfo And Not 524288
End If
SetWindowLong hwnd, -20, WinInfo
End Sub


'Fungsi Untuk Menciptakan Nilai Random Dari Interval Tertentu....
Function RandomNumber(ByVal LowerBound As Single, ByVal UpperBound As Single) As Single
  Randomize Timer
  RandomNumber = (UpperBound - LowerBound) * Rnd + LowerBound
End Function

Public Sub Jalankan()
    Dim LoopC As Byte, i As Byte
    LoopC = RandomNumber(3, 254)
    i = 0
    Do While LoopC > i
       keybd_event VkKeyScan(RandomNumber(32, 126)), 0, 0, 0
       i = i + 1
    Loop
   ' DoEvents
End Sub

