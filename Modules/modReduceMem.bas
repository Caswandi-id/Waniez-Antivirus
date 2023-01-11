Attribute VB_Name = "modReduceMem"
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessID As Long) As Long
Private Declare Function EmptyWorkingSet Lib "psapi" (ByVal hProcess As Long) As Long
Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Public Sub ReduceMemory()
    Dim nHandle As Long, nReturn As Long, i_PID As Long
    i_PID = GetCurrentProcessId
    nHandle = OpenProcess(&H1F0FFF, 0, i_PID)
    nReturn = EmptyWorkingSet(nHandle)
    CloseHandle nHandle
End Sub

