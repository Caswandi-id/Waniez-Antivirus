Attribute VB_Name = "ModLockDrive"
Option Explicit

Private Const INVALID_HANDLE_VALUE = -1, FILE_ATTRIBUTE_ARCHIVE = &H20, FILE_ATTRIBUTE_DIRECTORY = &H10, FILE_ATTRIBUTE_HIDDEN = &H2, FILE_ATTRIBUTE_NORMAL = &H80, FILE_ATTRIBUTE_READONLY = &H1, FILE_ATTRIBUTE_SYSTEM = &H4, FILE_ATTRIBUTE_TEMPORARY = &H100
Private Declare Function GetFileAttributes Lib "kernel32" Alias "GetFileAttributesW" (ByVal lpFileName As Long) As Long
Private Declare Function SetFileAttributes Lib "kernel32" Alias "SetFileAttributesW" (ByVal lpFileName As Long, ByVal dwFileAttributes As Long) As Long
Private Declare Function DeleteFile Lib "kernel32" Alias "DeleteFileW" (ByVal lpFileName As Long) As Long
Private Declare Function CreateDirectory Lib "kernel32" Alias "CreateDirectoryW" (ByVal lpPathName As String, lpSecurityAttributes As SECURITY_ATTRIBUTES) As Long
Private Declare Function RemoveDirectory Lib "kernel32" Alias "RemoveDirectoryW" (ByVal lpPathName As String) As Long
Private Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long
Private Type SECURITY_ATTRIBUTES
        nLength As Long
        lpSecurityDescriptor As Long
        bInheritHandle As Long
End Type

Private Declare Function GetVolumeInformation Lib "kernel32" Alias "GetVolumeInformationA" (ByVal lpRootPathName As String, ByVal lpVolumeNameBuffer As String, ByVal nVolumeNameSize As Long, lpVolumeSerialNumber As Long, lpMaximumComponentLength As Long, lpFileSystemFlags As Long, ByVal lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Long) As Long
Public Function DriveLabel(ByVal sDrive As String) As String
    Const clMaxLen As Long = 100
    Dim lSerial      As Long, sDriveName As String * clMaxLen, sFileSystemName As String * clMaxLen
    sDrive = Left$(sDrive, 1) & ":\"
    'dapatkan info drive
    If GetVolumeInformation(sDrive, sDriveName, clMaxLen, lSerial, 0, 0, sFileSystemName, clMaxLen) Then
        DriveLabel = Left$(sDriveName, InStr(1, sDriveName, vbNullChar) - 1)
    Else
        DriveLabel = ""
    End If
    If DriveLabel <> "" Then Exit Function
    Select Case GetDriveType(sDrive)
    Case 3: DriveLabel = "Local Disk"
    Case 5: DriveLabel = "CD/DVD-Drive"
    End Select
End Function

Public Function DriveList(syarat As String, hasil As Collection)
    Dim DriveNum      As Integer, DriveType As Long, CekKunci As Boolean
    DriveNum = 64: Set hasil = New Collection
    Do Until DriveNum > 90
        DriveNum = DriveNum + 1: DriveType = GetDriveType(Chr$(DriveNum) & ":\")
        If InStr(syarat, CStr(DriveType)) <> 0 Then
            If DriveNum = 65 Then GoTo lanjutkan:
            hasil.Add UCase$(Chr$(DriveNum)), Chr$(DriveNum)
        End If
lanjutkan:
    Loop
End Function
Public Function HapusVirus(sPathDel As String) As Long
    On Error Resume Next
    SetFileAttributes StrPtr(sPathDel), FILE_ATTRIBUTE_NORMAL
    DeleteFile StrPtr(sPathDel)
End Function
Public Function GetSelect(lst As ucListView) As Long
Dim i As Long
If lst.ListItems.Count = 0 Then GetSelect = 0
For i = 1 To lst.ListItems.Count
    If lst.ListItems.Item(i).Selected = True Then
    GetSelect = i
    Exit Function
    End If
Next
End Function
Public Sub KunciFD(fdPath As String)
    On Error Resume Next
    Dim lenFD As SECURITY_ATTRIBUTES
    lenFD.nLength = Len(lenFD)
    Dim fAman As String: fAman = fdPath & ":\autorun.inf\"
    HapusVirus fdPath & ":\autorun.inf"
    CreateDirectory StrConv(fAman, vbUnicode), lenFD
    CreateDirectory StrConv(fAman & "kunci . \", vbUnicode), lenFD
    CreateDirectory StrConv(fAman & "Wan'iez Lock\", vbUnicode), lenFD
    CreateDirectory StrConv(fAman & "con\", vbUnicode), lenFD: CreateDirectory StrConv(fAman & "con\aux\", vbUnicode), lenFD: CreateDirectory StrConv(fAman & "con\aux\nul\", vbUnicode), lenFD
    CreateDirectory StrConv(fdPath & ":\Wan'iez Lock\", vbUnicode), lenFD
    SetFileAttributes StrPtr(fAman), FILE_ATTRIBUTE_HIDDEN Or FILE_ATTRIBUTE_READONLY Or FILE_ATTRIBUTE_SYSTEM
   ' If AdaFile(fAman & "desktop.ini") = False Then WriteINI ".ShellClassInfo", "CLSID", "{2559a1f2-21d7-11d4-bdaf-00c04f60b9f0}", fAman & "desktop.ini"
End Sub
Public Sub BukaFD(fdPath As String)
    On Error Resume Next
    Dim fAman As String: fAman = fdPath & ":\autorun.inf\"
    SetFileAttributes StrPtr(fAman), FILE_ATTRIBUTE_NORMAL
    'fUnic = fdPath & ":\ " & ChrW(9833) & " Limav " & ChrW(9833) & " \"
    RemoveDirectory StrConv(fAman & "kunci . \", vbUnicode)
    'MkDir fAman & " " & ChrW(9833) & " Limav " & ChrW(9833) & " \"
    RemoveDirectory StrConv(fAman & "con\aux\nul\", vbUnicode): RemoveDirectory StrConv(fAman & "con\aux\", vbUnicode): RemoveDirectory StrConv(fAman & "con\", vbUnicode)
    RemoveDirectory StrConv(fAman & "Wan'iez Lock\", vbUnicode)
    'CreateObject("scripting.filesystemobject").createfolder fAman & "con\aux\nul\"
    HapusVirus fAman & "desktop.ini"
    RemoveDirectory fAman
End Sub

