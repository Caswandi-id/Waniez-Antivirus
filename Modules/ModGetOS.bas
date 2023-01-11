Attribute VB_Name = "ModGetOS"
Private Declare Function GetVersionExA Lib "kernel32.dll" _
       (lpVersionInformation As OSVERSIONINFO) As Long
 
Private Type OSVERSIONINFO
   dwOSVersionInfoSize As Long
   dwMajorVersion      As Long
   dwMinorVersion      As Long
   dwBuildNumber       As Long
   dwPlatformId        As Long
   szCSDVersion        As String * 128
End Type

' little note
'PlatformId = 2
'major 6 minor 0 = vista
'major 6 minor 1 = winseven
'major 5 minor 0 = windows 2000
'major 5 minor 1 = winxp
'major 5 minor 2 = windows 2003
'major 4         = winnt

Function GetBasicOS() As String
Dim OSINFO As OSVERSIONINFO
Dim RetVal As Long

       OSINFO.dwOSVersionInfoSize = 148
       OSINFO.szCSDVersion = Space$(128)
       RetVal = GetVersionExA(OSINFO)
           
       With OSINFO
       
       Select Case .dwPlatformId
         Case 1
           Select Case .dwMinorVersion
             Case 0
                GetBasicOS = "Windows 95"
             Case 10
                GetBasicOS = "Windows 98"
             Case 90
                GetBasicOS = "Windows ME"
             End Select
         Case 2 ' Windows NT begin from here (until Win 7)
           Select Case .dwMajorVersion
             Case 3
                GetBasicOS = "Windows NT 3.51"
             Case 4
                GetBasicOS = "Windows NT 4.0"
             Case 5
                If .dwMinorVersion = 0 Then
                   GetBasicOS = "Windows 2000"
                Else
                   GetBasicOS = "Windows XP"
                End If
             Case 6 ' Vista and 7 begin here
                If .dwMinorVersion = 0 Then
                   GetBasicOS = "Windows Vista"
                ElseIf .dwMinorVersion = 1 Then
                   GetBasicOS = "Windows 7"
                   Else
                   GetBasicOS = "Windows 8"
                End If
                
             End Select
          Case Else
            GetBasicOS = "Failed"
        End Select
 
        End With
 
End Function

' Basic
' information about OS version begin from Windows NT saved on regitry with path :
' HKEY_LOCAL_MACHINE\Software\Microsoft\Windows NT\CurrentVersion
Public Function GetOS()
Dim szSPVersion As String
Dim szBuildNum As String
Dim szVersion  As String
Dim szProdukName As String
Dim szBasicOs    As String

szBasicOs = GetBasicOS

' In Vista/7 we using value on EditionID
' In XP or before we using value CSDVersion
If szBasicOs = "Windows Vista" Or szBasicOs = "Windows 7" Then
   szSPVersion = GetSTRINGValue(SingkatanKey("HKLM"), "Software\Microsoft\Windows NT\CurrentVersion", "EditionID")
Else
   szSPVersion = GetSTRINGValue(SingkatanKey("HKLM"), "Software\Microsoft\Windows NT\CurrentVersion", "CSDVersion")
End If

szBuildNum = GetSTRINGValue(SingkatanKey("HKLM"), "Software\Microsoft\Windows NT\CurrentVersion", "CurrentBuildNumber")
szVersion = GetSTRINGValue(SingkatanKey("HKLM"), "Software\Microsoft\Windows NT\CurrentVersion", "CurrentVersion")
szProdukName = GetSTRINGValue(SingkatanKey("HKLM"), "Software\Microsoft\Windows NT\CurrentVersion", "ProductName")


'FrmAbout.Lblinfo.Caption = "Basic OS: " & szBasicOs & ", Edition: " & szSPVersion & ", Version Number: " & szVersion & ", Build Number: " & szBuildNum & ", Product Name: " & szProdukName & ", Processor: " & Environ("PROCESSOR_ARCHITECTURE") & ""
frmMain.LbWindows(0) = "Computer Name : " & Environ("COMPUTERNAME")
frmMain.LbWindows(1) = "User Name :" & Environ("USERNAME")
frmMain.LbWindows(2) = "Basic OS : " & szBasicOs & ", Edition: " & szSPVersion & ", Version Number: " & szVersion & ", Build Number: " & szBuildNum & ", Product Name: " & szProdukName & ", Processor: " & Environ("PROCESSOR_ARCHITECTURE") & ""
'MsgBox "Long Description : " & szProdukName & " Version : " & szVersion & _
       "." & szBuildNum & " " & szSPVersion
End Function



