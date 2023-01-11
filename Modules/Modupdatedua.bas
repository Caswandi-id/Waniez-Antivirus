Attribute VB_Name = "basupdatedua"
'download without inet
Private Declare Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" (ByVal pCaller As Long, ByVal szURL As String, ByVal szFileName As String, ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long
Private Declare Function DoFileDownload Lib "shdocvw.dll" (ByVal lpszFile As String) As Long
'get readfile from internet without download
Const scUserAgent = "API-Guide test program"
Const INTERNET_OPEN_TYPE_DIRECT = 1
Const INTERNET_OPEN_TYPE_PROXY = 3
Const INTERNET_FLAG_RELOAD = &H80000000
Const sURL = "http://waniez.p.ht/download/upx/"

Private Declare Function InternetOpen Lib "wininet" Alias "InternetOpenA" (ByVal sAgent As String, ByVal lAccessType As Long, ByVal sProxyName As String, ByVal sProxyBypass As String, ByVal _
lFlags As Long) As Long
Private Declare Function InternetCloseHandle Lib "wininet" (ByVal hInet As Long) As Integer
Private Declare Function InternetReadFile Lib "wininet" (ByVal hFile As Long, ByVal sBuffer As String, ByVal lNumBytesToRead As Long, lNumberOfBytesRead As Long) As Integer
Private Declare Function InternetOpenUrl Lib "wininet" Alias "InternetOpenUrlA" (ByVal hInternetSession As Long, ByVal lpszUrl As String, ByVal lpszHeaders As String, ByVal dwHeadersLength As Long, ByVal dwFlags As Long, ByVal dwContext As Long) As Long
'check connection
Private Const FLAG_ICC_FORCE_CONNECTION = &H1
Private Declare Function InternetCheckConnection Lib "wininet.dll" Alias "InternetCheckConnectionA" (ByVal lpszUrl As String, ByVal dwFlags As Long, ByVal dwReserved As Long) As Long

Private Function ReadInternet(URL As String) As String
    Dim hOpen      As Long, hFile As Long, sBuffer As String, ret As Long
    sBuffer = Space$(10)
    hOpen = InternetOpen(scUserAgent, INTERNET_OPEN_TYPE_DIRECT, vbNullString, vbNullString, 0)
    hFile = InternetOpenUrl(hOpen, URL, vbNullString, ByVal 0&, INTERNET_FLAG_RELOAD, ByVal 0&)
    InternetReadFile hFile, sBuffer, 10, ret
    InternetCloseHandle hFile
    InternetCloseHandle hOpen
    ReadInternet = sBuffer
End Function
Private Function downloadfile(URL As String, LocalFilename As String) As Boolean
    Dim lngRetVal      As Long
    lngRetVal = URLDownloadToFile(0, URL, LocalFilename, 0, 0)
    If lngRetVal = 0 Then downloadfile = True
End Function
Private Function DoDownload(URL As String)
    DoFileDownload StrConv(URL, vbUnicode)
End Function

'new pindah
Public Function CekKoneksi(Server As String) As Boolean
    If InternetCheckConnection(Server, FLAG_ICC_FORCE_CONNECTION, 0&) = 0 Then
        CekKoneksi = False
    Else
        CekKoneksi = True
    End If
End Function

