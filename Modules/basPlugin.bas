Attribute VB_Name = "basPlugin"
Option Explicit
Public Type PG_PLUGIN_GENERAL_INFORMATION
    szPluginName                        As String 'misal: ---> IGuaNaByteFix
    szPluginDescription                 As String 'misal: ---> Bla-Bla-Bla fixer
    szPluginAuthor                      As String 'misal: ---> IGuaNa
    szPluginAuthorSite                  As String 'misal: ---> www.iguana.com
    szPluginAuthorEMail                 As String 'misal: ---> iguana@gmail.com
    szPluginAuthorAddress               As String 'misal: ---> somewhere
    szPluginValidationID                As String 'misal: ---> 3B616451572CCA2DF5FB379501000000004D3000 --->[reverse:CRC32 untuk semua section yg dapat dieksekusi]+[reverse:CRC32 untuk semua section yg tidak dapat dieksekusi]+[reverse:total CRC32 yg dapat dieksekusi terhadap struktur value total CRC32 yg tidak dapat dieksekusi]+[reverse:2 ^ jumlah section]+[reverse:panjang semua section yg ada isinya].
    szPluginStartupPathW                As String 'misal: ---> C:\Program Files\CMC\Plugins\IGByteFix.dll
    nPluginCharacteristic               As Long 'misal: ---> 0 = tidak aktif di memori, 1 = aktif di memori.
    nPluginVersionMajor                 As Long 'misal:versi 1.5.400 ---> 1
    nPluginVersionMinor                 As Long 'misal:versi 1.5.400 ---> 5
    nPluginVersionRevision              As Long 'misal:versi 1.5.400 ---> 400
    nPluginIsRunAsActiveX               As Long 'sebenarnya bernama safe-multithreaded, bernilai 1 bila dapat dipergunakan secara multithreading (untuk std-dll), bernilai 0 untuk non-safe-multithreading (untuk activex-dll).
End Type

Private Type MY_PLUGIN_GENERAL_INFORMATION '---1548 bytes.
    szPluginName                        As String * 128 '---128 Character unicode = 256 bytes. 'misal: ---> IGuaNaByteFix
    szPluginDescription                 As String * 128 '---128 Character unicode = 256 bytes. 'misal: ---> Bla-Bla-Bla fixer
    szPluginAuthor                      As String * 128 '---128 Character unicode = 256 bytes. 'misal: ---> IGuaNa
    szPluginAuthorSite                  As String * 128 '---128 Character unicode = 256 bytes. 'misal: ---> www.iguana.com
    szPluginAuthorEMail                 As String * 128 '---128 Character unicode = 256 bytes. 'misal: ---> iguana@gmail.com
    szPluginAuthorAddress               As String * 128 '---128 Character unicode = 256 bytes. 'misal: ---> SomeWhere
    nPluginVersionMajor                 As Long 'misal:versi 1.5.400 ---> 1
    nPluginVersionMinor                 As Long 'misal:versi 1.5.400 ---> 5
    nPluginVersionRevision              As Long 'misal:versi 1.5.400 ---> 400
End Type

Private Const szStdLibFuncGetPluginInfo     As String = "CPgGetInfo0" '---CMC Plugin Get Informations
Private Const szStdLibFuncPluginGo          As String = "CPgGo0" '---CMC Plugin Go (lakukan sesuai keahlian masing-masing).
Private Const szStdLibFuncDllGetClassObject As String = "DllGetClassObject" '---fungsi yg selalu ada pada activex dll.
Private Const szStdLibFuncDllRegisterServer As String = "DllRegisterServer" '---fungsi yg selalu ada pada activex dll untuk registrasi.

Private Const szStdPluginClassName          As String = "ClsCMCPlugin_0"

Private Type IMAGE_NT_HEADERS
    SignatureLow            As Integer '2 "PE"
    SignatureHigh           As Integer '2
    FileHeader              As IMAGE_FILE_HEADER '20
End Type


Private Const IMAGE_SCN_MEM_EXECUTE     As Long = &H20000000    '---Section is executable.
Private Const IMAGE_SCN_MEM_READ        As Long = &H40000000    '---Section is readable.
Private Const IMAGE_SCN_MEM_WRITE       As Long = &H80000000    '---Section is writeable.

Private Const MAX_PATH                  As Long = 260               '00-FF
Private Const MAX_BUFFER                As Long = (MAX_PATH * 2)    '00 00 - FF FF

Private Const SYNCHRONIZE = &H100000    'penting! sinkronisasi data dan akses dengan proses lain.
Private Const READ_CONTROL = &H20000    'penting! ijin untuk mengoperasikan file.
Private Const FILE_READ_DATA = (&H1)    'penting! operasi: membaca file.
Private Const FILE_WRITE_DATA = (&H2)   'penting! operasi: menulis file.

Private Const FILE_SHARE_READ = &H1     'dapat diakses baca oleh proses lain.
Private Const FILE_SHARE_WRITE = &H2    'dapat diakses tulis oleh proses lain.
Private Const FILE_SHARE_DELETE = &H4   'dapat diakses hapus oleh proses lain.

Private Const FILE_ATTRIBUTE_DIRECTORY = &H10
Private Const FILE_ATTRIBUTE_NORMAL = &H80 'untuk file standar.

'operasi alternatif untuk file yang akan dibuat ataupun dibuka:
Private Const FILE_DISPOSE_CREATE_NEW = 1           'hanya akan membuat file baru. bila file sudah ada sebelumnya, fungsi gagal.
Private Const FILE_DISPOSE_CREATE_ALWAYS = 2        'hapus file yang lama (bila ada), dan akan membuat file yang baru.
Private Const FILE_DISPOSE_OPEN_EXISTING = 3        'hanya akan membuka file yang sudah ada, bila file tidak ada, fungsi gagal.
Private Const FILE_DISPOSE_OPEN_ALWAYS = 4          'membuka file yang ada (bila ada), dan akan membuat file yang baru bila file belum ada.
Private Const FILE_DISPOSE_TRUNCATE_EXISTING = 5    'membuka file yang sudah ada, dan menghapus semua isinya terlebih dahulu. fungsi gagal bila file tidak ada.


Private Const LMEM_FIXED = &H0
Private Const LMEM_MOVEABLE = &H2
Private Const LMEM_NOCOMPACT = &H10
Private Const LMEM_NODISCARD = &H20
Private Const LMEM_ZEROINIT = &H40
Private Const LMEM_MODIFY = &H80
Private Const LMEM_DISCARDABLE = &HF00
Private Const LMEM_VALID_FLAGS = &HF72
Private Const LMEM_INVALID_HANDLE = &H8000

Private Const LHND = LMEM_MOVEABLE + LMEM_ZEROINIT
Private Const lptr = LMEM_FIXED + LMEM_ZEROINIT


Private Type FILETIME
    dwLowDateTime       As Long
    dwHighDateTime      As Long
End Type

Private Type WIN32_FIND_DATA_W
    dwFileAttributes    As Long
    ftCreationTime      As FILETIME
    ftLastAccessTime    As FILETIME
    ftLastWriteTime     As FILETIME
    nFileSizeHigh       As Long
    nFileSizeLow        As Long
    dwReserved0         As Long
    dwReserved1         As Long
    cszFileName         As String * 260
    cszAlternFileName   As String * 14
End Type

Private Declare Function FindFirstFileW Lib "kernel32.dll" (ByVal pv_lpFileName As Long, ByVal pv_lpFindFileData As Long) As Long
Private Declare Function FindNextFileW Lib "kernel32.dll" (ByVal hFindFile As Long, ByVal pv_lpFindFileData As Long) As Long
Private Declare Function FindClose Lib "kernel32.dll" (ByVal hFindFile As Long) As Long

Private Declare Function GetFileAttributesW Lib "kernel32.dll" (ByVal pv_lpFileName As Long) As Long

Private Declare Function CreateFileW Lib "kernel32.dll" (ByVal lpFileName As Long, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, lpSecurityAttributes As Any, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long

Private Declare Function GetFileSize Lib "kernel32.dll" (ByVal hFile As Long, lpFileSizeHigh As Long) As Long

Private Declare Function SetFilePointer Lib "kernel32.dll" (ByVal hFile As Long, ByVal lDistanceToMove As Long, lpDistanceToMoveHigh As Long, ByVal dwMoveMethod As Long) As Long

Private Declare Function ReadFile Lib "kernel32.dll" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, lpOverlapped As Any) As Long
Private Declare Function WriteFile Lib "kernel32.dll" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, lpOverlapped As Any) As Long

Private Declare Function CloseHandle Lib "kernel32.dll" (ByVal hObject As Long) As Long

Private Declare Function LocalAlloc Lib "kernel32.dll" (ByVal uFlags As Long, ByVal lBytes As Long) As Long
Private Declare Function LocalSize Lib "kernel32.dll" (ByVal hMem As Long) As Long
Private Declare Function LocalFree Lib "kernel32.dll" (ByVal hMem As Long) As Long

Private Declare Function LoadResource Lib "kernel32.dll" (ByVal hInstance As Long, ByVal hResInfo As Long) As Long
Private Declare Function FindResourceW Lib "kernel32.dll" (ByVal hInstance As Long, ByVal pv_lpName As Long, ByVal pv_lpType As Long) As Long

Private Declare Sub RtlMoveMemory Lib "NTDLL.DLL" (ByVal pDestBuffer As Long, ByVal pSourceBuffer As Long, ByVal nBufferLengthToMove As Long) '<---sebenarnya namanya kurang sesuai, karena yang dilakukan adalah menyalin (copy) isi dari src ke dst.
Private Declare Sub RtlFillMemory Lib "NTDLL.DLL" (ByVal pDestBuffer As Long, ByVal nDestLengthToFill As Long, ByVal nByteNumber As Long) '<---harusnya byte,tapi memori 32 bit, jadi nggak apa-apa, asal tetap bernilai antara 0 sampai 255.
Private Declare Sub RtlZeroMemory Lib "NTDLL.DLL" (ByVal pDestBuffer As Long, ByVal nDestLengthToFillWithZeroBytes As Long) '<---reset isi dst yaitu mengisinya dengan bytenumber = 0.

Private Declare Function LoadLibraryW Lib "kernel32.dll" (ByVal pv_lpLibFileName As Long) As Long
Private Declare Function GetModuleHandleW Lib "kernel32.dll" (ByVal pv_lpModuleName As Long) As Long
Private Declare Function GetProcAddress Lib "kernel32.dll" (ByVal hModule As Long, ByVal szlpProcName As String) As Long
Private Declare Function FreeLibrary Lib "kernel32.dll" (ByVal hLibModule As Long) As Long

Private Declare Function CreateThread Lib "kernel32.dll" (ByVal lpThreadAttributes As Long, ByVal dwStackSize As Long, ByVal StartAddress As Long, ByVal lpParameter As Long, ByVal dwCreationFlags As Long, ByRef lpThreadId As Long) As Long

'---waduh, vb6 nggak punya fungsi untuk "Call by Pointer", jadi pakai cara-akal-akalan ajah, deh:
Private Declare Function CallWindowProcW Lib "user32.dll" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

'---penghitung crc32 tercepat dlm sebuah sistem 32 bit:
Private Declare Function RtlComputeCrc32 Lib "NTDLL.DLL" (ByVal dwInitialize As Long, ByVal pBufferData As Long, ByVal pBufferDataLen As Long) As Long


'---format MSFT, q suka cara brutal untuk belajar, langsung ke raw data, lama tapi diharapkan lebih nyambung :)
Private Const MSFT_SIGNATURE = &H5446534D          '  "MSFT"
Private Const SLTG_SIGNATURE = &H47544C53          '  "SLTG"

Private Type TLB_HEADER
    Magic1              As Long
    Magic2              As Long
    oGUID               As Long
    LCID                As Long
    LCID2               As Long
    fVar                As Long
    Version             As Long
    Flags               As Long
    nTypeInfo           As Long
    HelpStr             As Long
    HelpStrCnt          As Long
    HelpCntxt           As Long
    nName               As Long
    nChars              As Long
    oName               As Long
    HelpFile            As Long
    CustDat             As Long
    Res1                As Long
    Res2                As Long
    oDispatch           As Long
    nImpInfos           As Long
    'oFileName          As Long '---offset to typelib file name in string table
End Type

Type MSFT_NAMEINTRO
    hRefType            As Long
    NextHash            As Long
    cName               As Long
End Type

Private Type MSFT_SEGDESC
    Offs                As Long
    nLen                As Long
    Res01               As Long
    Res02               As Long
End Type

Type MSFT_SEGDIR
    pTypInfo            As MSFT_SEGDESC
    pImpInfo            As MSFT_SEGDESC
    pImpFiles           As MSFT_SEGDESC
    pRefer              As MSFT_SEGDESC
    pLibs               As MSFT_SEGDESC
    pGUID               As MSFT_SEGDESC
    Unk01               As MSFT_SEGDESC
    pNames              As MSFT_SEGDESC
    pStrings            As MSFT_SEGDESC
    pTypDesc            As MSFT_SEGDESC
    pArryDesc           As MSFT_SEGDESC
    pCustData           As MSFT_SEGDESC
    pCDGuids            As MSFT_SEGDESC
    Unk02               As MSFT_SEGDESC
    Unk03               As MSFT_SEGDESC
End Type


Public Function PgGetPluginDefaultValidationID(ByVal szPluginFileNameW As String) As String
On Error Resume Next '---result:validationId yg seharusnya dipakai untuk plugin tersebut.
Dim hFile           As Long
    hFile = CreateFileW(StrPtr(szPluginFileNameW), SYNCHRONIZE Or READ_CONTROL Or FILE_READ_DATA, FILE_SHARE_READ, ByVal 0, FILE_DISPOSE_OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, 0)
    If hFile <= 0 Then
        PgGetPluginDefaultValidationID = vbNullString
        GoTo LBL_TERAKHIR
    End If
Dim nFileLen        As Long
    nFileLen = GetFileSize(hFile, ByVal 0)
Dim pBuffer         As Long
Dim nBufferLength   As Long
Dim nCallResult     As Long
Dim nRetLength      As Long
Dim IDOSH           As IMAGE_DOS_HEADER
Dim IMNTH           As IMAGE_NT_HEADERS
Dim ISECH()         As IMAGE_SECTION_HEADER
Dim pVarPos         As Long
Dim nVarData        As Long
Dim CTurn           As Long
Dim nCRC32Exec      As Long
Dim nCRC32Norm      As Long
Dim nCompAll        As Long
Dim nLenAllSections As Long
    If nFileLen <= Len(IDOSH) Then
        PgGetPluginDefaultValidationID = vbNullString
        GoTo LBL_FREE_OBJECTS
    End If
    nBufferLength = nFileLen
    pBuffer = LocalAlloc(lptr, nBufferLength)
    If pBuffer <= 0 Then
        PgGetPluginDefaultValidationID = vbNullString
        GoTo LBL_FREE_OBJECTS
    End If
    nCallResult = ReadFile(hFile, ByVal pBuffer, nFileLen, nRetLength, ByVal 0)
    If nRetLength <> nFileLen Then
        PgGetPluginDefaultValidationID = vbNullString
        GoTo LBL_FREE_OBJECTS
    End If
    pVarPos = 0
    Call RtlMoveMemory(VarPtr(IDOSH), pBuffer + pVarPos, Len(IDOSH))
    If IDOSH.e_magic <> &H5A4D Then
        PgGetPluginDefaultValidationID = vbNullString
        GoTo LBL_FREE_OBJECTS
    End If
    If IDOSH.e_lfanew >= (nFileLen - Len(IMNTH)) Then
        PgGetPluginDefaultValidationID = vbNullString
        GoTo LBL_FREE_OBJECTS
    End If
    pVarPos = IDOSH.e_lfanew
    Call RtlMoveMemory(VarPtr(IMNTH), pBuffer + pVarPos, Len(IMNTH))
    If IMNTH.SignatureLow <> &H4550 Then
        PgGetPluginDefaultValidationID = vbNullString
        GoTo LBL_FREE_OBJECTS
    End If
    If IMNTH.FileHeader.NumberOfSections <= 0 Then
        PgGetPluginDefaultValidationID = vbNullString
        GoTo LBL_FREE_OBJECTS
    End If
    ReDim ISECH(IMNTH.FileHeader.NumberOfSections - 1) As IMAGE_SECTION_HEADER
    pVarPos = pVarPos + Len(IMNTH) + IMNTH.FileHeader.SizeOfOptionalHeader
    If pVarPos >= (nFileLen - (Len(ISECH(0)) * IMNTH.FileHeader.NumberOfSections)) Then
        PgGetPluginDefaultValidationID = vbNullString
        GoTo LBL_FREE_OBJECTS
    End If
    Call RtlMoveMemory(VarPtr(ISECH(0)), pBuffer + pVarPos, (Len(ISECH(0)) * IMNTH.FileHeader.NumberOfSections))
    For CTurn = 0 To (IMNTH.FileHeader.NumberOfSections - 1)
        If ISECH(CTurn).SizeOfRawData > 0 Then
            If (ISECH(CTurn).PointerToRawData + ISECH(CTurn).SizeOfRawData) <= nFileLen Then
                If (ISECH(CTurn).Characteristics And IMAGE_SCN_MEM_EXECUTE) = IMAGE_SCN_MEM_EXECUTE Then
                    nCRC32Exec = RtlComputeCrc32(nCRC32Exec, pBuffer + ISECH(CTurn).PointerToRawData, ISECH(CTurn).SizeOfRawData)
                Else
                    nCRC32Norm = RtlComputeCrc32(nCRC32Norm, pBuffer + ISECH(CTurn).PointerToRawData, ISECH(CTurn).SizeOfRawData)
                End If
            End If
            nLenAllSections = nLenAllSections + ISECH(CTurn).SizeOfRawData
        End If
    Next
    nCompAll = RtlComputeCrc32(nCRC32Exec, VarPtr(nCRC32Norm), Len(nCRC32Norm))
    PgGetPluginDefaultValidationID = Left$(StrReverse(PgSetHexStringLen(Hex$(nCRC32Exec), 8)) & StrReverse(PgSetHexStringLen(Hex$(nCRC32Norm), 8)) & StrReverse(PgSetHexStringLen(Hex$(nCompAll), 8)) & StrReverse(PgSetHexStringLen(Hex$(2 ^ IMNTH.FileHeader.NumberOfSections), 8)) & StrReverse(PgSetHexStringLen(Hex$(nLenAllSections), 8)), 40)
LBL_FREE_OBJECTS:
    Call RtlZeroMemory(VarPtr(IDOSH), Len(IDOSH))
    Call RtlZeroMemory(VarPtr(IMNTH), Len(IMNTH))
    Erase ISECH()
    If pBuffer > 0 Then
        Call LocalFree(pBuffer)
        pBuffer = 0
    End If
    If hFile > 0 Then
        Call CloseHandle(hFile)
        hFile = 0
    End If
LBL_TERAKHIR:
    If err.Number <> 0 Then
        err.Clear
    End If
End Function

Public Function PgGetPluginRealValidationID(ByVal szPluginFileNameW As String) As String
On Error Resume Next '---result:validationId yg terpasang pada plugin saat ini.
Dim hFile           As Long
    hFile = CreateFileW(StrPtr(szPluginFileNameW), SYNCHRONIZE Or READ_CONTROL Or FILE_READ_DATA, FILE_SHARE_READ, ByVal 0, FILE_DISPOSE_OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, 0)
    If hFile <= 0 Then
        PgGetPluginRealValidationID = vbNullString
        GoTo LBL_TERAKHIR
    End If
Dim nFileLen        As Long
    nFileLen = GetFileSize(hFile, ByVal 0)
Dim pBuffer         As Long
Dim nBufferLength   As Long
Dim nCallResult     As Long
Dim nRetLength      As Long
Dim IDOSH           As IMAGE_DOS_HEADER
Dim IMNTH           As IMAGE_NT_HEADERS
Dim ISECH()         As IMAGE_SECTION_HEADER
Dim pVarPos         As Long
Dim nVarData        As Long
Dim CTurn           As Long
Dim nDefaultSize    As Long
Dim nLenAllSections As Long
    If nFileLen <= Len(IDOSH) Then
        PgGetPluginRealValidationID = vbNullString
        GoTo LBL_FREE_OBJECTS
    End If
    nBufferLength = nFileLen
    pBuffer = LocalAlloc(lptr, nBufferLength)
    If pBuffer <= 0 Then
        PgGetPluginRealValidationID = vbNullString
        GoTo LBL_FREE_OBJECTS
    End If
    nCallResult = ReadFile(hFile, ByVal pBuffer, nFileLen, nRetLength, ByVal 0)
    If nRetLength <> nFileLen Then
        PgGetPluginRealValidationID = vbNullString
        GoTo LBL_FREE_OBJECTS
    End If
    pVarPos = 0
    Call RtlMoveMemory(VarPtr(IDOSH), pBuffer + pVarPos, Len(IDOSH))
    If IDOSH.e_magic <> &H5A4D Then
        PgGetPluginRealValidationID = vbNullString
        GoTo LBL_FREE_OBJECTS
    End If
    If IDOSH.e_lfanew >= (nFileLen - Len(IMNTH)) Then
        PgGetPluginRealValidationID = vbNullString
        GoTo LBL_FREE_OBJECTS
    End If
    pVarPos = IDOSH.e_lfanew
    Call RtlMoveMemory(VarPtr(IMNTH), pBuffer + pVarPos, Len(IMNTH))
    If IMNTH.SignatureLow <> &H4550 Then
        PgGetPluginRealValidationID = vbNullString
        GoTo LBL_FREE_OBJECTS
    End If
    If IMNTH.FileHeader.NumberOfSections <= 0 Then
        PgGetPluginRealValidationID = vbNullString
        GoTo LBL_FREE_OBJECTS
    End If
    ReDim ISECH(IMNTH.FileHeader.NumberOfSections - 1) As IMAGE_SECTION_HEADER
    pVarPos = pVarPos + Len(IMNTH) + IMNTH.FileHeader.SizeOfOptionalHeader
    If pVarPos >= (nFileLen - (Len(ISECH(0)) * IMNTH.FileHeader.NumberOfSections)) Then
        PgGetPluginRealValidationID = vbNullString
        GoTo LBL_FREE_OBJECTS
    End If
    Call RtlMoveMemory(VarPtr(ISECH(0)), pBuffer + pVarPos, (Len(ISECH(0)) * IMNTH.FileHeader.NumberOfSections))
    For CTurn = 0 To (IMNTH.FileHeader.NumberOfSections - 1)
        If (ISECH(CTurn).PointerToRawData + ISECH(CTurn).SizeOfRawData) >= nDefaultSize Then
            nDefaultSize = ISECH(CTurn).PointerToRawData + ISECH(CTurn).SizeOfRawData
        End If
    Next
    '---cek apakah masih ada ruang sisa:
    If nDefaultSize >= nFileLen Then '---korupsi atau tidak ada sisa.
        PgGetPluginRealValidationID = vbNullString
        GoTo LBL_FREE_OBJECTS
    End If
    If nFileLen >= (nDefaultSize + 40) Then '---validationID butuh 40 bytes.
        PgGetPluginRealValidationID = String$(40, 0) '---terlalu berlebih tapi nggak apa-apa,lah, buat jaga-jaga memori.
        Call RtlMoveMemory(StrPtr(PgGetPluginRealValidationID), pBuffer + nDefaultSize, 40)
        PgGetPluginRealValidationID = Left$(StrConv(PgGetPluginRealValidationID, vbUnicode), 40)
    Else '---ada overlay tapi tidak mencukupi untuk memuat info validationID plugin.
        PgGetPluginRealValidationID = vbNullString
        GoTo LBL_FREE_OBJECTS
    End If
LBL_FREE_OBJECTS:
    Call RtlZeroMemory(VarPtr(IDOSH), Len(IDOSH))
    Call RtlZeroMemory(VarPtr(IMNTH), Len(IMNTH))
    Erase ISECH()
    If pBuffer > 0 Then
        Call LocalFree(pBuffer)
        pBuffer = 0
    End If
    If hFile > 0 Then
        Call CloseHandle(hFile)
        hFile = 0
    End If
LBL_TERAKHIR:
    If err.Number <> 0 Then
        err.Clear
    End If
End Function

Public Function PgSetPluginValidationID(ByVal szPluginFileNameW As String) As String
On Error Resume Next '--result:validationId yg telah dipasang pada plugin.
'---ambil standarisasinya:
Dim szDefaultValidityID         As String
    szDefaultValidityID = PgGetPluginDefaultValidationID(szPluginFileNameW)
    If Len(szDefaultValidityID) <= 0 Then
        PgSetPluginValidationID = vbNullString
        GoTo LBL_TERAKHIR
    End If
'---coba pasang validity-id ke file target:
Dim hFile           As Long
    hFile = CreateFileW(StrPtr(szPluginFileNameW), SYNCHRONIZE Or READ_CONTROL Or FILE_READ_DATA Or FILE_WRITE_DATA, FILE_SHARE_READ, ByVal 0, FILE_DISPOSE_OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, 0)
    If hFile <= 0 Then
        PgSetPluginValidationID = vbNullString
        GoTo LBL_TERAKHIR
    End If
Dim nFileLen        As Long
    nFileLen = GetFileSize(hFile, ByVal 0)
Dim pBuffer         As Long
Dim nBufferLength   As Long
Dim nCallResult     As Long
Dim nRetLength      As Long
Dim IDOSH           As IMAGE_DOS_HEADER
Dim IMNTH           As IMAGE_NT_HEADERS
Dim ISECH()         As IMAGE_SECTION_HEADER
Dim pVarPos         As Long
Dim nVarData        As Long
Dim CTurn           As Long
Dim nDefaultSize    As Long
Dim nLenAllSections As Long
    If nFileLen <= Len(IDOSH) Then
        PgSetPluginValidationID = vbNullString
        GoTo LBL_FREE_OBJECTS
    End If
    nBufferLength = nFileLen
    pBuffer = LocalAlloc(lptr, nBufferLength)
    If pBuffer <= 0 Then
        PgSetPluginValidationID = vbNullString
        GoTo LBL_FREE_OBJECTS
    End If
    nCallResult = ReadFile(hFile, ByVal pBuffer, nFileLen, nRetLength, ByVal 0)
    If nRetLength <> nFileLen Then
        PgSetPluginValidationID = vbNullString
        GoTo LBL_FREE_OBJECTS
    End If
    pVarPos = 0
    Call RtlMoveMemory(VarPtr(IDOSH), pBuffer + pVarPos, Len(IDOSH))
    If IDOSH.e_magic <> &H5A4D Then
        PgSetPluginValidationID = vbNullString
        GoTo LBL_FREE_OBJECTS
    End If
    If IDOSH.e_lfanew >= (nFileLen - Len(IMNTH)) Then
        PgSetPluginValidationID = vbNullString
        GoTo LBL_FREE_OBJECTS
    End If
    pVarPos = IDOSH.e_lfanew
    Call RtlMoveMemory(VarPtr(IMNTH), pBuffer + pVarPos, Len(IMNTH))
    If IMNTH.SignatureLow <> &H4550 Then
        PgSetPluginValidationID = vbNullString
        GoTo LBL_FREE_OBJECTS
    End If
    If IMNTH.FileHeader.NumberOfSections <= 0 Then
        PgSetPluginValidationID = vbNullString
        GoTo LBL_FREE_OBJECTS
    End If
    ReDim ISECH(IMNTH.FileHeader.NumberOfSections - 1) As IMAGE_SECTION_HEADER
    pVarPos = pVarPos + Len(IMNTH) + IMNTH.FileHeader.SizeOfOptionalHeader
    If pVarPos >= (nFileLen - (Len(ISECH(0)) * IMNTH.FileHeader.NumberOfSections)) Then
        PgSetPluginValidationID = vbNullString
        GoTo LBL_FREE_OBJECTS
    End If
    Call RtlMoveMemory(VarPtr(ISECH(0)), pBuffer + pVarPos, (Len(ISECH(0)) * IMNTH.FileHeader.NumberOfSections))
    For CTurn = 0 To (IMNTH.FileHeader.NumberOfSections - 1)
        If (ISECH(CTurn).PointerToRawData + ISECH(CTurn).SizeOfRawData) >= nDefaultSize Then
            nDefaultSize = ISECH(CTurn).PointerToRawData + ISECH(CTurn).SizeOfRawData
        End If
    Next
    '---cek apakah masih ada ruang sisa:
    If nFileLen < nDefaultSize Then '---korupsi atau mencukupi standar ukuran file.
        PgSetPluginValidationID = vbNullString
        GoTo LBL_FREE_OBJECTS
    End If
    Call SetFilePointer(hFile, nDefaultSize, 0, 0)
    nCallResult = WriteFile(hFile, ByVal StrPtr(StrConv(szDefaultValidityID, vbFromUnicode)), 40, nRetLength, ByVal 0)
    If nRetLength <> 40 Then
        PgSetPluginValidationID = vbNullString
        GoTo LBL_FREE_OBJECTS
    End If
    PgSetPluginValidationID = szDefaultValidityID
LBL_FREE_OBJECTS:
    Call RtlZeroMemory(VarPtr(IDOSH), Len(IDOSH))
    Call RtlZeroMemory(VarPtr(IMNTH), Len(IMNTH))
    Erase ISECH()
    If pBuffer > 0 Then
        Call LocalFree(pBuffer)
        pBuffer = 0
    End If
    If hFile > 0 Then
        Call CloseHandle(hFile)
        hFile = 0
    End If
LBL_TERAKHIR:
    If err.Number <> 0 Then
        err.Clear
    End If
End Function

Public Function PgEnumeratePluginFiles(ByVal szPluginDirectoryW As String, ByRef OutPutPlugsList() As PG_PLUGIN_GENERAL_INFORMATION) As Long
On Error Resume Next '--result:jumlah file plugin yg valid pada suatu plugin directory (folder).
Dim nValEax             As Long
    nValEax = GetFileAttributesW(StrPtr(szPluginDirectoryW))
    If nValEax = -1 Then
        PgEnumeratePluginFiles = -1 '---invalid, sekaligus menunjukkan kalau bukan folder.
        GoTo LBL_TERAKHIR
    End If
    If (nValEax And FILE_ATTRIBUTE_DIRECTORY) <> FILE_ATTRIBUTE_DIRECTORY Then
        PgEnumeratePluginFiles = -1 '---invalid, sekaligus menunjukkan kalau bukan folder.
        GoTo LBL_TERAKHIR
    End If
    Erase OutPutPlugsList()
    If Right$(szPluginDirectoryW, 1) <> ChrW$(92) Then
        szPluginDirectoryW = szPluginDirectoryW & ChrW$(92)
    End If
Dim WFDW                As WIN32_FIND_DATA_W
Dim hFindHandle         As Long
Dim nPluginCounter      As Long
Dim bPluginWasLoaded    As Boolean
Dim hLibPluginHandle    As Long
Dim pLibPluginFuncAddr  As Long
Dim nFuncCallRet        As Long
Dim MYPGINFO            As MY_PLUGIN_GENERAL_INFORMATION
Dim szFileName          As String
Dim nTrimNullPos        As String
Dim szDefaultValidID    As String
Dim szRealValidID       As String
Dim szBufXName          As String
Dim ActiveXObject       As Object
    hFindHandle = FindFirstFileW(StrPtr(szPluginDirectoryW & ChrW$(42)), VarPtr(WFDW))
    If hFindHandle <= 0 Then
        PgEnumeratePluginFiles = -1 '---invalid, sekaligus menunjukkan kalau bukan folder.
        GoTo LBL_TERAKHIR
    End If
    nPluginCounter = 0 '---preset.
LBL_QUERY_LOOP_FIND_NEXT:
    If (WFDW.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) <> FILE_ATTRIBUTE_DIRECTORY Then '---pastikan hanya attribute file.
        nTrimNullPos = InStr(1, WFDW.cszFileName, ChrW$(0), vbBinaryCompare)
        If nTrimNullPos > 0 Then
            szFileName = Left$(WFDW.cszFileName, nTrimNullPos - 1)
        Else
            szFileName = WFDW.cszFileName
        End If
        '---cek apakah merupakan plugin yg telah tervalidasi?
        szFileName = szPluginDirectoryW & szFileName
        szDefaultValidID = PgGetPluginDefaultValidationID(szFileName)
        If Len(szDefaultValidID) > 0 Then
            szRealValidID = PgGetPluginRealValidationID(szFileName)
            If Len(szRealValidID) > 0 Then
                If szRealValidID = szDefaultValidID Then
                    '---sepertinya sudah divalidasi, tapi...cek dulu apakah punya fungsi startup plugin:
                    hLibPluginHandle = GetModuleHandleW(StrPtr(szFileName))
                    If hLibPluginHandle > 0 Then '---nggak usah ditutup handle library-nya:
                        bPluginWasLoaded = True
                        pLibPluginFuncAddr = GetProcAddress(hLibPluginHandle, szStdLibFuncGetPluginInfo)
                        If pLibPluginFuncAddr = 0 Then '---loh,kok nggak ada fungsi mendapatkan deskripsi,sih?
                            '---kemungkinan lain: adalah VB-ActiveX-DLL:
                            '---secara umum (biasa-nya sih) activex dll dari vb6 mengeksport fungsi DllGetClassObject:
                            pLibPluginFuncAddr = GetProcAddress(hLibPluginHandle, szStdLibFuncDllGetClassObject)
                            If pLibPluginFuncAddr = 0 Then
                                '---wah,dianggap nggak valid,nih,segera tutup library, unload dari memori!
                                Call FreeLibrary(hLibPluginHandle)
                                hLibPluginHandle = 0 '---reset.
                            Else
                                'MsgBox "Memori:getinfo:VB_ACTIVE-X_DLL?"
                                '---cek activex library:
                                szBufXName = PgForceGetPluginNameW(hLibPluginHandle)
                                If Len(szBufXName) > 0 Then
                                    szBufXName = szBufXName & ChrW$(46) & szStdPluginClassName
                                    Set ActiveXObject = CreateObject(szBufXName)
                                    If err.Number <> 0 Then '--ada error, kemungkinan belum registrasi active-x-nya:
                                        '---@@@@@:coba registrasi dulu:
                                        err.Clear '---bersihkan error.
                                        pLibPluginFuncAddr = GetProcAddress(hLibPluginHandle, szStdLibFuncDllRegisterServer)
                                        If pLibPluginFuncAddr = 0 Then
                                            Set ActiveXObject = Nothing '---kalau nothing,apapun hasilnya diharapkan tetap 0.
                                        Else
                                            nFuncCallRet = CallWindowProcW(pLibPluginFuncAddr, 0, 0, 0, 0)
                                            Set ActiveXObject = CreateObject(szBufXName) '---apapun hasilnya,dicoba dulu.
                                        End If
                                    End If
                                    nFuncCallRet = ActiveXObject.CPgGetInfo0
                                    If nFuncCallRet > 0 Then    '---cari tahu info kalau ada alamat ke struktur yg benar:
                                        Call RtlMoveMemory(VarPtr(MYPGINFO), nFuncCallRet, LenB(MYPGINFO))
                                        ReDim Preserve OutPutPlugsList(nPluginCounter) As PG_PLUGIN_GENERAL_INFORMATION
                                        With OutPutPlugsList(nPluginCounter)
                                            .szPluginValidationID = szRealValidID
                                            .szPluginStartupPathW = szFileName
                                            .szPluginName = PgTrimNullW(MYPGINFO.szPluginName)
                                            .szPluginDescription = PgTrimNullW(MYPGINFO.szPluginDescription)
                                            .szPluginAuthorSite = PgTrimNullW(MYPGINFO.szPluginAuthorSite)
                                            .szPluginAuthorEMail = PgTrimNullW(MYPGINFO.szPluginAuthorEMail)
                                            .szPluginAuthorAddress = PgTrimNullW(MYPGINFO.szPluginAuthorAddress)
                                            .szPluginAuthor = PgTrimNullW(MYPGINFO.szPluginAuthor)
                                            .nPluginVersionRevision = MYPGINFO.nPluginVersionRevision
                                            .nPluginVersionMinor = MYPGINFO.nPluginVersionMinor
                                            .nPluginVersionMajor = MYPGINFO.nPluginVersionMajor
                                            .nPluginCharacteristic = 1 '--sudah ada di memori.
                                            .nPluginIsRunAsActiveX = 1 '---relatif tidak aman bila dipergunakan sebagai multithreading.
                                        End With
                                        nPluginCounter = nPluginCounter + 1
                                    End If
                                    Set ActiveXObject = Nothing
                                End If
                                '---tidak usah unload dll dari memori.
                            End If
                            '---[@#@#@:kode di atas tidak usah unload dll dari memori,siapa tahu masih sedang digunakan].
                        Else '---fungsi ditemukan,coba cari tahu deskripsi pluginnya:
                            '#####:
                            nFuncCallRet = CallWindowProcW(pLibPluginFuncAddr, 0, 0, 0, 0)
                                If nFuncCallRet > 0 Then    '---cari tahu info kalau ada alamat ke struktur yg benar:
                                    Call RtlMoveMemory(VarPtr(MYPGINFO), nFuncCallRet, LenB(MYPGINFO))
                                    ReDim Preserve OutPutPlugsList(nPluginCounter) As PG_PLUGIN_GENERAL_INFORMATION
                                    With OutPutPlugsList(nPluginCounter)
                                        .szPluginValidationID = szRealValidID
                                        .szPluginStartupPathW = szFileName
                                        .szPluginName = PgTrimNullW(MYPGINFO.szPluginName)
                                        .szPluginDescription = PgTrimNullW(MYPGINFO.szPluginDescription)
                                        .szPluginAuthorSite = PgTrimNullW(MYPGINFO.szPluginAuthorSite)
                                        .szPluginAuthorEMail = PgTrimNullW(MYPGINFO.szPluginAuthorEMail)
                                        .szPluginAuthorAddress = PgTrimNullW(MYPGINFO.szPluginAuthorAddress)
                                        .szPluginAuthor = PgTrimNullW(MYPGINFO.szPluginAuthor)
                                        .nPluginVersionRevision = MYPGINFO.nPluginVersionRevision
                                        .nPluginVersionMinor = MYPGINFO.nPluginVersionMinor
                                        .nPluginVersionMajor = MYPGINFO.nPluginVersionMajor
                                        .nPluginCharacteristic = 1
                                        .nPluginIsRunAsActiveX = 0 '---relatif aman bila dipergunakan sebagai multithreading.
                                    End With
                                    nPluginCounter = nPluginCounter + 1
                                End If
                            '---[@#@#@:kode di atas tidak usah unload dll dari memori,siapa tahu masih sedang digunakan].
                        End If
                    '----------------------------------------------------------------load:
                    Else '---harus tutup kembali DLL:
                        bPluginWasLoaded = False
                        hLibPluginHandle = LoadLibraryW(StrPtr(szFileName))
                        If hLibPluginHandle > 0 Then
                            pLibPluginFuncAddr = GetProcAddress(hLibPluginHandle, szStdLibFuncGetPluginInfo)
                            If pLibPluginFuncAddr = 0 Then '---loh,kok nggak ada fungsi mendapatkan deskripsi,sih?
                                '---kemungkinan lain: adalah VB-ActiveX-DLL:
                                '---secara umum (biasa-nya sih) activex dll dari vb6 mengeksport fungsi DllGetClassObject:
                                pLibPluginFuncAddr = GetProcAddress(hLibPluginHandle, szStdLibFuncDllGetClassObject)
                                If pLibPluginFuncAddr = 0 Then
                                    '---wah,dianggap nggak valid,nih,segera tutup library, unload dari memori![@@@:kalau memang awalnya tidak dimuat,segera unload dari memori]:
                                    Call FreeLibrary(hLibPluginHandle)
                                    hLibPluginHandle = 0 '---reset.
                                Else
                                    'MsgBox "Load:getinfo:VB_ACTIVE-X_DLL?"
                                    '---cek activex library:
                                    szBufXName = PgForceGetPluginNameW(hLibPluginHandle)
                                    If Len(szBufXName) > 0 Then
                                        szBufXName = szBufXName & ChrW$(46) & szStdPluginClassName
                                        Set ActiveXObject = CreateObject(szBufXName)
                                        If err.Number <> 0 Then '--ada error, kemungkinan belum registrasi active-x-nya:
                                            '---@@@@@:coba registrasi dulu:
                                            err.Clear '---bersihkan error.
                                            pLibPluginFuncAddr = GetProcAddress(hLibPluginHandle, szStdLibFuncDllRegisterServer)
                                            If pLibPluginFuncAddr = 0 Then
                                                Set ActiveXObject = Nothing '---kalau nothing,apapun hasilnya diharapkan tetap 0.
                                            Else
                                                nFuncCallRet = CallWindowProcW(pLibPluginFuncAddr, 0, 0, 0, 0)
                                                Set ActiveXObject = CreateObject(szBufXName) '---apapun hasilnya,dicoba dulu.
                                            End If
                                        End If
                                        nFuncCallRet = ActiveXObject.CPgGetInfo0
                                        If nFuncCallRet > 0 Then    '---cari tahu info kalau ada alamat ke struktur yg benar:
                                            Call RtlMoveMemory(VarPtr(MYPGINFO), nFuncCallRet, LenB(MYPGINFO))
                                            ReDim Preserve OutPutPlugsList(nPluginCounter) As PG_PLUGIN_GENERAL_INFORMATION
                                            With OutPutPlugsList(nPluginCounter)
                                                .szPluginValidationID = szRealValidID
                                                .szPluginStartupPathW = szFileName
                                                .szPluginName = PgTrimNullW(MYPGINFO.szPluginName)
                                                .szPluginDescription = PgTrimNullW(MYPGINFO.szPluginDescription)
                                                .szPluginAuthorSite = PgTrimNullW(MYPGINFO.szPluginAuthorSite)
                                                .szPluginAuthorEMail = PgTrimNullW(MYPGINFO.szPluginAuthorEMail)
                                                .szPluginAuthorAddress = PgTrimNullW(MYPGINFO.szPluginAuthorAddress)
                                                .szPluginAuthor = PgTrimNullW(MYPGINFO.szPluginAuthor)
                                                .nPluginVersionRevision = MYPGINFO.nPluginVersionRevision
                                                .nPluginVersionMinor = MYPGINFO.nPluginVersionMinor
                                                .nPluginVersionMajor = MYPGINFO.nPluginVersionMajor
                                                .nPluginCharacteristic = 0
                                                .nPluginIsRunAsActiveX = 1 '---relatif tidak aman bila dipergunakan sebagai multithreading.
                                            End With
                                            nPluginCounter = nPluginCounter + 1
                                        End If
                                        Set ActiveXObject = Nothing
                                    End If
                                    '---test untuk menghemat alokasi memori[@@@:kalau memang awalnya tidak dimuat,segera unload dari memori]:
                                    Call FreeLibrary(hLibPluginHandle)
                                    hLibPluginHandle = 0 '---reset.
                                End If
                            Else '---fungsi ditemukan,coba cari tahu deskripsi pluginnya:
                                '#####:
                                nFuncCallRet = CallWindowProcW(pLibPluginFuncAddr, 0, 0, 0, 0)
                                If nFuncCallRet > 0 Then    '---cari tahu info kalau ada alamat ke struktur yg benar:
                                    Call RtlMoveMemory(VarPtr(MYPGINFO), nFuncCallRet, LenB(MYPGINFO))
                                    ReDim Preserve OutPutPlugsList(nPluginCounter) As PG_PLUGIN_GENERAL_INFORMATION
                                    With OutPutPlugsList(nPluginCounter)
                                        .szPluginValidationID = szRealValidID
                                        .szPluginStartupPathW = szFileName
                                        .szPluginName = PgTrimNullW(MYPGINFO.szPluginName)
                                        .szPluginDescription = PgTrimNullW(MYPGINFO.szPluginDescription)
                                        .szPluginAuthorSite = PgTrimNullW(MYPGINFO.szPluginAuthorSite)
                                        .szPluginAuthorEMail = PgTrimNullW(MYPGINFO.szPluginAuthorEMail)
                                        .szPluginAuthorAddress = PgTrimNullW(MYPGINFO.szPluginAuthorAddress)
                                        .szPluginAuthor = PgTrimNullW(MYPGINFO.szPluginAuthor)
                                        .nPluginVersionRevision = MYPGINFO.nPluginVersionRevision
                                        .nPluginVersionMinor = MYPGINFO.nPluginVersionMinor
                                        .nPluginVersionMajor = MYPGINFO.nPluginVersionMajor
                                        .nPluginCharacteristic = 0
                                        .nPluginIsRunAsActiveX = 0 '---relatif aman bila dipergunakan sebagai multithreading.
                                    End With
                                    nPluginCounter = nPluginCounter + 1
                                End If
                                '---tutup n unload kembali plugin setelah ambil info:[@@@:kalau memang awalnya tidak dimuat,segera unload dari memori]:
                                Call FreeLibrary(hLibPluginHandle)
                                hLibPluginHandle = 0 '---reset.
                            End If
                        Else '---loh,plugin apa bukan,nih?kok mencurigakan?nggak bisa di-load di memori?
                            '---nggak ngapa-ngapain, apa yg harus dilakukan? :(
                        End If
                    End If
                End If
            End If
        End If
    End If
    If FindNextFileW(hFindHandle, VarPtr(WFDW)) <> 0 Then
        GoTo LBL_QUERY_LOOP_FIND_NEXT
    End If
LBL_CLOSE_FIND_OBJECT:
    If hFindHandle > 0 Then
        Call FindClose(hFindHandle)
        hFindHandle = 0
    End If
    PgEnumeratePluginFiles = nPluginCounter
LBL_TERAKHIR:
    If err.Number <> 0 Then
        err.Clear
    End If
End Function

Public Function PgLoadAndRunPlugin(ByVal szPluginFileNameW As String, ByVal bRunInNewThread As Boolean) As Long
On Error Resume Next '--result:0=gagal; 1=berhasil,Load dan Jalankan di thread yg sama; 2=berhasil,sudah Load sebelumnya dan sekarang cuman Jalankan di thread yg sama; 3=berhasil,Load dan Jalankan di thread baru; 4=berhasil,sudah Load sebelumnya dan sekarang cuman Jalankan di thread baru.
Dim szDefaultValidID    As String
Dim szRealValidID       As String
Dim hLibPluginHandle    As Long
Dim pLibPluginFuncAddr  As Long
Dim nFuncCallRet        As Long
Dim bPluginWasLoaded    As Boolean
Dim nNewThreadID        As Long
Dim szBufXName          As String
Dim ActiveXObject       As Object
    szDefaultValidID = PgGetPluginDefaultValidationID(szPluginFileNameW)
        If Len(szDefaultValidID) > 0 Then
            szRealValidID = PgGetPluginRealValidationID(szPluginFileNameW)
            If Len(szRealValidID) > 0 Then
                If szRealValidID = szDefaultValidID Then
                    '---sepertinya sudah divalidasi,sekarang cek apakah sudah jalan di memori:
                    hLibPluginHandle = GetModuleHandleW(StrPtr(szPluginFileNameW))
                    If hLibPluginHandle > 0 Then '---sudah jalan di memori:
                        bPluginWasLoaded = True
                        pLibPluginFuncAddr = GetProcAddress(hLibPluginHandle, szStdLibFuncPluginGo)
                        If pLibPluginFuncAddr = 0 Then '---loh,kok nggak ada fungsi untuk jalankan plugin,sih?
                            '---kemungkinan lain: adalah VB-ActiveX-DLL:
                            '---secara umum (biasa-nya sih) activex dll dari vb6 mengeksport fungsi DllGetClassObject:
                            pLibPluginFuncAddr = GetProcAddress(hLibPluginHandle, szStdLibFuncDllGetClassObject)
                            If pLibPluginFuncAddr = 0 Then
                                '---wah,dianggap nggak valid,nih,segera tutup library, unload dari memori!
                                Call FreeLibrary(hLibPluginHandle)
                                hLibPluginHandle = 0 '---reset.
                                PgLoadAndRunPlugin = 0 '---gagal.
                            Else
                                'MsgBox "Memori:Execute:VB_ACTIVE-X_DLL?"
                                '---cek activex library:hanya digunakan sebagai single-threaded:
                                szBufXName = PgForceGetPluginNameW(hLibPluginHandle)
                                If Len(szBufXName) > 0 Then
                                    szBufXName = szBufXName & ChrW$(46) & szStdPluginClassName
                                    Set ActiveXObject = CreateObject(szBufXName)
                                    If err.Number <> 0 Then '--ada error, kemungkinan belum registrasi active-x-nya:
                                        '---@@@@@:coba registrasi dulu:
                                        err.Clear '---bersihkan error.
                                        pLibPluginFuncAddr = GetProcAddress(hLibPluginHandle, szStdLibFuncDllRegisterServer)
                                        If pLibPluginFuncAddr = 0 Then
                                            Set ActiveXObject = Nothing '---kalau nothing,apapun hasilnya diharapkan tetap 0.
                                        Else
                                            nFuncCallRet = CallWindowProcW(pLibPluginFuncAddr, 0, 0, 0, 0)
                                            Set ActiveXObject = CreateObject(szBufXName) '---apapun hasilnya,dicoba dulu.
                                        End If
                                    End If
                                    nFuncCallRet = ActiveXObject.CPgGo0
                                    If nFuncCallRet > 0 Then    '---cari tahu info kalau ada alamat ke struktur yg benar:
                                        PgLoadAndRunPlugin = 2 '---sudah di memori+singlethreading.
                                    Else
                                        PgLoadAndRunPlugin = 0 '---sepertinya nggak jalan, atau mengembalikan hasil 0.
                                    End If
                                    Set ActiveXObject = Nothing
                                End If
                                '---tidak usah unload dll dari memori.
                            End If
                        Else '---fungsi ditemukan,coba jalankan plugin-nya:
                                If bRunInNewThread = True Then
                                    nFuncCallRet = CreateThread(0, 0, pLibPluginFuncAddr, 0, 0, nNewThreadID)
                                    If nFuncCallRet <> 0 Then
                                        Call CloseHandle(nFuncCallRet)
                                        nFuncCallRet = 0
                                        PgLoadAndRunPlugin = 4 '---sudah di memori+multithreading.
                                    Else
                                        PgLoadAndRunPlugin = 0 '---sepertinya nggak jalan.
                                    End If
                                Else
                                    nFuncCallRet = CallWindowProcW(pLibPluginFuncAddr, 0, 0, 0, 0)
                                    PgLoadAndRunPlugin = 2 '---sudah di memori+singlethreading.
                                End If
                        End If
                    Else '---belum jalan di memori:
                        bPluginWasLoaded = False
                        hLibPluginHandle = LoadLibraryW(StrPtr(szPluginFileNameW))
                        If hLibPluginHandle > 0 Then
                            pLibPluginFuncAddr = GetProcAddress(hLibPluginHandle, szStdLibFuncPluginGo)
                            If pLibPluginFuncAddr = 0 Then '---loh,kok nggak ada fungsi mendapatkan deskripsi,sih?
                                '---kemungkinan lain: adalah VB-ActiveX-DLL:
                                '---secara umum (biasa-nya sih) activex dll dari vb6 mengeksport fungsi DllGetClassObject:
                                pLibPluginFuncAddr = GetProcAddress(hLibPluginHandle, szStdLibFuncDllGetClassObject)
                                If pLibPluginFuncAddr = 0 Then
                                    '---wah,dianggap nggak valid,nih,segera tutup library, unload dari memori![@@@:kalau memang awalnya tidak dimuat,segera unload dari memori]:
                                    Call FreeLibrary(hLibPluginHandle)
                                    hLibPluginHandle = 0 '---reset.
                                    PgLoadAndRunPlugin = 0 '---gagal.
                                Else
                                    'MsgBox "Load:Execute:VB_ACTIVE-X_DLL?"
                                    '---cek activex library:hanya digunakan sebagai single-threaded:
                                    szBufXName = PgForceGetPluginNameW(hLibPluginHandle)
                                    If Len(szBufXName) > 0 Then
                                        szBufXName = szBufXName & ChrW$(46) & szStdPluginClassName
                                        Set ActiveXObject = CreateObject(szBufXName)
                                        If err.Number <> 0 Then '--ada error, kemungkinan belum registrasi active-x-nya:
                                            '---@@@@@:coba registrasi dulu:
                                            err.Clear '---bersihkan error.
                                            pLibPluginFuncAddr = GetProcAddress(hLibPluginHandle, szStdLibFuncDllRegisterServer)
                                            If pLibPluginFuncAddr = 0 Then
                                                Set ActiveXObject = Nothing '---kalau nothing,apapun hasilnya diharapkan tetap 0.
                                            Else
                                                nFuncCallRet = CallWindowProcW(pLibPluginFuncAddr, 0, 0, 0, 0)
                                                Set ActiveXObject = CreateObject(szBufXName) '---apapun hasilnya,dicoba dulu.
                                            End If
                                        End If
                                        nFuncCallRet = ActiveXObject.CPgGo0
                                        If nFuncCallRet > 0 Then    '---cari tahu info kalau ada alamat ke struktur yg benar:
                                            PgLoadAndRunPlugin = 1 '---baru dimuat+singlethreading.
                                        Else
                                            PgLoadAndRunPlugin = 0 '---sepertinya nggak jalan, atau mengembalikan hasil 0.
                                        End If
                                        Set ActiveXObject = Nothing
                                    End If
                                    '---nggak usah unload dll dari memori.
                                End If
                            Else '---fungsi ditemukan,coba jalankan plugin-nya:
                                If bRunInNewThread = True Then
                                    nFuncCallRet = CreateThread(0, 0, pLibPluginFuncAddr, 0, 0, nNewThreadID)
                                    If nFuncCallRet <> 0 Then
                                        Call CloseHandle(nFuncCallRet)
                                        nFuncCallRet = 0
                                        PgLoadAndRunPlugin = 3 '---baru dimuat+multithreading.
                                    Else
                                        PgLoadAndRunPlugin = 0 '---sepertinya nggak jalan.
                                    End If
                                Else
                                    nFuncCallRet = CallWindowProcW(pLibPluginFuncAddr, 0, 0, 0, 0)
                                    PgLoadAndRunPlugin = 1 '---baru dimuat+singlethreading.
                                End If
                            End If
                        Else '---loh,plugin apa bukan,nih?kok mencurigakan?nggak bisa di-load di memori?
                            '---nggak ngapa-ngapain, apa yg harus dilakukan? :(
                            PgLoadAndRunPlugin = 0 '---gagal.
                        End If
                    End If
                End If
            End If
        End If
        
LBL_TERAKHIR:
    If err.Number <> 0 Then
        err.Clear
    End If
End Function

Private Function PgSetHexStringLen(ByRef szHexString As String, ByVal HexStringLength As Long) As String
On Error Resume Next '--result:validationId yg telah dipasang pada plugin.
    If Len(szHexString) < HexStringLength Then
        PgSetHexStringLen = String$(HexStringLength - Len(szHexString), 48) & szHexString
    Else
        PgSetHexStringLen = szHexString
    End If
LBL_TERAKHIR:
    If err.Number <> 0 Then
        err.Clear
    End If
End Function

Private Function PgTrimNullW(ByRef szStringToTrim As String) As String
On Error Resume Next '--result:string yg telah dipotong nullcharnya (bila ada).
Dim nNullCharPos        As Long
    nNullCharPos = InStr(1, szStringToTrim, ChrW$(0), vbBinaryCompare)
    If nNullCharPos > 0 Then
        PgTrimNullW = Left$(szStringToTrim, nNullCharPos - 1)
    Else
        PgTrimNullW = szStringToTrim
    End If
LBL_TERAKHIR:
    If err.Number <> 0 Then
        err.Clear
    End If
End Function

Private Function PgForceGetPluginNameW(ByVal hLibModule As Long) As String
On Error Resume Next '--result:string nama plugin,misal:"MyNicePlugin".
Dim hToResourceInfo     As Long
Dim pToLoadedResource   As Long
Dim LEAX                As Long
Dim TLBHeader           As TLB_HEADER
Dim TLBSegDir           As MSFT_SEGDIR
Dim TLBNameIntro        As MSFT_NAMEINTRO
Dim szTypeLibName       As String
Dim nTypeLibNameLen     As Long
    hToResourceInfo = FindResourceW(hLibModule, StrPtr("#1"), StrPtr("TYPELIB")) '---cari resource tipe "TYPELIB": dgn nama "#1" saja.
    If hToResourceInfo = 0 Then
        PgForceGetPluginNameW = vbNullString
        GoTo LBL_TERAKHIR
    End If
    '---alamat isi data resource typelib:
    pToLoadedResource = LoadResource(hLibModule, hToResourceInfo)
    If pToLoadedResource = 0 Then
        PgForceGetPluginNameW = vbNullString
        GoTo LBL_TERAKHIR
    End If
    '---cari tahu header typelib:
    LEAX = 0: Call RtlMoveMemory(VarPtr(LEAX), pToLoadedResource, Len(LEAX))
    If LEAX <> MSFT_SIGNATURE Then '---fokus ke tipe MSFT saja.
        PgForceGetPluginNameW = vbNullString
        GoTo LBL_TERAKHIR
    End If
    Call RtlMoveMemory(VarPtr(TLBHeader), pToLoadedResource, Len(TLBHeader))
    '---cari tahu direktori msft:
    Call RtlMoveMemory(VarPtr(TLBSegDir), pToLoadedResource + (Len(TLBHeader) + (4 * (TLBHeader.nTypeInfo)) + IIf((TLBHeader.fVar And &H100), 4, 0)), Len(TLBSegDir))
    '---cari tahu nama typelib:
    If TLBHeader.oName < 0 Then '---pastikan indeks-nya base0 positif.
        PgForceGetPluginNameW = vbNullString
        GoTo LBL_TERAKHIR
    End If
    Call RtlMoveMemory(VarPtr(TLBNameIntro), pToLoadedResource + (TLBSegDir.pNames.Offs + TLBHeader.oName), Len(TLBNameIntro))
    nTypeLibNameLen = (TLBNameIntro.cName And &HFF)
    szTypeLibName = String$(nTypeLibNameLen \ 2, 0) '---format ansi.
    Call RtlMoveMemory(StrPtr(szTypeLibName), pToLoadedResource + (TLBSegDir.pNames.Offs + TLBHeader.oName + Len(TLBNameIntro)), nTypeLibNameLen)
    szTypeLibName = StrConv(szTypeLibName, vbUnicode) '---format unicode.
    PgForceGetPluginNameW = szTypeLibName
    szTypeLibName = vbNullString
LBL_TERAKHIR:
    If err.Number <> 0 Then
        err.Clear
    End If
End Function




