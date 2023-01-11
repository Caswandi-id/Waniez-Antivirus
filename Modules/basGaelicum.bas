Attribute VB_Name = "basGaelicum"
Option Explicit
'********************************************************************************'
'---ini adalah modul [simpel] untuk men-disinfeksi virus: W32\Tenga-a/Gaelicum.
'---catatan:
'---virus Tenga-a/Gaelicum = virus ep semi-statis.
'---lokasi ep (Alhamdulillah) tidak di-enkripsi,tersimpan pada kodevirus:ep+3
'---misal:tenga-a_inf_hello.exe_.text Entry Point (0x00403010)--->OEP:(0x00401000)
'0x403011: 60                     PUSHAD
'0x403012: B900100000             MOV         ECX,0x1000 ;---di sini penyimpan posisi OEP. 0x1000 + OptionalHeader.ImageBase.
'0x403017: E800000000             CALL        0x40301C           ; 0x40301C
'0x40301C: 5F                     POP         EDI                ; <==0x00403017(*-0x5)
'0x40301D: 4F                     DEC         EDI                ; <==0x00403026(*+0x9)
'0x40301E: 6631FF                 XOR         DI,DI
'0x403021: 66813F4D5A             CMP         WORD PTR [EDI],0x5A4D
'0x403026: 75F5                   JNE         0x40301D           ; (*-0x9)
'0x403028: 01F9                   ADD         ECX,EDI
'0x40302A: 89E5                   MOV         EBP,ESP
'0x40302C: 894D20                 MOV         DWORD PTR [EBP+0x20],ECX
'0x40302F: FC CLD
'0x403030: E880000000             CALL        0x4030B5           ; 0x4030B5
'...next-opcodes.
'---bugs yg ditemukan [setelah disinfeksi]:
'---(-)program berbasis dot-net tidak dapat berjalan lancar lagi.akan muncul tulisan spt mencari runtime-nya saat startup :(
'---(-)program yg menggunakan signature-stamp spt verisign, tidak dapat diverifikasi lagi stamp-nya.
'---(-)program yg menggunakan pengecekan checksum akan file diri-nya sendiri, akan muncul pesan spt checksum error.
'--->silahkan dikembangkan n diperbaiki apabila menemukan bugs yg lainnya.trims.
'---oleh:ari pambudi [pamzlogic] 2010.
'********************************************************************************'

Private Type IMAGE_DOS_HEADER
    e_magic                 As Integer ' Magic number "MZ"
    e_cblp                  As Integer ' Bytes on last page of file
    e_cp                    As Integer ' Pages in file
    e_crlc                  As Integer ' Relocations
    e_cparhdr               As Integer ' Size of header in paragraphs
    e_minalloc              As Integer ' Minimum extra paragraphs needed
    e_maxalloc              As Integer ' Maximum extra paragraphs needed
    e_ss                    As Integer ' Initial (relative) SS value
    e_sp                    As Integer ' Initial SP value
    e_csum                  As Integer ' Checksum
    e_ip                    As Integer ' Initial IP value
    e_cs                    As Integer ' Initial (relative) CS value
    e_lfarlc                As Integer ' File address of relocation table
    e_ovno                  As Integer ' Overlay number
    e_res(0 To 3)           As Integer ' Reserved words
    e_oemid                 As Integer ' OEM identifier (for e_oeminfo)
    e_oeminfo               As Integer ' OEM information; e_oemid specific
    e_res2(0 To 9)          As Integer ' Reserved words
    e_lfanew                As Long ' File address of new exe header
End Type

Private Type IMAGE_FILE_HEADER '---20 bytes.
    Machine                 As Integer 'tipe processor yg dibutuhkan.
    NumberOfSections        As Integer 'jumlah section,sectionheaders & section data terletak setelah peheader.
    TimeDateStamp           As Long 'dalam detik mulai 1 januari 1970.
    PointerToSymbolTable    As Long '
    NumberOfSymbols         As Long '
    SizeOfOptionalHeader    As Integer 'ukuran IMAGE_OPTIONAL_HEADER_XX,untuk 32 bit n 64 bit berbeda.
    Characteristics         As Integer '
End Type

Private Type IMAGE_NT_HEADERS
    SignatureLow            As Integer '2 "PE"
    SignatureHigh           As Integer '2
    FileHeader              As IMAGE_FILE_HEADER '20
End Type

Private Type IMAGE_DATA_DIRECTORY_32
    VirtualAddress          As Long
    nSize                   As Long
End Type

Private Type IMAGE_OPTIONAL_HEADER_32
    '---Standard fields:
    Magic                       As Integer
    MajorLinkerVersion          As Byte
    MinorLinkerVersion          As Byte
    SizeOfCode                  As Long
    SizeOfInitializedData       As Long
    SizeOfUninitializedData     As Long
    AddressOfEntryPoint         As Long 'PEHDR+40
    BaseOfCode                  As Long
    BaseOfData                  As Long
    '---NT additional fields:
    ImageBase                   As Long
    SectionAlignment            As Long
    FileAlignment               As Long
    MajorOperatingSystemVersion As Integer
    MinorOperatingSystemVersion As Integer
    MajorImageVersion           As Integer
    MinorImageVersion           As Integer
    MajorSubsystemVersion       As Integer
    MinorSubsystemVersion       As Integer
    Win32VersionValue           As Long
    SizeOfImage                 As Long
    SizeOfHeaders               As Long
    CheckSum                    As Long
    Subsystem                   As Integer
    DllCharacteristics          As Integer
    SizeOfStackReserve          As Long
    SizeOfStackCommit           As Long
    SizeOfHeapReserve           As Long
    SizeOfHeapCommit            As Long
    LoaderFlags                 As Long
    NumberOfRvaAndSizes         As Long
    DataDirectory(0 To 15)      As IMAGE_DATA_DIRECTORY_32 '%IMAGE_DIRECTORY_ENTRY_EXPORT       =  0   ' Export Directory
End Type

Private Type IMAGE_SECTION_HEADER '---40bytes.
    SectionName(7)          As Byte '---8 bytes.
    VirtualSize             As Long
    VirtualAddress          As Long
    SizeOfRawData           As Long
    PointerToRawData        As Long
    PointerToRelocations    As Long
    PointerToLinenumbers    As Long
    NumberOfRelocations     As Integer
    NumberOfLinenumbers     As Integer
    Characteristics         As Long
End Type



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

Private Const FILE_TYPE_UNKNOWN = &H0
Private Const FILE_TYPE_DISK = &H1
Private Const FILE_TYPE_CHAR = &H2
Private Const FILE_TYPE_PIPE = &H3
Private Const FILE_TYPE_REMOTE = &H8000

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





Private Declare Function CreateFileW Lib "kernel32.dll" (ByVal lpFileName As Long, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, lpSecurityAttributes As Any, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long

Private Declare Function GetFileType Lib "kernel32.dll" (ByVal hFile As Long) As Long
Private Declare Function GetFileSize Lib "kernel32.dll" (ByVal hFile As Long, lpFileSizeHigh As Long) As Long

Private Declare Function SetFilePointer Lib "kernel32.dll" (ByVal hFile As Long, ByVal lDistanceToMove As Long, lpDistanceToMoveHigh As Long, ByVal dwMoveMethod As Long) As Long
Private Declare Function SetEndOfFile Lib "kernel32.dll" (ByVal hFile As Long) As Long

Private Declare Function ReadFile Lib "kernel32.dll" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, lpOverlapped As Any) As Long
Private Declare Function WriteFile Lib "kernel32.dll" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, lpOverlapped As Any) As Long

Private Declare Function CloseHandle Lib "kernel32.dll" (ByVal hObject As Long) As Long

Private Declare Function LocalAlloc Lib "kernel32.dll" (ByVal uFlags As Long, ByVal lBytes As Long) As Long
Private Declare Function LocalSize Lib "kernel32.dll" (ByVal hMem As Long) As Long
Private Declare Function LocalFree Lib "kernel32.dll" (ByVal hMem As Long) As Long

Private Declare Sub RtlMoveMemory Lib "NTDLL.DLL" (ByVal pDestBuffer As Long, ByVal pSourceBuffer As Long, ByVal nBufferLengthToMove As Long) '<---sebenarnya namanya kurang sesuai, karena yang dilakukan adalah menyalin (copy) isi dari src ke dst.
Private Declare Sub RtlFillMemory Lib "NTDLL.DLL" (ByVal pDestBuffer As Long, ByVal nDestLengthToFill As Long, ByVal nByteNumber As Long) '<---harusnya byte,tapi memori 32 bit, jadi nggak apa-apa, asal tetap bernilai antara 0 sampai 255.
Private Declare Sub RtlZeroMemory Lib "NTDLL.DLL" (ByVal pDestBuffer As Long, ByVal nDestLengthToFillWithZeroBytes As Long) '<---reset isi dst yaitu mengisinya dengan bytenumber = 0.



'---kalau variabel hFileHandle > 0, akan digunakan file handlenya, selain itu akan pakai szFileName.
Public Function CheckNFixFH_V_W32_Tenga_a(ByVal szFileName As String, ByVal hFileHandle As Long, ByVal bTryToFix As Boolean) As Long
On Error Resume Next
Dim nFunctionResult         As Long
Dim bInpuFromFileHandle     As Boolean '---kalau input dari file handle, jangan tutup file handle tersebut.
LBL_CHECK_INPUT_MODE:
    If hFileHandle > 0 Then
        bInpuFromFileHandle = True
        GoTo LBL_GET_FROM_HANDLE
    End If
    If bTryToFix = True Then
        hFileHandle = CreateFileW(StrPtr(szFileName), SYNCHRONIZE Or READ_CONTROL Or FILE_READ_DATA Or FILE_WRITE_DATA, FILE_SHARE_READ, ByVal 0, FILE_DISPOSE_OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, 0)
    Else
        hFileHandle = CreateFileW(StrPtr(szFileName), SYNCHRONIZE Or READ_CONTROL Or FILE_READ_DATA, FILE_SHARE_READ, ByVal 0, FILE_DISPOSE_OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, 0)
    End If
    If hFileHandle <= 0 Then '-1=INVALID_HANDLE_VALUE; 0=FUNGSI CREATEFILEW TIDAK TERSEDIA :(
        nFunctionResult = -1 'FILE TIDAK DAPAT DIAKSES,KOSONG, ATAU BUKAN FILE!
        GoTo LBL_BROADCAST_RESULT
    End If
LBL_GET_FROM_HANDLE:
Dim nFileLen                As Long
Dim LEAX                    As Long
    nFileLen = GetFileSize(hFileHandle, ByVal 0)
    If nFileLen <= 0 Then '---error ataupun kosong.
        nFunctionResult = -1 'FILE TIDAK DAPAT DIAKSES,KOSONG, ATAU BUKAN FILE!
        GoTo LBL_FREE_OBJECTS
    End If
LBL_VERIFY_FILE:
Dim ExpandValue             As Long
Dim nCallResult             As Long
Dim nRetLength              As Long
Dim IDOSH                   As IMAGE_DOS_HEADER
Dim IMNTH                   As IMAGE_NT_HEADERS
Dim IOH32                   As IMAGE_OPTIONAL_HEADER_32
Dim ISECH()                 As IMAGE_SECTION_HEADER
Dim gVEPSecNumber           As Long
Dim pVarPos                 As Long
Dim nVarLen                 As Long
Dim nVarData                As Long
Dim CTurn                   As Long
Dim bThreatFound            As Boolean
Dim btCPattern()            As Byte
    If nFileLen <= Len(IDOSH) Then
        nFunctionResult = 0 'BUKAN,INI BUKAN VIRUS YANG DIMAKSUD! DIANGGAP BERSIH!
        GoTo LBL_FREE_OBJECTS
    End If
    Call SetFilePointer(hFileHandle, 0, 0, 0)
    nCallResult = ReadFile(hFileHandle, IDOSH, Len(IDOSH), nRetLength, ByVal 0)
    If nRetLength <> Len(IDOSH) Then
        nFunctionResult = 0 'BUKAN,INI BUKAN VIRUS YANG DIMAKSUD! DIANGGAP BERSIH!
        GoTo LBL_FREE_OBJECTS
    End If
    If IDOSH.e_magic <> &H5A4D Then '<> "MZ"
        nFunctionResult = 0 'BUKAN,INI BUKAN VIRUS YANG DIMAKSUD! DIANGGAP BERSIH!
        GoTo LBL_FREE_OBJECTS
    End If
    If IDOSH.e_lfanew <= 0 Then
        nFunctionResult = 0 'BUKAN,INI BUKAN VIRUS YANG DIMAKSUD! DIANGGAP BERSIH!
        GoTo LBL_FREE_OBJECTS
    End If
    If nFileLen <= (IDOSH.e_lfanew + Len(IMNTH)) Then
        nFunctionResult = 0 'BUKAN,INI BUKAN VIRUS YANG DIMAKSUD! DIANGGAP BERSIH!
        GoTo LBL_FREE_OBJECTS
    End If
    Call SetFilePointer(hFileHandle, IDOSH.e_lfanew, 0, 0)
    nCallResult = ReadFile(hFileHandle, IMNTH, Len(IMNTH), nRetLength, ByVal 0)
    If nRetLength <> Len(IMNTH) Then
        nFunctionResult = 0 'BUKAN,INI BUKAN VIRUS YANG DIMAKSUD! DIANGGAP BERSIH!
        GoTo LBL_FREE_OBJECTS
    End If
    If IMNTH.SignatureLow <> &H4550 Then '<> "PE"
        nFunctionResult = 0 'BUKAN,INI BUKAN VIRUS YANG DIMAKSUD! DIANGGAP BERSIH!
        GoTo LBL_FREE_OBJECTS
    End If
    If IMNTH.FileHeader.SizeOfOptionalHeader <> Len(IOH32) Then '<> 224 bytes.
        nFunctionResult = 0 'BUKAN,INI BUKAN VIRUS YANG DIMAKSUD! DIANGGAP BERSIH!
        GoTo LBL_FREE_OBJECTS
    End If
    If IMNTH.FileHeader.NumberOfSections <= 0 Then 'di-luar batas kewajaran.
        nFunctionResult = 0 'BUKAN,INI BUKAN VIRUS YANG DIMAKSUD! DIANGGAP BERSIH!
        GoTo LBL_FREE_OBJECTS
    End If
    '---lolos uji valid PE 32 bit, lanjutkan:
    '---analisis optional header:
    Call SetFilePointer(hFileHandle, IDOSH.e_lfanew + Len(IMNTH), 0, 0)
    nCallResult = ReadFile(hFileHandle, IOH32, Len(IOH32), nRetLength, ByVal 0)
    If nRetLength <> Len(IOH32) Then
        nFunctionResult = 0 'BUKAN,INI BUKAN VIRUS YANG DIMAKSUD! DIANGGAP BERSIH!
        GoTo LBL_FREE_OBJECTS
    End If
    '---analisis per-section:
    ReDim ISECH(IMNTH.FileHeader.NumberOfSections - 1) As IMAGE_SECTION_HEADER '--aloaksikan ruang penyimpanan section headers.
    If nFileLen <= (IDOSH.e_lfanew + Len(IMNTH) + Len(IOH32) + (Len(ISECH(0)) * IMNTH.FileHeader.NumberOfSections)) Then
        nFunctionResult = 0 'BUKAN,INI BUKAN VIRUS YANG DIMAKSUD! DIANGGAP BERSIH!
        GoTo LBL_FREE_OBJECTS
    End If
    Call SetFilePointer(hFileHandle, IDOSH.e_lfanew + Len(IMNTH) + Len(IOH32), 0, 0)
    nCallResult = ReadFile(hFileHandle, ISECH(0), Len(ISECH(0)) * IMNTH.FileHeader.NumberOfSections, nRetLength, ByVal 0)
    If nRetLength <> Len(ISECH(0)) * IMNTH.FileHeader.NumberOfSections Then
        nFunctionResult = 0 'BUKAN,INI BUKAN VIRUS YANG DIMAKSUD! DIANGGAP BERSIH!
        GoTo LBL_FREE_OBJECTS
    End If
    For CTurn = 0 To (IMNTH.FileHeader.NumberOfSections - 1)
        If ISECH(CTurn).SizeOfRawData > 0 Then '---pastikan hanya yg ada isinya saja:
            If (IOH32.AddressOfEntryPoint >= ISECH(CTurn).VirtualAddress) And (IOH32.AddressOfEntryPoint < (ISECH(CTurn).VirtualAddress + ISECH(CTurn).SizeOfRawData)) Then
                pVarPos = ISECH(CTurn).PointerToRawData + (IOH32.AddressOfEntryPoint - ISECH(CTurn).VirtualAddress)
                If nFileLen <= pVarPos Then
                    bThreatFound = False
                    nFunctionResult = 0 'BUKAN,INI BUKAN VIRUS YANG DIMAKSUD! DIANGGAP BERSIH!
                    GoTo LBL_FREE_OBJECTS
                End If
                nVarLen = (ISECH(CTurn).PointerToRawData + ISECH(CTurn).SizeOfRawData) - pVarPos '---panjang dari ep ke akhir section tersebut.
                gVEPSecNumber = CTurn '---number Base0.
                '---cek spesifik ep:[per-virus-variant]:
                nVarData = 32 '---ambil 32 bytes awal.
                ReDim btCPattern(nVarData - 1) As Byte
                Call SetFilePointer(hFileHandle, pVarPos, 0, 0)
                If nVarLen < nVarData Then
                    nVarData = nVarLen
                End If
                nCallResult = ReadFile(hFileHandle, btCPattern(0), nVarData, nRetLength, ByVal 0)
                If nRetLength <> nVarData Then
                    bThreatFound = False
                    nFunctionResult = 0 'BUKAN,INI BUKAN VIRUS YANG DIMAKSUD! DIANGGAP BERSIH!
                    GoTo LBL_FREE_OBJECTS
                End If
                '---cek pattern (versi sederhana, tanpa optimisasi):
                If btCPattern(0) = &H52 And _
                    btCPattern(1) = &H60 And _
                    btCPattern(2) = &HB9 And _
                    btCPattern(7) = &HE8 And _
                    btCPattern(12) = &H5F And _
                    btCPattern(13) = &H4F And _
                    btCPattern(14) = &H66 And _
                    btCPattern(15) = &H31 And _
                    btCPattern(16) = &HFF And _
                    btCPattern(17) = &H66 And _
                    btCPattern(18) = &H81 Then
                        bThreatFound = True '---menurut byte pattern adalah:[cocok].
                End If
                '--------------------------------------;
                Exit For '---sudah,jangan terlalu lama berputar, 'ntar pusing :)
            End If
        End If
    Next
    If bTryToFix = False Then
        If bThreatFound = True Then
            nFunctionResult = 1 'YA! INI ADALAH FILE VIRUS! HATI-HATI!
        Else
            nFunctionResult = 0 'BUKAN,INI BUKAN VIRUS YANG DIMAKSUD! DIANGGAP BERSIH!
        End If
        GoTo LBL_FREE_OBJECTS
    Else '---mau lanjut tapi nggak ada yg dilanjutin, ya akhiri saja :)
        If bThreatFound = True Then
            nFunctionResult = 1 'preset: YA, INI FILE VIRUS[...].
        Else
            nFunctionResult = 0 'BUKAN,INI BUKAN VIRUS YANG DIMAKSUD! DIANGGAP BERSIH!
            GoTo LBL_FREE_OBJECTS
        End If
    End If
    '---yg melewati baris ini dianggap file ber-virus dan minta dibersihkan:
Dim pAddrOEP                As Long
    Call RtlMoveMemory(VarPtr(pAddrOEP), VarPtr(btCPattern(3)), Len(pAddrOEP)) '---ambil 4 bytes.
    '*+"V":[50]---tanda/flag-IGNORE.
    '*Optional.AddressOfEntryPoint-OK.
    '*Optional.SizeOfImage-OK.
    '*Section.SizeOfRawData-OK.
Dim pBuffer                 As Long
Dim nBufferLength           As Long
Dim pInBuf2                 As Long
Dim nInBufLen2              As Long
    nBufferLength = nFileLen
    pBuffer = LocalAlloc(lptr, nBufferLength)
    If pBuffer <= 0 Then
        nFunctionResult = 1 'YA, INI FILE VIRUS TAPI GAGAL DIBERSIHKAN :(
        GoTo LBL_FREE_OBJECTS
    End If
    '---backup file ke memori[sementara] sebelum di-reset:
    Call SetFilePointer(hFileHandle, 0, 0, 0) '---mulai dari awal lagi:
    nCallResult = ReadFile(hFileHandle, ByVal pBuffer, nBufferLength, nRetLength, ByVal 0)
    If nRetLength <> nBufferLength Then
        nFunctionResult = 1 'YA, INI FILE VIRUS TAPI GAGAL DIBERSIHKAN :(
        GoTo LBL_FREE_OBJECTS
    End If
    '---reset isi file:
    Call SetFilePointer(hFileHandle, 0, 0, 0) '---mulai dari awal lagi:
    Call SetEndOfFile(hFileHandle)
    '---tulis ulang dos-header[tanpa perubahan]:
    Call SetFilePointer(hFileHandle, 0, 0, 0)
    nCallResult = WriteFile(hFileHandle, IDOSH, Len(IDOSH), nRetLength, ByVal 0)
    If nRetLength <> Len(IDOSH) Then
        nFunctionResult = 1 'YA, INI FILE VIRUS TAPI GAGAL DIBERSIHKAN :(
        GoTo LBL_FREE_OBJECTS
    End If
    '---tulis ulang nt-header[tanpa perubahan]:
    Call SetFilePointer(hFileHandle, IDOSH.e_lfanew, 0, 0)
    nCallResult = WriteFile(hFileHandle, IMNTH, Len(IMNTH), nRetLength, ByVal 0)
    If nRetLength <> Len(IMNTH) Then
        nFunctionResult = 1 'YA, INI FILE VIRUS TAPI GAGAL DIBERSIHKAN :(
        GoTo LBL_FREE_OBJECTS
    End If
    '---tulis ulang optional-header[dengan perubahan]:
    IOH32.AddressOfEntryPoint = pAddrOEP '---tulis entrypoint yg sebenarnya [oep].
    IOH32.SizeOfImage = (((IOH32.SizeOfHeaders - 1) \ IOH32.SectionAlignment) + 1) * IOH32.SectionAlignment '---diisi dgn ukuran alokasi headers di memori terlebih dahulu(pembulatan terhadap section alignment).
    CTurn = 0 '---reset counter.
    For CTurn = 0 To (IMNTH.FileHeader.NumberOfSections - 1)
        IOH32.SizeOfImage = IOH32.SizeOfImage + ((((ISECH(CTurn).VirtualSize - 1) \ IOH32.SectionAlignment) + 1) * IOH32.SectionAlignment) '---diisi dgn total atau penambahan keseluruhan ukuran section di memori(pembulatan terhadap section alignment).
    Next
    Call SetFilePointer(hFileHandle, IDOSH.e_lfanew + Len(IMNTH), 0, 0)
    nCallResult = WriteFile(hFileHandle, IOH32, Len(IOH32), nRetLength, ByVal 0)
    If nRetLength <> Len(IOH32) Then
        nFunctionResult = 1 'YA, INI FILE VIRUS TAPI GAGAL DIBERSIHKAN :(
        GoTo LBL_FREE_OBJECTS
    End If
    '---pasang kembali semua 'section-header'[dengan perubahan]:
    '---hapus kode virus di section tempat kode virus berada(sebenarnya nggak dihapus nggak apa-apa,tapi ini biar ngirit media penyimpanan):
    ISECH(gVEPSecNumber).SizeOfRawData = pVarPos - ISECH(gVEPSecNumber).PointerToRawData
    Call SetFilePointer(hFileHandle, IDOSH.e_lfanew + Len(IMNTH) + Len(IOH32), 0, 0)
    nCallResult = WriteFile(hFileHandle, ISECH(0), Len(ISECH(0)) * IMNTH.FileHeader.NumberOfSections, nRetLength, ByVal 0)
    If nRetLength <> (Len(ISECH(0)) * IMNTH.FileHeader.NumberOfSections) Then
        nFunctionResult = 1 'YA, INI FILE VIRUS TAPI GAGAL DIBERSIHKAN :(
        GoTo LBL_FREE_OBJECTS
    End If
    '---pasang kembali [?]:ini header/tanda/apaan,sih? dari dulu saya nggak paham bagian ini:
    CTurn = 0 '---reset counter.
    nVarData = 0 '---reset---:variabel nVarData lagi nganggur, jadi bisa di-daur ulang :)
    For CTurn = 0 To (IMNTH.FileHeader.NumberOfSections - 1)
        If ISECH(CTurn).SizeOfRawData > 0 Then
            If nVarData > 0 Then
                If ISECH(CTurn).PointerToRawData < nVarData Then
                    nVarData = ISECH(CTurn).PointerToRawData '---masukkan niai yg lebih kecil.
                End If
            Else
                nVarData = ISECH(CTurn).PointerToRawData '---untuk pertama kali, langsung masukkan ajah.
            End If
        End If
    Next
    '---cek apakah ada jeda antara akhir dari section header dgn awal dari section data:
    CTurn = IDOSH.e_lfanew + Len(IMNTH) + Len(IOH32) + (Len(ISECH(0)) * IMNTH.FileHeader.NumberOfSections)
    If nVarData > CTurn Then
        Call SetFilePointer(hFileHandle, CTurn, 0, 0)
        nCallResult = WriteFile(hFileHandle, ByVal (pBuffer + CTurn), nVarData - CTurn, nRetLength, ByVal 0)
        If nRetLength <> (nVarData - CTurn) Then
            nFunctionResult = 1 'YA, INI FILE VIRUS TAPI GAGAL DIBERSIHKAN :(
            GoTo LBL_FREE_OBJECTS
        End If
    End If
    '---pasang kembali tiap section di tempat yg benar:
    CTurn = 0 '---reset counter.
    For CTurn = 0 To (IMNTH.FileHeader.NumberOfSections - 1)
        If CTurn <> gVEPSecNumber Then '---bila bukan section yg kena infeksi virus, biarkan apa adanya:
            Call SetFilePointer(hFileHandle, ISECH(CTurn).PointerToRawData, 0, 0)
            nCallResult = WriteFile(hFileHandle, ByVal (pBuffer + ISECH(CTurn).PointerToRawData), ISECH(CTurn).SizeOfRawData, nRetLength, ByVal 0)
            If nRetLength <> ISECH(CTurn).SizeOfRawData Then
                nFunctionResult = 1 'YA, INI FILE VIRUS TAPI GAGAL DIBERSIHKAN :(
                GoTo LBL_FREE_OBJECTS
            End If
        Else '---bila section yg kena infeksi virus, sesuaikan dulu file-alignment-nya, bila bukan kode akan di-reset ke 0[zero]:
            '---variabel nVarData lagi nganggur, jadi bisa di-daur ulang :)
            nVarData = (((ISECH(CTurn).SizeOfRawData - 1) \ IOH32.FileAlignment) + 1) * IOH32.FileAlignment '---diisi dgn ukuran alokasi headers di file-fisik terlebih dahulu(pembulatan terhadap file-alignment).
            '---isi dgn zero-bytes:
            nInBufLen2 = nVarData
            pInBuf2 = LocalAlloc(lptr, nInBufLen2)
            If pInBuf2 <= 0 Then
                Call SetFilePointer(hFileHandle, ISECH(CTurn).PointerToRawData, 0, 0)
                nCallResult = WriteFile(hFileHandle, 0, nVarData, nRetLength, ByVal 0) '---akan diisi dgn data sampah dari memori heap.
                If nRetLength <> nVarData Then
                    nFunctionResult = 1 'YA, INI FILE VIRUS TAPI GAGAL DIBERSIHKAN :(
                    GoTo LBL_FREE_OBJECTS
                End If
            Else '---akan diisi dgn data zero-bytes teralokasi:
                Call SetFilePointer(hFileHandle, ISECH(CTurn).PointerToRawData, 0, 0)
                nCallResult = WriteFile(hFileHandle, ByVal pInBuf2, nInBufLen2, nRetLength, ByVal 0) '---akan diisi dgn data sampah dari memori heap.
                If nRetLength <> nInBufLen2 Then
                    '--jangan lupa untuk tutup memori yg teralokasi:
                    If pInBuf2 > 0 Then
                        Call LocalFree(pInBuf2)
                        pInBuf2 = 0
                    End If
                    '---akhiri:
                    nFunctionResult = 1 'YA, INI FILE VIRUS TAPI GAGAL DIBERSIHKAN :(
                    GoTo LBL_FREE_OBJECTS
                End If
                If pInBuf2 > 0 Then
                    Call LocalFree(pInBuf2)
                    pInBuf2 = 0
                End If
            End If
            '---isi dgn kode sebenarnya:
            Call SetFilePointer(hFileHandle, ISECH(CTurn).PointerToRawData, 0, 0)
            nCallResult = WriteFile(hFileHandle, ByVal (pBuffer + ISECH(CTurn).PointerToRawData), ISECH(CTurn).SizeOfRawData, nRetLength, ByVal 0)
            If nRetLength <> ISECH(CTurn).SizeOfRawData Then
                nFunctionResult = 1 'YA, INI FILE VIRUS TAPI GAGAL DIBERSIHKAN :(
                GoTo LBL_FREE_OBJECTS
            End If
        End If
    Next
    '---yg berhasil melewati baris ini dianggap berhasil men-dis-infeksi virus:
    nFunctionResult = 2 'YA, INI FILE VIRUS DAN BERHASIL DIBERSIHKAN!
    '---selesai membersihkan virus,sekarang bersihkan objects:
LBL_FREE_OBJECTS:
    Call RtlZeroMemory(VarPtr(IDOSH), Len(IDOSH))
    Call RtlZeroMemory(VarPtr(IMNTH), Len(IMNTH))
    Call RtlZeroMemory(VarPtr(IOH32), Len(IOH32))
    Erase ISECH()
    Erase btCPattern()
    If pBuffer > 0 Then
        Call LocalFree(pBuffer)
        pBuffer = 0
    End If
    If hFileHandle > 0 Then
        If bInpuFromFileHandle = False Then
            Call CloseHandle(hFileHandle)
            hFileHandle = 0
        End If
    End If
LBL_BROADCAST_RESULT:
    CheckNFixFH_V_W32_Tenga_a = nFunctionResult
LBL_TERAKHIR:
    If err.Number <> 0 Then
        err.Clear
    End If
End Function



Public Function DisInfectTenga(zFileName As String) As Long
Dim RTFU    As Long

RTFU = CheckNFixFH_V_W32_Tenga_a(zFileName, 0, True)
If RTFU = 2 Then
    DisInfectTenga = 1
Else
    DisInfectTenga = 0
End If
End Function


Public Function DisInfectRunouce(zFileName As String) As Long
Dim RTFU    As Long

RTFU = CheckNFixFH_V_W32_Runouce_a(zFileName, 0, True)
If RTFU = 2 Then
    DisInfectRunouce = 1
Else
    DisInfectRunouce = 0
End If
End Function


'-------------------------------------------------------------------------------------------------
' Disinfect Runouce
' OEP ada pada data VEP, offset 17
' A.M Hirin

Function CheckNFixFH_V_W32_Runouce_a(ByVal szFileName As String, ByVal hFileHandle As Long, ByVal bTryToFix As Boolean) As Long
On Error Resume Next
Dim nFunctionResult         As Long
Dim bInpuFromFileHandle     As Boolean '---kalau input dari file handle, jangan tutup file handle tersebut.
LBL_CHECK_INPUT_MODE:
    If hFileHandle > 0 Then
        bInpuFromFileHandle = True
        GoTo LBL_GET_FROM_HANDLE
    End If
    If bTryToFix = True Then
        hFileHandle = CreateFileW(StrPtr(szFileName), SYNCHRONIZE Or READ_CONTROL Or FILE_READ_DATA Or FILE_WRITE_DATA, FILE_SHARE_READ, ByVal 0, FILE_DISPOSE_OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, 0)
    Else
        hFileHandle = CreateFileW(StrPtr(szFileName), SYNCHRONIZE Or READ_CONTROL Or FILE_READ_DATA, FILE_SHARE_READ, ByVal 0, FILE_DISPOSE_OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, 0)
    End If
    If hFileHandle <= 0 Then '-1=INVALID_HANDLE_VALUE; 0=FUNGSI CREATEFILEW TIDAK TERSEDIA :(
        nFunctionResult = -1 'FILE TIDAK DAPAT DIAKSES,KOSONG, ATAU BUKAN FILE!
        GoTo LBL_BROADCAST_RESULT
    End If
LBL_GET_FROM_HANDLE:
Dim nFileLen                As Long
Dim LEAX                    As Long
    nFileLen = GetFileSize(hFileHandle, ByVal 0)
    If nFileLen <= 0 Then '---error ataupun kosong.
        nFunctionResult = -1 'FILE TIDAK DAPAT DIAKSES,KOSONG, ATAU BUKAN FILE!
        GoTo LBL_FREE_OBJECTS
    End If
LBL_VERIFY_FILE:
Dim ExpandValue             As Long
Dim nCallResult             As Long
Dim nRetLength              As Long
Dim IDOSH                   As IMAGE_DOS_HEADER
Dim IMNTH                   As IMAGE_NT_HEADERS
Dim IOH32                   As IMAGE_OPTIONAL_HEADER_32
Dim ISECH()                 As IMAGE_SECTION_HEADER
Dim gVEPSecNumber           As Long
Dim pVarPos                 As Long
Dim nVarLen                 As Long
Dim nVarData                As Long
Dim CTurn                   As Long
Dim bThreatFound            As Boolean
Dim btCPattern()            As Byte
Dim imgBase                 As Long

    If nFileLen <= Len(IDOSH) Then
        nFunctionResult = 0 'BUKAN,INI BUKAN VIRUS YANG DIMAKSUD! DIANGGAP BERSIH!
        GoTo LBL_FREE_OBJECTS
    End If
    Call SetFilePointer(hFileHandle, 0, 0, 0)
    nCallResult = ReadFile(hFileHandle, IDOSH, Len(IDOSH), nRetLength, ByVal 0)
    If nRetLength <> Len(IDOSH) Then
        nFunctionResult = 0 'BUKAN,INI BUKAN VIRUS YANG DIMAKSUD! DIANGGAP BERSIH!
        GoTo LBL_FREE_OBJECTS
    End If
    If IDOSH.e_magic <> &H5A4D Then '<> "MZ"
        nFunctionResult = 0 'BUKAN,INI BUKAN VIRUS YANG DIMAKSUD! DIANGGAP BERSIH!
        GoTo LBL_FREE_OBJECTS
    End If
    If IDOSH.e_lfanew <= 0 Then
        nFunctionResult = 0 'BUKAN,INI BUKAN VIRUS YANG DIMAKSUD! DIANGGAP BERSIH!
        GoTo LBL_FREE_OBJECTS
    End If
    If nFileLen <= (IDOSH.e_lfanew + Len(IMNTH)) Then
        nFunctionResult = 0 'BUKAN,INI BUKAN VIRUS YANG DIMAKSUD! DIANGGAP BERSIH!
        GoTo LBL_FREE_OBJECTS
    End If
    Call SetFilePointer(hFileHandle, IDOSH.e_lfanew, 0, 0)
    nCallResult = ReadFile(hFileHandle, IMNTH, Len(IMNTH), nRetLength, ByVal 0)
    If nRetLength <> Len(IMNTH) Then
        nFunctionResult = 0 'BUKAN,INI BUKAN VIRUS YANG DIMAKSUD! DIANGGAP BERSIH!
        GoTo LBL_FREE_OBJECTS
    End If
    If IMNTH.SignatureLow <> &H4550 Then '<> "PE"
        nFunctionResult = 0 'BUKAN,INI BUKAN VIRUS YANG DIMAKSUD! DIANGGAP BERSIH!
        GoTo LBL_FREE_OBJECTS
    End If
    If IMNTH.FileHeader.SizeOfOptionalHeader <> Len(IOH32) Then '<> 224 bytes.
        nFunctionResult = 0 'BUKAN,INI BUKAN VIRUS YANG DIMAKSUD! DIANGGAP BERSIH!
        GoTo LBL_FREE_OBJECTS
    End If
    If IMNTH.FileHeader.NumberOfSections <= 0 Then 'di-luar batas kewajaran.
        nFunctionResult = 0 'BUKAN,INI BUKAN VIRUS YANG DIMAKSUD! DIANGGAP BERSIH!
        GoTo LBL_FREE_OBJECTS
    End If
    '---lolos uji valid PE 32 bit, lanjutkan:
    '---analisis optional header:
    Call SetFilePointer(hFileHandle, IDOSH.e_lfanew + Len(IMNTH), 0, 0)
    nCallResult = ReadFile(hFileHandle, IOH32, Len(IOH32), nRetLength, ByVal 0)
    If nRetLength <> Len(IOH32) Then
        nFunctionResult = 0 'BUKAN,INI BUKAN VIRUS YANG DIMAKSUD! DIANGGAP BERSIH!
        GoTo LBL_FREE_OBJECTS
    End If
    '---analisis per-section:
    ReDim ISECH(IMNTH.FileHeader.NumberOfSections - 1) As IMAGE_SECTION_HEADER '--aloaksikan ruang penyimpanan section headers.
    If nFileLen <= (IDOSH.e_lfanew + Len(IMNTH) + Len(IOH32) + (Len(ISECH(0)) * IMNTH.FileHeader.NumberOfSections)) Then
        nFunctionResult = 0 'BUKAN,INI BUKAN VIRUS YANG DIMAKSUD! DIANGGAP BERSIH!
        GoTo LBL_FREE_OBJECTS
    End If
    Call SetFilePointer(hFileHandle, IDOSH.e_lfanew + Len(IMNTH) + Len(IOH32), 0, 0)
    nCallResult = ReadFile(hFileHandle, ISECH(0), Len(ISECH(0)) * IMNTH.FileHeader.NumberOfSections, nRetLength, ByVal 0)
    If nRetLength <> Len(ISECH(0)) * IMNTH.FileHeader.NumberOfSections Then
        nFunctionResult = 0 'BUKAN,INI BUKAN VIRUS YANG DIMAKSUD! DIANGGAP BERSIH!
        GoTo LBL_FREE_OBJECTS
    End If
    For CTurn = 0 To (IMNTH.FileHeader.NumberOfSections - 1)
        If ISECH(CTurn).SizeOfRawData > 0 Then '---pastikan hanya yg ada isinya saja:
            If (IOH32.AddressOfEntryPoint >= ISECH(CTurn).VirtualAddress) And (IOH32.AddressOfEntryPoint < (ISECH(CTurn).VirtualAddress + ISECH(CTurn).SizeOfRawData)) Then
                pVarPos = ISECH(CTurn).PointerToRawData + (IOH32.AddressOfEntryPoint - ISECH(CTurn).VirtualAddress)
                If nFileLen <= pVarPos Then
                    bThreatFound = False
                    nFunctionResult = 0 'BUKAN,INI BUKAN VIRUS YANG DIMAKSUD! DIANGGAP BERSIH!
                    GoTo LBL_FREE_OBJECTS
                End If
                nVarLen = (ISECH(CTurn).PointerToRawData + ISECH(CTurn).SizeOfRawData) - pVarPos '---panjang dari ep ke akhir section tersebut.
                gVEPSecNumber = CTurn '---number Base0.
                '---cek spesifik ep:[per-virus-variant]:
                nVarData = 32 '---ambil 32 bytes awal.
                ReDim btCPattern(nVarData - 1) As Byte
                Call SetFilePointer(hFileHandle, pVarPos, 0, 0)
                If nVarLen < nVarData Then
                    nVarData = nVarLen
                End If
                nCallResult = ReadFile(hFileHandle, btCPattern(0), nVarData, nRetLength, ByVal 0)
                If nRetLength <> nVarData Then
                    bThreatFound = False
                    nFunctionResult = 0 'BUKAN,INI BUKAN VIRUS YANG DIMAKSUD! DIANGGAP BERSIH!
                    GoTo LBL_FREE_OBJECTS
                End If
                '---cek pattern (versi sederhana, tanpa optimisasi):
                If btCPattern(0) = &H60 And _
                    btCPattern(1) = &HE8 And _
                    btCPattern(6) = &H8B And _
                    btCPattern(7) = &H74 And _
                    btCPattern(10) = &HE8 And _
                    btCPattern(15) = &H61 And _
                    btCPattern(16) = &H68 Then
                    bThreatFound = True                        '---menurut byte pattern adalah:[cocok].
                    'MsgBox "Benar"
                End If
                '--------------------------------------;
                Exit For '---sudah,jangan terlalu lama berputar, 'ntar pusing :)
            End If
        End If
    Next

    If bTryToFix = False Then
        If bThreatFound = True Then
            nFunctionResult = 1 'YA! INI ADALAH FILE VIRUS! HATI-HATI!
        Else
            nFunctionResult = 0 'BUKAN,INI BUKAN VIRUS YANG DIMAKSUD! DIANGGAP BERSIH!
        End If
        GoTo LBL_FREE_OBJECTS
    Else '---mau lanjut tapi nggak ada yg dilanjutin, ya akhiri saja :)
        If bThreatFound = True Then
            nFunctionResult = 1 'preset: YA, INI FILE VIRUS[...].
        Else
            nFunctionResult = 0 'BUKAN,INI BUKAN VIRUS YANG DIMAKSUD! DIANGGAP BERSIH!
            GoTo LBL_FREE_OBJECTS
        End If
    End If
    '---yg melewati baris ini dianggap file ber-virus dan minta dibersihkan:
Dim pAddrOEP                As Long
    Call RtlMoveMemory(VarPtr(pAddrOEP), VarPtr(btCPattern(17)), Len(pAddrOEP)) '---ambil 4 bytes.
    '*+"V":[50]---tanda/flag-IGNORE.
    '*Optional.AddressOfEntryPoint-OK.
    '*Optional.SizeOfImage-OK.
    '*Section.SizeOfRawData-OK.
Dim pBuffer                 As Long
Dim nBufferLength           As Long
Dim pInBuf2                 As Long
Dim nInBufLen2              As Long
    nBufferLength = nFileLen
    pBuffer = LocalAlloc(lptr, nBufferLength)
    If pBuffer <= 0 Then
        nFunctionResult = 1 'YA, INI FILE VIRUS TAPI GAGAL DIBERSIHKAN :(
        GoTo LBL_FREE_OBJECTS
    End If
    '---backup file ke memori[sementara] sebelum di-reset:
    Call SetFilePointer(hFileHandle, 0, 0, 0) '---mulai dari awal lagi:
    nCallResult = ReadFile(hFileHandle, ByVal pBuffer, nBufferLength, nRetLength, ByVal 0)
    If nRetLength <> nBufferLength Then
        nFunctionResult = 1 'YA, INI FILE VIRUS TAPI GAGAL DIBERSIHKAN :(
        GoTo LBL_FREE_OBJECTS
    End If
    '---reset isi file:
    Call SetFilePointer(hFileHandle, 0, 0, 0) '---mulai dari awal lagi:
    Call SetEndOfFile(hFileHandle)
    '---tulis ulang dos-header[tanpa perubahan]:
    Call SetFilePointer(hFileHandle, 0, 0, 0)
    nCallResult = WriteFile(hFileHandle, IDOSH, Len(IDOSH), nRetLength, ByVal 0)
    If nRetLength <> Len(IDOSH) Then
        nFunctionResult = 1 'YA, INI FILE VIRUS TAPI GAGAL DIBERSIHKAN :(
        GoTo LBL_FREE_OBJECTS
    End If
    '---tulis ulang nt-header[tanpa perubahan]:
    Call SetFilePointer(hFileHandle, IDOSH.e_lfanew, 0, 0)
    nCallResult = WriteFile(hFileHandle, IMNTH, Len(IMNTH), nRetLength, ByVal 0)
    If nRetLength <> Len(IMNTH) Then
        nFunctionResult = 1 'YA, INI FILE VIRUS TAPI GAGAL DIBERSIHKAN :(
        GoTo LBL_FREE_OBJECTS
    End If
    '---tulis ulang optional-header[dengan perubahan]:
    IOH32.AddressOfEntryPoint = pAddrOEP - IOH32.ImageBase  '---tulis entrypoint yg sebenarnya [oep] RVA.
    IOH32.SizeOfImage = (((IOH32.SizeOfHeaders - 1) \ IOH32.SectionAlignment) + 1) * IOH32.SectionAlignment '---diisi dgn ukuran alokasi headers di memori terlebih dahulu(pembulatan terhadap section alignment).
    CTurn = 0 '---reset counter.
    For CTurn = 0 To (IMNTH.FileHeader.NumberOfSections - 1)
        IOH32.SizeOfImage = IOH32.SizeOfImage + ((((ISECH(CTurn).VirtualSize - 1) \ IOH32.SectionAlignment) + 1) * IOH32.SectionAlignment) '---diisi dgn total atau penambahan keseluruhan ukuran section di memori(pembulatan terhadap section alignment).
    Next
    Call SetFilePointer(hFileHandle, IDOSH.e_lfanew + Len(IMNTH), 0, 0)
    nCallResult = WriteFile(hFileHandle, IOH32, Len(IOH32), nRetLength, ByVal 0)
    If nRetLength <> Len(IOH32) Then
        nFunctionResult = 1 'YA, INI FILE VIRUS TAPI GAGAL DIBERSIHKAN :(
        GoTo LBL_FREE_OBJECTS
    End If
    '---pasang kembali semua 'section-header'[dengan perubahan]:
    '---hapus kode virus di section tempat kode virus berada(sebenarnya nggak dihapus nggak apa-apa,tapi ini biar ngirit media penyimpanan):
    ISECH(gVEPSecNumber).SizeOfRawData = pVarPos - ISECH(gVEPSecNumber).PointerToRawData
    ISECH(gVEPSecNumber).Characteristics = ISECH(gVEPSecNumber).Characteristics And Not &H20000000
    Call SetFilePointer(hFileHandle, IDOSH.e_lfanew + Len(IMNTH) + Len(IOH32), 0, 0)
    nCallResult = WriteFile(hFileHandle, ISECH(0), Len(ISECH(0)) * IMNTH.FileHeader.NumberOfSections, nRetLength, ByVal 0)
    If nRetLength <> (Len(ISECH(0)) * IMNTH.FileHeader.NumberOfSections) Then
        nFunctionResult = 1 'YA, INI FILE VIRUS TAPI GAGAL DIBERSIHKAN :(
        GoTo LBL_FREE_OBJECTS
    End If
    '---pasang kembali [?]:ini header/tanda/apaan,sih? dari dulu saya nggak paham bagian ini:
    CTurn = 0 '---reset counter.
    nVarData = 0 '---reset---:variabel nVarData lagi nganggur, jadi bisa di-daur ulang :)
    For CTurn = 0 To (IMNTH.FileHeader.NumberOfSections - 1)
        If ISECH(CTurn).SizeOfRawData > 0 Then
            If nVarData > 0 Then
                If ISECH(CTurn).PointerToRawData < nVarData Then
                    nVarData = ISECH(CTurn).PointerToRawData '---masukkan niai yg lebih kecil.
                End If
            Else
                nVarData = ISECH(CTurn).PointerToRawData '---untuk pertama kali, langsung masukkan ajah.
            End If
        End If
    Next
    '---cek apakah ada jeda antara akhir dari section header dgn awal dari section data:
    CTurn = IDOSH.e_lfanew + Len(IMNTH) + Len(IOH32) + (Len(ISECH(0)) * IMNTH.FileHeader.NumberOfSections)
    If nVarData > CTurn Then
        Call SetFilePointer(hFileHandle, CTurn, 0, 0)
        nCallResult = WriteFile(hFileHandle, ByVal (pBuffer + CTurn), nVarData - CTurn, nRetLength, ByVal 0)
        If nRetLength <> (nVarData - CTurn) Then
            nFunctionResult = 1 'YA, INI FILE VIRUS TAPI GAGAL DIBERSIHKAN :(
            GoTo LBL_FREE_OBJECTS
        End If
    End If
    '---pasang kembali tiap section di tempat yg benar:
    CTurn = 0 '---reset counter.
    For CTurn = 0 To (IMNTH.FileHeader.NumberOfSections - 1)
        If CTurn <> gVEPSecNumber Then '---bila bukan section yg kena infeksi virus, biarkan apa adanya:
            Call SetFilePointer(hFileHandle, ISECH(CTurn).PointerToRawData, 0, 0)
            nCallResult = WriteFile(hFileHandle, ByVal (pBuffer + ISECH(CTurn).PointerToRawData), ISECH(CTurn).SizeOfRawData, nRetLength, ByVal 0)
            If nRetLength <> ISECH(CTurn).SizeOfRawData Then
                nFunctionResult = 1 'YA, INI FILE VIRUS TAPI GAGAL DIBERSIHKAN :(
                GoTo LBL_FREE_OBJECTS
            End If
        Else '---bila section yg kena infeksi virus, sesuaikan dulu file-alignment-nya, bila bukan kode akan di-reset ke 0[zero]:
            '---variabel nVarData lagi nganggur, jadi bisa di-daur ulang :)
            nVarData = (((ISECH(CTurn).SizeOfRawData - 1) \ IOH32.FileAlignment) + 1) * IOH32.FileAlignment '---diisi dgn ukuran alokasi headers di file-fisik terlebih dahulu(pembulatan terhadap file-alignment).
            '---isi dgn zero-bytes:
            nInBufLen2 = nVarData
            pInBuf2 = LocalAlloc(lptr, nInBufLen2)
            If pInBuf2 <= 0 Then
                Call SetFilePointer(hFileHandle, ISECH(CTurn).PointerToRawData, 0, 0)
                nCallResult = WriteFile(hFileHandle, 0, nVarData, nRetLength, ByVal 0) '---akan diisi dgn data sampah dari memori heap.
                If nRetLength <> nVarData Then
                    nFunctionResult = 1 'YA, INI FILE VIRUS TAPI GAGAL DIBERSIHKAN :(
                    GoTo LBL_FREE_OBJECTS
                End If
            Else '---akan diisi dgn data zero-bytes teralokasi:
                Call SetFilePointer(hFileHandle, ISECH(CTurn).PointerToRawData, 0, 0)
                nCallResult = WriteFile(hFileHandle, ByVal pInBuf2, nInBufLen2, nRetLength, ByVal 0) '---akan diisi dgn data sampah dari memori heap.
                If nRetLength <> nInBufLen2 Then
                    '--jangan lupa untuk tutup memori yg teralokasi:
                    If pInBuf2 > 0 Then
                        Call LocalFree(pInBuf2)
                        pInBuf2 = 0
                    End If
                    '---akhiri:
                    nFunctionResult = 1 'YA, INI FILE VIRUS TAPI GAGAL DIBERSIHKAN :(
                    GoTo LBL_FREE_OBJECTS
                End If
                If pInBuf2 > 0 Then
                    Call LocalFree(pInBuf2)
                    pInBuf2 = 0
                End If
            End If
            '---isi dgn kode sebenarnya:
            Call SetFilePointer(hFileHandle, ISECH(CTurn).PointerToRawData, 0, 0)
            nCallResult = WriteFile(hFileHandle, ByVal (pBuffer + ISECH(CTurn).PointerToRawData), ISECH(CTurn).SizeOfRawData, nRetLength, ByVal 0)
            If nRetLength <> ISECH(CTurn).SizeOfRawData Then
                nFunctionResult = 1 'YA, INI FILE VIRUS TAPI GAGAL DIBERSIHKAN :(
                GoTo LBL_FREE_OBJECTS
            End If
        End If
    Next
    '---yg berhasil melewati baris ini dianggap berhasil men-dis-infeksi virus:
    nFunctionResult = 2 'YA, INI FILE VIRUS DAN BERHASIL DIBERSIHKAN!
    '---selesai membersihkan virus,sekarang bersihkan objects:
LBL_FREE_OBJECTS:
    Call RtlZeroMemory(VarPtr(IDOSH), Len(IDOSH))
    Call RtlZeroMemory(VarPtr(IMNTH), Len(IMNTH))
    Call RtlZeroMemory(VarPtr(IOH32), Len(IOH32))
    Erase ISECH()
    Erase btCPattern()
    If pBuffer > 0 Then
        Call LocalFree(pBuffer)
        pBuffer = 0
    End If
    If hFileHandle > 0 Then
        If bInpuFromFileHandle = False Then
            Call CloseHandle(hFileHandle)
            hFileHandle = 0
        End If
    End If
LBL_BROADCAST_RESULT:
    CheckNFixFH_V_W32_Runouce_a = nFunctionResult
LBL_TERAKHIR:
    If err.Number <> 0 Then
        err.Clear
    End If
End Function

