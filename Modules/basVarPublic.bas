Attribute VB_Name = "basVarPublic"
Public RegNode      As Boolean ' Penampung Node Registry
Public ProsesNode   As Boolean ' Penampung Node Proses dan Service
Public StartUpNode  As Boolean ' Penampung Node Startup
Public WinNode      As Boolean
Public DocNode      As Boolean
Public ProgNode     As Boolean


Public JailExt      As String
Public FolderJail   As String
'Public FolderAv   As String
'Public FolderHelp   As String
'Public FolderPlugin   As String
'Public FolderSign   As String
'Public FolderSignx   As String
'Public FolderSounds   As String

Public StatScan     As String ' status scan
Public FolToScan    As Long
Public FileToScan   As Long
Public WithBuffer   As Boolean
Public BERHENTI     As Boolean

Public TmpIsiFileSkrg As String
Public nSalityGet     As String ' penerima nilai dari fungsi CekKemungkinanSality
Public nPEHeurGet     As String ' penerima nilai dari PE Heur
Public nGetAlman      As Boolean ' Penerima status cek alaman cara baru
Public nGetAlmanB     As Boolean ' Penerima status cek alaman cara baru (var B)

Public InfoFound    As Long
Public VirusFound   As Long
Public FileFound    As Long
Public FileCheck    As Long
Public FileNotCheck As Long
Public nHiddenObj   As Long
Public hGlobal      As Long
Public nSizeGlobal  As Long


Public xSectionAkhir    As String
Public xNamaSectionAkhir As String
Public xSectionAkhir2    As String
Public xNamaSectionAkhir2 As String
Public xSectionJum      As String

Public LastFlashVolume As Long

Public nVirusTmp       As Long ' jumlah virus TMP

Public nRealSizePE  As Long ' menampung ukuran file PE asli yang sedang di cek (hampir sama dengan nSizeGlobal)

Public nErrorReg     As Long ' Jumlah registry yang bermasalah
Public nRegVal       As Long ' Jumlah registry yang di scan

' Penanganan Ceksum dan NamaVirus PE
Public sMD5(15, 400)          As String
Public sNamaVirus(15, 400)    As String ' karena dua dimensi harus deklarasi dulu elemenya biar ga repot

' Penanganan Ceksum dan NamaVirus non PE
Public sMD5nonPE(15, 400)          As String
Public sNamaVirusnonPE(15, 400)    As String ' karena dua dimensi harus deklarasi dulu elemenya biar ga repot

' Penanganan Ceksum dan NamaVirus milik User Atau ArrS masuk sini ajh
Public sMD5User(20)          As String
Public sNamaVirusUser(20)    As String ' karena dua dimensi harus deklarasi dulu elemenya biar ga repot

Public JumVirusUser     As Long      ' Jumlah Database virus User atau ArrS

Public JumVirus              As Long      ' Jumlah Database seluruhnya (PE dan Non PE)
Public JumlahVirus()         As Long      ' Jumlah virus masing-masing DB PE
Public JumlahVirusNonPE()    As Long      ' Jumlah virus masing-masing DB non PE

Public VirStatus    As Boolean ' Status apakah virus atau bukan selama discan dengan ceksum + heuristic (jika diset)
Public IsPE32EXE    As Boolean ' menampung status file yang sdg di cek adalah PE EXE 32

Public LangUsed     As String ' Penampung FileBhasa yang digunakan (bisa default@1/2)

Public BufferCtlDwonload As Boolean ' true klo proses doenload error atau selesai
