Attribute VB_Name = "basLang"

Public a_bahasa(100)   As String
Public b_bahasa(100)   As String
Public c_bahasa(100)   As String
Public d_bahasa(100)   As String
Public e_bahasa(100)   As String
Public f_bahasa(100)   As String
Public g_bahasa(100)   As String
Public h_bahasa(100)   As String
Public i_bahasa(100)   As String
Public j_bahasa(100)   As String


Public Sub InitLanguange(LangName As String)
Dim LangPath As String
LangPath = GetFilePath(App_FullPathW(False)) & "\lang"

Select Case LangName
    Case "Default@1": Call InitInternalLang ' ENG
   ' Case "Default@2": Call InitInternalLang2 ' IND
    Case Else: GoTo LBL_CARI2
End Select

GoTo LBL_AKHIR

LBL_CARI2:
If ValidFile(LangPath & "\" & LangName) = True Then ' jika ada
   If ReadExternalLang(LangPath & "\" & LangName) = False Then
      GoTo LBL_FALSE
   End If
Else
  GoTo LBL_FALSE
End If

'GoTo LBL_AKHIR ' klo gagal init extrenal (bisa saja bahasa extrenal tidak valid)

LBL_FALSE:
    Call InitInternalLang ' ENG ' balik default jika gagal
   ' MsgBox j_bahasa(21), vbExclamation

LBL_AKHIR:
   Call WriteLangToInterface ' tulis ke interface
End Sub

Private Sub InitInternalLang()
'[A] Bahasa Tombol ==============================================================================================

a_bahasa(0) = "Summary"
a_bahasa(1) = "Scan Computer"
a_bahasa(2) = "Additional Protection"
a_bahasa(3) = "Virus Chest"
a_bahasa(4) = "Update"

a_bahasa(5) = "Start Scan"
a_bahasa(6) = "Skip Buffer.."
a_bahasa(7) = "Abort Scan"
a_bahasa(8) = "Fix Checked"
a_bahasa(9) = "Fix All Object"
a_bahasa(10) = "Properties"
a_bahasa(11) = "Explore"
a_bahasa(12) = "Quarantine All"
a_bahasa(13) = "Fix All"
a_bahasa(14) = "Remove All"
a_bahasa(15) = "Add Path"
a_bahasa(16) = "Add File"
a_bahasa(17) = "Execute Plugin"
a_bahasa(18) = "Log Scan"
a_bahasa(19) = "Add"
a_bahasa(20) = "Cancel"
a_bahasa(21) = "Add Quarantine"
a_bahasa(22) = "Remove Selected"
a_bahasa(23) = "Restore.."
a_bahasa(24) = "Restore To.."
a_bahasa(25) = "Result >>"
a_bahasa(26) = "Update Now"
a_bahasa(27) = "Ignore"
a_bahasa(28) = "Delete"
a_bahasa(29) = "Enable"
a_bahasa(30) = "Disable"
a_bahasa(31) = "Upload Virus"

a_bahasa(32) = "Save"
a_bahasa(33) = "Edit"
a_bahasa(34) = "Apply"

a_bahasa(35) = "Ok"
'a_bahasa(36) = "Login"
'a_bahasa(37) = "Change Password"
'a_bahasa(38) = "Remove Password"
'a_bahasa(39) = "Read Me"
'a_bahasa(40) = "Refresh Quarantine"
'a_bahasa(41) = "Save Color"
'a_bahasa(42) = "Stop"
'a_bahasa(43) = "Download Update"
'a_bahasa(44) = "Search"
'a_bahasa(45) = "More >>"
'a_bahasa(46) = "Delete All"
'a_bahasa(47) = "Enable"
'a_bahasa(48) = "Disable"
'a_bahasa(49) = ""
'a_bahasa(50) = "Fix All"

'[B] Bahasa

'b_bahasa(0) = "Path Scan"
'b_bahasa(1) = "Malware"
b_bahasa(2) = "Registry"
b_bahasa(3) = "Hidden"
b_bahasa(4) = "Information"

'b_bahasa(13) = "Virus Chest"


'[C] Bahasa Frem=============================================================================================================

'c_bahasa(0) = "Scan Option"
'c_bahasa(1) = "Application Configuration"
c_bahasa(2) = "Process List"
c_bahasa(3) = "Module List"
'c_bahasa(4) = "Malware Temporary"
'c_bahasa(5) = "List of Temporary"
'c_bahasa(6) = "Quarantine"
'c_bahasa(7) = "List of Detected Malware"
'c_bahasa(8) = "Software Information"
'c_bahasa(9) = " Internal Virus Detector"
'c_bahasa(10) = "Thanks for using Wan'iez Antivirus AntiVirus, support us with send new malware sampel to Wan'iez AntiVirus Team"
'c_bahasa(11) = "Have a problem with Malware? or your Wan'iez AntiVirus not detect Malware that infected your computer! Send the Malware sample to our e-mail address Caswandi14@gmail.com we will analyze the malware and add it into Wan'iez database. Any problem with Wan'iez AntiVirus, or have a suggestion, you can tell us at: [+6285853618213]"
   
'[D] Bahasa Label============================================================================================================

d_bahasa(0) = "Info Scan"
d_bahasa(1) = "Processed"
d_bahasa(2) = "Time"
d_bahasa(3) = "Founded"
d_bahasa(4) = "Checked"
d_bahasa(5) = "ByPassed"
d_bahasa(6) = "Malware"
d_bahasa(7) = "Registry"
d_bahasa(8) = "Information"
d_bahasa(9) = "[Ready]"
d_bahasa(10) = "Scan Registry !"
d_bahasa(11) = "Scan Service !"
d_bahasa(12) = "Scan Process !"
d_bahasa(13) = "Scan Startup !"
d_bahasa(14) = "Scan Root !"
d_bahasa(15) = "Scanning "
d_bahasa(16) = "Aborted"
d_bahasa(17) = "Finished"
    
d_bahasa(18) = "Language Selected"
d_bahasa(19) = "Language ID"
d_bahasa(20) = "Language Author"
d_bahasa(21) = "Dont Give me Warning about Threat in this Path"
d_bahasa(22) = "I'am sure this is normal file, dont catch as a malware file(s) below"
d_bahasa(23) = "I'am sure this is normal value. Don't catch as a bad value, value(s) below"
d_bahasa(24) = "Avalaible Plugin(s)"
d_bahasa(25) = "Avalaible Language"
d_bahasa(26) = "Plugin Selected"
d_bahasa(27) = "Plugin Author"
d_bahasa(28) = "Plugin Description"
d_bahasa(29) = "Malware Path"
d_bahasa(30) = "Malware Name"
'd_bahasa(31) = "Engine Version"
'd_bahasa(32) = "Build Number"
'd_bahasa(33) = "Build Date"
'd_bahasa(34) = "Reg Database"
'd_bahasa(35) = "Worm Signature"
'd_bahasa(36) = "Virus Signature"
'd_bahasa(37) = "Machine"
'd_bahasa(47) = "Total File(s) should be remove : "
    
'd_bahasa(38) = " object(s)"
'd_bahasa(39) = "Congratulation to you for using Wan'iez AntiVirus with all of its advatages and disadvantages...For finding new update database klick Update Now"
'd_bahasa(40) = "Please Wait..."
'd_bahasa(41) = "Enable Use Heuristic"
'd_bahasa(42) = "Read Me"
'd_bahasa(43) = "Next=>"
'd_bahasa(44) = "<=Back"
'd_bahasa(45) = "To get serial key Wan'iez PRO, you must send serial AV : "
'd_bahasa(46) = " to the following number : 1.(+6285853618213)= Heru's, number : 2.(+6287736190060)= Ahlul's, number : 3.(+6285865929068)= Rachmat's, with your name and your address for our user data. Key PRO will be send max in 1 X 24 hour after we received the message. Key PRO can use forever although there is Wan'iez update and new feature added"
   
'd_bahasa(48) = "Registry Mechanic"
'd_bahasa(49) = "Make Password"
'd_bahasa(50) = "Repeat the Password"
'd_bahasa(51) = "Because there are many Wan'iez user sugestion about Donation for Wan'iez AntiVirus so I think this is important for Wan'iez AntiVirus in the future. So I make bank account to save the donation. If you, or your friend want to donate us, so hope Wan'iez AntiVirus can be better."
'd_bahasa(52) = "Every One want to Donate Wan'iez AntiVirus, can send your donation to our bank account below"
'd_bahasa(53) = "NAME : LESTARI"
'd_bahasa(54) = "BRI BANK Unit Sukorejo :"
'd_bahasa(55) = "Bank account number : 6499.01.000802.53.6"
'd_bahasa(56) = "Registry Mechanic [1]"
'd_bahasa(57) = "Registry Mechanic [2]"
'd_bahasa(58) = "PRO TRIAL feature has been run  " ' & jumlah & " Kali" & " Sisa pemakaian " & sisa & " Kali"
'd_bahasa(59) = " Time"
'd_bahasa(60) = " the remainder of using "
d_bahasa(61) = "Wan'iez Real Time Protection Found " '[ " & frmRTP.lvRTP.ListItems.Count & " ] Malware!!!"
'd_bahasa(62) = "Maximum Protection"
'd_bahasa(63) = "Not Protection"
'd_bahasa(64) = "Green"
'd_bahasa(65) = "Red"
'd_bahasa(66) = "Blue"
'd_bahasa(67) = "Software Information"
'd_bahasa(68) = "System Information"

'[E] Bahasa Listview Luar====================================================================================================

e_bahasa(0) = "Malware Name"
e_bahasa(1) = "Object Path"
e_bahasa(2) = "Size [B]"
e_bahasa(3) = "Information"
    
e_bahasa(4) = "Value Name"
e_bahasa(5) = "Key Path"
    
e_bahasa(6) = "Object Name"
e_bahasa(7) = "File Name"
    
e_bahasa(8) = "Process Name"
e_bahasa(9) = "Startup"
e_bahasa(10) = "ParentPID"
e_bahasa(11) = "Update Path"
e_bahasa(12) = "Hidden"
e_bahasa(13) = "In Debug"
e_bahasa(14) = "Locked"

e_bahasa(15) = "Virus Name"
e_bahasa(16) = "Original Path"
e_bahasa(17) = "In jail"

e_bahasa(18) = "Attributes"
e_bahasa(19) = "Hash File Worm"

e_bahasa(20) = "Drive"
e_bahasa(21) = "Status"
e_bahasa(22) = "Type"
    
'[F] Listview Dalam =========================================================================================================

f_bahasa(0) = "Hidden"
f_bahasa(1) = "Suspected With"
f_bahasa(2) = "Suspected File"
'f_bahasa(3) = "Bad PE File"
'f_bahasa(4) = "Dirty PE File"
'f_bahasa(5) = "Contain too much additonal bytes - Potensial Dropper/Installer (please send to us if you also suspect it)"
f_bahasa(6) = "Infected file"
f_bahasa(7) = "Virus file"
f_bahasa(8) = "Malware Startup"
f_bahasa(9) = "Useless Value, Should be deleted"
f_bahasa(10) = "Value Deleted"
f_bahasa(11) = "String Fixed"
f_bahasa(12) = "DWORD Fixed"
f_bahasa(13) = "File Normalized"
f_bahasa(14) = "Folder Normalized"
f_bahasa(15) = "Restored"
f_bahasa(16) = "Sent to jail !"
f_bahasa(17) = "Sent to jail but fail remove source !!"
f_bahasa(18) = "Fail sent to jail and remove source !!"
f_bahasa(19) = "File System"
f_bahasa(20) = "File is needed by system to run normally !"
f_bahasa(21) = "Contain too much additional bytes (maybe your data), "
f_bahasa(22) = "Shortcut not target"
'f_bahasa(23) = "Virus By.User"
f_bahasa(24) = "Suspect by user database"
    
'[G] Menu Editor ============================================================================================================
     
g_bahasa(0) = "Hide Wan'iez ! user interface"
g_bahasa(1) = "Open Wan'iez ! user interface"
g_bahasa(2) = "Wan'iez Protection Control"
g_bahasa(3) = "Enable Protection"
g_bahasa(4) = "Disable for 1 minutes"
g_bahasa(5) = "Disable for 10 minutes"
g_bahasa(6) = "Run On Startup"
g_bahasa(7) = "Setting"
g_bahasa(8) = "About"
g_bahasa(9) = "Exit"

g_bahasa(10) = "Properties"
g_bahasa(11) = "Refresh Process"
g_bahasa(12) = "Kill Process"
g_bahasa(13) = "Restart Process"
g_bahasa(14) = "Pause Process"
g_bahasa(15) = "Resume Process"

g_bahasa(16) = "Clear Quarantine"
g_bahasa(17) = "Kill Selected"
g_bahasa(18) = "Submit For Analysis"

g_bahasa(19) = "Locked"
g_bahasa(20) = "UnLocked"

'[H] Bahasa Check Box =======================================================================================================
    
h_bahasa(0) = "Filter file (by pass file with certain extensions)"
h_bahasa(1) = "Use Heuristic to suspect malware"
h_bahasa(2) = "Detect useless registry value (XP only)"
h_bahasa(3) = "Detect hidden object (file and folder)"
h_bahasa(4) = "Give strange  information while scanning"
h_bahasa(5) = "Place Application on Top"
    
h_bahasa(6) = "Launch Wan'iez AntiVirus at computer startup"
h_bahasa(7) = "Wan'iez Antivirus Real Time Protection"
h_bahasa(8) = "Auto Check Online Update"
h_bahasa(9) = "Auto Scan automatically when Flashdisk plugged"
h_bahasa(10) = "Create Context Menu"
h_bahasa(11) = "Splash screen on start-up"
h_bahasa(12) = "Create shortcut to desktop"
h_bahasa(13) = "Check All"
h_bahasa(14) = "Scan File(s) at copy and exstract"
    
h_bahasa(15) = "Password Protection"
h_bahasa(16) = "Real Time Protection"
h_bahasa(17) = "Virus Chest"
h_bahasa(18) = "Additional Protection"
h_bahasa(19) = "Update"
h_bahasa(20) = "Delete Virus By User"
h_bahasa(21) = "Exit Application"

    
'[I] Bahasa Pesan MSGBOX ====================================================================================================
    
i_bahasa(0) = "All Quarantina killed !"
i_bahasa(1) = "File is already exist. Do you want to over write?"
i_bahasa(2) = "Quarantina restore to"
i_bahasa(3) = "Original path is not avalaible  - use custom path to restore Quarantina !"
i_bahasa(4) = "Success unload selected module"
i_bahasa(5) = "Fail to unload selected module"
i_bahasa(6) = "Maybe work well after application restarted !"
i_bahasa(7) = "Select a file to be add user database !"
i_bahasa(8) = "Please terminate scanning process first !"
i_bahasa(9) = "Are you sure to unload selected module?"
i_bahasa(10) = "Process with PID"
i_bahasa(11) = "was terminated succesfully !"
i_bahasa(12) = "cannot be terminated !"
i_bahasa(13) = "was restarted succesfully !"
i_bahasa(14) = "cannot be restarted !"
'i_bahasa(15) = "Restart Application for apply all change"
'i_bahasa(16) = "success added as new temporary malware sample !"
'i_bahasa(17) = "Your new malware name"
'i_bahasa(18) = "Scan computer to view the result"
i_bahasa(19) = "Are you sure to clear all Quarantina in Wan'iez jail ?"
i_bahasa(20) = "Fail added as new malware"
i_bahasa(21) = "Are you sure to kill selected Quarantinas"
i_bahasa(22) = "Are you sure to restore selected Quarantina"
i_bahasa(23) = "Wan'iez fail to get file system list, it can make Wan'iez delete virus file although needed by your system"
'i_bahasa(24) = "Wan'iez Protector is turn ON, your system are protected by Wan'iez AntiVirus+ now"
'i_bahasa(25) = "Wan'iez Protector is turn OFF, Wan'iez AntiVirus+ is rest to protect your system"
i_bahasa(26) = "Information"
i_bahasa(27) = "Caution"
'i_bahasa(28) = "Your computer isn't connected to the internet"
'i_bahasa(29) = "It's not recommended to disable scanning in "
'i_bahasa(30) = "this will cause scanning won't be perfec! "
'i_bahasa(31) = "Are you sure you want to disable it? "
'i_bahasa(32) = "Please select object will be scan"
i_bahasa(33) = "Unfortunately, Your computer system is "
i_bahasa(34) = "NOT GOOD "
i_bahasa(35) = "BAD "
i_bahasa(36) = "GOOD "
i_bahasa(37) = "Wan'iez AntiVirus Found: "
i_bahasa(38) = "Computer Status "
i_bahasa(39) = "Infected Registry(s) "
i_bahasa(40) = "Malware File(s) "
i_bahasa(41) = "Hidden File(s) "
i_bahasa(42) = "Fortunately, your computer system is FINE NO MALWARE DETECTED."
'i_bahasa(43) = "Your current Wan'iez-RTP (Wan'iez Real Time Protection) is not active in your system"
'i_bahasa(44) = "Do you want activate it now?"
i_bahasa(45) = "You don't have Wan'iez-RTP (Wan'iez Real Time Protection) in your system"
'i_bahasa(46) = "Do you want to activate it now?"
i_bahasa(47) = "Your current Wan'iez-RTP (Wan'iez Real Time Protection) is OLD version"
'i_bahasa(48) = "Do you want to update it with the NEW version?"
'i_bahasa(49) = ""

'[J] Bahasa Yang Lainya =====================================================================================================

j_bahasa(0) = "Scan Module !"
j_bahasa(1) = "Service Fail Destroyed"
j_bahasa(2) = "Service Destroyed"
j_bahasa(3) = "Service"
j_bahasa(4) = "Found in Reg-Startup"
'j_bahasa(5) = "Virus Startup"
j_bahasa(6) = "Found in Explorer-Startup"
j_bahasa(7) = "in Memory [Terminated+Locked]"
j_bahasa(8) = "[Unload From Process]"
j_bahasa(9) = "My-Virus"
j_bahasa(10) = "System Area"
j_bahasa(11) = "Process + Service"
j_bahasa(12) = "This program can not be run"
j_bahasa(13) = "checked item"
j_bahasa(14) = "selected item"
j_bahasa(15) = "all item"
j_bahasa(16) = "We are sorry WanUI.dll Not found"
j_bahasa(17) = "We are sorry WanSM.dll Not found"
j_bahasa(18) = "We are sorry WanUDB.dll Not found"
j_bahasa(19) = "Some items of detected malware has not been fixed."
j_bahasa(20) = "You will lost information if you continue scanning process, Are you sure ?"
'j_bahasa(21) = "Fail read external language !"
j_bahasa(22) = "Cannot enumerate modules from selected process"
'j_bahasa(23) = "Remove Selected"
j_bahasa(24) = "Fail to get Update information !"
j_bahasa(25) = "No New Update avalaible for your Wan'iez version"
j_bahasa(26) = "Getting Update Info ...."
'j_bahasa(27) = ""
j_bahasa(28) = "Some of database is fail to read"
j_bahasa(29) = "Updating"
j_bahasa(30) = "Downloading"
j_bahasa(31) = "Done.."
j_bahasa(32) = "Update Component Canceled.."
j_bahasa(33) = "Checking Update.."
j_bahasa(34) = "Cancel Update"
j_bahasa(35) = "[Special Paths]"
j_bahasa(36) = "Language Used"
j_bahasa(37) = "Buffering"
j_bahasa(38) = "file(s)"
j_bahasa(39) = "folder(s)"
j_bahasa(40) = "You are not use WinXP OS, please turn off"
j_bahasa(41) = "File is Corupted"
j_bahasa(42) = "Download back to http:/www.waniez.p.ht"
j_bahasa(43) = "File Found"
j_bahasa(44) = "File Scanned"
j_bahasa(45) = "File Not Scanned"
j_bahasa(46) = "Virus Found"
'j_bahasa(47) = "Value Scanned"
j_bahasa(48) = "Bad Value"
'j_bahasa(49) = "End Time"
'j_bahasa(50) = "Enable Context Menu -Scan With"
'j_bahasa(51) = "Scan With"
j_bahasa(52) = "Author Email"
j_bahasa(53) = "Author Site"
j_bahasa(54) = "Verification Code"
j_bahasa(55) = "No Plugin Avalaible - Get Wan'iez plugin at"
j_bahasa(56) = "Are you sure to execute selected plugin"
j_bahasa(57) = "Run as new Thread"
j_bahasa(58) = "Plugin can't be executed"
'j_bahasa(59) = "Found Virus In Your Memori"
j_bahasa(60) = "Removable drive detected"
j_bahasa(61) = "detect new removable drive inserted, would you like to scan with"
'j_bahasa(62) = "Scanning removable drive inserted"
j_bahasa(63) = "Wan'iez AntiVirus found 'RECYCLER' folder in "
j_bahasa(64) = "This folder is recognized as the potential threat and cntains virus! you are suggested to click YES to remove and immune it"
j_bahasa(65) = "RECYCLER folder in "
j_bahasa(66) = "Failed to clean and imumne it"
j_bahasa(67) = "RECYCLER folder in "
j_bahasa(68) = "Succes to be cleaned and imumned """

 End Sub


Private Sub WriteLangToInterface()
On Error Resume Next
Call LoadDll
With frmMain '===============================================================================================================

'Button
.LAv(0).Caption = a_bahasa(0): .LAv(1).Caption = a_bahasa(1): .LAv(2).Caption = a_bahasa(2): .LAv(3).Caption = a_bahasa(3): .LAv(4).Caption = a_bahasa(4)

.cmdStartScan.Caption = a_bahasa(5)
.BtnResult.Caption = a_bahasa(25)
.cmdStartScan.Caption = a_bahasa(5)
.cmdFixMalware.Caption = a_bahasa(8)
.cmdFixMalwareAll.Caption = a_bahasa(9)
.cmdFixReg.Caption = a_bahasa(8)
.cmdFixRegAll.Caption = a_bahasa(9)
.cmdFixHidden.Caption = a_bahasa(8)
.cmdFixHiddenall.Caption = a_bahasa(9)
.cmdProperties.Caption = a_bahasa(10)
.cmdExplore.Caption = a_bahasa(11)
.cmdRemovePath.Caption = a_bahasa(14)
.cmdRemovePath1.Caption = a_bahasa(22)
.cmdRemExcFile.Caption = a_bahasa(14)
.cmdRemExcFile1.Caption = a_bahasa(22)
.cmdRemExcReg.Caption = a_bahasa(14)
.cmdRemExcReg1.Caption = a_bahasa(22)
.cmdAddExcFolder.Caption = a_bahasa(15)
.cmdAddExcFile.Caption = a_bahasa(16)
.cmdExecutePlug.Caption = a_bahasa(17)
.cmdAddVirus.Caption = a_bahasa(19)
.cmdCancel.Caption = a_bahasa(20)
.CmdRestore.Caption = a_bahasa(23)
.CmdRestoreto.Caption = a_bahasa(24)
.cmdupload.Caption = a_bahasa(31)
.cmdCheckUpdate.Caption = a_bahasa(26)
.BtnFix.Caption = a_bahasa(13)
.CmnKarantine.Caption = a_bahasa(21)
.BTNlOG.Caption = a_bahasa(18)
.CmdQuarAll.Caption = a_bahasa(12)
.cmdRemmoveCk.Caption = a_bahasa(28)
'Label
.frProses.Caption = c_bahasa(2)
.frModule.Caption = c_bahasa(3)
.lbStatus1.Caption = d_bahasa(0)
.lblProcessed.Caption = d_bahasa(1)
.lbTime1.Caption = d_bahasa(2)
.lbFileFound1.Caption = d_bahasa(3)
.lbFileCheck1.Caption = d_bahasa(4)
.lbBypass1.Caption = d_bahasa(5)
.lbMalware1.Caption = d_bahasa(6)
.lbHidden1.Caption = b_bahasa(3) ' ambil orang lain
.lbRegistry1.Caption = d_bahasa(7)
.lblInfor.Caption = d_bahasa(8)
.lbStatus.Caption = d_bahasa(9)
.lbMalware.Caption = "0": .lbHidden.Caption = "0": .lbReg.Caption = "0": .lbInfo.Caption = "0"
.lblExceptFolder.Caption = d_bahasa(21)
.lblExceptFile.Caption = d_bahasa(22)
.lblExceptReg.Caption = d_bahasa(23)
.lblAvalaiblePlug.Caption = d_bahasa(24)
.lblPlugSelect.Caption = d_bahasa(26)
.lblPlugAut.Caption = d_bahasa(27)
.lblPlugAutEmail.Caption = j_bahasa(52)
.lblPlugAutSite.Caption = j_bahasa(53)
.lblPlugVer.Caption = j_bahasa(54)
.lblPlugDesc.Caption = d_bahasa(28)
.lblMalwarePath.Caption = d_bahasa(29)
.lblMalwareName.Caption = d_bahasa(30)
'CheckBox
.ckScan(0).Caption = h_bahasa(13): .ckScan(1).Caption = h_bahasa(13): .ckScan(2).Caption = h_bahasa(13)
End With

With FrmAbout '===============================================================================================================
'label
.Title.Caption = "Wan'iez Antivirus"
.LbInfoAv.Caption = "Information about your Wan'iez security application"

End With

With FrmRTP '=================================================================================================================
'Button
.cmdFixAllRtp.Caption = a_bahasa(13)
.cmdFixRtp2.Caption = a_bahasa(8)
.cmdIgnore.Caption = a_bahasa(27)
.mnQuar.Caption = a_bahasa(12)
'CheckBox
.CkAll.Caption = h_bahasa(13)
'label
'.Title.Caption = ""
End With

With FrmConfig '==============================================================================================================
'CheckBox
.ck1.Caption = h_bahasa(0)
.ck2.Caption = h_bahasa(1)
.ck3.Caption = h_bahasa(2)
.ck4.Caption = h_bahasa(3)
.ck5.Caption = h_bahasa(4)
.ck6.Caption = h_bahasa(5)
.ck7.Caption = h_bahasa(6)
.ck8.Caption = h_bahasa(7)
.ck9.Caption = h_bahasa(8)
.ck10.Caption = h_bahasa(9)
.ck11.Caption = h_bahasa(11)
.ck12.Caption = h_bahasa(10) & " Wan'iez Antivirus"
.Ck14.Caption = h_bahasa(14)
.ck15.Caption = h_bahasa(12)
.CkPass.Caption = h_bahasa(15)
.CkPotection.Caption = h_bahasa(16)
.CkQuar.Caption = h_bahasa(17)
.CkAdditional.Caption = h_bahasa(18)
.CkUpdate.Caption = h_bahasa(19)
.CkDellUs.Caption = h_bahasa(20)
.CkExit.Caption = h_bahasa(21)
'Label
.lblLangUsed.Caption = j_bahasa(36)
.lblLangSel.Caption = d_bahasa(18)
.lblLangID.Caption = d_bahasa(19)
.lblLangAut.Caption = d_bahasa(20)
.lblAvalaibleLang.Caption = d_bahasa(25)
'Button
.CmdSave.Caption = a_bahasa(32)
.CmdEdit.Caption = a_bahasa(33)
.cmdok.Caption = a_bahasa(34)
End With

With FrmSysTray '=============================================================================================================
'Menu Editor
.mnCScan.Caption = g_bahasa(0): .mnEPro2.Caption = g_bahasa(2): .mnEPro.Caption = g_bahasa(3):
.mndisable.Caption = g_bahasa(4): .mndisable2.Caption = g_bahasa(5): .mnRun.Caption = g_bahasa(6):
.mnseting.Caption = g_bahasa(7): .mnAbout.Caption = g_bahasa(8): .mnExit.Caption = g_bahasa(9)

.mnRefresh.Caption = g_bahasa(11): .mnKillPro.Caption = g_bahasa(12): .mnRestartPro.Caption = g_bahasa(13)
.mnPausePro.Caption = g_bahasa(14): .mnResumePro.Caption = g_bahasa(15): .mnProProperties.Caption = g_bahasa(10)

.mncleanjail.Caption = g_bahasa(16): .mnkiljail.Caption = g_bahasa(17): .mnsubmit.Caption = g_bahasa(18)

.mnLock.Caption = g_bahasa(19): .mnunLock.Caption = g_bahasa(20)
End With

With FrmFDmsk
'button
.CmdScanFD(0).Caption = a_bahasa(5): .CmdScanFD(1).Caption = a_bahasa(5)
End With

With FrmPassword
.cmdok.Caption = a_bahasa(35)
End With

End Sub


Private Function ReadExternalLang(sFile As String) As Boolean
Dim IsiFile     As String
Dim SplitterA() As String
Dim SplitterB() As String
Dim SBlok(9)    As String
Dim iCounter    As Long
Dim CutterA     As Long
Dim CutterB     As Long

On Error GoTo LBL_FALSE

If ValidFile(sFile) = False Then GoTo LBL_FALSE
       
IsiFile = ReadUnicodeFile(sFile)
CutterA = InStr(IsiFile, "**/ BEGIN LANG")

' "[Block-0]" : 28
IsiFile = Mid$(IsiFile, CutterA)
CutterA = InStr(IsiFile, "[Block-0]") + 10
If CutterA = 0 Then GoTo LBL_FALSE

SBlok(0) = Mid$(IsiFile, CutterA)
SplitterA = Split(SBlok(0), Chr$(13))

iCounter = 0
For iCounter = 0 To 28
    SplitterB = Split(SplitterA(iCounter), "=")
    a_bahasa(iCounter) = SplitterB(1)
Next
'-----------------------------------

' "[Block-1]" : 15
IsiFile = Mid$(IsiFile, CutterA)
CutterA = InStr(IsiFile, "[Block-1]") + 10
If CutterA = 0 Then GoTo LBL_FALSE

SBlok(1) = Mid$(IsiFile, CutterA)
SplitterA = Split(SBlok(1), Chr$(13))

iCounter = 0
For iCounter = 0 To 15
    SplitterB = Split(SplitterA(iCounter), "=")
    b_bahasa(iCounter) = SplitterB(1)
Next
'-----------------------------------
   
' "[Block-2]" : 09
IsiFile = Mid$(IsiFile, CutterA)
CutterA = InStr(IsiFile, "[Block-2]") + 10
If CutterA = 0 Then GoTo LBL_FALSE

SBlok(2) = Mid$(IsiFile, CutterA)
SplitterA = Split(SBlok(2), Chr$(13))

iCounter = 0
For iCounter = 0 To 9
    SplitterB = Split(SplitterA(iCounter), "=")
    c_bahasa(iCounter) = SplitterB(1)
Next
'-----------------------------------
   
' "[Block-3]" : 38
IsiFile = Mid$(IsiFile, CutterA)
CutterA = InStr(IsiFile, "[Block-3]") + 10
If CutterA = 0 Then GoTo LBL_FALSE

SBlok(3) = Mid$(IsiFile, CutterA)
SplitterA = Split(SBlok(3), Chr$(13))

iCounter = 0
For iCounter = 0 To 38
    SplitterB = Split(SplitterA(iCounter), "=")
    d_bahasa(iCounter) = SplitterB(1)
Next
'-----------------------------------
   
' "[Block-4]" : 17
IsiFile = Mid$(IsiFile, CutterA)
CutterA = InStr(IsiFile, "[Block-4]") + 10
If CutterA = 0 Then GoTo LBL_FALSE

SBlok(4) = Mid$(IsiFile, CutterA)
SplitterA = Split(SBlok(4), Chr$(13))

iCounter = 0
For iCounter = 0 To 17
    SplitterB = Split(SplitterA(iCounter), "=")
    e_bahasa(iCounter) = SplitterB(1)
Next
'-----------------------------------
   
' "[Block-5]" : 21
IsiFile = Mid$(IsiFile, CutterA)
CutterA = InStr(IsiFile, "[Block-5]") + 10
If CutterA = 0 Then GoTo LBL_FALSE

SBlok(5) = Mid$(IsiFile, CutterA)
SplitterA = Split(SBlok(5), Chr$(13))

iCounter = 0
For iCounter = 0 To 21
    SplitterB = Split(SplitterA(iCounter), "=")
    f_bahasa(iCounter) = SplitterB(1)
Next
'-----------------------------------

' "[Block-6]" : 15
IsiFile = Mid$(IsiFile, CutterA)
CutterA = InStr(IsiFile, "[Block-6]") + 10
If CutterA = 0 Then GoTo LBL_FALSE

SBlok(6) = Mid$(IsiFile, CutterA)
SplitterA = Split(SBlok(6), Chr$(13))

iCounter = 0
For iCounter = 0 To 15
    SplitterB = Split(SplitterA(iCounter), "=")
    g_bahasa(iCounter) = SplitterB(1)
Next
'-----------------------------------

' "[Block-7]" : 10
IsiFile = Mid$(IsiFile, CutterA)
CutterA = InStr(IsiFile, "[Block-7]") + 10
If CutterA = 0 Then GoTo LBL_FALSE

SBlok(7) = Mid$(IsiFile, CutterA)
SplitterA = Split(SBlok(7), Chr$(13))

iCounter = 0
For iCounter = 0 To 10
    SplitterB = Split(SplitterA(iCounter), "=")
    h_bahasa(iCounter) = SplitterB(1)
Next
'-----------------------------------

' "[Block-8]" : 27
IsiFile = Mid$(IsiFile, CutterA)
CutterA = InStr(IsiFile, "[Block-8]") + 10
If CutterA = 0 Then GoTo LBL_FALSE

SBlok(8) = Mid$(IsiFile, CutterA)
SplitterA = Split(SBlok(8), Chr$(13))

iCounter = 0
For iCounter = 0 To 27
    SplitterB = Split(SplitterA(iCounter), "=")
    i_bahasa(iCounter) = SplitterB(1)
Next
'-----------------------------------

' "[Block-9]" : 62
IsiFile = Mid$(IsiFile, CutterA)
CutterA = InStr(IsiFile, "[Block-9]") + 10
If CutterA = 0 Then GoTo LBL_FALSE

SBlok(9) = Mid$(IsiFile, CutterA)
SplitterA = Split(SBlok(9), Chr$(13))

iCounter = 0
For iCounter = 0 To 62
    SplitterB = Split(SplitterA(iCounter), "=")
    j_bahasa(iCounter) = SplitterB(1)
Next
'-----------------------------------

ReadExternalLang = True

Exit Function

LBL_FALSE:
ReadExternalLang = False
End Function



Public Function EnumLangAvalaible(sPath As String, LstOut As ListBox) As Boolean
Dim JumLngFile   As Long
Dim iCounter     As Long
Dim StrLngFile() As String
Dim ArrHeadL     As String

JumLngFile = GetFile(sPath, StrLngFile)

LstOut.Clear

LstOut.AddItem "=> ENGLISH | Default@1"
'LstOut.AddItem "=> INDONESIA | Default@2"

For iCounter = 0 To JumLngFile - 1
    ArrHeadL = ReadHeaderLang(StrLngFile(iCounter))
    If Len(ArrHeadL) > 0 Then LstOut.AddItem "=> " & GetLangName(ArrHeadL) & " | " & GetFileName(StrLngFile(iCounter))
Next
End Function

Private Function ReadHeaderLang(sFileLang As String) As String
Dim IsiFile     As String
Dim SplitterA() As String
Dim SplitterB() As String

Dim CutterA As Long
If ValidFile(sFileLang) = False Then Exit Function


IsiFile = ReadUnicodeFile(sFileLang)
CutterA = InStr(IsiFile, "[WANIEZ LANG]")

If CutterA = 0 Then Exit Function
CutterA = InStr(IsiFile, "**/ INIT HEADER INFO") + 21
IsiFile = Mid$(IsiFile, CutterA)
SplitterA = Split(IsiFile, Chr$(13))

SplitterB = Split(SplitterA(0), "=") 'ID
ReadHeaderLang = ReadHeaderLang & SplitterB(1) & "\"

SplitterB = Split(SplitterA(1), "=") 'NAME
ReadHeaderLang = ReadHeaderLang & SplitterB(1) & "\"

SplitterB = Split(SplitterA(2), "=") 'AUT
ReadHeaderLang = ReadHeaderLang & SplitterB(1) & "\"


End Function

Public Sub WriteLngInfoToLabel(strSelected As String, LBID As Label, LNAME As Label, LAUT As Label)
Dim sTmp        As String
Dim ArrLHead    As String
Dim sPath       As String
Dim SplitterA() As String
Dim sNameTmp    As String

' Karena letak filenya namabahasa | namafile
sTmp = Mid$(strSelected, InStr(strSelected, "| ") + 2)

Select Case sTmp
    Case "Default@1": sNameTmp = "Default@1": GoTo LBL_KECUALI
   ' Case "Default@2": sNameTmp = "Default@2": GoTo LBL_KECUALI
End Select

sPath = GetFilePath(App_FullPathW(False)) & "\lang"
sPath = sPath & "\" & sTmp

If ValidFile(sPath) = False Then GoTo LBL_FALSE

ArrLHead = ReadHeaderLang(sPath)

SplitterA = Split(ArrLHead, "\")

LBID.Caption = ": " & SplitterA(0)
LNAME.Caption = ": " & SplitterA(1)
LAUT.Caption = ": " & SplitterA(2)

Exit Sub

LBL_FALSE:
    LBID.Caption = ": -"
    LNAME.Caption = ": -"
    LAUT.Caption = ": -"

Exit Sub
LBL_MINANG:
    LBID.Caption = ": Built-in"
    LNAME.Caption = ": " & sNameTmp
    LAUT.Caption = ": Canvas Team"
    Exit Sub
LBL_ANANG:
    LBID.Caption = ": Built-in"
    LNAME.Caption = ": " & sNameTmp
    LAUT.Caption = ": Canvas Team"
    Exit Sub
LBL_KECUALI:
    LBID.Caption = ": Built-in"
    LNAME.Caption = ": " & sNameTmp
    LAUT.Caption = ": Canvas Team"
End Sub

Private Function GetLangName(LangHead As String) As String
Dim SplitA()  As String

SplitA = Split(LangHead, "\")

GetLangName = SplitA(1)
End Function



'---- Untuk dipakai saat load config
Public Function getNameLangFromFile(sFileLangToRead As String) As String
On Error GoTo LBL_DEF
    getNameLangFromFile = GetLangName(ReadHeaderLang(sFileLangToRead))
     
Exit Function

LBL_DEF:
 getNameLangFromFile = "Default@n"
End Function

