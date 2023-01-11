Attribute VB_Name = "BasALOG"
Public Function awal()
       LogPrint "==============================Result Scan Wan'iez Antivirus========================="
       LogPrint "------------------------------------------------------------------------------------"
       LogPrint "==>  " + "Day & Date      : " + Format(Now, "dddd, dd mmmm yyyy")
       LogPrint "==>  " + "Start Scan  : " + Format(Now, "hh:mm:ss")
       'LogPrint "==>  " + "Waktu Mulai Scan  : " + Format(Now, "hh:mm:ss")
       LogPrint "==>  " + "Reg Database      " + FrmAbout.lbRegDataBase
       LogPrint "==>  " + "Signature database: " & CStr(JumVirus) & " Worm, " & JumlahVirusINT & " Virus, " & JumlahVirusOUT & " Malware User + Heuristic"
       LogPrint "------------------------------------------------------------------------------------"
       LogPrint "=====================================Start Scan=============================="
       LogPrint "                                                                              "
End Function
Public Function akir()
       LogPrint "                                                                              "
       LogPrint "===================================Scan Finish=============================="
       LogPrint "------------------------------------------------------------------------------------"
       LogPrint "==>  " + d_bahasa(17) + ": " + Format(Now, "hh:mm:ss")
       LogPrint "==>  " + j_bahasa(43) + ": " & FileFound
       LogPrint "==>  " + j_bahasa(44) + ": " & FileCheck
       LogPrint "==>  " + j_bahasa(45) + ": " & FileNotCheck
       LogPrint "==>  " + j_bahasa(46) + ": " & frmMain.lvMalware.ListItems.Count
       LogPrint "==>  " + b_bahasa(3) + ": " & frmMain.lvHidden.ListItems.Count
       LogPrint "==>  " + j_bahasa(48) + ": " & nErrorReg
       LogPrint "==>  " + b_bahasa(4) + ": " & frmMain.lvInfo.ListItems.Count
       LogPrint "==>  " + frmMain.lbStatus22
      ' LogPrint "==>  " + FrmAbout.Lblinfo
       LogPrint "-------------------------------------------------------------------------------------"
       LogPrint "========================================Wan'iez======================================"
       LogPrint "                                                                               "
       
End Function
Public Sub LogPrint(sMessage As String)
'проца записи в файл исходного текста макровируса
Dim nFile As Integer
Dim ffile As String
nFile = FreeFile
'Open ffile For Append As #nFile
ffile = App.path & "\LOG.scan"
Open ffile For Append Access Write Lock Read Write As #nFile
Print #nFile, sMessage
Close #nFile
End Sub

