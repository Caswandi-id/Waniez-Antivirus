Attribute VB_Name = "basPluginAkses"
Dim sPluginPath()     As String
Dim sPluginName()     As String
Dim sPluginAuthor()   As String
Dim sPluginAutEmail() As String
Dim sPluginAutSite()  As String
Dim sPluginValID()    As String
Dim sPluginDesc()     As String



Public Sub EnumPlugin(szFolderPlugin As String, LstOut As ListBox)
Dim nPluginCount            As Long
Dim ArPlugsList()           As PG_PLUGIN_GENERAL_INFORMATION
Dim CDCounter               As Long
Dim HeadS                   As String


    
    HeadS = ": "
    nPluginCount = PgEnumeratePluginFiles(szFolderPlugin, ArPlugsList())
    
    
    LstOut.Clear
    
    If nPluginCount = 0 Then
        LstOut.AddItem j_bahasa(55) & "www.waniez.p.ht"
        frmMain.cmdExecutePlug.Enabled = False
        'frmMain.LbPlugin.Caption = "Avalaible Plugin - " & lstPlugin.ListCount
        Exit Sub
    ElseIf nPluginCount < 0 Then
        Exit Sub
    Else
        ReDim sPluginPath(nPluginCount - 1) As String
        ReDim sPluginName(nPluginCount - 1) As String
        ReDim sPluginAuthor(nPluginCount - 1) As String
        ReDim sPluginAutEmail(nPluginCount - 1) As String
        ReDim sPluginAutSite(nPluginCount - 1) As String
        ReDim sPluginValID(nPluginCount - 1) As String
        ReDim sPluginDesc(nPluginCount - 1) As String
        
        With frmMain
             .lblPlugSelect1.Caption = HeadS
             .lblPlugAut1.Caption = HeadS
             .lblPlugAutEmail1.Caption = HeadS
             .lblPlugAutSite1.Caption = HeadS
             .lblPlugVer1.Caption = HeadS
             .lblPlugDesc1.Caption = HeadS
        End With

        For CDCounter = 0 To (nPluginCount - 1)
            sPluginPath(CDCounter) = ArPlugsList(CDCounter).szPluginStartupPathW
            sPluginName(CDCounter) = ArPlugsList(CDCounter).szPluginName
            sPluginAuthor(CDCounter) = ArPlugsList(CDCounter).szPluginAuthor
            sPluginAutEmail(CDCounter) = ArPlugsList(CDCounter).szPluginAuthorEMail
            sPluginAutSite(CDCounter) = ArPlugsList(CDCounter).szPluginAuthorSite
            sPluginValID(CDCounter) = ArPlugsList(CDCounter).szPluginValidationID
            sPluginDesc(CDCounter) = ArPlugsList(CDCounter).szPluginDescription

            LstOut.AddItem "-> " & ArPlugsList(CDCounter).szPluginStartupPathW
        Next
        frmMain.cmdExecutePlug.Enabled = True
    End If
    
    Erase ArPlugsList
End Sub

Public Sub RetrievePlugInfo(PlugIndek As Long, LstRead As ListBox, lblPlugName As Label, lblAut As Label, lblAutEmail As Label, lblAutSite As Label, lblVerCode As Label, lblDesc As Label)
Dim HeadS                   As String
Dim hIcon                   As Long

    
    HeadS = ": "
    If PlugIndek >= 0 And ValidFile(sPluginPath(PlugIndek)) = True Then
       lblPlugName.Caption = HeadS & sPluginName(PlugIndek)
       lblAut.Caption = HeadS & sPluginAuthor(PlugIndek)
       lblAutEmail.Caption = HeadS & sPluginAutEmail(PlugIndek)
       lblAutSite.Caption = HeadS & sPluginAutSite(PlugIndek)
       lblVerCode.Caption = HeadS & sPluginValID(PlugIndek)
       lblDesc.Caption = HeadS & sPluginDesc(PlugIndek)
       
       ' gambar Iconya
      ' CopiFile sPluginPath(PlugIndek), "C:\$$$$$.exe", False
      ' RetrieveIcon "C:\$$$$$.exe", frmMain.picPlugin, ricnLarge
      ' HapusFile "C:\$$$$$.exe"
    End If
    
End Sub


Public Sub RunPlugin(PlugIndek As Long)
Dim RetRun                  As Long
Dim szPluginFileNameW       As String

            If PlugIndek >= 0 Then
               szPluginFileNameW = sPluginPath(PlugIndek)
               If ValidFile(szPluginFileNameW) = False Then Exit Sub
               If MsgBox(j_bahasa(56) & " (" & sPluginName(PlugIndek) & ") ?", vbInformation + vbYesNo) = vbYes Then
                  If MsgBox(j_bahasa(57) & " ?", vbExclamation + vbYesNo) = vbYes Then
                     RetRun = PgLoadAndRunPlugin(szPluginFileNameW, True)
                     If RetRun = 0 Then
                        MsgBox j_bahasa(58) & " !", vbExclamation
                     End If
                  Else
                     RetRun = PgLoadAndRunPlugin(szPluginFileNameW, True)
                     If RetRun = 0 Then
                        MsgBox j_bahasa(58) & " !", vbExclamation
                     End If
                  End If
               End If
            End If

End Sub

