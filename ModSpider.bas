Attribute VB_Name = "ModSpider"
'============================================='
' Phenix Spider                               '
'                                             '
' Author: Dominic 'Phenix' Black              '
' Email: phenix@sg15.com                      '
'                                             '
' Feel free to use the code in you appication '
' as long as you give credit to me.           '
'============================================='

Global SpiderOnline As Boolean
Public Sub Toggle()
    If SpiderOnline = False Then
        MainLoop
    Else
        SpiderOnline = False
        MsgBox "Phenix Spider will finishing scanning the curernt page," + vbNewLine + "and then will stop.", vbInformation
    End If
End Sub


Public Sub MainLoop()
    Dim page As String
    Dim pageURL As String
    Dim FoundPage As Boolean
    Dim i As Long
    Dim i2 As Integer
    Dim endOfHREF As Boolean
    Dim phrasedURL As String
    Dim PageRow As Long
    
    On Error Resume Next
    
    FrmMain.CmdToggle.Caption = "Pause"
    FrmStatistics.CmdToggle.Caption = "Pause"
    
    If FrmMain.CmdOptions.Visible = True Then
        FrmMain.CmdOptions.Visible = False
        FrmMain.txtMemory.Enabled = False
        FrmMain.TxtUrl.Enabled = False
        LogURL FrmMain.TxtUrl.Text
    End If
        
    SpiderOnline = True
    
    While SpiderOnline = True
        DoEvents
        
        FoundPage = False
        
        For i = 1 To FrmStatistics.LstURL.ListItems.Count
            If FrmStatistics.LstURL.ListItems.item(i).SubItems(1) = "No" Then
                pageURL = FrmStatistics.LstURL.ListItems.item(i).Text
                PageRow = i
                FoundPage = True
                Exit For
            End If
        Next
                
        If FoundPage = True Then
            Debug.Print "Scanning " & pageURL
            page = FrmMain.Inet1.OpenURL(pageURL)
            
            FrmStatistics.ProgressBar1.Value = 0
            FrmStatistics.ProgressBar1.Max = Len(page)
            FrmStatistics.lblPage = pageURL
            FrmMain.Caption = "Phenix Spider :: " & pageURL
            
            For i = 1 To Len(page)
                
                DoEvents
                
                ' " link
                If Mid(page, i, 6) = "href=" & Chr(34) Then
                    endOfHREF = False
                    
                    i2 = 1
                    
                    While endOfHREF = False
                        If Mid(page, i + i2 + 7, 1) = Chr(34) Then
                            endOfHREF = True
                            phrasedURL = Mid(page, i + 6, i2 + 1) ' - (i + 6)
                            PhraseHREF phrasedURL, pageURL
                        End If
                        i2 = i2 + 1
                    Wend
                End If
                
                DoEvents
                
                ' ' link
                If Mid(page, i, 6) = "href=" & Chr(39) Then
                    endOfHREF = False
                    
                    i2 = 1
                    
                    While endOfHREF = False
                        If Mid(page, i + i2 + 7, 1) = Chr(39) Then
                            endOfHREF = True
                            phrasedURL = Mid(page, i + 6, i2 + 1) ' - (i + 6)
                            PhraseHREF phrasedURL, pageURL
                        End If
                        i2 = i2 + 1
                    Wend
                End If
                
                If FrmOptions.OptFind = True And Mid(page, i, Len(FrmOptions.txtSearch)) = FrmOptions.txtSearch Then
                    LogFind pageURL
                End If
                
                DoEvents
                
                FrmStatistics.ProgressBar1.Value = i
            Next
            
            FrmStatistics.LstURL.ListItems.item(PageRow).SubItems(1) = "Yes"
            FrmStatistics.lblScanned = FrmStatistics.lblScanned + 1
            FrmMain.lblScanned = FrmStatistics.lblScanned
        Else
            Debug.Print "No More Addresses Left"
            SpiderOnline = False
            FrmMain.CmdToggle.Visible = False
            FrmStatistics.CmdToggle.Visible = False
            MsgBox "Ran out of URL's to Crawl", vbCritical
        End If
    Wend
        
    FrmMain.CmdToggle.Caption = "Resume"
    FrmStatistics.CmdToggle.Caption = "Resume"
End Sub

Public Sub PhraseHREF(href As String, baseURL As String)
    Dim DontLog As Boolean
    Dim NewBase As String
    Dim i As Integer
    
    DontLog = False
    
        If InStr(1, href, ".exe", vbTextCompare) <> 0 Then DontLog = True
        If InStr(1, href, ".zip", vbTextCompare) <> 0 Then DontLog = True
        If InStr(1, href, ".rar", vbTextCompare) <> 0 Then DontLog = True
        If InStr(1, href, ".ace", vbTextCompare) <> 0 Then DontLog = True
        If InStr(1, href, ".jpg", vbTextCompare) <> 0 Then DontLog = True
        If InStr(1, href, ".gif", vbTextCompare) <> 0 Then DontLog = True
        If InStr(1, href, ".jpeg", vbTextCompare) <> 0 Then DontLog = True
        If InStr(1, href, ".png", vbTextCompare) <> 0 Then DontLog = True
        If InStr(1, href, ".bmp", vbTextCompare) <> 0 Then DontLog = True
        If InStr(1, href, ".wmv", vbTextCompare) <> 0 Then DontLog = True
        If InStr(1, href, ".mp3", vbTextCompare) <> 0 Then DontLog = True
        If InStr(1, href, ".wav", vbTextCompare) <> 0 Then DontLog = True
        If InStr(1, href, ".mov", vbTextCompare) <> 0 Then DontLog = True
        If InStr(1, href, "javascript", vbTextCompare) <> 0 Then DontLog = True
        If InStr(1, href, ".swf", vbTextCompare) <> 0 Then DontLog = True
        If InStr(1, href, "#", vbTextCompare) <> 0 Then DontLog = True
    
    DoEvents
    
    If DontLog = False Then
        If Mid(href, 1, 7) = "mailto:" Then
            ' Log Email
            If FrmOptions.OptEmail.Value = True Then LogEmail Mid(href, 8, Len(href) - 7)
        ElseIf Mid(href, 1, 7) = "http://" Then
            ' Log URL
            If Right(href, 1) = "/" Then href = Mid(href, 1, Len(href) - 1)
            LogURL href
        ElseIf Mid(href, 1, 6) = "ftp://" Then
            'FTP Not supported
        ElseIf Mid(href, 1, 8) = "https://" Then
            ' HTTPS not supported
        ElseIf Mid(href, 1, 6) = "mms://" Then
            ' MMS not supported
        ElseIf Mid(href, 1, 6) = "irc://" Then
            ' IRC not supported
        Else
            ' Must be a URL which isnt 'Real'
            DoEvents
            
            If Mid(href, 1, 1) = "/" Or Mid(href, 1, 1) = "\" Then
                baseURL = Replace(baseURL, "//", "!*!")
                For i = 1 To Len(baseURL)
                    If Mid(baseURL, i, 1) <> "/" Then
                        NewBase = NewBase + Mid(baseURL, i, 1)
                    Else
                        Exit For
                    End If
                Next
                NewBase = Replace(NewBase, "!*!", "//")
                LogURL (NewBase & href)
            Else
                For i = 1 To Len(baseURL)
                    If Mid(baseURL, Len(baseURL) - i + 1, 1) = "/" Then
                        NewBase = Mid(baseURL, 1, Len(baseURL) - i + 1)
                        Exit For
                    End If
                Next
                
                NewBase = Replace(NewBase, "!*!", "//")
                
                If NewBase = "" Then NewBase = baseURL + "/"
                LogURL NewBase & href
                DoEvents
            End If
        End If
    End If
End Sub


Public Sub LogURL(URL As String)
    Dim item As ListItem
    Dim i As Integer
    Dim NotSearched As Long
    Dim foundURL As Boolean
    
    foundURL = False
    
    NotSearched = 0
    
    URL = Replace(URL, "!*!", "//")
    
    For i = 1 To FrmStatistics.LstURL.ListItems.Count
        If FrmStatistics.LstURL.ListItems.item(i).Text = URL Then
            foundURL = True
            Exit For
        End If
        
        If FrmStatistics.LstURL.ListItems.item(i).SubItems(1) = "No" Then NotSearched = NotSearched + 1
    Next
    
    If foundURL = False And (NotSearched <= FrmMain.txtMemory Or FrmMain.txtMemory = -1) Then
        Debug.Print "  -> Found URL: " & URL
        Set item = FrmStatistics.LstURL.ListItems.Add(, , URL)
        item.SubItems(1) = "No"
        FrmStatistics.lblUrl = FrmStatistics.lblUrl + 1
        If FrmOptions.ChkSave.Value = 1 Then
                Print #3, URL
        End If
    End If
End Sub

Public Sub LogEmail(Address As String)
    Dim item As ListItem
    Dim i As Integer
    Dim foundEmail As Boolean
    
    foundEmail = False
    
    For i = 1 To FrmStatistics.LstEmail.ListItems.Count
        If FrmStatistics.LstEmail.ListItems.item(i).Text = Address Then
            foundEmail = True
            Exit For
        End If
    Next
    
    If foundEmail = False Then
        Debug.Print "  -> Found Email: " & Address
        Set item = FrmStatistics.LstEmail.ListItems.Add(, , Address)
        FrmStatistics.lblEmail = FrmStatistics.lblEmail + 1
        If FrmOptions.ChkSave.Value = 1 Then
                Print #2, Address
        End If
    End If
End Sub

Public Sub LogFind(URL As String)
    Dim item As ListItem
    Dim i As Integer
    Dim foundURL As Boolean
    
    foundURL = False
    
    For i = 1 To FrmStatistics.LstEmail.ListItems.Count
        If FrmStatistics.LstEmail.ListItems.item(i).Text = URL Then
            foundURL = True
            Exit For
        End If
    Next
    
    If foundURL = False Then
        Debug.Print "  -> Found Something: " & URL
        Set item = FrmStatistics.LstEmail.ListItems.Add(, , URL)
        FrmStatistics.lblEmail = FrmStatistics.lblEmail + 1
        If FrmOptions.ChkSave.Value = 1 Then
                Print #1, URL
        End If
    End If
End Sub
