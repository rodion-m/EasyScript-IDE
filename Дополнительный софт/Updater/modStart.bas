Attribute VB_Name = "modStart"
Option Explicit

Sub Main()
    On Error Resume Next
    If Dir$(App.Path & "\Update.ini") <> "" Then
        GlobalName = ReadIni("Name")
        GlobalVersion = CDbl("0," & Replace(ReadIni("Version", 0), ".", ""))
        InstallerName = ReadIni("InstallerName")
        FULLUPDATE_FILENAME_URL = DOWNLOADS_PATH_URL & InstallerName
        frmMain.Caption = "RodySoft " & GlobalName & " Updater"
    End If
    If Left$(Command$, 8) = "/success" Then ' После обновы сообщаю о том, что все хорошо
        Dim msgChoose As VbMsgBoxResult
        If Len(Command$) > 9 Then
            Dim sBaseName As String
            sBaseName = Right$(Command$, Len(Command$) - 9)
            If IsFileExists(App.Path & "\" & sBaseName) Then
                WriteIni "Validated", "True" ' После обновления снова записываю валидацию, ибо затирается
                Shell App.Path & "\" & sBaseName & " /restore", vbNormalFocus
            End If
        End If
        SetWindowPos frmMain.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE
        msgChoose = ApiMsg("Обновление программы " & GlobalName & " успешно завершено!" & vbCrLf & "Программа распространяется совершенно бесплатно. Вы можете поддержать проект, написав о нем в своем блоге или на своей страничке в социальной сети." & vbCrLf & "Желаете ли вы поддержать проект?", "RodySoft Updater", vbInformation + vbYesNo)
        If msgChoose = vbYes Then ShellExecute 0, "Open", SERVER_URL & "donate/", "", "", vbNormalFocus
        End
    ElseIf Left$(Command$, 10) = "/uninstall" Then
        SendInfoToServer "uninstall", , True
        End
    End If
    If Command$ <> "" Then
        ExecuteCommandLine
    End If
    If Not CommandLine.HideMode And Not CommandLine.bSendInfo Then
        frmMain.Show
        SetWindowPos frmMain.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE
    End If
End Sub

Sub ExecuteCommandLine()
    Const hm = "hm", au = "au", bn = "bn", pn = "pn", hw = "hw", si = "si" ' SendInfo
    Dim Commands() As String, strExParam As String, i As Long
    Commands() = Split(Command$, "/")
    With CommandLine
        For i = 1 To UBound(Commands)
            strExParam = ""
            If Len(Commands(i)) > 3 Then
                strExParam = Right$(Commands(i), Len(Commands(i)) - 3)
                If Right$(strExParam, 1) = " " Then strExParam = Left$(strExParam, Len(strExParam) - 1)
                If Right$(strExParam, 1) = Chr$(34) And Left$(strExParam, 1) = Chr$(34) Then strExParam = Mid$(strExParam, 2, Len(strExParam) - 2)
            End If
            Select Case LCase(Left$(Commands(i), 2))
                Case hm: .HideMode = True: .AutoStart = True
                Case hw: .hwnd = strExParam
                Case pn: .ProductName = strExParam
                Case bn: .BaseName = strExParam & ".exe"
                Case au: .AutoStart = True
                Case si: .bSendInfo = True
            End Select
        Next i
        If .bSendInfo Then frmBugReport.Show
        If .AutoStart Then frmMain.StartUpdate: End
    End With
End Sub
