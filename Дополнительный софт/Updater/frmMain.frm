VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form frmMain 
   BackColor       =   &H00F1E4D9&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "RodySoft Updater"
   ClientHeight    =   1545
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6585
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   12
      Charset         =   204
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1545
   ScaleWidth      =   6585
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin SHDocVwCtl.WebBrowser Browser 
      Height          =   855
      Left            =   5160
      TabIndex        =   6
      Top             =   1320
      Visible         =   0   'False
      Width           =   2655
      ExtentX         =   4683
      ExtentY         =   1508
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin ComctlLib.ProgressBar ProgressBar 
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   873
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Закрыть"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4440
      TabIndex        =   4
      Top             =   840
      Width           =   2000
   End
   Begin ComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   1290
      Width           =   6585
      _ExtentX        =   11615
      _ExtentY        =   450
      Style           =   1
      SimpleText      =   "Программа запущена..."
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "Начать обновление "
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   2000
   End
   Begin VB.CommandButton cmdStop 
      Caption         =   "Остановить"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2280
      TabIndex        =   1
      Top             =   840
      Width           =   2000
   End
   Begin VB.Frame Frame 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   0
      TabIndex        =   5
      Top             =   720
      Width           =   6855
      Begin VB.Line Line 
         BorderColor     =   &H00808080&
         X1              =   0
         X2              =   6840
         Y1              =   0
         Y2              =   0
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private blnStopDownload As Boolean, BAT_Body As String, strUpdatedProducts As String, Programms As FileInfo, sNewVersionsContent As String

Dim WithEvents Downloader As clsRSWinHTTP
Attribute Downloader.VB_VarHelpID = -1

Sub Downloader_OnResponseStart(ByVal Status As Long, ByVal ContentType As String)
    ProgressBar.Value = 5
End Sub

Sub Downloader_Progress(ByVal Progress As Long, ByVal ProgressMax As Long)
    If ProgressMax < Progress Then ProgressMax = Progress
    If ProgressMax = 0 Then ProgressMax = 1
    ProgressBar.Value = (Progress / ProgressMax) * 100
    DoEvents
End Sub

Sub Downloader_Complete(ByVal Status As Long, ByVal StatusText As String)
    ProgressBar.Value = 100
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdStart_Click()
    Dim clean As COMMAND_PROMT_ARGUMENTS
    CommandLine = clean
    StartUpdate
End Sub

Private Sub cmdStop_Click()
    cmdStop.Enabled = False
    blnStopDownload = True
    Downloader.Abort
End Sub

Private Sub Form_Initialize()
    'If App.PrevInstance Then End
    InitCommonControls
End Sub

Private Sub Form_Load()
    Set Downloader = New clsRSWinHTTP
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Downloader.Abort
    Set Downloader = Nothing
    End
End Sub

Sub SetStatus(ByVal strStatus As String)
    StatusBar.SimpleText = strStatus
End Sub

Function GetNewVersionText() As String
    GetNewVersionText = GetValue("UpdateText", "", sNewVersionsContent)
    GetNewVersionText = Replace(GetNewVersionText, "[crlf]", vbCrLf)
    If Len(GetNewVersionText) > 0 Then GetNewVersionText = vbCrLf & GetNewVersionText
End Function

Sub CheckUpdates()
    Dim FilesInfo() As FileInfo, Files() As String, n As Long
    GetFilesList App.Path, Files, ".exe"
    ReDim FilesInfo(UBound(Files))
    For n = 0 To UBound(Files)
        GetFileInfo Files(n), FilesInfo(n)
        If FilesInfo(n).BaseName <> "Uninstall.exe" And FilesInfo(n).ProductName <> "" And FilesInfo(n).ProductVersion <> "" Then CheckNewVersion FilesInfo(n)
        If blnStopDownload Then Exit Sub
    Next
End Sub

Sub StartUpdate(Optional ByVal ProductName As String)
    Me.MousePointer = vbHourglass
    cmdStop.Enabled = True
    cmdStart.Enabled = False
    blnStopDownload = False
    strUpdatedProducts = ""
    ProgressBar.Value = 5!
    DoEvents
    SetStatus "Проверка подключения к интернету..."
    If ReadIni("Validated", "", , , "Update.ini") = "" Then
        ' первый старт
		
    End If
    If blnStopDownload Then GoTo EndSub
    If Not Downloader.DownloadToString(GETINFO_URL & "versions&app=" & CommandLine.BaseName, sNewVersionsContent, False, "GET") Then
        CheckHide
        SetStatus "Ошибка: Не удалось подключиться к серверу..."
        ApiMsg "Не удается подключиться к серверу. Пожалуйста, проверьте подключение к интернету.", "Ошибка подключения", vbExclamation
        GoTo EndSub
    End If
    If blnStopDownload Then GoTo EndSub
    SetStatus "Поиск новой версии..."
    If blnStopDownload Then GoTo EndSub
    If GlobalVersion > 0 And GlobalName <> "" Then
        If GlobalVersion < GetNewVersion(GlobalName) Then
            ' Если принципиально новая версия
            HideModeMessage
            Call FullUpdate
            GoTo EndSub
        End If
    End If
    ' Сперва проверяюется обновление на модуль, который открыл апдейтер
    If CommandLine.BaseName <> "" Then
        Dim Parrent As FileInfo
        GetFileInfo App.Path & "\" & CommandLine.BaseName, Parrent
        If Len(Parrent.ProductName) > 0 Then
            If Not CheckNewVersion(Parrent) Then GoTo ItsLastVersion
        End If
    End If
    If blnStopDownload Then GoTo EndSub
    CheckUpdates
    If blnStopDownload Then GoTo EndSub
    If BAT_Body <> "" Then
        ' Если хоть один продукт имеет новую версию
        AddLineToString BAT_Body, "start " & App.EXEName & ".exe /success " & CommandLine.BaseName
        ShowDownloadCompleteMessage
        RunBat BAT_Body
    Else
ItsLastVersion:
        ' Установлена последняя версия
        CheckHide
        ApiMsg "У вас установлена последняя версия программы", "Обновление не требуется", , CommandLine.hwnd
        Unload Me
    End If
EndSub:
    Me.MousePointer = vbNormal
    cmdStart.Enabled = True
    cmdStop.Enabled = False
    ProgressBar.Value = 0
    CheckHide
End Sub

Private Function CheckNewVersion(ProgrammInfo As FileInfo) As Boolean
    With ProgrammInfo
        If InStr(1, strUpdatedProducts, .ProductName) > 0 Then Exit Function ' Проверка не был ли этот продукт обновлен
        SetStatus "Обновляемый модуль: " & .ProductName & "..."
        If IsNewVersionAvailable(ProgrammInfo) Then
            HideModeMessage
            If IsUserAnAdmin = 0 Then ShellExecute 0, "runas", App.Path & "\" & App.EXEName & ".exe", "/au", App.Path, vbNormalFocus: End
            If Not Downloader.DownloadToFile(DOWNLOADS_PATH_URL & .BaseName, App.Path & "\" & .BaseName & ".update") Then ShowError: blnStopDownload = True
            If blnStopDownload Then Exit Function
            AddUpdate .BaseName
            strUpdatedProducts = strUpdatedProducts & .ProductName
            CheckNewVersion = True
        End If
    End With
End Function

Private Function GetVersion(ProgrammInfo As FileInfo) As Double
    Dim Ver As String
    Ver = Replace(ProgrammInfo.ProductVersion, ".", "")
    If (IsNumeric(Ver)) Then
        GetVersion = CDbl("0," & Ver)
    Else
        GetVersion = 0
    End If
End Function

Function GetNewVersion(ByVal ProductName As String) As Double
    Dim NewVer As String
    NewVer = GetValue(ProductName, 0, sNewVersionsContent)
    If NewVer = "" Then NewVer = "0" 'Этого софта нет в списке для обновлений
    GetNewVersion = CDbl("0," & Replace(NewVer, ".", ""))
End Function

Private Function IsNewVersionAvailable(ProgrammInfo As FileInfo) As Boolean
    If GetVersion(ProgrammInfo) < GetNewVersion(ProgrammInfo.ProductName) Then
        IsNewVersionAvailable = True
    End If
End Function

Sub AddUpdate(ByVal BaseName As String)
    If BAT_Body = "" Then BAT_Body = "cd " & App.Path
    AddLineToString BAT_Body, "taskkill /f /im " & BaseName
    AddLineToString BAT_Body, "ping -n 3 localhost" 'пауза 2 секунды
    AddLineToString BAT_Body, "del /q " & BaseName
    AddLineToString BAT_Body, "ren " & BaseName & ".update" & " " & BaseName
End Sub

Sub RunBat(ByVal strBody As String)
    AddLineToString strBody, "del %0"
    CreateFile App.Path & "\update.bat", strBody
    Shell App.Path & "\update.bat", vbHide
    End
End Sub

'Function UpdateDownloadsCouter(ByVal BaseName As String) As Boolean
'    Downloader.DownloadToString DOWNLOADS_COUNTER_URL & BaseName, , False
'    UpdateDownloadsCouter = (Len(Downloader.Data) > 0)
'End Function

Sub HideModeMessage()
    If CommandLine.HideMode Then
        If IsWindowVisible(CommandLine.hwnd) = 0 Then Unload Me ' Если окно было закрыто
        If MessageBox(CommandLine.hwnd, "Доступна новая версия программы. Загрузить сейчас?" & GetNewVersionText, "RodySoft — Обновление", vbQuestion + vbYesNo) = vbNo Then
            Unload Me
        End If
        CommandLine.HideMode = False
    End If
    If Me.Visible = False Then Me.Show
End Sub

Sub ShowError()
    CheckHide
    SetStatus "Сбой во время обновления"
    If ApiMsg("Ошибка во время загрузки обновления", "Ошибка", vbExclamation + vbRetryCancel) = vbRetry Then Call StartUpdate
End Sub

Function UpdateUsers() As Boolean
    Downloader.DownloadToString SENDINFO_URL & "adduser", , False
    UpdateUsers = (Downloader.Data = "OK")
End Function

Sub FullUpdate()
    Dim FileName As String, q As String, sParams As String
    q = Chr$(34)
    If IsUserAnAdmin = 0 Then ShellExecute 0, "runas", App.Path & "\" & App.EXEName & ".exe", "/au", App.Path, vbNormalFocus: End
    SetStatus "Загрузка обновления..."
    FileName = Environ$("temp") & "\" & InstallerName
    If Downloader.DownloadToFile(FULLUPDATE_FILENAME_URL, FileName) Then
        'Shell FileName & " /s /p=" & q & App.Path & q & "\", vbHide
        sParams = "/s /p=" & GetShortFileName(App.Path)
        BAT_Body = q & FileName & q & " " & sParams
        AddLineToString BAT_Body, "start " & App.EXEName & ".exe /success " & CommandLine.BaseName
        ShowDownloadCompleteMessage
        RunBat BAT_Body
        Unload Me
    Else
        Debug.Print FULLUPDATE_FILENAME_URL
        If ApiMsg("Ошибка во время загрузки обновления", "Ошибка", vbExclamation + vbRetryCancel) = vbRetry Then Call FullUpdate
        SetStatus "Сбой во время обновления"
    End If
End Sub

Sub CheckHide()
    If CommandLine.HideMode Then Unload Me
End Sub

Sub ShowDownloadCompleteMessage()
    ApiMsg "Загрузка обновлений завершена. Во время обновления модули EasyScript будут закрыты. Перед тем, как продолжить, сохраните несохраненные данные.", "EasyScript Updater"
End Sub
