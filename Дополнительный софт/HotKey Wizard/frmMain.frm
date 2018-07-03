VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmMain 
   Caption         =   "RodySoft EasyScript Manager"
   ClientHeight    =   4680
   ClientLeft      =   6405
   ClientTop       =   2775
   ClientWidth     =   11295
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4680
   ScaleWidth      =   11295
   Begin ComctlLib.ListView List 
      Height          =   4095
      Left            =   0
      TabIndex        =   3
      Top             =   10
      Width           =   11295
      _ExtentX        =   19923
      _ExtentY        =   7223
      View            =   3
      Arrange         =   1
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      OLEDropMode     =   1
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   16777215
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      OLEDropMode     =   1
      NumItems        =   7
      BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Key             =   ""
         Object.Tag             =   "tNumber"
         Text            =   "#"
         Object.Width           =   706
      EndProperty
      BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   1
         Key             =   ""
         Object.Tag             =   "tCaption"
         Text            =   "Название"
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   2
         Key             =   ""
         Object.Tag             =   "tPath"
         Text            =   "Путь"
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(4) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   3
         Key             =   ""
         Object.Tag             =   "tHotKey"
         Text            =   "Горячая клавиша"
         Object.Width           =   3352
      EndProperty
      BeginProperty ColumnHeader(5) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   4
         Key             =   ""
         Object.Tag             =   "tState"
         Text            =   "Состояние"
         Object.Width           =   2118
      EndProperty
      BeginProperty ColumnHeader(6) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   5
         Key             =   ""
         Object.Tag             =   "tCount"
         Text            =   "Запущено экземпляров"
         Object.Width           =   3529
      EndProperty
      BeginProperty ColumnHeader(7) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   6
         Key             =   ""
         Object.Tag             =   "tAlreadyRuning"
         Text            =   "Если уже запущен"
         Object.Width           =   2823
      EndProperty
   End
   Begin VB.Timer tmrUpdate 
      Interval        =   1
      Left            =   5040
      Top             =   4080
   End
   Begin VB.CommandButton cmdKill 
      Caption         =   "Завершить скрипт"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2400
      TabIndex        =   5
      Top             =   4200
      Width           =   2175
   End
   Begin VB.CommandButton cmdRun 
      Caption         =   "Запустить скрипт"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   4200
      Width           =   2175
   End
   Begin VB.CommandButton cmdDel 
      Caption         =   "Удалить"
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
      Left            =   9960
      TabIndex        =   2
      Top             =   4200
      Width           =   1215
   End
   Begin VB.CommandButton cmdChange 
      Caption         =   "Изменить"
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
      Left            =   8640
      TabIndex        =   1
      Top             =   4200
      Width           =   1215
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Добавить скрипт"
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
      Left            =   6480
      TabIndex        =   0
      Top             =   4200
      Width           =   2055
   End
   Begin ComctlLib.ImageList ImageList 
      Left            =   4320
      Top             =   3960
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   5
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":2EFA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":324C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":359E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":38F0
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":3ACA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&Файл"
      Begin VB.Menu mnuNew 
         Caption         =   "&Новый список"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuOpenList 
         Caption         =   "&Открыть список"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuSaveList 
         Caption         =   "&Сохранить список"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuSaveAs 
         Caption         =   "&Сохранить список как"
      End
      Begin VB.Menu sep5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAutoAdd 
         Caption         =   "&Автоматически заполнить список"
      End
      Begin VB.Menu sep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTray 
         Caption         =   "&Свернуть в трей"
      End
      Begin VB.Menu sep7 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "&Выход"
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "&Опции"
      WindowList      =   -1  'True
      Begin VB.Menu mnuRun 
         Caption         =   "&Запустить скрипт"
      End
      Begin VB.Menu mnuKill 
         Caption         =   "&Завершить скрипт"
      End
      Begin VB.Menu sep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "&Отправить в редактор"
      End
      Begin VB.Menu mnuOpenInDir 
         Caption         =   "&Показать в папке"
      End
      Begin VB.Menu mnuOpenEditor 
         Caption         =   "&Создать новый скрипт в редакторе"
      End
      Begin VB.Menu sep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuChange 
         Caption         =   "&Изменить"
      End
      Begin VB.Menu mnuDel 
         Caption         =   "&Удалить"
      End
      Begin VB.Menu mnuAdd 
         Caption         =   "&Добавить"
      End
   End
   Begin VB.Menu mnuSettings 
      Caption         =   "&Настройки"
      Begin VB.Menu mnuAutoRun 
         Caption         =   "&Автозапуск при старте Windows"
      End
      Begin VB.Menu mnuEqvHotKeys 
         Caption         =   "&Разрешить использование одинаковых горячих клавиш"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuTrayOn 
         Caption         =   "&Сворачивать в трей"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuKillAll 
         Caption         =   "&Завершить все скрипты"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Справка"
      Begin VB.Menu mnuCheckUpdate 
         Caption         =   "&Проверить обновление"
      End
      Begin VB.Menu mnuSendError 
         Caption         =   "&Отправить информацию об ошибке"
      End
      Begin VB.Menu sep6 
         Caption         =   "-"
      End
      Begin VB.Menu mnuShowHelp 
         Caption         =   "&Помощь"
      End
      Begin VB.Menu sep8 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "О программе"
      End
   End
   Begin VB.Menu mnuTrayMenu 
      Caption         =   "{Tray}"
      Visible         =   0   'False
      Begin VB.Menu mnuTShow 
         Caption         =   "&Показать EasyScript Manager"
      End
      Begin VB.Menu sep4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTExit 
         Caption         =   "&Выход"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim WithEvents TrayIcon As cBalloonSysTray
Attribute TrayIcon.VB_VarHelpID = -1

Private Sub cmdAdd_Click()
    ShowCreate
End Sub

Private Sub cmdChange_Click()
    ShowChange
End Sub

Private Sub cmdDel_Click()
    Dim clean As HotKey
    UpdateSubItemsArray
    If MsgBox("Удалить из списка скрипт " & Chr$(34) & SubItemsArray(COLUMN_TITLE) & Chr$(34) & "?", vbYesNo + vbQuestion, "Вопрос") = vbYes Then
        HKSArray(GetIndex(GetSelectedItem)) = clean
        List.ListItems.Remove GetSelectedItem
    End If
End Sub

Private Sub cmdKill_Click()
    KillProcesses GetIndex(GetSelectedItem)
End Sub

Private Sub cmdRun_Click()
    cmdRun.Enabled = False
    HotKeyHandler GetRealIndex(GetSelectedItem)
    UpdateControls
End Sub

Private Sub Form_Initialize()
    If App.PrevInstance Then
        Dim PrevHwnd As Long
        PrevHwnd = ReadIni("PrevHwnd", 0, , dNumeric)
        If PrevHwnd > 0 Then
            WriteIni "Command", Command$
            PostMessage PrevHwnd, 0&, 0&, 0&
            End
        End If
    End If
    InitCommonControls
End Sub

Private Sub Form_Load()
    #If COMPILED Then
        MainFormPrevProc = SetWindowLong(Me.hwnd, GWL_WNDPROC, AddressOf wpMainForm)
    #End If
    Set TrayIcon = New cBalloonSysTray
    Call AddToSysTray
    Set mdlMain.List = frmMain.List
    Form_Resize
    LoadSettings
    LoadList
    UpdateStatus
    UpdateControls
    If Command$ <> "/autorun" Then Show Else Me.Visible = False
    If IsFileExists(App.path & "\" & SETTINGS_FILENAME) = False _
    And IsFileExists(App.path & "\List.rshw") = False _
    And IsFileMaskExistsInDir(ScriptsPath) And Left$(Command$, 4) <> "/add" Then
        If MsgBox("Это первый запуск программы. Хотите ли вы автоматически загрузить список скриптов из папки Scripts?", _
        vbYesNo + vbInformation, "EasyScript Manager") = vbYes Then
            If AutoAdd(ScriptsPath, "Ctrl", STR_Just_Kill) Then MsgBox "Все скрипты успешно загружены. Чтобы изменить комбинацию горячих клавиш, выберите скрипт и нажмите 'Изменить'", vbInformation, "Операция выполнена"
        End If
    End If
    LoadCommandLine
    UpdateAllHotKeys
    WriteIni "PrevHwnd", Me.hwnd
    ProgrammLoaded = True
    ShowMe
    CheckUpdate (True)
End Sub

Sub LoadCommandLine(Optional ByVal strLine As String)
    If strLine = "" Then strLine = Command$
    If Left$(strLine, 1) = Chr$(34) And Right$(strLine, 1) = Chr$(34) Then strLine = Mid$(strLine, 2, Len(strLine) - 2)
    If IsFileExists(strLine) Then
        ScriptsList_FN = strLine
        LoadList
    End If
    If Left$(strLine, 4) = "/add" Then
        Dim sScriptName As String
        sScriptName = Right$(strLine, Len(strLine) - 5)
        If Left$(sScriptName, 1) = Chr$(34) And Right$(sScriptName, 1) = Chr$(34) Then sScriptName = Mid$(sScriptName, 2, Len(sScriptName) - 2)
        AddScriptFileToList sScriptName
    End If
End Sub

Sub AddToSysTray()
    With TrayIcon
        .OwnerForm = Me
        .OwnerMenuObject = mnuTrayMenu
        .DefaultMenuObject = mnuTShow
        .IconHandle = TrayIcon.IconHandle = ImageList.ListImages.Item(ICON_SYSTRAY_NOTRUNNING).Picture.Handle
        .ToolTip = App.ProductName
        .ShowTrayIcon
    End With
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    TrayIcon.CallEvents X
End Sub

Private Sub List_OLEDragDrop(Data As ComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim n As Long
    For n = 1 To Data.Files.Count
        If Right$(Data.Files(n), 3) = ".es" Then AddScriptFileToList Data.Files(n)
    Next n
End Sub

Private Sub mnuCheckUpdate_Click()
    CheckUpdate
End Sub

Private Sub mnuEqvHotKeys_Click()
    mnuEqvHotKeys.Checked = Not mnuEqvHotKeys.Checked
End Sub

Private Sub mnuOpenEditor_Click()
    StartProgramm "Editor.exe", "/cl"
End Sub

Private Sub mnuSendError_Click()
    SendInfoToServer , "bugreport", hwnd
End Sub

Private Sub mnuShowHelp_Click()
    ShowHelp "Manager", frmMain
End Sub

Private Sub TrayIcon_SysTrayMouseUp(ByVal eButton As MouseButtonConstants)
    If eButton = vbLeftButton Then ShowMe
    If eButton = vbRightButton Then TrayIcon.ShowMenu
End Sub

Sub ShowMe()
    If Me.Enabled Then
        Me.Visible = True
        Me.Show
        ZOrder
        ShowWindow Me.hwnd
    End If
End Sub

Private Sub Form_Resize()
    Static PrevWidth As Single, PrevHeight As Single, blnValide As Boolean
    Dim XOffset As Single, YOffset As Single
    If Me.WindowState = 1 Then Exit Sub
    If blnValide Then
        XOffset = Width - PrevWidth
        YOffset = Height - PrevHeight
    End If
    List.Width = List.Width + XOffset
    List.Height = List.Height + YOffset
    For Each Arr In Me
        If TypeOf Arr Is CommandButton Then
            Arr.Top = Arr.Top + YOffset
            If Arr.hwnd = cmdAdd.hwnd Or Arr.hwnd = cmdChange.hwnd Or Arr.hwnd = cmdDel.hwnd Then Arr.Left = Arr.Left + XOffset
        End If
    Next
    PrevWidth = Width
    PrevHeight = Height
    blnValide = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Cancel = 1
    ToTray True
End Sub

Private Sub List_ColumnClick(ByVal ColumnHeader As ComctlLib.ColumnHeader)
    List.SortKey = ColumnHeader.Index - 1
    If List.SortKey = 1 Then List.SortKey = 3 ' **** This column changes the key
    List.SortOrder = (List.SortOrder - 1) * -1
    List.Sorted = True
End Sub

Private Sub List_DblClick()
    ShowChange
End Sub

Sub ShowChange()
    If List.ListItems.Count > 0 Then
        If List.SelectedItem.Selected Then frmChangeItem.ChangeItem chkmChange
    Else
        ShowCreate
    End If
End Sub

Sub ShowCreate()
    frmChangeItem.ChangeItem chkmCreate
End Sub

Private Sub List_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete And cmdDel.Enabled Then cmdDel_Click
End Sub

Private Sub List_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    UpdateControls
    If Button = 2 Then
        PopupMenu mnuOptions, , , , mnuAdd
    End If
End Sub

Private Sub mnuAbout_Click()
    frmAbout.Show vbModal, Me
End Sub

Private Sub mnuAdd_Click()
    cmdAdd_Click
End Sub

Private Sub mnuAutoAdd_Click()
    frmAutoAdd.Show vbModal, Me
End Sub

Private Sub mnuAutoRun_Click()
    If mnuAutoRun.Checked Then
        If Autorun_Delete Then
            mnuAutoRun.Checked = False
        End If
    Else
        If Autorun_Add Then
            mnuAutoRun.Checked = True
        End If
    End If
End Sub

Private Sub mnuChange_Click()
    cmdChange_Click
End Sub

Private Sub mnuDel_Click()
    cmdDel_Click
End Sub

Private Sub mnuEdit_Click()
    Me.MousePointer = vbHourglass
    UpdateSubItemsArray
    APIShell Editor_FN, SubItemsArray(COLUMN_FILENAME), vbNormalFocus
    Me.MousePointer = vbNormal
End Sub

Private Sub mnuExit_Click()
    Shutdown
End Sub

Private Sub mnuKill_Click()
    cmdKill_Click
End Sub

Private Sub mnuNew_Click()
    If List.ListItems.Count = 0 Then Exit Sub
    If MsgBox("Вы точно хотите очистить весь список?", vbQuestion + vbYesNo, "") = vbYes Then
        List.ListItems.Clear
        ScriptsList_FN = App.path & "\List.rshw"
        UpdateAllHotKeys
    End If
End Sub

Private Sub mnuOpenInDir_Click()
    MousePointer = vbHourglass
    UpdateSubItemsArray
    Call ShowFileInDir(SubItemsArray(COLUMN_FILENAME))
    MousePointer = vbNormal
End Sub

Private Sub mnuOpenList_Click()
    Dim Temp As typeFileStruct
    Temp = OpenFile("Выберите список скриптов", "", ScriptsList_FN, "Списки скриптов EasyScript (*.rshw)" & Chr$(0) & "*.rshw", Me.hwnd)
    If Temp.FileName <> "" Then
        ScriptsList_FN = Temp.FileName
        LoadList
        SaveSettings
    End If
End Sub

Private Sub mnuOptions_Click()
    UpdateControls
End Sub

Private Sub mnuRun_Click()
    HotKeyHandler GetRealIndex(GetSelectedItem)
End Sub

Private Sub mnuSaveAs_Click()
    Dim Temp As typeFileStruct
    Temp = SaveFile("Сохранить список как", "", ScriptsList_FN, "Списки скриптов EasyScript (*.rshw)" & Chr$(0) & "*.rshw", Me.hwnd)
    If Temp.FileName <> "" Then
        If CreateFile(Temp.FileName, "") Then
            ScriptsList_FN = Temp.FileName
            SaveList
            SaveSettings
        Else
            MsgBox "Не удается сохранить файл", vbExclamation, "Ошибка"
        End If
    End If
End Sub

Private Sub mnuSaveList_Click()
    SaveList
    SaveSettings
End Sub

Private Sub mnuTExit_Click()
    Shutdown
End Sub

Private Sub mnuTray_Click()
    ToTray
End Sub

Sub ToTray(Optional ByVal blnExit As Boolean)
    If blnExit Then TrayIcon.ShowBalloonTip "EasyScript Manager все еще работает", "EasyScript Manager", NIIF_INFO
    Hide
End Sub

Private Sub mnuTrayOn_Click()
    mnuTrayOn.Checked = Not mnuTrayOn.Checked
End Sub

Private Sub mnuTShow_Click()
    ShowMe
End Sub

Private Sub tmrUpdate_Timer()
    CheckHotKeys
    If Me.WindowState = 1 And mnuTrayOn.Checked And Me.Visible Then
        Sleep 100
        ToTray
    End If
    If blnRuning Then
        TrayIcon.IconHandle = ImageList.ListImages.Item(ICON_SYSTRAY_RUNNING).Picture.Handle
        If TrayIcon.ToolTip <> App.ProductName & " - Скрипт запущен" Then TrayIcon.ToolTip = App.ProductName & " - Скрипт запущен"
    Else
        TrayIcon.IconHandle = ImageList.ListImages.Item(ICON_SYSTRAY_NOTRUNNING).Picture.Handle
        If TrayIcon.ToolTip <> App.ProductName Then TrayIcon.ToolTip = App.ProductName
    End If
End Sub

Sub UpdateControls()
    Dim RealIndex As Long
    RealIndex = GetSelectedItem
    cmdKill.Caption = "Завершить скрипт"
    If RealIndex = 0 Then
        cmdRun.Enabled = False
        cmdKill.Enabled = False
        cmdChange.Enabled = False
        cmdDel.Enabled = False
    Else
        UpdateSubItemsArray (RealIndex)
        Select Case SubItemsArray(COLUMN_STATUS)
            Case STATUS_RUNNING
                'запущен
                cmdKill.Enabled = True
                cmdChange.Enabled = False
                cmdDel.Enabled = False
                If SubItemsArray(COLUMN_IFRUNING) = STR_Just_Run Then
                    cmdRun.Enabled = True
                Else
                    cmdRun.Enabled = False
                End If
                If HKSArray(GetIndex(RealIndex)).CopysRuningCount > 1 Then
                    cmdKill.Caption = "Завершить скрипты"
                End If
            Case STATUS_NOTRUNNING
                'не запущен
                cmdRun.Enabled = True
                cmdKill.Enabled = False
                cmdChange.Enabled = True
                cmdDel.Enabled = True
            Case STATUS_BAD
                'ошибка
                cmdRun.Enabled = False
                cmdKill.Enabled = False
                cmdChange.Enabled = True
                cmdDel.Enabled = True
        End Select
    End If
    mnuOpenInDir = CBool(RealIndex)
    mnuEdit = CBool(RealIndex)
    mnuRun = cmdRun.Enabled
    mnuKill = cmdKill.Enabled
    mnuAdd = cmdAdd.Enabled
    mnuChange = cmdChange.Enabled
    mnuDel = cmdDel.Enabled
End Sub

Sub Shutdown()
    TrayIcon.HideTrayIcon
    Unload frmAbout
    Unload frmAutoAdd
    Unload frmChangeItem
    Hide
    SaveSettings
    SaveList
    End
End Sub
