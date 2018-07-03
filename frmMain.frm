VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00F1E4D9&
   Caption         =   "RodySoft EasyScript Editor"
   ClientHeight    =   6210
   ClientLeft      =   5595
   ClientTop       =   4395
   ClientWidth     =   9225
   BeginProperty Font 
      Name            =   "Book Antiqua"
      Size            =   9.75
      Charset         =   204
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Main"
   ScaleHeight     =   414
   ScaleMode       =   3  'Пиксель
   ScaleWidth      =   615
   Begin VB.TextBox txtCoords 
      BorderStyle     =   0  'Нет
      Enabled         =   0   'False
      ForeColor       =   &H00008000&
      Height          =   255
      Left            =   7560
      Locked          =   -1  'True
      MousePointer    =   1  'Указатель
      TabIndex        =   11
      TabStop         =   0   'False
      Text            =   "X:1000 Y:1000"
      ToolTipText     =   "Кординаты курсора в реальном времени"
      Top             =   5880
      Width           =   1335
   End
   Begin VB.PictureBox picToolTip 
      Appearance      =   0  'Плоска
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000018&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   262
      Left            =   2880
      ScaleHeight     =   15
      ScaleMode       =   3  'Пиксель
      ScaleWidth      =   79
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   3840
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox Arrow 
      Appearance      =   0  'Плоска
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'Нет
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   54
      Picture         =   "frmMain.frx":1042
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Timer tmrUpdate 
      Interval        =   1
      Left            =   120
      Top             =   3000
   End
   Begin VB.Timer tmrCheckHKs 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   126
      Top             =   4080
   End
   Begin VB.ListBox lstToolTip 
      Appearance      =   0  'Плоска
      BackColor       =   &H80000018&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      ItemData        =   "frmMain.frx":13FA
      Left            =   3906
      List            =   "frmMain.frx":1404
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   2394
      Visible         =   0   'False
      Width           =   1148
   End
   Begin VB.TextBox txtStatus 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Вертикаль
      TabIndex        =   6
      TabStop         =   0   'False
      ToolTipText     =   "Поле статуса"
      Top             =   5535
      Width           =   9230
   End
   Begin VB.Frame frmNum 
      Appearance      =   0  'Плоска
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'Нет
      Caption         =   "Frame1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1638
      Left            =   504
      TabIndex        =   2
      Top             =   2142
      Visible         =   0   'False
      Width           =   420
      Begin VB.Label num 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Прозрачно
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Book Antiqua"
            Size            =   9.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   266
         Index           =   0
         Left            =   126
         TabIndex        =   5
         Top             =   504
         Visible         =   0   'False
         Width           =   98
      End
      Begin VB.Label num 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Прозрачно
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Book Antiqua"
            Size            =   9.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   1
         Left            =   126
         TabIndex        =   4
         Top             =   0
         Width           =   98
      End
      Begin VB.Label num 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Прозрачно
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "Book Antiqua"
            Size            =   9.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   2
         Left            =   126
         TabIndex        =   3
         Top             =   252
         Width           =   98
      End
   End
   Begin VB.TextBox txtMain 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'Нет
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5535
      Left            =   360
      MultiLine       =   -1  'True
      OLEDropMode     =   1  'Вручную
      ScrollBars      =   3  'Оба
      TabIndex        =   0
      ToolTipText     =   "Нажмите ""Действия > Начать запись"", либо вводите команды самостоятельно"
      Top             =   0
      Width           =   8850
   End
   Begin VB.Frame frmStatusColor 
      Appearance      =   0  'Плоска
      BackColor       =   &H00F1E4D9&
      BorderStyle     =   0  'Нет
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   375
   End
   Begin VB.Image ArrowOLD 
      Height          =   225
      Left            =   120
      Picture         =   "frmMain.frx":140E
      Top             =   120
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Label lblForMeasure 
      AutoSize        =   -1  'True
      Caption         =   "Label1"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   2880
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&Файл"
      Begin VB.Menu mnuNew 
         Caption         =   "&Новый"
         Shortcut        =   ^N
      End
      Begin VB.Menu s1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "&Открыть"
         Shortcut        =   ^O
      End
      Begin VB.Menu s2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSave 
         Caption         =   "&Сохранить"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuSaveAs 
         Caption         =   "&Сохранить как"
      End
      Begin VB.Menu s3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuBuildExe 
         Caption         =   "&Создать script.exe"
      End
      Begin VB.Menu s4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "&Выход"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Правка"
      Begin VB.Menu mnuShowMarker 
         Caption         =   "&Показать маркер координат"
      End
      Begin VB.Menu s13 
         Caption         =   "-"
      End
      Begin VB.Menu mnuUndo 
         Caption         =   "&Шаг назад"
         Shortcut        =   ^Z
      End
      Begin VB.Menu mnuRedo 
         Caption         =   "&Шаг вперед"
         Shortcut        =   ^Y
      End
      Begin VB.Menu s10 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCopy 
         Caption         =   "&Копировать"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuCut 
         Caption         =   "&Вырезать"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuPaste 
         Caption         =   "&Вставить"
         Shortcut        =   ^V
      End
      Begin VB.Menu mnuDel 
         Caption         =   "&Удалить"
         Shortcut        =   {DEL}
      End
      Begin VB.Menu s8 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFind 
         Caption         =   "&Найти"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuFindNext 
         Caption         =   "&Найти далее"
         Shortcut        =   {F3}
      End
      Begin VB.Menu mnuReplace 
         Caption         =   "&Заменить"
         Shortcut        =   ^H
      End
      Begin VB.Menu s9 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSelectAll 
         Caption         =   "&Выделить все"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuCleanAll 
         Caption         =   "&Очистить все"
      End
   End
   Begin VB.Menu mnuActions 
      Caption         =   "&Действия"
      Begin VB.Menu mnuExecute 
         Caption         =   "&Выполнить скрипт"
         Enabled         =   0   'False
      End
      Begin VB.Menu s5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRec 
         Caption         =   "&Начать запись"
      End
   End
   Begin VB.Menu mnuInsts 
      Caption         =   "&Инструменты"
      Begin VB.Menu mnuOpenVESE 
         Caption         =   "&Открыть EasyScript Visual Editor"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuOpenHKM 
         Caption         =   "&Открыть EasyScript Manager"
      End
      Begin VB.Menu mnuAddScriptToManager 
         Caption         =   "&Добавить скрипт в EasyScript Manager"
      End
      Begin VB.Menu sep13 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRunMFC 
         Caption         =   "&Запустить Most Fastest Clicker"
      End
   End
   Begin VB.Menu mnuSettings 
      Caption         =   "&Настройки"
      Begin VB.Menu mnuProgrammSettings 
         Caption         =   "&Настройки программы"
      End
      Begin VB.Menu mnuScriptSettings 
         Caption         =   "&Настройки скрипта"
      End
      Begin VB.Menu s7 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuCustomForm 
         Caption         =   "&Собственные константы и команды"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuInfo 
      Caption         =   "&Справка"
      Begin VB.Menu mnuCheckUpdate 
         Caption         =   "&Проверить обновление"
      End
      Begin VB.Menu mnuSendError 
         Caption         =   "&Отправить информацию об ошибке"
      End
      Begin VB.Menu s11 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelp 
         Caption         =   "&Помощь"
         Shortcut        =   {F1}
      End
      Begin VB.Menu s12 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "&О программе"
      End
   End
   Begin VB.Menu mnuToolTip 
      Caption         =   "[Подсказка]"
      Visible         =   0   'False
      Begin VB.Menu mnuSelectFunc 
         Caption         =   "&Выбрать {Enter}"
      End
      Begin VB.Menu mnuFuncInfo 
         Caption         =   "&Информация о команде"
      End
      Begin VB.Menu mnuHideToolTip 
         Caption         =   "&Скрыть подсказки"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ************************************** Отличный вариант делать 2 версии - одну полную со всеми приблудами, а вторую урезанную для компиляции скриптов
Option Explicit
Dim LinesCount As Long, i As Integer, n As Long, SelStartLine As Long, _
sngWidthChange As Single, sngWidth As Single, sngHeightChange As Single, sngHeight As Single, _
ScriptVal As String, blnProgrammLoaded As Boolean, blnIsAStep As Boolean, blnUpdateTextOff As Boolean

Dim FuncInLine As Functions, PrevLine As Long, PrevSelStart As Long

Private Const LISTBOX_ITEM_HEIGHT = 16
Private Const LISBOX_ONE_SYMBOL_LENGH = 8
Private Const TEXTBOX_LINE_HEIGHT = 16

Private Sub Form_Click()
    HideMarker
    HideToolTip
End Sub

Sub UpdateStatusColorPos()
    frmStatusColor.Width = txtMain.Left
    frmStatusColor.Height = txtStatus.Top
End Sub

Private Sub Form_Initialize()
    InitCommonControls
End Sub

Private Sub Form_Load()
    #If COMPILED Then
        MainFormPrevProc = SetWindowLong(Me.hwnd, GWL_WNDPROC, AddressOf wpMainForm)
        ScriptTextBoxPrevProc = SetWindowLong(txtMain.hwnd, GWL_WNDPROC, AddressOf wpScriptTextBox)
    #End If
    frmNum.Left = txtMain.Left - frmNum.Width - 20
    sngWidth = Me.Width
    sngHeight = Me.Height
    mnuSaveAs.Caption = "Сохранить как" & vbTab & "Ctrl+Shift+S"
    LoadDefaultScriptSettings
    LoadSettings
    SetAll
    PrevSelStart = -1
    CheckResourceCode
    If Command$ <> "" Then
        ExecuteCommandLine
    End If
    UpdateStatusColorPos
    If frmPSettings.chkShowButtons Then ShowButtons
    ShowMe
    CheckUpdate (True)
End Sub

Sub CheckResourceCode()
    Dim ResourceCode As String
    ResourceCode = GetResourceCode
    If Len(ResourceCode) > 3 Then
        LoadScript , , , ResourceCode
        CommandLine.HideMode = True
        ExecuteScript LoadedFile.ScriptBody
        End
    End If
End Sub

Sub ShowButtons()
    frmButtons.Left = Me.Left + Me.Width - frmButtons.Width - (Screen.TwipsPerPixelX * GetScrollWidth)
    frmButtons.Top = Me.Top + (Screen.TwipsPerPixelY * txtMain.Height) - frmButtons.Height + (Screen.TwipsPerPixelX * GetSystemMetrics(SM_CYCAPTION)) + (Screen.TwipsPerPixelX * GetScrollHeight)
    frmButtons.Show , Me
End Sub

Sub ExecuteCommandLine()
    Const hm = "hm", fn = "fn", se = "se", ex = "ex", co = "co", dh = "dh", au = "au", cl = "cl" ' Clean - запустить редактор, не загружая файла
    Dim Commands() As String, strExParam As String, c As Long
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
                Case hm: .HideMode = True: .DisableHKs = True: .AutoStart = True: .Exit = True: .SkipErrors = True
                Case se: .SkipErrors = True
                Case ex: .Exit = True
                Case dh: If strExParam = "0" Then .DisableHKs = False Else .DisableHKs = True
                Case co: If IsNumeric(strExParam) Then .Count = CLng(strExParam)
                Case fn: .FileName = strExParam
                Case au: .AutoStart = True
                Case cl: NewScript
            End Select
        Next i
        If .FileName = "" And InStr(Command$, "/") = 0 Then .FileName = Command$
        If Left$(.FileName, 1) = Chr$(34) And Right$(.FileName, 1) = Chr$(34) Then .FileName = Mid$(.FileName, 2, Len(.FileName) - 2)
        If Len(.FileName) > 3 And InStr(1, .FileName, "\") = 0 Then .FileName = App.Path & "\" & .FileName
        If IsFileExists(.FileName) Then
            LoadScript .FileName
        ElseIf Len(.FileName) > 0 Then
            ShowNotFound .FileName
            Exit Sub
        End If
        Settings.bSkipErrors = .SkipErrors
        If Not .AutoStart Then Exit Sub
        If .HideMode Then Me.Visible = False Else ShowMe
        If .Count = 0 Then .Count = 1
        For c = 1 To .Count
            If IsValidScript(txtMain) Then
                If .HideMode Then ExecuteScript txtMain Else mnuExecute_Click
            End If
        Next c
        If .Exit Or .HideMode Then End
    End With
End Sub

Sub ShowMe()
    Me.Show
    Me.Visible = True
    ShowWindow Me.hwnd
    ZOrder
    Me.SetFocus
    blnProgrammLoaded = True
    UpdateAll True
    SetCoordsPos
End Sub

Private Sub Form_LostFocus()
    lstToolTip.Visible = False
    picToolTip.Visible = False
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Static PrevX As Single, PrevY As Single
    If Button = 1 And WindowState = 0 Then
        Me.Left = Me.Left + (X - PrevX)
        Me.Top = Me.Top + (Y - PrevY)
    Else
        PrevX = X
        PrevY = Y
    End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If Left$(Me.Caption, 1) = "*" Then
        Select Case MsgBox("Сохранить изменения в файле " & LoadedFile.BaseName & "?", vbYesNoCancel + vbQuestion, "Файл изменен")
            Case vbYes: mnuSave_Click
            Case vbNo
            Case Else: Cancel = 1: Exit Sub
        End Select
    End If
    For Each arr In Forms
        If arr.Visible Then arr.Hide
    Next
    blnCancel = True
    blnExecuting = False
    blnRec = False
    DoEvents
    SaveSettings
    End
End Sub

Private Sub Form_Resize()
    Static antirecurse As Boolean
    If antirecurse Then Exit Sub
    antirecurse = True
    If WindowState = 0 Or WindowState = 2 Then
        sngWidthChange = Me.Width - sngWidth
        sngHeightChange = Me.Height - sngHeight
        With txtMain
            .Height = .Height + (sngHeightChange / Screen.TwipsPerPixelY)
            .Width = .Width + (sngWidthChange / Screen.TwipsPerPixelX)
        End With
        With txtStatus
            .Top = .Top + (sngHeightChange / Screen.TwipsPerPixelY)
            .Width = .Width + (sngWidthChange / Screen.TwipsPerPixelX)
        End With
        SetCoordsPos
        sngHeight = Me.Height
        sngWidth = Me.Width
    End If
    If Not blnUpdateTextOff And blnProgrammLoaded Then
        UpdateStatusColorPos
        If lstToolTip.Visible Then UpdateToolTip
    End If
    antirecurse = False
End Sub

Sub SetCoordsPos()
    With txtCoords
        .Left = txtStatus.Width - .Width - GetScrollWidth - 3
        .Top = txtStatus.Top + txtStatus.Height - .Height - 1
    End With
End Sub

Private Sub lstToolTip_Click()
    txtMain.SetFocus
End Sub

Private Sub lstToolTip_DblClick()
    SelectFunc lstToolTip.Text
End Sub

Private Sub lstToolTip_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        LeftClick
        PopupMenu mnuToolTip, , , , mnuSelectFunc
    End If
End Sub

Private Sub mnuAbout_Click()
    frmAbout.Show vbModal, Me
End Sub

Private Sub mnuAddScriptToManager_Click()
    If LoadedFile.FileName = "" Then
        If MsgBox("Скрипт не сохранен, поэтому его невозможно добавить в список EasyScript Manager. Сохранить скрипт?", vbQuestion + vbYesNo, "Сохраните скрипт") = vbNo Then Exit Sub
        Call mnuSave_Click
    End If
    If Len(LoadedFile.FileName) > 0 Then StartProgramm "Manager.exe", "/add " & LoadedFile.FileName
End Sub

Private Sub mnuCheckUpdate_Click()
    CheckUpdate
End Sub

Private Sub mnuCleanAll_Click()
    txtMain.Text = ""
End Sub

Private Sub mnuCopy_Click()
    On Error Resume Next
    If txtMain.SelLength = 0 Then Exit Sub
    Clipboard.Clear
    Clipboard.SetText txtMain.SelText
End Sub

Private Sub mnuBuildExe_Click()
    'On Error Resume Next
    Dim bStatus As Boolean, FileName As String, Content As String
    If txtMain = "" Then MsgBox "Скрипт пуст", vbExclamation: Exit Sub
    Me.MousePointer = vbHourglass
    txtMain.MousePointer = vbHourglass
    Save
    If LoadedFile.Content <> "" Then
        FileName = LoadedFile.Path & LoadedFile.BaseTitle & ".exe"
        Content = LoadedFile.Content
    Else
        FileName = App.Path & "\Script.exe"
        Content = txtMain
    End If
    bStatus = BuildExe(Content, FileName)
    If bStatus Then
        MsgBox "Скрипт создан... Нажмите ОК, чтобы открыть папку.", vbInformation, "Готово!"
        Shell "explorer /select," & FileName, vbNormalFocus
    Else
        MsgBox "Не удается создать скрипт! Возможно, файл занят другой программой.", vbExclamation:
    End If
    Me.MousePointer = vbNormal
    txtMain.MousePointer = vbNormal
End Sub

Private Sub mnuCustomForm_Click()
    frmCustom.Show , Me
End Sub

Sub UpdateScriptInfo()
    If LoadedFile.FileName <> "" Then
        If LoadedFile.Path = ScriptsPath Then
            Me.Caption = LoadedFile.BaseName & " - " & "EasyScript Editor"
        Else
            Me.Caption = LoadedFile.FileName & " - " & "EasyScript Editor"
        End If
        If Not Compare(LoadedFile.ScriptBody, txtMain) Then Me.Caption = "* " & Me.Caption
        mnuBuildExe.Caption = "Создать " & LoadedFile.BaseTitle & ".exe"
    Else
        Me.Caption = "RodySoft EasyScript Editor"
        mnuBuildExe.Caption = "Создать Script.exe"
    End If
End Sub

Private Sub mnuCut_Click()
    On Error Resume Next
    If txtMain.SelLength = 0 Then Exit Sub
    Clipboard.Clear
    Clipboard.SetText txtMain.SelText
    SetScriptText AccurateReplace(txtMain, "", txtMain.SelStart + 1, txtMain.SelStart + txtMain.SelLength)
    txtMain.SelLength = 0
    PrevSelStart = -1
    UpdateAll
End Sub

Private Sub mnuDel_Click()
    If txtMain.SelLength > 0 Then
        SetScriptText AccurateReplace(txtMain, "", txtMain.SelStart + 1, txtMain.SelStart + txtMain.SelLength), , 0
    ElseIf Len(txtMain) > txtMain.SelStart Then
        If Mid$(txtMain, txtMain.SelStart + 1, 1) = vbCr Then
            SetScriptText Left$(txtMain, txtMain.SelStart) & Right$(txtMain, Len(txtMain) - txtMain.SelStart - 2), , 0
        Else
            SetScriptText Left$(txtMain, txtMain.SelStart) & Right$(txtMain, Len(txtMain) - txtMain.SelStart - 1), , 0
        End If
    End If
    PrevSelStart = -1
    UpdateAll
End Sub

Private Sub mnuExit_Click()
    Unload Me
End Sub

Private Sub mnuFind_Click()
    frmFind.ShowMe
End Sub

Private Sub mnuFindNext_Click()
    frmFind.FindNext
End Sub

Private Sub mnuHelp_Click()
    If FuncInLine > 0 Then ShowHelp FuncInLine, frmMain Else ShowHelp "EasyScript Editor", frmMain
End Sub

Private Sub mnuFuncInfo_Click()
    ShowHelp GetFunc(lstToolTip.Text, True), frmMain
End Sub

Private Sub mnuHideToolTip_Click()
    If MsgBox("Вы точно хотите скрыть список команд? После этого действия, включить подскакзки можно только в настройках.", vbQuestion + vbYesNo, "Данное действие требует подтверждения") = vbYes Then
        lstToolTip.Visible = False
        frmPSettings.chkShowListTT.Value = 0
        CacheProgrammSettings
    End If
End Sub

Private Sub mnuNew_Click()
    If txtMain <> "" Then
        If MsgBox("Вы точно хотите создать новый скрипт?", vbQuestion + vbYesNo, "EasyScript Editor") = vbNo Then Exit Sub
    End If
    NewScript
End Sub

Sub NewScript()
    Dim clear_file As typeFileStruct, clear_settings As mySettings
    blnUpdateTextOff = True
    StopAll
    LoadedFile = clear_file
    PrevSelStart = 0
    txtMain = ""
    History.Step = 0
    ReDim History.Text(0)
    ReDim History.SelStart(0)
    LoadDefaultScriptSettings
    blnUpdateTextOff = False
    UpdateAll
End Sub

Private Sub mnuOpen_Click()
    Dim TempLoaded As typeFileStruct, Path As String
    StopAll
    If LoadedFile.Path <> "" Then
        Path = LoadedFile.Path
    Else
        Path = ScriptsPath
    End If
    TempLoaded = OpenFile("Открыть скрипт", Path, LoadedFile.FileName, "Скрипты EasyScript (*.es)" & Chr$(0) & "*.es" & Chr$(0) & "Все файлы (*.*)" & Chr$(0) & "*.*" & Chr$(0), Me.hwnd)
    If TempLoaded.FileName <> "" Then
        'файл выбран
        LoadedFile = TempLoaded
        If Dir(LoadedFile.FileName, vbNormal) = "" Then CreateFile LoadedFile.FileName, ""
        LoadScript LoadedFile.FileName
        SaveSettings
    End If
    UpdateAll
End Sub

Private Sub mnuOpenHKM_Click()
    StartProgramm "Manager.exe"
End Sub

Private Sub mnuPaste_Click()
    Call Paste
End Sub

Sub Paste()
    On Error Resume Next
    Dim ClipText As String, tmp As String
    If txtMain.Locked Then Exit Sub
    ClipText = Clipboard.GetText
    tmp = txtMain
    If txtMain.SelLength > 0 Then
        tmp = AccurateReplace(tmp, "", txtMain.SelStart + 1, txtMain.SelStart + txtMain.SelLength)
    End If
    If txtMain.SelStart < Len(txtMain) Then
        tmp = Left$(tmp, txtMain.SelStart) & ClipText & Right$(tmp, Len(tmp) - (txtMain.SelStart))
    Else
        tmp = Left$(tmp, txtMain.SelStart + 1) & ClipText '& Right$(tmp, Len(tmp) - (txtMain.SelStart))
    End If
    txtMain.SelLength = 0
    SetScriptText tmp, txtMain.SelStart + Len(ClipText)
    PrevSelStart = -1
    UpdateAll
End Sub

Private Sub mnuProgrammSettings_Click()
'    Dim X As Long, Y As Long
'    GetCursorPos X, Y
'    If Me.Top > 3000 And Me.Left > 3000 Then
'        X = (Screen.TwipsPerPixelX * X) - frmPSettings.Width \ 2
'        Y = (Screen.TwipsPerPixelY * Y) - frmPSettings.Height \ 2
'    Else
'        X = (Screen.TwipsPerPixelX * X)
'        Y = (Screen.TwipsPerPixelY * Y)
'    End If
'    frmPSettings.Left = X
'    frmPSettings.Top = Y
    frmPSettings.Show vbModal, Me
End Sub

Public Sub mnuRec_Click()
    If mnuRec.Enabled = False Then Exit Sub
    If blnRec Then
        blnRec = False
        SetNormalState
    Else
        SetRecState
        StartRec
    End If
End Sub

Public Sub SetRecState()
    If frmPSettings.chkMinimizeOnRecord Then WindowState = 1
    blnExecuting = False
    blnRec = False
    HideToolTip
    HideMarker
    frmStatusColor.BackColor = RecColor
    mnuExecute.Enabled = False
    UpdateHotKeys , "Остановить запись"
    txtMain.Locked = True
    UpdateControlsState
    txtCoords = ""
End Sub

Sub HideToolTip()
    PrevSelStart = -1
    lstToolTip.Visible = False
    picToolTip.Visible = False
End Sub

Public Sub SetExecState()
    If frmPSettings.chkMinimize Then WindowState = 1
    blnExecuting = False
    blnRec = False
    HideToolTip
    HideMarker
    frmStatusColor.BackColor = RunColor
    UpdateHotKeys "Остановить скрипт", "Начать запись"
    mnuRec.Enabled = False
    txtMain.Locked = True
    ScriptVal = txtMain
    mnuCleanAll.Enabled = False
    txtCoords = ""
End Sub

Public Sub SetNormalState()
'    blnExecuting = False
'    blnRec = False
    Arrow.Visible = False
    mnuCleanAll.Enabled = True
    mnuRec.Enabled = True
    mnuExecute.Enabled = True
    txtMain.Locked = False
    If bDownloaderOn Then WHR.Abort
    frmStatusColor.BackColor = DefaultColor
    mnuExecute.Enabled = True
    UpdateHotKeys "Выполнить скрипт", "Начать запись"
    UpdateControlsState
    If frmPSettings.chkMaximize Then ShowWindow hwnd: txtMain.SetFocus
End Sub

Public Sub mnuExecute_Click()
    UpdateControlsState
    If mnuExecute.Enabled = False Then Exit Sub
    If blnExecuting Then
        blnExecuting = False
        SetNormalState
    Else
        SetExecState
        ExecuteScript txtMain, txtMain, Arrow, txtStatus
    End If
End Sub

Private Sub mnuRedo_Click()
    If History.Step = UBound(History.Text) Then Exit Sub
    blnIsAStep = True
    History.Step = History.Step + 1
    txtMain.Text = History.Text(History.Step)
    txtMain.SelStart = History.SelStart(History.Step)
    blnIsAStep = False
End Sub

Private Sub mnuReplace_Click()
    frmReplace.ShowMe
End Sub

Private Sub mnuRunMFC_Click()
    Dim Point As POINTAPI
    If FuncInLine = cClick Or FuncInLine = cSetCursorPos Then Point = GetPoint
    StartProgramm "MFC.exe", Point.X & "," & Point.Y
End Sub

Private Sub mnuSave_Click()
    Save
End Sub

Function Save()
    StopAll
    If LoadedFile.FileName = "" Then mnuSaveAs_Click: Exit Function
Save:
    If SaveScript(LoadedFile.FileName) Then
        LoadedFile.ScriptBody = txtMain
        SaveSettings
    Else
        If MsgBox("Невозможно сохранить файл", vbExclamation + vbRetryCancel, "Ошибка") = vbRetry Then
            GoTo Save
        End If
    End If
    UpdateScriptInfo
End Function

Public Sub mnuSaveAs_Click()
    Dim SavedFN As typeFileStruct
    StopAll
    If LoadedFile.Path = "" Then LoadedFile.Path = ScriptsPath
    SavedFN = SaveFile("Сохранить скрипт как", LoadedFile.Path, LoadedFile.FileName, "Скрипты EasyScript (*.es)" & Chr$(0) & "*.es" & Chr$(0) & "Все файлы (*.*)" & Chr$(0) & "*.*" & Chr$(0), Me.hwnd)
    If SavedFN.FileName = "" Then Exit Sub
lSave:
    If SaveScript(SavedFN.FileName) Then
        LoadedFile = SavedFN
        LoadedFile.ScriptBody = txtMain
        Save
    Else
        If MsgBox("Невозможно сохранить файл", vbExclamation + vbRetryCancel, "Ошибка") = vbRetry Then
            GoTo lSave
        End If
    End If
'    Else
'        ' иначе задать вопрос о перезаписи
'        Select Case MsgBox("Файл с таким именем уже существует. Заменить?", vbQuestion + vbYesNo, "Подтвердить сохранение в виде")
'            Case vbYes
'                GoTo Save
'        End Select
'    End If
    UpdateScriptInfo
End Sub

Private Sub mnuScriptSettings_Click()
    frmSSettings.Show vbModal, Me
End Sub

Private Sub mnuSelectAll_Click()
    txtMain.SelStart = 0
    txtMain.SelLength = Len(txtMain)
End Sub

Private Sub mnuSelectFunc_Click()
    SelectFunc
End Sub

Private Sub mnuSendError_Click()
    SendInfoToServer , "bugreport", hwnd
End Sub

Private Sub mnuShowMarker_Click()
    ShowMarker , True
End Sub

Private Sub mnuUndo_Click()
    History.Step = History.Step - 1
    txtMain = History.Text(History.Step)
    txtMain.SelStart = History.SelStart(History.Step)
    UpdateAll
End Sub

Private Sub picToolTip_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    picToolTip.Visible = False
    txtMain.SetFocus
End Sub

Private Sub tmrCheckHKs_Timer()
    If blnCancel Then Exit Sub
    CheckHotKeys
End Sub

Private Sub tmrUpdate_Timer()
    If Not blnProgrammLoaded Then Exit Sub
    If Not blnRec And Not blnExecuting Then
        UpdateCoords
        If GetForegroundWindow <> Me.hwnd Then HideToolTip
'        If GetForegroundWindow <> Me.hwnd And GetForegroundWindow <> frmMarker.hwnd Then HideMarker
    End If
End Sub

Sub UpdateAll(Optional blnIsLoad As Boolean)
    If Not blnProgrammLoaded Or Not IsMeForeground Then Exit Sub
    If blnExecuting Then
        txtMain = ScriptVal
    ElseIf Not blnRec And Not blnExecuting Then
        'UpdateMainText
        UpdateHistory
        UpdateControlsState
        UpdateStatus
        UpdateScriptInfo
        If Not blnIsLoad Then
            If (Settings.bShowListTT Or Settings.bShowTextTT) And (txtMain.SelStart <> PrevSelStart) Then UpdateToolTip
            PrevSelStart = txtMain.SelStart
            If Settings.bShowMarker Then ShowMarker FuncInLine
        End If
    End If
    If Len(txtMain) > 5 Then txtMain.ToolTipText = "Нажмите F1, чтобы получить информацию о команде"
End Sub

Sub UpdateHistory()
    If blnIsAStep Then Exit Sub
    If (Not Not History.Text) = 0 Then
        ReDim History.Text(0)
        ReDim History.SelStart(0)
    End If
    If History.Text(History.Step) <> txtMain Then
        History.Step = History.Step + 1
        ReDim Preserve History.Text(History.Step)
        ReDim Preserve History.SelStart(History.Step)
        History.SelStart(History.Step) = txtMain.SelStart
        History.Text(History.Step) = txtMain
    End If
End Sub

Sub Undo()
    If History.Step = 0 Then Exit Sub
    blnIsAStep = True
    History.Step = History.Step - 1
    txtMain.Text = History.Text(History.Step)
    txtMain.SelStart = History.SelStart(History.Step)
    blnIsAStep = False
End Sub

Sub UpdateMainText()
    Dim tmp As Variant
    tmp = txtMain
    For Each arr In ClickV()
        If InStr(1, tmp, arr) Then ReplaceOutQuotes tmp, arr, arr, atFunc
    Next
    For Each arr In WaitV()
        If InStr(1, tmp, arr) Then ReplaceOutQuotes tmp, arr, arr, atFunc
    Next
    For Each arr In PressV()
        If InStr(1, tmp, arr) Then ReplaceOutQuotes tmp, arr, arr, atFunc
    Next
    For Each arr In ExitV()
        If InStr(1, tmp, arr) Then ReplaceOutQuotes tmp, arr, arr, atFunc
    Next
    For Each arr In StartScriptV()
        If InStr(1, tmp, arr) Then ReplaceOutQuotes tmp, arr, arr, atFunc
    Next
    For Each arr In SetCursorPosV()
        If InStr(1, tmp, arr) Then ReplaceOutQuotes tmp, arr, arr, atFunc
    Next
    For Each arr In LoopV()
        If InStr(1, tmp, arr) Then ReplaceOutQuotes tmp, arr, arr, atFunc
    Next
    For Each arr In MsgBoxV()
        If InStr(1, tmp, arr) Then ReplaceOutQuotes tmp, arr, arr, atFunc
    Next
    For Each arr In InputBoxV()
        If InStr(1, tmp, arr) Then ReplaceOutQuotes tmp, arr, arr, atFunc
    Next
    For Each arr In PrintV()
        If InStr(1, tmp, arr) Then ReplaceOutQuotes tmp, arr, arr, atFunc
    Next
    For Each arr In CreateFileV()
        If InStr(1, tmp, arr) Then ReplaceOutQuotes tmp, arr, arr, atFunc
    Next
    For Each arr In WriteFileV()
        If InStr(1, tmp, arr) Then ReplaceOutQuotes tmp, arr, arr, atFunc
    Next
    For Each arr In ReadFileV()
        If InStr(1, tmp, arr) Then ReplaceOutQuotes tmp, arr, arr, atFunc
    Next
    For Each arr In SetWindowV()
        If InStr(1, tmp, arr) Then ReplaceOutQuotes tmp, arr, arr, atFunc
    Next
    For Each arr In ShowWindowV()
        If InStr(1, tmp, arr) Then ReplaceOutQuotes tmp, arr, arr, atFunc
    Next
    For Each arr In ShellV()
        If InStr(1, tmp, arr) Then ReplaceOutQuotes tmp, arr, arr, atFunc
    Next
    For Each arr In StopV()
        If InStr(1, tmp, arr) Then ReplaceOutQuotes tmp, arr, arr, atFunc
    Next
    For Each arr In SkipErrorsV()
        If InStr(1, tmp, arr) Then ReplaceOutQuotes tmp, arr, arr, atFunc
    Next
    For Each arr In ChangeResolutionV()
        If InStr(1, tmp, arr) Then ReplaceOutQuotes tmp, arr, arr, atFunc
    Next
    SetScriptText tmp
    If Me.hwnd = GetForegroundWindow Then txtMain.SetFocus
End Sub

Sub UpdateToolTip(Optional JustGetData As Boolean, Optional GetTextStartTo As Long, Optional GetTextEndTo As Long, Optional GetTextTo As String, Optional GetToolTipTypeTo As ToolTipType)
    Static PrevParamNumber As Long, PrevToolTipText As String
    Dim tmp As String, Pos(1 To 11) As Double, LineStart As Long, blnHideList As Boolean, _
    PosIndex As Long, ToolTipType As ToolTipType, StopVariantes() As String, _
    Text As String, TextStartPos As Long, line As Variant, SelStartInLine As Long, HelpList(), _
    ParamNumber As Long, ParamType As ParamType, NewSelStart As Long, MaxParams As Long, _
    ptCaretPos As POINTAPI, X As Long, Y As Long, MaxLengh As Long, PrevListIndex As Long
    If txtMain <> "" Then
        LineStart = InStrRev(txtMain, vbLf, tSelStart)
        tmp = Mid$(txtMain, LineStart + 1, tSelStart - LineStart)
        tmp = ReplaceInQuotesSTR(EscapeChars(tmp), ")", "")
        If InStr(1, ReplaceInQuotesSTR(EscapeChars(tmp), ")", ""), ")") > 0 Then ' если скобка закрыта
            lstToolTip.Visible = False
            PrevSelStart = 0
            picToolTip.Visible = False
            Exit Sub
        End If
        line = GetLineByCharNumber(txtMain, tSelStart)
        StopVariantes = Split(" , ( ) + - / \ * = ^", " ")
        For i = 1 To 10
            Pos(i) = InStrRev(tmp, StopVariantes(i)) + LineStart
        Next i
    End If
    ' **************************************************************************************************
    SelStartInLine = tSelStart - GetLargestNum(Pos)
    TextStartPos = GetLargestNum(Pos) - LineStart + 1
    If TextStartPos = 1 Then ToolTipType = tttFunc Else ToolTipType = tttParam
    For n = 1 To Len(line)
        Text = Mid$(line, TextStartPos, n)
        If FindInArraySTR(StopVariantes, Right$(Text, 1)) Or Right$(Text, 1) = vbCr Then
            If Text <> "" Then Text = Left$(Text, Len(Text) - 1)
            Exit For
        End If
    Next n
    If JustGetData Then
        GetTextStartTo = GetLargestNum(Pos)
        GetTextEndTo = GetTextStartTo + Len(Text)
        GetTextTo = Text
        GetToolTipTypeTo = ToolTipType
        Exit Sub
    End If
    ReDim HelpList(0)
    If GetCount(tmp, Chr$(34)) Mod 2 = 1 Then tmp = tmp & Chr$(34): blnHideList = True
    tmp = ReplaceInQuotesSTR(tmp, ",", "")
    ParamNumber = GetCount(tmp, ",") + 1
    'If GetCount(tmp, Chr$(34)) = 1 Then ParamNumber = 1
    If ToolTipType = tttFunc Then
        GetFunctions HelpList, Text
    Else
        ' Все параметры и желательно переменные тут
        If Left$(Text, 1) = Chr$(34) Then blnHideList = True: GoTo lSetToolTipText
        If Not Settings.bShowListTT Then GoTo lSetToolTipText
        ' ********************** Определение какого типа параметр *********************
        Select Case FuncInLine
            Case cClick
                Select Case ParamNumber
                    Case 3: ParamType = ptMouseEvent
                    Case 4: ParamType = ptMouseKey
                    Case 6: ParamType = ptYesOrNo
                End Select
            Case cMsgBox
                Select Case ParamNumber
                    Case 3: ParamType = ptMsgBoxStyle
                    Case 4: ParamType = ptYesOrNo
                End Select
            Case cShell
                Select Case ParamNumber
                    Case 2: ParamType = ptYesOrNo
                    Case 3: ParamType = ptYesOrNo
                    Case 4: ParamType = ptAppWinStyle
                End Select
            Case cShowWindow
                If ParamNumber = 2 Then ParamType = ptYesOrNo
            Case cSetWindow
                If ParamNumber = 2 Then ParamType = ptYesOrNo
            Case cStartScript
                If ParamNumber = 3 Then ParamType = ptYesOrNo
            Case cSkipErrors
                If ParamNumber = 1 Then ParamType = ptYesOrNo
            Case cKillProcess
                If ParamNumber = 2 Then ParamType = ptYesOrNo
            Case cCloseWindow
                If ParamNumber = 2 Or ParamNumber = 3 Then ParamType = ptYesOrNo
            Case cDownloadFile
                If ParamNumber = 3 Then ParamType = ptHTTPmethod
                If ParamNumber = 8 Then ParamType = ptYesOrNo
            Case cIf
                If ParamNumber = 2 Or ParamNumber = 3 Then ParamType = ptFunction
            Case cVBScript
                If ParamNumber = 3 Then ParamType = ptYesOrNo
        End Select
        ' *************************************************************
        Select Case ParamType
            Case ptYesOrNo: AddToArray HelpList, Text, "Да", "Нет"
            Case ptMouseEvent: AddToArray HelpList, Text, "Клик", "Вниз", "Вверх"
            Case ptMouseKey: AddToArray HelpList, Text, "Левая", "Правая", "Средняя"
            Case ptMsgBoxStyle: AddToArray HelpList, Text, "Информация", "Ошибка", "Критическая ошибка"
            Case ptAppWinStyle: AddToArray HelpList, Text, "Нормально", "Свернуто", "Развернуто", "Скрыто"
            Case ptHTTPmethod: AddToArray HelpList, Text, "GET", "POST", "HEAD"
            Case ptFunction: GetFunctions HelpList, Text
        End Select
     End If
    If UBound(HelpList) = 0 Then blnHideList = True: GoTo lSetToolTipText
    If FindInArray(HelpList, Text) And Text <> "" Then blnHideList = True: GoTo lSetToolTipText
    PrevListIndex = lstToolTip.ListIndex
    lstToolTip.Clear
    For i = 1 To UBound(HelpList)
        lstToolTip.AddItem (HelpList(i))
        If MaxLengh < Len(HelpList(i)) Then MaxLengh = Len(HelpList(i))
    Next i
lSetToolTipText:
    If MaxLengh < 3 Then MaxLengh = 3
    If PrevSelStart <> txtMain.SelStart Then lstToolTip.ListIndex = 0 Else lstToolTip.ListIndex = PrevListIndex
    NewSelStart = GetLargestNum(Pos, PosIndex)
    GetCaretPos ptCaretPos
    If ToolTipType = tttFunc Then ptCaretPos.X = 0
    X = ptCaretPos.X + txtMain.Left
    Y = ptCaretPos.Y + TEXTBOX_LINE_HEIGHT
    lstToolTip.Left = X
    lstToolTip.Top = Y
    If lstToolTip.Top + ((lstToolTip.ListCount + 1) * LISTBOX_ITEM_HEIGHT) < Me.ScaleHeight Then
        lstToolTip.Width = GetTextWidth(lstToolTip, Space$(MaxLengh)) + 7
        lstToolTip.Height = (lstToolTip.ListCount + 1) * LISTBOX_ITEM_HEIGHT
    Else
        lstToolTip.Width = GetTextWidth(lstToolTip, Space$(MaxLengh)) + GetScrollWidth
        lstToolTip.Height = Me.ScaleHeight - lstToolTip.Top - 5
    End If
    If Not blnHideList And Settings.bShowListTT Then
        lstToolTip.Visible = True
        lstToolTip.ZOrder 0
    Else
        lstToolTip.Visible = False
        PrevSelStart = 0
    End If
    '  *****************************************************************************  PIC TEXT TOOL TIP *********************************
    Dim FuncCaption As String, ParamsStart As Long, ToolTipText As String, OptStartFrom As Long
    ParamsStart = InStr(1, tmp, "(")
    If ParamsStart > 0 Then
        FuncCaption = Mid$(tmp, 1, ParamsStart - 1)
    End If
    If FuncInLine > 0 And ToolTipType = tttParam And ParamsStart > 0 Then
        Call SetToolTipText(ToolTipText, OptStartFrom)
        MaxParams = GetCount(ToolTipText, ",") + 1
        'If ParamNumber > MaxParams Then ParamNumber = MaxParams: blnBadParams = True
        ToolTipText = FuncCaption & "(" & ToolTipText & ")"
        If ToolTipText = PrevToolTipText And ParamNumber = PrevParamNumber And picToolTip.Visible Then Exit Sub
        PrevToolTipText = ToolTipText: PrevParamNumber = ParamNumber
        picToolTip.Left = txtMain.Left '- 0.5
        picToolTip.Top = ptCaretPos.Y + TEXTBOX_LINE_HEIGHT
        DrawToolTip picToolTip, ToolTipText, ParamNumber, OptStartFrom
        If Settings.bShowTextTT Then picToolTip.Visible = True
    Else
        picToolTip.Visible = False
    End If
    '  ***************************************************************************************************************************************
Exit Sub
lHide:
    lstToolTip.Visible = False
    PrevSelStart = 0
End Sub

Sub SetToolTipText(ByRef GetTo As String, ByRef OptStartFrom As Long)
        Select Case FuncInLine
            Case cClick: GetTo = "X, Y, событие, кнопка мыши, кол-во повторов, вернуть курсор на место, ожидать до, ожидать после, ограничение по x, ограничение по y"
            Case cPress: GetTo = "клавиша*, количество повторов, ожидать до, ожидать после"
            Case cWait: GetTo = "время ожидания, интервал рандомизации"
            Case cExit: GetTo = ""
            Case cStartScript: GetTo = "имя скрипта*, кол-во повторов, ждать завершения"
            Case cSetCursorPos: GetTo = "X*, Y*, ожидать до, ожидать после"
            Case cLoop: GetTo = "кол-во повторов, на какую строку прыгнуть, задержка перед прыжком"
            Case cMsgBox: GetTo = "текст сообщения*, заголовок, стиль сообщения, поверх всех окон"
            Case cInputBox: GetTo = "текст, заголовок, переменная для получения информации*"
            Case cPrint: GetTo = "текст отладки*"
            Case cCreateFile: GetTo = "имя файла*, содержимое"
            Case cWriteFile: GetTo = "имя файла*, содержимое"
            Case cReadFile: GetTo = "имя файла*, переменная для получения информации, начиная с какой строки, количество строк"
            Case cSetWindow: GetTo = "заголовок, если окно не обнаружено показать ошибку"
            Case cShowWindow: GetTo = "заголовок, если окно не обнаружено показать ошибку"
            Case cShell: GetTo = "полное имя программы*, дожидаться появления окна, дожидаться завершения процесса, стиль окна при запуске"
            Case cStop: GetTo = ""
            Case cSkipErrors: GetTo = "да или нет"
            Case cChangeResolution: GetTo = "ширина, высота"
            Case cDeleteFile: GetTo = "путь к файлу*"
            Case cKillProcess: GetTo = "имя процесса*, принудительно"
            Case cDeleteDir: GetTo = "путь к папке*"
            Case cCreateDir: GetTo = "путь к папке*"
            Case cCloseWindow: GetTo = "заголовок окна*, усиленно, если окно не обнаружено показать ошибку"
            Case cDownloadFile: GetTo = "URL адрес*, путь к файлу или имя переменной, метод, заголовки через &&, POST-данные, Cookies, пееременная для получаения ответных Cookie, включить редирект"
            Case cIf: GetTo = "условие*, если условие верно — функции через &&*, если условие неверно"
            Case cVBScript: GetTo = "код VBScript*, переменная для получения результатов, является ли код новой функцией"
            Case cGoTo: GetTo = "номер строки*"
        End Select
        OptStartFrom = GetCount(GetTo, "*") + 1
        GetTo = Replace(GetTo, "*", "")
        GetTo = Replace(GetTo, ", ", ",")
End Sub

Sub UpdateControlsState(Optional DisableControls As Boolean)
    On Error Resume Next
    If FuncInLine = cClick Or FuncInLine = cSetCursorPos Then
        mnuShowMarker.Visible = True: s13.Visible = True
    Else
        mnuShowMarker.Visible = False: s13.Visible = False
    End If
    If IsValidScript(txtMain) Then
        mnuExecute = True
        mnuBuildExe = True
    Else
        mnuExecute.Enabled = False
        mnuBuildExe = False
    End If
    If blnExecuting Then
        mnuExecute = True
        mnuRec = False
    ElseIf blnRec Then
        mnuExecute = False
        mnuRec = True
    End If
    If blnExecuting Or blnRec Or DisableControls Then
        mnuFile = False
        mnuEdit = False
        mnuSettings = False
        mnuInfo = False
        mnuInsts = False
    Else
        mnuFile = True
        mnuEdit = True
        mnuSettings = True
        mnuInfo = True
        mnuInsts = True
    End If
    mnuUndo = History.Step
    If (Not Not History.Text) <> 0 Then
        If History.Step = UBound(History.Text) Then mnuRedo.Enabled = False Else mnuRedo.Enabled = True
    Else
        mnuRedo.Enabled = False
    End If
    If txtMain = "" Then
        mnuFind = False
        mnuFindNext = False
        mnuReplace = False
        mnuSelectAll = False
        mnuCleanAll = False
    Else
        mnuFind = True
        mnuFindNext = True
        mnuReplace = True
        mnuSelectAll = True
        mnuCleanAll = True
    End If
    mnuCopy = txtMain.SelLength
    mnuCut = txtMain.SelLength
    mnuDel = txtMain.SelLength Or (Len(txtMain) > txtMain.SelStart)
    mnuPaste = Len(Clipboard.GetText(1)) > 0
End Sub

Sub UpdateCoords()
    Dim X As Long, Y As Long
    GetCursorPos X, Y
    txtCoords = "X:" & X & "; Y:" & Y
End Sub

Sub UpdateStatus()
    Dim Point As POINTAPI
    SelStartLine = GetLineNumberByCharNumber(txtMain, txtMain.SelStart)
    FuncInLine = GetFunc(GetScriptLine, True)
    PrevLine = SelStartLine
    txtStatus = "Строка: " & SelStartLine '& " || " & "Символ: " & txtMain.SelStart
    If FuncInLine = cClick Or FuncInLine = cSetCursorPos Then
        Point = GetPoint
        If Point.X > 0 And Point.Y > 0 Then txtStatus = txtStatus & vbCrLf & "Координаты на текущей строке: X=" & Point.X & "; Y=" & Point.Y
    End If
    UpdateCoords
'    Static antirecurse As Boolean
'    If antirecurse Then Exit Sub
'    antirecurse = True
'    LinesCount = GetLinesCount(txtMain)
'    If LinesCount > 900 Then StopInput
'    SelStartLine = GetLineNumberByCharNumber(txtMain, txtMain.SelStart)
'    Caption = SelStartLine
'    'If SelStartLine >= MaxLines Then
'        frmNum.Top = -JustPositive(SelStartLine - MaxLines) * 252.45
'    'End If
'    frmNum.Height = LinesCount * 252.45
'    For i = 1 To LinesCount
'        If i > num.UBound Then Load num(i)
'        num(i).Caption = CStr(i)
'        num(i).Top = (i - 1) * 252.45
'        num(i).Left = frmNum.Width - num(i).Width - 20
'        num(i).Visible = True
'    Next i
'    antirecurse = False
End Sub

Sub StopInput()
    Dim Last_CrLf As Long, strMain As String
    strMain = txtMain
    Last_CrLf = 1
    For n = 1 To 900
        Last_CrLf = InStr(Last_CrLf + 1, strMain, vbCrLf)
    Next n
    If Last_CrLf = 0 Then Last_CrLf = Len(txtMain)
    txtMain = Left(txtMain, Last_CrLf - 2)
    txtMain.SelStart = Len(txtMain)
    LinesCount = 900
End Sub

Private Sub txtCoords_GotFocus()
    HideCaret txtCoords.hwnd
End Sub

Private Sub txtCoords_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    HideCaret txtCoords.hwnd
End Sub

Private Sub txtMain_Change()
    If Not blnUpdateTextOff Then UpdateAll
End Sub

Private Sub txtMain_Click()
    frmFind.SetFalsePressed
    If Not blnUpdateTextOff Then UpdateAll
End Sub

Sub UpdateFuncInfo(Optional IsClick As Boolean)
'    Dim Pos As Long, LineStart As Long, LineEnd As Long, Script As String, line As String, Func As String
'    Script = txtMain
'    Script = vbCrLf & Script & vbCrLf
'    Pos = txtMain.SelStart + 2
'    LineStart = InStrRev(Script, vbCrLf, Pos)
'    LineEnd = InStr(Pos, Script, vbCrLf)
'    line = Mid(Script, LineStart + 2, LineEnd - LineStart - 2)
'    ConvertScript line
'    Func = GetFunc(line)
'    If Len(line) < 4 Then GoTo ShowFunc Else frmHelp.ShowHelp typParam, Func
'    UpdateMarker line, IsClick
'    txtMain.ToolTipText = line
'    Exit Sub
'ShowFunc:
'    frmHelp.ShowHelp typFunc, "ShowFuncs" ' параметр передается не как действие, а как текст для NowHelp, чтобы хелп по 100 раз не перерисовывался
End Sub

Sub SelectFunc(Optional ByVal Item As String, Optional UpdateOff As Boolean, Optional IsDelimiter As Boolean)
    Dim TextStart As Long, TextEnd As Long, ToolTipType As ToolTipType, strLine As String, lngMinusVal As Long
    If lstToolTip.Text = "" Then Exit Sub
    If Item = "" Then Item = lstToolTip.Text
    UpdateToolTip True, TextStart, TextEnd, , ToolTipType
    strLine = GetLineByCharNumber(txtMain, txtMain.SelStart)
    If ToolTipType = tttFunc And Mid$(txtMain, TextEnd + 1, 1) <> " " Then
        If InStr(1, strLine, "(") = 0 Then
            If FindInArray(FuncsFirstTextParam(), GetFunc(Item, True)) Then ' ЕСЛИ параметр текстовый
                Item = Item & "(" & Chr$(34) & Chr$(34) & ")"
                lngMinusVal = 2
            Else
                Item = Item & "()"
                lngMinusVal = 1
            End If
        Else
            lngMinusVal = -1
        End If
    ElseIf IsDelimiter Then
        Item = Item & ","
    End If
    SetScriptText AccurateReplace(txtMain, Item, TextStart + 1, TextEnd), TextStart + Len(Item) - lngMinusVal
    If Not blnUpdateTextOff And Not UpdateOff Then UpdateAll
End Sub

Sub ChangeFuncParametrs(ByVal bIsMouse As Boolean, ParamArray Params())
    On Error GoTo EndSub
    Static PrevUpdateTime As Long
    If GetTickCount - PrevUpdateTime < 10 Then Exit Sub  ' Каждые 10 мс задаю
    PrevUpdateTime = GetTickCount
    Dim LineStart As Long, LineEnd As Long, line As String, buff As String, FullParams() As String, i As Long
    If bIsMouse Then If FuncInLine <> cClick And FuncInLine <> cSetCursorPos Then Exit Sub
    line = GetLineByCharNumber(txtMain, tSelStart, LineStart, LineEnd)
    line = ReplaceInQuotesSTR(line, "(", "")
    line = ReplaceInQuotesSTR(line, ")", "")
    buff = fMid(line, InStr(1, line, "(") + 1, InStr(1, line, ")"))
    FullParams = Split(buff, ",")
    If UBound(FullParams) < UBound(Params) Then ReDim FullParams(UBound(Params))
    For i = 0 To UBound(Params)
        If bIsMouse And i = 1 And Left$(FullParams(i), 1) = " " Then Params(i) = " " & Params(i)
        FullParams(i) = Params(i)
    Next i
    SetScriptText AccurateReplace(txtMain, Join(FullParams, ","), LineStart + InStr(1, line, "("), LineEnd - IIf(LineEnd = Len(txtMain), 1, 2)), LineStart
EndSub:
    UpdateStatus
End Sub

Sub InsertFunc(ByVal sFunc As String)
    Dim sText As String, lStart As Long, lLineEnd As Long, IsCleanFunc As Long
    lStart = txtMain.SelStart
    sText = txtMain
    If Right(sFunc, 1) <> ")" Then sFunc = sFunc & "()": IsCleanFunc = 1
    If sText = "" Then
        SetScriptText sFunc, Len(sFunc) - 1
        UpdateAll
        Exit Sub
    End If
    If Mid$(sText, lStart + 1, 1) <> vbCr And lStart < Len(sText) Then
        GetLineByCharNumber sText, lStart, , lLineEnd
        lStart = IIf(lLineEnd = Len(sText), lLineEnd, lLineEnd - 1)
    End If
    If Mid$(sText, lStart, 1) = vbLf And Mid$(sText, lStart + 1, 1) = vbCr Then ' ЕСЛИ пустая строка и после нее еще одна пустая строка
        SetScriptText Left$(sText, lStart) & sFunc & Right$(sText, Len(sText) - lStart), lStart + InStr(1, sFunc, ")") - IsCleanFunc
    ElseIf Mid$(sText, lStart, 1) = vbLf And lStart = Len(sText) Then
        SetScriptText Left$(sText, lStart) & sFunc & Right$(sText, Len(sText) - lStart), lStart + InStr(1, sFunc, ")") - IsCleanFunc ' ЕСЛИ пустая строка и после нее ничего нет
    ElseIf Mid$(sText, lStart, 1) <> vbLf And lStart = Len(sText) Then ' ЕСЛИ строка не пуста и после нее ничего нет
        sFunc = vbCrLf & sFunc
        SetScriptText Left$(sText, lStart) & sFunc & Right$(sText, Len(sText) - lStart), lStart + InStr(1, sFunc, ")") - IsCleanFunc
    Else ' Все остальное: если строка не пуста и после неё что-то есть
        sFunc = vbCrLf & sFunc
        SetScriptText Left$(sText, lStart) & sFunc & Right$(sText, Len(sText) - lStart), lStart + InStr(1, sFunc, ")") - IsCleanFunc
    End If
    UpdateAll
End Sub

Private Sub txtMain_DblClick()
    If Not blnUpdateTextOff Then UpdateAll
End Sub

Private Sub txtMain_KeyDown(KeyCode As Integer, Shift As Integer)
    CheckHotKeys
    If blnRec Or blnExecuting Then KeyCode = 0: Exit Sub
    If lstToolTip.Visible Then
        Select Case KeyCode
            Case vbKeyUp
                If lstToolTip.ListIndex > 0 Then lstToolTip.ListIndex = lstToolTip.ListIndex - 1
                KeyCode = 0
            Case vbKeyDown
                If lstToolTip.ListIndex < (frmMain.lstToolTip.ListCount - 1) Then lstToolTip.ListIndex = lstToolTip.ListIndex + 1
                KeyCode = 0
            Case vbKeyReturn
                If lstToolTip.ListIndex = -1 Then Exit Sub
                blnUpdateTextOff = True
                SelectFunc , True
                KeyCode = 0
            Case vbKeySpace
                If Not SpaceEqv Then Exit Sub
                blnUpdateTextOff = True
                SelectFunc , True
                KeyCode = 0
            Case vbKeyEscape
                'HideToolTipAtOnce
                KeyCode = 0
        End Select
    End If
End Sub

Function SpaceEqv() As Boolean
    Dim CrLfStartPos As Long, Text As String
    CrLfStartPos = InStrRev(txtMain, vbCrLf, tSelStart)
    If CrLfStartPos = 0 Then CrLfStartPos = -1
    Text = Mid$(txtMain, CrLfStartPos + 2, tSelStart - CrLfStartPos)
    If GetCount(Text, " ") = GetCount(lstToolTip.Text, " ") Or Mid$(lstToolTip.Text, Len(Text) + 1, 1) <> " " Then SpaceEqv = True
End Function

Private Sub txtMain_KeyPress(KeyAscii As Integer)
    If lstToolTip.Visible Then ' Тут пока только работа с тултипом
        Select Case KeyAscii
            Case vbKeyReturn
                If lstToolTip.ListIndex = -1 Then Exit Sub
                KeyAscii = 0
                UpdateAll
            Case vbKeySpace
                If Not SpaceEqv Then Exit Sub
                KeyAscii = 0
                UpdateAll
            Case Asc(",")
                If FuncInLine > 0 Then
                    SelectFunc , , True
                    KeyAscii = 0
                End If
            Case Asc("(")
                SelectFunc , , True
                KeyAscii = 0
        End Select
'    ElseIf KeyAscii = vbKeyReturn Then ' ************************ Временно убрал ф-цию автомат закрытия скобки ***********
'        If FuncInLine > 0 Then
'            If Mid$(txtMain, txtMain.SelStart, 1) <> ")" Then
'                CloseBracket
'                KeyAscii = 0
'            End If
'        End If
    End If
    blnUpdateTextOff = False
End Sub

Sub CloseBracket()
    Dim LineStart As Long, LineEnd As Long, line As String
    line = GetScriptLine(LineStart, LineEnd)
    If Mid$(GetScript(line), Len(FuncInLine) + 1, 1) = "(" And (Mid$(txtMain, LineEnd, 1) = vbCr Or LineEnd = Len(txtMain)) Then line = line & ")"
    line = line & vbCrLf
    SetScriptText AccurateReplace(txtMain, line, LineStart, LineEnd - IIf(LineEnd = Len(txtMain), 0, 1)), txtMain.SelStart + 3
    UpdateAll
End Sub

Sub SetScriptText(ByVal Text As String, Optional NewSelStart As Long = -1, Optional NewSelLengh As Long = -1)
    On Error Resume Next
    If txtMain = Text Then Exit Sub
    blnUpdateTextOff = True
    If NewSelStart = -1 Then NewSelStart = txtMain.SelStart
    If NewSelLengh = -1 Then NewSelLengh = txtMain.SelLength
    txtMain.Enabled = False
    txtMain = Text
    txtMain.SelStart = NewSelStart
    txtMain.SelLength = NewSelLengh
    txtMain.Enabled = True
    txtMain.SetFocus
    blnUpdateTextOff = False
End Sub

Function GetScriptLine(Optional GetLineStartTo As Long, Optional GetLineEndTo As Long) As String
    GetScriptLine = GetLineByCharNumber(txtMain, txtMain.SelStart, GetLineStartTo, GetLineEndTo)
End Function

Private Sub txtMain_KeyUp(KeyCode As Integer, Shift As Integer)
    If Not lstToolTip.Visible Then
        If KeyCode = vbKeyUp Or KeyCode = vbKeyDown Then
            UpdateAll
        End If
    End If
End Sub

Private Sub txtMain_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    LoadScript Data.Files.Item(1)
End Sub

Private Sub txtStatus_GotFocus()
    HideMarker
    HideToolTip
End Sub
