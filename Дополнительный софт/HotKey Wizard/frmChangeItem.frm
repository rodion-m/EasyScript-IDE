VERSION 5.00
Begin VB.Form frmChangeItem 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Создать комбинацию клавиш быстрого доступа"
   ClientHeight    =   2955
   ClientLeft      =   -15
   ClientTop       =   270
   ClientWidth     =   6240
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   204
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmChangeItem.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2955
   ScaleWidth      =   6240
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdAutoAdd 
      Caption         =   "Автоматическое заполнение списка"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   18
      ToolTipText     =   "Открыть окно автоматического заполнения списка"
      Top             =   2520
      Width           =   3015
   End
   Begin VB.CommandButton cmdSetCaptionByFileTitle 
      Height          =   255
      Left            =   5760
      TabIndex        =   16
      TabStop         =   0   'False
      ToolTipText     =   "Задать имя по заголовку скрипта"
      Top             =   120
      Width           =   375
   End
   Begin VB.Frame Frame1 
      Caption         =   "Действие если скрипт уже работает"
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   11.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   15
      Top             =   1560
      Width           =   6015
      Begin VB.ComboBox cmbIfAlreadyRuning 
         Height          =   360
         ItemData        =   "frmChangeItem.frx":058A
         Left            =   120
         List            =   "frmChangeItem.frx":058C
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   360
         Width           =   5820
      End
   End
   Begin VB.TextBox txtTitle 
      Alignment       =   2  'Центровка
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   285
      Left            =   1320
      TabIndex        =   1
      Text            =   "Выберите скрипт"
      Top             =   120
      Width           =   4335
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Отмена"
      Height          =   375
      Left            =   5040
      TabIndex        =   9
      Top             =   2520
      Width           =   1095
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Создать"
      Default         =   -1  'True
      Height          =   375
      Left            =   3840
      TabIndex        =   8
      Top             =   2520
      Width           =   1095
   End
   Begin VB.ComboBox Key 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   4
      ItemData        =   "frmChangeItem.frx":058E
      Left            =   5130
      List            =   "frmChangeItem.frx":0590
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   1080
      Width           =   1022
   End
   Begin VB.ComboBox Key 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   1
      ItemData        =   "frmChangeItem.frx":0592
      Left            =   1320
      List            =   "frmChangeItem.frx":0594
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   1080
      Width           =   1022
   End
   Begin VB.ComboBox Key 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   3
      ItemData        =   "frmChangeItem.frx":0596
      Left            =   3840
      List            =   "frmChangeItem.frx":0598
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   1080
      Width           =   1022
   End
   Begin VB.ComboBox Key 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   2
      ItemData        =   "frmChangeItem.frx":059A
      Left            =   2580
      List            =   "frmChangeItem.frx":059C
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   1080
      Width           =   1022
   End
   Begin VB.CommandButton mnuOpen 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5760
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   600
      Width           =   375
   End
   Begin VB.TextBox txtFileName 
      Alignment       =   2  'Центровка
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   285
      Left            =   1320
      MousePointer    =   1  'Указатель
      TabIndex        =   2
      Text            =   "Выберите скрипт"
      ToolTipText     =   "Кликните 2 раза чтобы открыть окно выбора"
      Top             =   600
      Width           =   4335
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Прозрачно
      Caption         =   "Название:"
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   12
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   5
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1140
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Прозрачно
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   13.5
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   2
      Left            =   4920
      TabIndex        =   14
      Top             =   1080
      Width           =   150
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Прозрачно
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   13.5
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   4
      Left            =   3630
      TabIndex        =   13
      Top             =   1080
      Width           =   150
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Прозрачно
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   13.5
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   3
      Left            =   2370
      TabIndex        =   12
      Top             =   1080
      Width           =   150
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Прозрачно
      Caption         =   "Клавиши:"
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   12
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   1
      Left            =   120
      TabIndex        =   11
      Top             =   1080
      Width           =   1125
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Прозрачно
      Caption         =   "Скрипт:"
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   12
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   0
      Left            =   120
      TabIndex        =   10
      Top             =   600
      Width           =   930
   End
End
Attribute VB_Name = "frmChangeItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Mode As ChangeHotKeyModes, blnStaticTitle As Boolean, Index As Long, FormLoaded As Boolean

Sub ChangeItem(ByVal pMode As ChangeHotKeyModes)
    Dim sKeys() As String
    Mode = pMode
    Select Case Mode
        Case chkmCreate
            Index = 0
            Me.Caption = "Создать комбинацию клавиш быстрого доступа"
            cmdSave.Caption = "Создать"
            sKeys = Split(GetFreeHotKey, "+")
            If UBound(sKeys) < 3 Or UBound(sKeys) > 3 Then ReDim Preserve sKeys(0 To 3)
            SetKeys sKeys(0), sKeys(1), sKeys(2), sKeys(3)
        Case chkmChange
            Index = GetIndex(GetSelectedItem) ' изменения идут по номеру
            txtTitle.ForeColor = &H0&
            txtFileName.ForeColor = &H0&
            Me.Caption = "Изменить комбинацию клавиш быстрого доступа"
            cmdSave.Caption = "Изменить"
            UpdateSubItemsArray GetSelectedItem
            txtFileName = SubItemsArray(COLUMN_FILENAME)
            If GetFileBaseTitle(txtFileName) <> SubItemsArray(COLUMN_TITLE) Then
                txtTitle = SubItemsArray(COLUMN_TITLE)
                blnStaticTitle = True
            End If
            sKeys = Split(SubItemsArray(COLUMN_HKS), "+")
            If UBound(sKeys) < 3 Or UBound(sKeys) > 3 Then ReDim Preserve sKeys(0 To 3)
            SetKeys sKeys(0), sKeys(1), sKeys(2), sKeys(3)
            If SubItemsArray(COLUMN_IFRUNING) <> "" Then cmbIfAlreadyRuning.Text = SubItemsArray(COLUMN_IFRUNING)
    End Select
    Me.Show vbModal, frmMain
End Sub

Private Sub cmdAutoAdd_Click()
    ShowFormEx frmAutoAdd, vbModal, frmMain, Me
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
    Dim sKeys As String
    For i = 1 To 4
        If Key(i) <> "нет" Then
            If sKeys = "" Then sKeys = Key(i) Else sKeys = sKeys & "+" & Key(i)
        End If
    Next i
    SortKeys sKeys
    If sKeys = "" Then
        If MsgBox("Комбинация клавиш не задана. Задать свободную горячую клавишу автоматически?", vbQuestion + vbYesNo, "") = vbYes Then
            sKeys = GetFreeHotKey
        Else
            Key(1).SetFocus
            Exit Sub
        End If
    End If
    If Not IsFileExists(txtFileName) Then
        ShowNotFound txtFileName
        Exit Sub
    End If
    If txtTitle = "" Then txtTitle = GetFileBaseTitle(txtFileName)
    If IsHotKeyExists(sKeys) <> Index Or frmMain.mnuEqvHotKeys.Checked Then
        If MsgBox("Такая комбинация клавиш уже используется. Задать свободную горячую клавишу автоматически?", vbQuestion + vbYesNo, "") = vbYes Then
            sKeys = GetFreeHotKey
        Else
            Exit Sub
        End If
    End If
    AddScript txtTitle, txtFileName, sKeys, cmbIfAlreadyRuning.Text, Index
    frmMain.UpdateControls
    Unload Me
End Sub

Private Sub cmdSetCaptionByFileTitle_Click()
    If txtFileName <> "Выберите скрипт" Then
        txtTitle = GetFileBaseTitle(txtFileName)
    End If
    blnStaticTitle = False
End Sub

Private Sub Form_Activate()
    If Index = 0 Then ShowOpen
    FormLoaded = True
End Sub

Private Sub Form_Load()
    AddKeysToComboBox Key(1), Key(2), Key(3), Key(4)
    cmbIfAlreadyRuning.AddItem STR_Just_Kill
    cmbIfAlreadyRuning.AddItem STR_Nothing
    cmbIfAlreadyRuning.AddItem STR_Kill_Run
    cmbIfAlreadyRuning.AddItem STR_Just_Run
    cmbIfAlreadyRuning.ItemData(0) = aeeJust_Kill
    cmbIfAlreadyRuning.ItemData(1) = aeeNothing
    cmbIfAlreadyRuning.ItemData(2) = aeeKill_Run
    cmbIfAlreadyRuning.ItemData(3) = aeeJust_Run
    
    cmbIfAlreadyRuning.Text = STR_Just_Kill
End Sub

Private Sub Key_Change(Index As Integer)
    If Key(Index).Text = "" Then Key(Index).Text = "нет"
End Sub

Private Sub mnuOpen_Click()
    ShowOpen
End Sub

Private Sub txtFileName_Change()
    With txtTitle
        If .ForeColor = &H808080 Then
            .ForeColor = &H0&
            .MousePointer = 0
        End If
        If .Text = "" Then blnStaticTitle = False
        If Not blnStaticTitle Then .Text = GetFileBaseTitle(txtFileName)
    End With
End Sub

Private Sub txtFileName_DblClick()
    ShowOpen
End Sub

Private Sub txtFileName_GotFocus()
    With txtFileName
        If .ForeColor = &H808080 Then
            .Text = ""
            .ForeColor = &H0&
            .MousePointer = 0
        End If
    End With
End Sub

Private Sub txtTitle_Click()
    If Not FormLoaded Then Exit Sub
    With txtTitle
        If .ForeColor = &H808080 Then
            .Text = ""
            .ForeColor = &H0&
            .MousePointer = 0
        End If
    End With
End Sub

Private Sub txtTitle_GotFocus()
    If Not FormLoaded Then Exit Sub
    With txtTitle
        If .ForeColor = &H808080 Then
            .Text = ""
            .ForeColor = &H0&
            .MousePointer = 0
        End If
    End With
End Sub

Private Sub txtTitle_KeyPress(KeyAscii As Integer)
    blnStaticTitle = True
End Sub

Sub CheckFile()
    
End Sub

Sub ShowOpen()
    Dim TempLoaded As typeFileStruct
    If IsFileExists(txtFileName) Then InitDir = GetPath(txtFileName)
    If InitDir = "" Then InitDir = ScriptsPath
    TempLoaded = OpenFile("Выберите скрипт", InitDir, txtFileName, "Скрипты EasyScript (*.es;*.exe)" & Chr$(0) & "*.es;*.exe" & Chr$(0) & "Все файлы (*.*)" & Chr$(0) & "*.*" & Chr$(0), Me.hwnd)
    If TempLoaded.FileName <> "" Then
        txtFileName = TempLoaded.FileName
        InitDir = TempLoaded.Path
        With txtFileName
            If .ForeColor = &H808080 Then
                .ForeColor = &H0&
                .MousePointer = 0
            End If
        End With
    End If
End Sub

Sub SetKeys(ParamArray sKeys())
    For i = 0 To UBound(sKeys)
        If sKeys(i) = "" Then sKeys(i) = "нет"
        Key(i + 1).Text = sKeys(i)
    Next
End Sub
