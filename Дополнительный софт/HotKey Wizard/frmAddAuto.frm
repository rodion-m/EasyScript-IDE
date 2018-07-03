VERSION 5.00
Begin VB.Form frmAutoAdd 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Параметры автоматического заполнения списка"
   ClientHeight    =   2910
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5445
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2910
   ScaleWidth      =   5445
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton mnuOpen 
      Caption         =   "..."
      Height          =   285
      Left            =   5040
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   120
      Width           =   375
   End
   Begin VB.CommandButton cmdGo 
      Caption         =   "Начать"
      Default         =   -1  'True
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
      Left            =   3120
      TabIndex        =   5
      ToolTipText     =   "zxc"
      Top             =   2520
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Отмена"
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
      Left            =   4320
      TabIndex        =   6
      Top             =   2520
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "Действие если скрипт уже запущен"
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
      Index           =   0
      Left            =   0
      TabIndex        =   13
      Top             =   1560
      Width           =   5415
      Begin VB.ComboBox cmbIfAlreadyRuning 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "frmAddAuto.frx":0000
         Left            =   120
         List            =   "frmAddAuto.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   360
         Width           =   5220
      End
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
      ItemData        =   "frmAddAuto.frx":0004
      Left            =   2580
      List            =   "frmAddAuto.frx":0006
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   600
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
      ItemData        =   "frmAddAuto.frx":0008
      Left            =   3840
      List            =   "frmAddAuto.frx":000A
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   600
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
      ItemData        =   "frmAddAuto.frx":000C
      Left            =   1320
      List            =   "frmAddAuto.frx":000E
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   600
      Width           =   1022
   End
   Begin VB.TextBox txtPath 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   1320
      MousePointer    =   1  'Указатель
      TabIndex        =   0
      Top             =   120
      Width           =   3615
   End
   Begin VB.Label lblExample 
      Appearance      =   0  'Плоска
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Фиксировано один
      Caption         =   "sample"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   120
      TabIndex        =   14
      ToolTipText     =   "Пример того, какие горячие клавиши будут заданы автоматически"
      Top             =   1080
      Width           =   5295
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Прозрачно
      BorderStyle     =   1  'Фиксировано один
      Caption         =   "#"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   11.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   5160
      TabIndex        =   12
      ToolTipText     =   "1-9, A-Z"
      Top             =   600
      Width           =   195
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
      Top             =   600
      Width           =   1125
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
      TabIndex        =   10
      Top             =   600
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
      TabIndex        =   9
      Top             =   600
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
      Index           =   2
      Left            =   4920
      TabIndex        =   8
      Top             =   600
      Width           =   150
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Прозрачно
      Caption         =   "Папка:"
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
      TabIndex        =   7
      ToolTipText     =   "Введите путь к скриптам"
      Top             =   120
      Width           =   780
   End
End
Attribute VB_Name = "frmAutoAdd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdGo_Click()
    Dim sKeys As String
    For i = 1 To 3
        If Key(i) <> "нет" Then
            If sKeys = "" Then sKeys = Key(i) Else sKeys = sKeys & "+" & Key(i)
        End If
    Next i
    If IsFileMaskExistsInDir(txtPath) Then
        If AutoAdd(txtPath, sKeys, cmbIfAlreadyRuning.Text) Then
            Hide
            MsgBox "Операция успешно завершена", vbInformation, "Операция выполнена"
            Unload Me
        End If
    Else
        MsgBox "Скрипты не обнаружены. Пожалуйста, проверьте правильно ли введен путь.", vbExclamation, ""
    End If
End Sub

Private Sub Form_Load()
    AddKeysToComboBox Key(1), Key(2), Key(3)
    cmbIfAlreadyRuning.AddItem STR_Just_Kill
    cmbIfAlreadyRuning.AddItem STR_Nothing
    cmbIfAlreadyRuning.AddItem STR_Kill_Run
    cmbIfAlreadyRuning.AddItem STR_Just_Run
    cmbIfAlreadyRuning.ItemData(0) = aeeJust_Kill
    cmbIfAlreadyRuning.ItemData(1) = aeeNothing
    cmbIfAlreadyRuning.ItemData(2) = aeeKill_Run
    cmbIfAlreadyRuning.ItemData(3) = aeeJust_Run
    cmbIfAlreadyRuning.Text = STR_Just_Kill
    SetKeys "Ctrl", "", ""
    txtPath = ScriptsPath
    UpdateExample
End Sub

Sub SetKeys(ParamArray sKeys())
    For i = 0 To UBound(sKeys)
        If sKeys(i) = "" Then sKeys(i) = "нет"
        Key(i + 1).Text = sKeys(i)
    Next i
End Sub

Sub UpdateExample()
    Dim sKeys As String, i As Long
    For i = 1 To 3
        If Key(i) <> "нет" Then
            If sKeys = "" Then sKeys = Key(i) Else sKeys = sKeys & "+" & Key(i)
        End If
    Next i
    lblExample = "Скрипт 1: [" & sKeys & "+1]" & "  ; Скрипт 2: [" & sKeys & "+2]"
End Sub

Private Sub Key_Click(Index As Integer)
    UpdateExample
End Sub

Private Sub Key_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    UpdateExample
End Sub

Private Sub Label1_DblClick()
    MsgBox "Данное поле не предназначено для изменения. Если Вам не ясно его назначение, нажмите 'Начать'.", vbInformation
End Sub

Private Sub mnuOpen_Click()
    txtPath = SelectDir(txtPath)
End Sub

Private Sub txtPath_Change()
    txtPath.ToolTipText = txtPath.Text
End Sub

Private Sub txtPath_DblClick()
    txtPath = SelectDir(txtPath)
End Sub
