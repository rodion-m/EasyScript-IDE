VERSION 5.00
Begin VB.Form frmSSettings 
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Настройки скрипта"
   ClientHeight    =   3930
   ClientLeft      =   30
   ClientTop       =   375
   ClientWidth     =   10185
   BeginProperty Font 
      Name            =   "Book Antiqua"
      Size            =   9.75
      Charset         =   204
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "SSettings"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3930
   ScaleWidth      =   10185
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Настройки мыши"
      Height          =   1980
      Left            =   5160
      TabIndex        =   11
      Top             =   0
      Width           =   5000
      Begin VB.ComboBox cmbMouseMode 
         Height          =   375
         ItemData        =   "frmSSettings.frx":0000
         Left            =   120
         List            =   "frmSSettings.frx":0010
         Style           =   2  'Dropdown List
         TabIndex        =   21
         Top             =   1440
         Width           =   4740
      End
      Begin VB.CheckBox chkDelayDownUp 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Увеличенная задержка между событиями Down и Up (если клики работают некорректно)"
         Height          =   735
         Left            =   120
         TabIndex        =   12
         Top             =   360
         Width           =   4815
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Центровка
         AutoSize        =   -1  'True
         BackStyle       =   0  'Прозрачно
         Caption         =   "Способ эмуляции мыши"
         BeginProperty Font 
            Name            =   "Book Antiqua"
            Size            =   11.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   3
         Left            =   960
         TabIndex        =   22
         Top             =   1080
         Width           =   2550
      End
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "?"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   9750
      TabIndex        =   10
      Top             =   3495
      Width           =   300
   End
   Begin VB.CommandButton cmdCancel 
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
      Height          =   390
      Left            =   8400
      TabIndex        =   9
      Top             =   3480
      Width           =   1260
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Сохранить"
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
      Left            =   7080
      TabIndex        =   8
      Top             =   3480
      Width           =   1260
   End
   Begin VB.CommandButton cmdSetAsDefault 
      Caption         =   "Назначить настройками по умолчанию"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   50
      TabIndex        =   7
      Top             =   3480
      Width           =   3855
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Специализированные настройки"
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   10.5
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1380
      Left            =   5160
      TabIndex        =   5
      Top             =   2040
      Width           =   4935
      Begin VB.CheckBox chkSpecial1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Перед кликом незначительно менять положение курсора (Word of Tanks)"
         Height          =   645
         Left            =   240
         TabIndex        =   6
         Top             =   480
         Width           =   4065
      End
   End
   Begin VB.Frame frameCommon 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Общие настройки"
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   11.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3420
      Left            =   50
      TabIndex        =   0
      Top             =   0
      Width           =   5050
      Begin VB.TextBox txtSpeed 
         Alignment       =   2  'Центровка
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4284
         MaxLength       =   4
         TabIndex        =   1
         Text            =   "1,0"
         Top             =   680
         Width           =   645
      End
      Begin VB.TextBox txtCoordsLimit 
         Alignment       =   2  'Центровка
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3480
         MaxLength       =   4
         TabIndex        =   16
         Text            =   "0"
         ToolTipText     =   "Введите 0 чтобы выключить"
         Top             =   2880
         Width           =   960
      End
      Begin VB.TextBox txtIntervalLimit 
         Alignment       =   2  'Центровка
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3480
         MaxLength       =   10
         TabIndex        =   14
         Text            =   "10"
         ToolTipText     =   "Введите 0 чтобы выключить"
         Top             =   2520
         Width           =   960
      End
      Begin VB.CheckBox chkFakeWait 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Искусственным образом делать мизерную задержку между действиями, чтобы обеспечить корректную работу скрипта"
         Height          =   675
         Left            =   120
         TabIndex        =   13
         Top             =   1680
         Value           =   1  'Отмечено
         Width           =   4620
      End
      Begin VB.CheckBox chkSkipErrors 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Пропускать ошибки при выполнении"
         BeginProperty Font 
            Name            =   "Book Antiqua"
            Size            =   10.5
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   120
         TabIndex        =   4
         Top             =   1260
         Width           =   4290
      End
      Begin VB.HScrollBar scrSpeed 
         Height          =   266
         LargeChange     =   100
         Left            =   126
         Max             =   10000
         Min             =   1
         SmallChange     =   10
         TabIndex        =   3
         Top             =   680
         Value           =   100
         Width           =   4046
      End
      Begin VB.CommandButton cmdSetNSpeed 
         Caption         =   "n"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4680
         TabIndex        =   18
         Top             =   720
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Прозрачно
         Caption         =   "Рандомизировать координаты на:   ±"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   20
         Top             =   2880
         Width           =   3330
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Прозрачно
         Caption         =   "Рандомизировать интервал на:         ±"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   19
         Top             =   2520
         Width           =   3315
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Прозрачно
         Caption         =   "пикс."
         Height          =   255
         Index           =   1
         Left            =   4480
         TabIndex        =   17
         Top             =   2880
         Width           =   480
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Прозрачно
         Caption         =   "мс."
         Height          =   255
         Index           =   0
         Left            =   4560
         TabIndex        =   15
         Top             =   2520
         Width           =   270
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Центровка
         AutoSize        =   -1  'True
         BackStyle       =   0  'Прозрачно
         Caption         =   "Скорость выполнения"
         BeginProperty Font 
            Name            =   "Book Antiqua"
            Size            =   11.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   0
         Left            =   1365
         TabIndex        =   2
         Top             =   320
         Width           =   2310
      End
   End
End
Attribute VB_Name = "frmSSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdHelp_Click()
    ShowHelp "Настройки скрипта"
End Sub

Private Sub cmdSave_Click()
    SetScrollPos
    ResetCommandLine
    CacheScriptSettings
    frmMain.Save
    Unload Me
End Sub

Private Sub cmdSetAsDefault_Click()
    SaveSettings True
End Sub

Private Sub cmdSetNSpeed_Click()
    SetSpeed "1,0"
End Sub

Private Sub Form_Load()
    SetNumericStyle txtIntervalLimit.hwnd
    SetNumericStyle txtCoordsLimit.hwnd
    SetBackColor Me, txtSpeed.hwnd, txtIntervalLimit.hwnd, txtCoordsLimit.hwnd, cmbMouseMode.hwnd
    cmbMouseMode.ItemData(0) = 0
    cmbMouseMode.ItemData(1) = 1
    cmbMouseMode.ItemData(2) = 2
    cmbMouseMode.ItemData(3) = 3
    LoadScriptSettingsFromCache
End Sub

Sub LoadScriptSettingsFromCache()
    With Settings
        SetSpeed .sngSpeed
        chkSkipErrors = CChk(.bSkipErrors)
        chkFakeWait.Value = CChk(.bFakeWait)
        txtIntervalLimit = .lngIntervalLimit
        txtCoordsLimit = .lngCoordsLimit
        chkDelayDownUp = CChk(.lngDelayDownUp > 1)
        cmbMouseMode.ListIndex = .MouseMode
        chkSpecial1 = CChk(.bSpecial1)
    End With
End Sub

Private Sub scrSpeed_Change()
    UpdateSpeedText
End Sub

Private Sub scrSpeed_Scroll()
    UpdateSpeedText
End Sub

Sub UpdateSpeedText()
    Dim Speed As Single
    Speed = scrSpeed.Value / 100
    If Speed < 1 Then
        txtSpeed = FormatNumber(Speed, 2)
        scrSpeed.SmallChange = 1
    ElseIf Speed < 10 Then
        txtSpeed = FormatNumber(Speed, 1)
        If Speed > 1 Then scrSpeed.SmallChange = 10 Else scrSpeed.SmallChange = 1
    Else
        txtSpeed = FormatNumber(Speed, 0)
        scrSpeed.SmallChange = 100
    End If
End Sub

Private Sub txtSpeed_LostFocus()
    SetScrollPos
End Sub

Sub SetScrollPos()
    txtSpeed = Replace(txtSpeed, ".", ",")
    If Not IsNumeric(txtSpeed) Then
        txtSpeed = 1
    End If
    If IsNumeric(txtSpeed) Then If txtSpeed <= 0 Then txtSpeed = 1
    If txtSpeed >= 100 Then txtSpeed = 100
    scrSpeed.Value = txtSpeed * 100
End Sub

Sub SetSpeed(ByVal Speed As String)
    txtSpeed = Speed
    SetScrollPos
End Sub
