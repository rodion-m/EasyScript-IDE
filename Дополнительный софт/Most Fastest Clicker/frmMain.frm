VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00F1E4D9&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "RS Most Fastest Clicker"
   ClientHeight    =   2475
   ClientLeft      =   9690
   ClientTop       =   4890
   ClientWidth     =   4590
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2475
   ScaleWidth      =   4590
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdMaxSpeed 
      Caption         =   "max"
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3600
      TabIndex        =   15
      ToolTipText     =   "Задать настройки максимальной скорости"
      Top             =   1320
      Width           =   735
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Выход"
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3000
      TabIndex        =   10
      Top             =   2160
      Width           =   1095
   End
   Begin VB.CommandButton cmdStop 
      Caption         =   "Стоп F11"
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1560
      TabIndex        =   9
      Top             =   2160
      Width           =   1335
   End
   Begin VB.CommandButton cmdInf 
      Caption         =   "i"
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
      Left            =   4200
      TabIndex        =   8
      ToolTipText     =   "Информация"
      Top             =   2160
      Width           =   255
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "Старт F12"
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   2160
      Width           =   1335
   End
   Begin VB.TextBox txtHM 
      Alignment       =   2  'Центровка
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2280
      MaxLength       =   10
      TabIndex        =   6
      Text            =   "0"
      ToolTipText     =   "0 - бесконечно"
      Top             =   600
      Width           =   2175
   End
   Begin VB.TextBox txtY 
      Alignment       =   2  'Центровка
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3000
      MaxLength       =   4
      TabIndex        =   1
      Text            =   "0"
      ToolTipText     =   "Нажмите F10, чтобы задать координаты"
      Top             =   120
      Width           =   855
   End
   Begin VB.TextBox txtX 
      Alignment       =   2  'Центровка
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1680
      MaxLength       =   4
      TabIndex        =   0
      Text            =   "0"
      ToolTipText     =   "Нажмите F10, чтобы задать координаты"
      Top             =   120
      Width           =   855
   End
   Begin VB.Timer tmrCheckHKS 
      Interval        =   1
      Left            =   3120
      Top             =   1920
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00F1E4D9&
      Caption         =   "Дополнительные настройки"
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   120
      TabIndex        =   11
      Top             =   960
      Width           =   4335
      Begin VB.ComboBox cmbMode 
         BeginProperty Font 
            Name            =   "Book Antiqua"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         ItemData        =   "frmMain.frx":0CCA
         Left            =   1800
         List            =   "frmMain.frx":0CD7
         Style           =   2  'Dropdown List
         TabIndex        =   12
         ToolTipText     =   "Нажмите 'max', чтобы задать максимальную скорость"
         Top             =   375
         Width           =   1620
      End
      Begin VB.CheckBox chkSetCPos 
         BackColor       =   &H00F1E4D9&
         Caption         =   "Пригвоздить курсор"
         BeginProperty Font 
            Name            =   "Book Antiqua"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   720
         Value           =   1  'Отмечено
         Width           =   2175
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Прозрачно
         Caption         =   "Событие:"
         BeginProperty Font 
            Name            =   "Book Antiqua"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   840
         TabIndex        =   13
         Top             =   375
         Width           =   840
      End
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      X1              =   120
      X2              =   4440
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Прозрачно
      Caption         =   "Координаты X:"
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   1470
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Прозрачно
      Caption         =   "Количество кликов:"
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   0
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   1935
   End
   Begin VB.Label Ff1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Прозрачно
      Caption         =   "F10"
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00917B34&
      Height          =   315
      Left            =   4005
      TabIndex        =   3
      ToolTipText     =   "Нажмите F10, чтобы задать координаты"
      Top             =   120
      Width           =   375
   End
   Begin VB.Label Label1y 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Прозрачно
      Caption         =   "Y:"
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   2760
      TabIndex        =   2
      Top             =   120
      Width           =   180
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim blnExecuting As Boolean, pt As POINTAPI

Private Enum ClickMode
    Click
    Down
    Up
End Enum

Private Sub cmdClose_Click()
    Unload Me
    End
End Sub

Private Sub cmdInf_Click()
    MsgBox "*** Назначение данной данной программы заключается в крайне быстром кликанье по 1 точке." & vbCrLf & "*** Количество кликов в секунду зависит от процессора." & vbCrLf & "*** Результаты теста показали, что программа делает в районе 10000 кликов в 1 секунду, что равно 10 кликам в 1 миллисекунду и 600000 кликам в минуту." & vbCrLf & "*** Для того, чтобы остановить процесс, на пару секунд зажмите клавишу F11." & vbCrLf & "*** Выбор правильного режима может в разы увеличить скорость кликера.", vbOKOnly + vbInformation, "Информация..."
End Sub

Private Sub cmdMaxSpeed_Click()
    cmbMode.Text = "Вниз"
    chkSetCPos.Value = 0
End Sub

Private Sub cmdStart_Click()
    blnExecuting = False
    If txtX = "" Then txtX = "0"
    If txtY = "" Then txtY = "0"
    If txtHM = "" Then txtHM = "0"
    If txtHM > (&H7FFFFFFF - 1) Then txtHM = &H7FFFFFFF - 1
    blnExecuting = True
    Call Start
End Sub

Private Sub cmdStop_Click()
    blnExecuting = False
End Sub

Sub Start()
    Dim Counter As Long, Count As Long, x As Long, y As Long, Mode As ClickMode, bSetCPos As Boolean
    Count = txtHM
    x = txtX
    y = txtY
    bSetCPos = chkSetCPos
    Mode = cmbMode.ItemData(cmbMode.ListIndex)
    SetCursorPos x, y
    Do While blnExecuting
        If bSetCPos Then SetCursorPos x, y
        If Mode = Click Or Mode = Down Then mouse_event MOUSEEVENTF_LEFTDOWN, x, y, 0, 0
        If Mode = Click Or Mode = Up Then mouse_event MOUSEEVENTF_LEFTUP, x, y, 0, 0
        Counter = Counter + 1
        If Counter >= Count And Count > 0 Then Exit Do
        If Counter Mod 2000 = 0 Then DoEvents: CheckHKs
    Loop
    blnExecuting = False
End Sub

Private Sub Form_Initialize()
    InitCommonControls
End Sub

Private Sub Form_Load()
    Call SetWindowLong(txtX.hWnd, GWL_STYLE, GetWindowLong(txtX.hWnd, GWL_STYLE) Or ES_NUMBER)
    Call SetWindowLong(txtY.hWnd, GWL_STYLE, GetWindowLong(txtY.hWnd, GWL_STYLE) Or ES_NUMBER)
    Call SetWindowLong(txtHM.hWnd, GWL_STYLE, GetWindowLong(txtHM.hWnd, GWL_STYLE) Or ES_NUMBER)
    cmbMode = "Клик"
    If Len(Command$) > 0 Then
        Dim buff() As String
        buff = Split(Command$, ",")
        If UBound(buff) = 1 Then
            If IsNumeric(buff(0)) Then txtX = buff(0)
            If IsNumeric(buff(0)) Then txtY = buff(1)
        End If
    End If
End Sub

Private Sub tmrCheckHKS_Timer()
    CheckHKs
End Sub

Sub CheckHKs()
    If GetAsyncKeyState(vbKeyF10) = -32767 Then
        GetCursorPosPT pt
        txtX = pt.x
        txtY = pt.y
    End If
    If GetAsyncKeyState(vbKeyF11) <> 0 Then
        blnExecuting = False
    End If
    If GetAsyncKeyState(vbKeyF12) = -32767 Then
        cmdStart_Click
    End If
End Sub
