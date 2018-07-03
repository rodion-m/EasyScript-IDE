VERSION 5.00
Begin VB.Form frmErrorBox 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Критическая ошибка"
   ClientHeight    =   3105
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6045
   BeginProperty Font 
      Name            =   "Book Antiqua"
      Size            =   12
      Charset         =   204
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3105
   ScaleWidth      =   6045
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSend 
      Caption         =   "Отправить"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   320
      Left            =   4800
      TabIndex        =   6
      ToolTipText     =   "Отправить информацию об ошибке разработчику"
      Top             =   2280
      Width           =   1215
   End
   Begin VB.TextBox txtInfo 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1125
      Left            =   126
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   4
      Top             =   1890
      Width           =   4550
   End
   Begin VB.CommandButton cmdCopy 
      Caption         =   "Копировать"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   320
      Left            =   4802
      TabIndex        =   1
      ToolTipText     =   "Копировать информацию об ошибке в буфер обмена"
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Закрыть"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   320
      Left            =   4788
      TabIndex        =   0
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Label lblDescription 
      Alignment       =   2  'Центровка
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Прозрачно
      Caption         =   "Описание ошибки"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   885
      Left            =   105
      TabIndex        =   5
      Top             =   120
      Width           =   4785
      WordWrap        =   -1  'True
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Прозрачно
      Caption         =   "Техническая информация:"
      Height          =   308
      Left            =   126
      TabIndex        =   3
      Top             =   1512
      Width           =   3108
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblStage 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Прозрачно
      Caption         =   "Этапы..."
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   224
      Left            =   126
      TabIndex        =   2
      Top             =   1106
      Width           =   5628
      WordWrap        =   -1  'True
   End
   Begin VB.Image img 
      Height          =   1050
      Left            =   4920
      Picture         =   "frmErrorBox.frx":0000
      Top             =   120
      Width           =   1050
   End
End
Attribute VB_Name = "frmErrorBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const LineHeight = 200

Public Sub ShowError()
    ShowWindow frmMain.hwnd
    frmErrorBox.Caption = "Критическая ошибка на строке №" & GlobalError.line
    lblDescription = GlobalError.Description & ":" & vbCrLf & LCase(Err.Description)
    lblStage.Top = lblDescription.Top + lblDescription.Height + 100
    'lblStage.Height = LineHeight * GlobalError.StageCount
    lblStage.Caption = Join(GlobalError.Stage, vbCrLf)
    Me.Height = Me.Height + (LineHeight * (GlobalError.StageCount - 1)) + 100
    lbl.Top = lbl.Top + (LineHeight * (GlobalError.StageCount - 1)) + 150
    txtInfo.Top = txtInfo.Top + (LineHeight * (GlobalError.StageCount - 1)) + 150
    cmdCopy.Top = cmdCopy.Top + (LineHeight * (GlobalError.StageCount - 1)) + 150
    cmdCancel.Top = cmdCancel.Top + (LineHeight * (GlobalError.StageCount - 1)) + 150
    cmdSend.Top = cmdSend.Top + (LineHeight * (GlobalError.StageCount - 1)) + 150
    txtInfo = "Номер: " & Err.Number & vbCrLf & _
    "Описание: " & Err.Description & vbCrLf & "Выражение: " & GlobalError.Expression
    Beep
    Me.Show vbModal, frmMain
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdCopy_Click()
    Clipboard.Clear
    Clipboard.SetText lblDescription & vbCrLf & vbCrLf & lblStage.Caption & vbCrLf & vbCrLf & txtInfo
End Sub

Private Sub cmdSend_Click()
    If MsgBox("Уверены ли в том, что данная ошибка — баг программы?", vbQuestion + vbYesNo, "Отправка отчета...") = vbNo Then Exit Sub
    SendInfoToServer Replace(lblDescription, vbCrLf, " ") & vbCrLf & lblStage.Caption & vbCrLf & vbCrLf & txtInfo, "bugreport", hwnd
End Sub

Private Sub Form_Load()
    cmdCancel.Top = txtInfo.Top + txtInfo.Height - cmdCancel.Height
End Sub
