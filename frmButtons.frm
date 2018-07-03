VERSION 5.00
Begin VB.Form frmButtons 
   BackColor       =   &H00F1E4D9&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Быстрая вставка команды"
   ClientHeight    =   1155
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5790
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1155
   ScaleWidth      =   5790
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdFunc 
      Caption         =   "Назначить окно"
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
      Index           =   6
      Left            =   1140
      TabIndex        =   8
      Top             =   480
      Width           =   2295
   End
   Begin VB.CommandButton cmdFunc 
      Caption         =   "Повторить"
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
      Index           =   4
      Left            =   3465
      TabIndex        =   7
      Top             =   120
      Width           =   2295
   End
   Begin VB.CommandButton cmdFunc 
      Caption         =   "Клик"
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
      Index           =   1
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton cmdFunc 
      Caption         =   "Ждать"
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
      Left            =   0
      TabIndex        =   6
      Top             =   480
      Width           =   1095
   End
   Begin VB.CommandButton cmdFunc 
      Caption         =   "Нажать"
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
      Index           =   2
      Left            =   0
      TabIndex        =   5
      Top             =   840
      Width           =   1095
   End
   Begin VB.CommandButton cmdFunc 
      Caption         =   "Запустить программу"
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
      Index           =   3
      Left            =   3465
      TabIndex        =   4
      Top             =   840
      Width           =   2295
   End
   Begin VB.CommandButton cmdFunc 
      Caption         =   "Запустить скрипт"
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
      Index           =   7
      Left            =   1140
      TabIndex        =   3
      Top             =   840
      Width           =   2295
   End
   Begin VB.CommandButton cmdFunc 
      Caption         =   "Переместить курсор"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   8
      Left            =   1140
      TabIndex        =   2
      Top             =   120
      Width           =   2295
   End
   Begin VB.CommandButton cmdFunc 
      Caption         =   "Показать окно"
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
      Index           =   5
      Left            =   3480
      TabIndex        =   1
      Top             =   480
      Width           =   2295
   End
   Begin VB.Menu mnuContext 
      Caption         =   "[Main]"
      Visible         =   0   'False
      Begin VB.Menu mnuInsert 
         Caption         =   "&Вставить команду"
      End
      Begin VB.Menu mnuInfo 
         Caption         =   "&Информация о команде"
      End
   End
End
Attribute VB_Name = "frmButtons"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim gIndex As Integer

Private Sub cmdFunc_Click(Index As Integer)
    frmMain.InsertFunc cmdFunc(Index).Caption
    frmMain.SetFocus
    frmMain.txtMain.SetFocus
End Sub

Private Sub cmdFunc_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    cmdFunc(Index).ToolTipText = "Вставить команду '" & cmdFunc(Index).Caption & "'"
End Sub

Private Sub cmdFunc_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        gIndex = Index
        PopupMenu mnuContext, , , , mnuInsert
    End If
End Sub

Private Sub Form_Activate()
    Alpha hwnd, 255
End Sub

Private Sub Form_Deactivate()
    Alpha hwnd, 150
End Sub

Private Sub Form_GotFocus()
    Alpha hwnd, 255
End Sub

Private Sub Form_LostFocus()
    Alpha hwnd, 150
End Sub

Private Sub mnuInfo_Click()
    ShowHelp cmdFunc(Index).Caption, Me
End Sub

Private Sub mnuInsert_Click()
    frmMain.InsertFunc cmdFunc(Index).Caption
    frmMain.SetFocus
    frmMain.txtMain.SetFocus
End Sub
