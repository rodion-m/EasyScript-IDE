VERSION 5.00
Begin VB.Form frmCustom 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Собственные константы и команды"
   ClientHeight    =   6045
   ClientLeft      =   105
   ClientTop       =   435
   ClientWidth     =   8400
   LinkTopic       =   "Form1"
   ScaleHeight     =   6045
   ScaleWidth      =   8400
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   518
      Left            =   6678
      TabIndex        =   2
      Top             =   4662
      Width           =   1148
   End
   Begin VB.TextBox txtFuncs 
      Height          =   5432
      Left            =   3724
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Оба
      TabIndex        =   3
      Text            =   "frmCustom.frx":0000
      Top             =   630
      Width           =   4676
   End
   Begin VB.TextBox txtConsts 
      Height          =   5432
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Оба
      TabIndex        =   0
      Text            =   "frmCustom.frx":0261
      Top             =   630
      Width           =   3542
   End
   Begin VB.Line DelimLine 
      BorderColor     =   &H00808080&
      BorderWidth     =   2
      X1              =   3780
      X2              =   3654
      Y1              =   756
      Y2              =   6048
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Прозрачно
      Caption         =   "Собственные функции:"
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   13.5
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Index           =   1
      Left            =   3720
      TabIndex        =   4
      ToolTipText     =   "Кликните, чтобы выбрать"
      Top             =   120
      Width           =   2985
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Прозрачно
      Caption         =   "Собственные константы:"
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   13.5
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Index           =   0
      Left            =   0
      TabIndex        =   1
      ToolTipText     =   "Кликните, чтобы выбрать"
      Top             =   120
      Width           =   3360
   End
End
Attribute VB_Name = "frmCustom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
SetCustomConsts txtConsts
SetCustomFunctions txtFuncs
End Sub

Private Sub Form_Load()
    Dim LinePos As Long
    LinePos = txtConsts.Width + (txtFuncs.Left - (txtConsts.Left + txtConsts.Width)) \ 2
    DelimLine.Y1 = txtConsts.Top
    DelimLine.Y2 = txtConsts.Top + txtConsts.Height
    DelimLine.X1 = LinePos
    DelimLine.X2 = LinePos
End Sub
