VERSION 5.00
Begin VB.Form frmReplace 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Замена"
   ClientHeight    =   2265
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6300
   BeginProperty Font 
      Name            =   "Book Antiqua"
      Size            =   11.25
      Charset         =   204
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2265
   ScaleWidth      =   6300
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox cmbWhere 
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
      ItemData        =   "frmReplace.frx":0000
      Left            =   2880
      List            =   "frmReplace.frx":000D
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   1800
      Width           =   1860
   End
   Begin VB.ComboBox cmbType 
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
      ItemData        =   "frmReplace.frx":0031
      Left            =   720
      List            =   "frmReplace.frx":003E
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   1800
      Width           =   1620
   End
   Begin VB.CommandButton cmdReplaceAll 
      Caption         =   "Заменить все"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   4800
      TabIndex        =   4
      Top             =   150
      Width           =   1455
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
      Height          =   300
      Left            =   4800
      TabIndex        =   5
      Top             =   480
      Width           =   1455
   End
   Begin VB.TextBox txtReplace 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   690
      Left            =   720
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Вертикаль
      TabIndex        =   1
      Top             =   990
      Width           =   3975
   End
   Begin VB.TextBox txtFind 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   690
      Left            =   720
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Вертикаль
      TabIndex        =   0
      Top             =   150
      Width           =   3975
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Прозрачно
      Caption         =   "Где:"
      Height          =   300
      Index           =   3
      Left            =   2400
      TabIndex        =   9
      Top             =   1800
      Width           =   420
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Прозрачно
      Caption         =   "Тип:"
      Height          =   300
      Index           =   2
      Left            =   120
      TabIndex        =   8
      Top             =   1800
      Width           =   465
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Прозрачно
      Caption         =   "Чем:"
      Height          =   300
      Index           =   1
      Left            =   120
      TabIndex        =   7
      Top             =   960
      Width           =   480
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Прозрачно
      Caption         =   "Что:"
      Height          =   300
      Index           =   0
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   450
   End
End
Attribute VB_Name = "frmReplace"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    Me.Hide
End Sub

Private Sub cmdReplaceAll_Click()
    Dim tmp As String
    If txtFind = "" Then Exit Sub
    tmp = frmMain.txtMain
    Select Case cmbWhere
        Case "Везде"
            If cmbType.Text = "Все" Then
                tmp = Replace(tmp, txtFind, txtReplace, , , vbTextCompare)
            Else
                tmp = ReplaceInQuotesSTR(tmp, txtFind, txtReplace, cmbType.ItemData(cmbType.ListIndex))
                tmp = ReplaceOutQuotesSTR(tmp, txtFind, txtReplace, cmbType.ItemData(cmbType.ListIndex))
            End If
        Case "В кавычках"
            tmp = ReplaceInQuotesSTR(tmp, txtFind, txtReplace, cmbType.ItemData(cmbType.ListIndex))
        Case "Вне кавычек"
            tmp = ReplaceOutQuotesSTR(tmp, txtFind, txtReplace, cmbType.ItemData(cmbType.ListIndex))
    End Select
     frmMain.SetScriptText tmp
     MsgBox "Операция успешно завершена", vbInformation
End Sub

Private Sub Form_Load()
    SetBackColor Me, txtFind.hwnd, txtReplace.hwnd, cmbType.hwnd, cmbWhere.hwnd
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Cancel = 1
    Me.Hide
End Sub

Sub ShowMe()
    If txtFind = "" Then txtFind = frmMain.txtMain.SelText
    If cmbType.Text = "" Then cmbType.Text = "Все"
    If cmbWhere.Text = "" Then cmbWhere.Text = "Везде"
    Me.Show , frmMain
End Sub
