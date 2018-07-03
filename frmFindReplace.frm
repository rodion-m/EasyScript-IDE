VERSION 5.00
Begin VB.Form frmFind 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Найти"
   ClientHeight    =   1395
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6105
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
   ScaleHeight     =   1395
   ScaleWidth      =   6105
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
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
      Left            =   4920
      TabIndex        =   5
      Top             =   480
      Width           =   1095
   End
   Begin VB.ComboBox cmbDest 
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
      ItemData        =   "frmFindReplace.frx":0000
      Left            =   1680
      List            =   "frmFindReplace.frx":000D
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   960
      Width           =   1140
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "Далее"
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
      Left            =   4920
      TabIndex        =   2
      Top             =   150
      Width           =   1095
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
      Left            =   1080
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Вертикаль
      TabIndex        =   0
      Top             =   150
      Width           =   3735
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Прозрачно
      Caption         =   "Направление:"
      Height          =   300
      Index           =   1
      Left            =   120
      TabIndex        =   4
      Top             =   960
      Width           =   1455
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Прозрачно
      Caption         =   "Искать:"
      Height          =   300
      Index           =   0
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   795
   End
End
Attribute VB_Name = "frmFind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim blnPressed As Boolean

Private Sub cmbDest_Change()
    blnPressed = False
End Sub

Private Sub cmbDest_Click()
    blnPressed = False
End Sub

Private Sub cmdCancel_Click()
    Me.Hide
End Sub

Private Sub cmdNext_Click()
    FindNext
End Sub

Sub FindNext()
    If txtFind = "" Then Exit Sub
    If (blnPressed And cmbDest.Text = "Все") Or cmbDest.Text = "Вниз" Then
        FindDown
    ElseIf cmbDest.Text = "Все" Then
        FindAll
    ElseIf cmbDest.Text = "Вверх" Then
        FindUp
    End If
    blnPressed = True
End Sub

Sub FindDown()
    With frmMain.txtMain
        n = InStr(.SelStart + 2, .Text, txtFind, vbTextCompare)
        If n = 0 Then ShowEndMsg: Exit Sub
        .SelStart = n - 1
        .SelLength = Len(txtFind)
        frmMain.SetFocus
        .SetFocus
    End With
End Sub

Sub FindUp()
    With frmMain.txtMain
        n = InStrRev(.Text, txtFind, .SelStart, vbTextCompare)
        If n = 0 Then ShowEndMsg: Exit Sub
        .SelStart = n - 1
        .SelLength = Len(txtFind)
        frmMain.SetFocus
        .SetFocus
    End With
End Sub

Sub FindAll()
    With frmMain.txtMain
        n = InStr(1, .Text, txtFind, vbTextCompare)
        If n = 0 Then ShowEndMsg: Exit Sub
        .SelStart = n - 1
        .SelLength = Len(txtFind)
        frmMain.SetFocus
        .SetFocus
    End With
End Sub

Sub ShowEndMsg()
    MsgBox "Поиск завершен", vbInformation
End Sub

Sub SetFalsePressed()
    blnPressed = False
End Sub

Sub ShowMe()
    blnPressed = False
    If txtFind = "" Then txtFind = frmMain.txtMain.SelText
    If cmbDest.Text = "" Then cmbDest.Text = "Все"
    Me.Show , frmMain
End Sub
