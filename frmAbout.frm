VERSION 5.00
Begin VB.Form frmAbout 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "О программе..."
   ClientHeight    =   5505
   ClientLeft      =   22560
   ClientTop       =   5055
   ClientWidth     =   4470
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   367
   ScaleMode       =   3  'Пиксель
   ScaleWidth      =   298
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Visible         =   0   'False
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "Обновление"
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
      Left            =   1560
      TabIndex        =   4
      ToolTipText     =   "Проверить обновление"
      Top             =   5040
      Width           =   1335
   End
   Begin VB.CommandButton cmdSite 
      Caption         =   "Web-сайт"
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
      Left            =   120
      TabIndex        =   2
      ToolTipText     =   "http://rodysoft.ru"
      Top             =   5040
      Width           =   1335
   End
   Begin VB.TextBox txtMain 
      Alignment       =   2  'Центровка
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'Нет
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Вертикаль
      TabIndex        =   1
      Top             =   2640
      Width           =   4455
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Закрыть"
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
      Left            =   3000
      TabIndex        =   0
      Top             =   5040
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'Нет
      Height          =   765
      Left            =   0
      TabIndex        =   3
      Top             =   4830
      Width           =   4815
   End
   Begin VB.Label lblMain 
      Caption         =   $"frmAbout.frx":000C
      Height          =   1215
      Left            =   1800
      TabIndex        =   5
      Top             =   720
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Внутренняя заливка
      Index           =   1
      X1              =   0
      X2              =   320
      Y1              =   320
      Y2              =   320
   End
   Begin VB.Image imgLogo 
      Height          =   2700
      Left            =   840
      Picture         =   "frmAbout.frx":0113
      Top             =   0
      Width           =   2700
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const SITE_URL = "http://rodysoft.ru/"

Private Declare Function ShellExecute Lib "shell32" Alias _
"ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As _
String, ByVal lpFile As String, ByVal lpParameters As String, _
ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Dim StaticText As String

Private Sub cmdOK_Click()
    Unload Me
End Sub

Private Sub cmdSite_Click()
    ShellExecute Me.hwnd, "Open", SITE_URL, "", "", vbNormalFocus
End Sub

Private Sub cmdUpdate_Click()
    CheckUpdate , Me.hwnd
End Sub

Private Sub Form_Deactivate()
    Unload Me
End Sub

Private Sub Form_Initialize()
    InitCommonControls
End Sub

Private Sub Form_Load()
    imgLogo.Left = Me.ScaleWidth / 2 - imgLogo.Width / 2
    txtMain = lblMain.Caption
End Sub

Private Sub txtMain_Change()
    txtMain = lblMain.Caption
End Sub
