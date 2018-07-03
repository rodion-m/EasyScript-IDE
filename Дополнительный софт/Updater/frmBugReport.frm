VERSION 5.00
Begin VB.Form frmBugReport 
   BackColor       =   &H00F1E4D9&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "RodySoft — Отправить информацию об ошибке"
   ClientHeight    =   3045
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8670
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3045
   ScaleWidth      =   8670
   StartUpPosition =   1  'CenterOwner
   Visible         =   0   'False
   Begin VB.TextBox txtInfo 
      Height          =   2415
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Оба
      TabIndex        =   2
      ToolTipText     =   "Описание ошибки"
      Top             =   120
      Width           =   8415
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "Отправить"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   2640
      Width           =   4155
   End
   Begin VB.CommandButton cmdCancel 
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
      Left            =   4440
      TabIndex        =   0
      Top             =   2640
      Width           =   4155
   End
End
Attribute VB_Name = "frmBugReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    End
End Sub

Private Sub cmdSend_Click()
    Static ar As Boolean
    If Len(txtInfo) < 15 Then Exit Sub
    If ar Then Exit Sub
    ar = True
    MousePointer = vbHourglass
    SendInfoToServer "bugreport", txtInfo, hwnd
    MousePointer = vbNormal
    ar = False
End Sub

Private Sub Form_Load()
    txtInfo = Replace(ReadIni("Send_Info"), "[crlf]", vbCrLf)
    WriteIni "Send_Info", ""
    SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE
End Sub
