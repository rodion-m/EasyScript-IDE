VERSION 5.00
Begin VB.Form frmDebug 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Окно отладки"
   ClientHeight    =   1485
   ClientLeft      =   105
   ClientTop       =   375
   ClientWidth     =   7815
   LinkTopic       =   "Debug"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1485
   ScaleWidth      =   7815
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtDebug 
      Appearance      =   0  'Плоска
      BorderStyle     =   0  'Нет
      ForeColor       =   &H000000C0&
      Height          =   1400
      Left            =   0
      Locked          =   -1  'True
      MousePointer    =   1  'Указатель
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Оба
      TabIndex        =   0
      Top             =   0
      Width           =   7700
   End
End
Attribute VB_Name = "frmDebug"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    SetSize
End Sub

Sub SetSize()
    txtDebug.Width = Me.ScaleWidth
    txtDebug.Height = Me.ScaleHeight
End Sub

Private Sub Form_Resize()
    SetSize
End Sub
