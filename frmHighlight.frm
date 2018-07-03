VERSION 5.00
Begin VB.Form frmMarker 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'Нет
   ClientHeight    =   480
   ClientLeft      =   -6060
   ClientTop       =   -6060
   ClientWidth     =   510
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   32
   ScaleMode       =   3  'Пиксель
   ScaleWidth      =   34
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.Shape shpPoint 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Непрозрачно
      Height          =   120
      Left            =   240
      Shape           =   3  'Круг
      Top             =   240
      Width           =   120
   End
   Begin VB.Shape shp 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Непрозрачно
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H000000C0&
      FillStyle       =   7  'Диагональное пересечение
      Height          =   795
      Left            =   -120
      Top             =   0
      Width           =   885
   End
   Begin VB.Menu mnuMain 
      Caption         =   "[Main]"
      Visible         =   0   'False
      Begin VB.Menu mnuHide 
         Caption         =   "&Скрыть маркер"
      End
      Begin VB.Menu mnuDisable 
         Caption         =   "&Отключить маркер"
      End
      Begin VB.Menu mnuAlwaysOnTop 
         Caption         =   "&Всегда видимый"
      End
      Begin VB.Menu s0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuChangeCoordinatesOnMove 
         Caption         =   "&При перемещении изменять координаты точки"
         Checked         =   -1  'True
      End
      Begin VB.Menu s1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuInsertClick 
         Caption         =   "&Вставить команду "
      End
      Begin VB.Menu mnuInsertSetCursorPos 
         Caption         =   "&Вставить команду "
      End
   End
End
Attribute VB_Name = "frmMarker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Activate()
    shpPoint.Left = MARKER_SIZE / 2 - shpPoint.Width / 2
    shpPoint.Top = MARKER_SIZE / 2 - shpPoint.Height / 2
    SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE
End Sub

Private Sub Form_DblClick()
    frmMain.InsertFunc "Клик(" & GetMarkerX & ", " & GetMarkerY & ")"
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyInsert Then frmMain.InsertFunc "Клик(" & GetMarkerX & ", " & GetMarkerY & ")"
End Sub

Private Sub Form_Load()
    Show
    frmMain.Show
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Static PrevX As Single, PrevY As Single
    If Button = 1 And WindowState = 0 Then
        Me.Left = Me.Left + (X - PrevX)
        Me.Top = Me.Top + (Y - PrevY)
        If mnuChangeCoordinatesOnMove.Checked Then frmMain.ChangeFuncParametrs True, GetMarkerX, GetMarkerY
    Else
        PrevX = X
        PrevY = Y
    End If
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        If mnuChangeCoordinatesOnMove.Checked Then frmMain.ChangeFuncParametrs True, GetMarkerX, GetMarkerY
    ElseIf Button = 2 Then
        mnuInsertClick.Caption = "Вставить команду 'Клик(" & GetMarkerX & ", " & GetMarkerY & ")'" & vbTab & "Ctrl+Insert"
        mnuInsertSetCursorPos.Caption = "Вставить команду 'Переместить курсор(" & GetMarkerX & ", " & GetMarkerY & ")'" & vbTab & "Alt+Insert"
        PopupMenu mnuMain, , , , mnuHide
    End If
End Sub

Function GetMarkerX() As Long
    GetMarkerX = (Me.Left / Screen.TwipsPerPixelX) + shpPoint.Left + shpPoint.Width / 2
End Function

Function GetMarkerY() As Long
    GetMarkerY = (Me.Top / Screen.TwipsPerPixelX) + shpPoint.Top + shpPoint.Height / 2
End Function

Private Sub mnuAlwaysOnTop_Click()
    mnuAlwaysOnTop.Checked = Not mnuAlwaysOnTop.Checked
End Sub

Private Sub mnuChangeCoordinatesOnMove_Click()
    mnuChangeCoordinatesOnMove.Checked = Not mnuChangeCoordinatesOnMove.Checked
End Sub

Private Sub mnuDisable_Click()
    frmPSettings.chkShowMarker = 0
    SaveSettings
    HideMarker
End Sub

Private Sub mnuHide_Click()
    HideMarker
End Sub

Private Sub mnuInsertClick_Click()
    frmMain.InsertFunc "Клик(" & GetMarkerX & ", " & GetMarkerY & ")"
End Sub

Private Sub mnuInsertSetCursorPos_Click()
    frmMain.InsertFunc "Переместить курсор(" & GetMarkerX & ", " & GetMarkerY & ")"
End Sub
