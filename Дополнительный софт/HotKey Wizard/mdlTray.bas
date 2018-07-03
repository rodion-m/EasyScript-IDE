Attribute VB_Name = "mdlTray"
Option Explicit

Declare Function Shell_NotifyIcon Lib "shell32.dll" Alias "Shell_NotifyIconA" (ByVal dwMessage As dwMess, lpData As NOTIFYICONDATA) As Long
                
Public Enum dwMess
    NIM_ADD = &H0&               ' Добавление иконки
    NIM_DELETE = &H2&            ' Удаление иконки
    NIM_MODIFY = &H1&            ' Изменение параметров иконки
End Enum

'Type NOTIFYICONDATA
'    cbSize As Long              ' Размер переменной типа NOTIFYICONDATA
'    hwnd As Long                ' Указатель окна создающего иконку
'    uID As Long                 ' Указатель на иконку в пределах приложения
'    uFlags As uF                ' Маска для следующих параметров
'    uCallbackMessage As WindowMessages ' Возвращаемое событие
'    hIcon As Long               ' Указатель на изображение для иконки
'    szTip As String * 64        ' Всплывающий над иконкой текст
'End Type

Private Const NIIFICON_INFO = 1&
Private Const NIF_INFO = &H10&

Private Type NOTIFYICONDATA
   cbSize As Long             ' 4
   hwnd As Long               ' 8
   uID As Long                ' 12
   uFlags As Long             ' 16
   uCallbackMessage As Long   ' 20
   hIcon As Long              ' 24
   szTip As String * 128      ' 152
   dwState As Long            ' 156
   dwStateMask As Long        ' 160
   szInfo As String * 256     ' 416
   uTimeOutOrVersion As Long  ' 420
   szInfoTitle As String * 64 ' 484
   dwInfoFlags As Long        ' 488
   guidItem As Long           ' 492
End Type

Public Enum uF
    NIF_MESSAGE = &H1           ' Значение имеет uCallbackMessage
    NIF_ICON = &H2              ' Значение имеет hIcon
    NIF_TIP = &H4               ' Значение имеет szTip
End Enum

Public Enum WindowMessages
    WM_MOUSEMOVE = &H200&
    WM_LBUTTONDOWN = &H201
    WM_LBUTTONUP = &H202
    WM_LBUTTONDBLCLK = &H203
    WM_RBUTTONDOWN = &H204
    WM_RBUTTONUP = &H205
    WM_RBUTTONDBLCLK = &H206
    WM_SETFOCUS = &H7
    WM_KEYDOWN = &H100
    WM_KEYFIRST = &H100
    WM_KEYLAST = &H108
    WM_KEYUP = &H101
End Enum

Dim NID As NOTIFYICONDATA
'-------------------------------------------------
Sub AddIcon()
    NID.cbSize = Len(NID)                       'Размер структуры
    NID.hwnd = frmMain.hwnd                       'Указатель на форму
    NID.uID = vbNull                            'Идентификатор иконки
    NID.uFlags = NIF_MESSAGE + NIF_ICON + NIF_TIP
    NID.uCallbackMessage = WM_MOUSEMOVE
    NID.hIcon = frmMain.Icon
    NID.szTip = App.ProductName & Chr$(0)
    Shell_NotifyIcon NIM_ADD, NID
End Sub

Sub DeleteIcon()
    Shell_NotifyIcon NIM_DELETE, NID
End Sub

Sub ShowMessage(ByVal strMessage As String, Optional strTitle As String)
    If strTitle = "" Then strTitle = App.ProductName
    NID.szInfo = strMessage & vbNullChar
    NID.szInfoTitle = strTitle & vbNullChar
    NID.uTimeOutOrVersion = 5000
    NID.dwInfoFlags = NIIFICON_INFO
    NID.uFlags = NIF_INFO
    Shell_NotifyIcon NIM_MODIFY, NID
End Sub
