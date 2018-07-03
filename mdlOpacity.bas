Attribute VB_Name = "mdlOpacity"
Option Explicit
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Sub SetWindowPos Lib "user32" _
    (ByVal hwnd As Long, ByVal hWndInsertAfter _
        As Long, ByVal X As Long, ByVal Y As Long, _
            ByVal cx As Long, ByVal cy As Long, _
                ByVal wFlags As Long)

'SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE
Public Const WS_EX_NOACTIVATE As Long = &H8000000
Public Const WS_BORDER As Long = &H800000
Public Const WS_DLGFRAME As Long = &H400000
Public Const WS_VISIBLE As Long = &H10000000

Public Const GWL_EXSTYLE As Long = (-20)
Public Const GWL_STYLE = (-16)

Public Const ES_NUMBER As Long = &H2000&
Public Const ES_READONLY As Long = &H800&

Public Const SWP_SHOWWINDOW As Long = &H40
Public Const SWP_NOSIZE = &H1
Public Const SWP_NOMOVE = &H2

Public Const HWND_BOTTOM As Long = 1
Public Const HWND_BROADCAST As Long = &HFFFF&
Public Const HWND_DESKTOP As Long = 0
Public Const HWND_MESSAGE As Long = -3
Public Const HWND_NOTOPMOST As Long = -2
Public Const HWND_TOP As Long = 0
Public Const HWND_TOPMOST As Long = -1


Private Declare Function SetLayeredWindowAttributes Lib "user32" _
(ByVal hwnd As Long, ByVal crKey As Long, _
ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Const LWA_COLORKEY = &H1
Private Const LWA_ALPHA = &H2
Private Const WS_EX_LAYERED = &H80000
'----- для формы окна --------:
Public Declare Function CreateEllipticRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long

Sub Alpha(Hwnds As Long, vals As Long) 'создание прозрачности у формы (описатель окна, степень прозрачности 0 - 255)
    Dim lStyle As Long 'объявляем числовую переменную
    If vals < 254 And vals > 0 Then 'если значение vals меньше 254, то
        lStyle = GetWindowLong(Hwnds, GWL_EXSTYLE) 'узнаем аттрибуты окна
        lStyle = lStyle Or WS_EX_LAYERED
        SetWindowLong Hwnds, GWL_EXSTYLE, lStyle 'устанавливаем аттрибуты окна
        SetLayeredWindowAttributes Hwnds, 0, vals, LWA_ALPHA 'устанавливаем прозрачность
    Else
        SetWindowLong Hwnds, GWL_EXSTYLE, 0 'устанавливаем обычные аттрибуты окна
    End If
End Sub

Sub DrawForm(hwnd As Long, size As Long)
    SetWindowRgn hwnd, CreateEllipticRgn(0, 0, size, size), True
End Sub
