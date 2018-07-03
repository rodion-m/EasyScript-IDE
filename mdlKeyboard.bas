Attribute VB_Name = "mdlKeyboard"
Option Explicit

'Узнаем заговоловок активного окна
Private Declare Function GetWindowTextA Lib "user32" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long

Declare Function GetKeyboardLayout Lib "user32" (ByVal dwLayout As Long) As Long
Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Declare Function ActivateKeyboardLayout Lib "user32.dll" (ByVal HKL As Long, ByVal Flags As Long) As Long
'
'Public Const ActivateEn = &H409
'Public Const ActivateRu = &H419

Public Const En = 67699721
Public Const Ru = 68748313

Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Long
Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Long

Public Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)

Public Const KEYEVENTF_EXTENDEDKEY = &H1
Public Const KEYEVENTF_KEYDOWN = &H0
Public Const KEYEVENTF_KEYUP = &H2
Public Const VK_STARTKEY = &H5B

'Альтернатива - SendKeys("KEY" + "KEY2"..)

'    Press VK_H, Down   ' press H
'    Press VK_H, Up   ' release H

' Проверка на сочитание QWERTY
'If CBool(GetAsyncKeyState(VK_Q) And &H8000) And _
'        CBool(GetAsyncKeyState(VK_W) And &H8000) And _
'        CBool(GetAsyncKeyState(VK_E) And &H8000) And _
'        CBool(GetAsyncKeyState(VK_R) And &H8000) And _
'        CBool(GetAsyncKeyState(VK_T) And &H8000) And _
'        CBool(GetAsyncKeyState(VK_Y) And &H8000) Then Text1 = "YES!QWERTY"

''ЕСЛИ английский, то
'    keybd_event VK_H, 0, KEYEVENTF_KEYDOWN, 0   ' press H
'    keybd_event VK_H, 0, KEYEVENTF_KEYUP, 0   ' release H
'    keybd_event VK_E, 0, 0, 0  ' press E
'    keybd_event VK_E, 0, KEYEVENTF_KEYUP, 0  ' release E
'    keybd_event VK_L, 0, 0, 0  ' press L
'    keybd_event VK_L, 0, KEYEVENTF_KEYUP, 0  ' release L
'    keybd_event VK_L, 0, 0, 0  ' press L
'    keybd_event VK_L, 0, KEYEVENTF_KEYUP, 0  ' release L
'    keybd_event VK_O, 0, 0, 0  ' press O
'    keybd_event VK_O, 0, KEYEVENTF_KEYUP, 0  ' release O
''ЕСЛИ же русский, то
'    keybd_event VK_H, 0, 0, 0   ' press р
'    keybd_event VK_H, 0, KEYEVENTF_KEYUP, 0   ' release р
'    keybd_event VK_E, 0, 0, 0  ' press E
'    keybd_event VK_E, 0, KEYEVENTF_KEYUP, 0  ' release у
'    keybd_event VK_L, 0, 0, 0  ' press L
'    keybd_event VK_L, 0, KEYEVENTF_KEYUP, 0  ' release д
'    keybd_event VK_L, 0, 0, 0  ' press L
'    keybd_event VK_L, 0, KEYEVENTF_KEYUP, 0  ' release д
'    keybd_event VK_O, 0, 0, 0  ' press O
'    keybd_event VK_O, 0, KEYEVENTF_KEYUP, 0  ' release щ

'Сочетание альт+Ф4
'    keybd_event VK_ALT, 0, 0, 0   ' press Alt
'    keybd_event VK_F4, 0, 0, 0   ' press F4
'    keybd_event VK_ALT, 0, KEYEVENTF_KEYUP, 0   ' release Alt
'    keybd_event VK_F4, 0, KEYEVENTF_KEYUP, 0   ' release F4
'
'Public Sub apiPress(ByVal Char As Byte, Optional ByVal DownOrUp As Long)
'    If DownOrUp = kPress Then
'        keybd_event Char, 0, Down, 0
'        Wait 10
'        keybd_event Char, 0, Up, 0
'    Else
'        keybd_event Char, 0, DownOrUp, 0
'    End If
'End Sub

Public Function GetLang() As String
    GetLang = GetKeyboardLayout(GetWindowThreadProcessId(GetForegroundWindow, 0))
End Function

Public Function GetMyLang() As String
    GetMyLang = GetKeyboardLayout(GetWindowThreadProcessId(frmMain.hwnd, 0))
End Function

Public Function GetForegroundWindowCaption() As String
    GetForegroundWindowCaption = Space(255)
    GetForegroundWindowCaption = Mid$(GetForegroundWindowCaption, 1, GetWindowTextA(GetForegroundWindow(), GetForegroundWindowCaption, 255)) 'обрезаем его, оставляя только нужную часть
    'If InStr(GetForegroundWindowCaption, "Program Manager") Then GetForegroundWindowCaption = ""
End Function

Public Function GetWindowCaption(hwnd As Long) As String
    GetWindowCaption = Space(255)
    GetWindowCaption = Mid$(GetWindowCaption, 1, GetWindowTextA(hwnd, GetWindowCaption, 255))
End Function

Public Function CheckMouse() As Boolean
    CheckMouse = Not CBool(GetAsyncKeyState(vbKeyLButton))
End Function

Public Function GetKey(vKey As KeyCodeConstants, Optional AnyKeyPos As Boolean) As Boolean
    If AnyKeyPos Then
        If GetAsyncKeyState(vKey) <> 0 Then If GetAsyncKeyState(vKey) <> 0 Then GetKey = True
    Else
        GetKey = (GetAsyncKeyState(vKey) = -32767)
    End If
End Function

Public Function GetNumLock() As Boolean
    GetNumLock = CBool(GetKeyState(vbKeyNumlock))
End Function

Public Function GetShift() As Boolean
    GetShift = CBool(GetAsyncKeyState(vbKeyShift) And &H8000)
End Function

Public Function GetCaps() As Boolean
    GetCaps = CBool(GetKeyState(vbKeyCapital))
End Function

Function GetUpCase() As Boolean
    GetUpCase = GetCaps Xor GetShift
End Function

Public Function GetKeys(vKeys() As KeyCodeConstants) As Boolean  'после проверки всегда необходимо делать слип
    If GetKeysCount(vKeys()) = 1 Then
        'если всего одна клавиша
        GetKeys = GetKey(vKeys(UBound(vKeys)))
    End If
    For i = LBound(vKeys) To UBound(vKeys)
        If (Not GetKey(vKeys(i), True)) And vKeys(i) > 0 Then
            Exit Function
        End If
    Next i
    GetKeys = True
End Function

Function GetKeysCount(Keys() As KeyCodeConstants) As Long
    For i = LBound(Keys) To UBound(Keys)
            If Keys(i) > 0 Then GetKeysCount = GetKeysCount + 1
    Next i
End Function
