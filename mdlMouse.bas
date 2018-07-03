Attribute VB_Name = "mdlMouse"
Option Explicit
Public Declare Function GetCursorPosPT& Lib "user32" Alias "GetCursorPos" (lpPoint As POINTAPI)

Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long
Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)

Public Const MOUSEEVENTF_LEFTDOWN = &H2
Public Const MOUSEEVENTF_LEFTUP = &H4
Public Const MOUSEEVENTF_MIDDLEDOWN = &H20
Public Const MOUSEEVENTF_MIDDLEUP = &H40
Public Const MOUSEEVENTF_RIGHTDOWN = &H8
Public Const MOUSEEVENTF_RIGHTUP = &H10
Public Const MOUSEEVENTF_MOVE = &H1

Dim Point As POINTAPI

Sub LeftClick(Optional X As Long, Optional Y As Long)
    LeftDown X, Y
    Wait Settings.lngDelayDownUp
    LeftUp X, Y
End Sub

Sub LeftDown(Optional X As Long, Optional Y As Long)
    If X = 0 And Y = 0 Then GetCursorPos X, Y
    mouse_event MOUSEEVENTF_LEFTDOWN, X, Y, 0, 0
End Sub

Sub LeftUp(Optional X As Long, Optional Y As Long)
    If X = 0 And Y = 0 Then GetCursorPos X, Y
    mouse_event MOUSEEVENTF_LEFTUP, X, Y, 0, 0
End Sub

Sub RightClick(Optional X As Long, Optional Y As Long)
    RightDown X, Y
    Wait Settings.lngDelayDownUp
    RightUp X, Y
End Sub

Sub RightDown(Optional X As Long, Optional Y As Long)
    If X = 0 And Y = 0 Then GetCursorPos X, Y
    mouse_event MOUSEEVENTF_RIGHTDOWN, X, Y, 0, 0
End Sub

Sub RightUp(Optional X As Long, Optional Y As Long)
    If X = 0 And Y = 0 Then GetCursorPos X, Y
    mouse_event MOUSEEVENTF_RIGHTUP, X, Y, 0, 0
End Sub

Sub GetCursorPos(X As Long, Y As Long)
    GetCursorPosPT Point
    X = Point.X
    Y = Point.Y
End Sub

Function CursorPos() As POINTAPI
    GetCursorPosPT CursorPos
End Function

Sub MiddleClick(Optional X As Long, Optional Y As Long)
    MiddleDown X, Y
    Wait Settings.lngDelayDownUp
    MiddleUp X, Y
End Sub

Sub MiddleDown(Optional X As Long, Optional Y As Long)
    If X = 0 And Y = 0 Then GetCursorPos X, Y
    mouse_event MOUSEEVENTF_MIDDLEDOWN, X, Y, 0, 0
End Sub

Sub MiddleUp(Optional X As Long, Optional Y As Long)
    If X = 0 And Y = 0 Then GetCursorPos X, Y
    mouse_event MOUSEEVENTF_MIDDLEUP, X, Y, 0, 0
End Sub

Sub MouseAction(Optional ByVal mKey As MouseKey = pLeft, Optional ByVal mEvent As MouseEvent = pClick, _
                            Optional ByVal Method As MouseModes = MM_mouse_event, Optional X As Long, Optional Y As Long)
    Dim blnIsClick As Boolean
    If X = 0 And Y = 0 Then GetCursorPos X, Y
    If mEvent = pClick Then blnIsClick = True: mEvent = pDown
    Select Case mKey
        Case pLeft
            Select Case mEvent
                Case pDown
                    Select Case Method
                        Case MM_mouse_event
                            'Left Down mouse_event
                            mouse_event MOUSEEVENTF_LEFTDOWN, X, Y, ByVal 0&, ByVal 0&
                        Case MM_keybd_event
                            'Left Down keybd_event
                            keybd_event vbKeyLButton, ByVal 0&, KEYEVENTF_KEYDOWN, ByVal 0&
                        Case MM_SendMessage
                            'Left Down SendMessage
                            SendMessage GetForegroundWindow, WM_LBUTTONDOWN, ByVal 0&, ByVal CLng(X + Y * &H10000)
                        Case MM_PostMessage
                            'Left Down PostMessage
                            PostMessage GetForegroundWindow, WM_LBUTTONDOWN, ByVal 0&, ByVal CLng(X + Y * &H10000)
                    End Select
                Case pUp
                    Select Case Method
                        Case MM_mouse_event
                            'Left Up mouse_event
                            mouse_event MOUSEEVENTF_LEFTUP, X, Y, ByVal 0&, ByVal 0&
                        Case MM_keybd_event
                            'Left Up keybd_event
                            keybd_event vbKeyLButton, ByVal 0&, KEYEVENTF_KEYUP, ByVal 0&
                        Case MM_SendMessage
                            'Left Up SendMessage
                            SendMessage GetForegroundWindow, WM_LBUTTONUP, ByVal 0&, ByVal CLng(X + Y * &H10000)
                        Case MM_PostMessage
                            'Left Up PostMessage
                            PostMessage GetForegroundWindow, WM_LBUTTONUP, ByVal 0&, ByVal CLng(X + Y * &H10000)
                    End Select
            End Select
        Case pRight
            Select Case mEvent
                Case pDown
                    Select Case Method
                        Case MM_mouse_event
                            'Right Down mouse_event
                            mouse_event MOUSEEVENTF_RIGHTDOWN, X, Y, ByVal 0&, ByVal 0&
                        Case MM_keybd_event
                            'Right Down keybd_event
                            keybd_event vbKeyRButton, 0, KEYEVENTF_KEYDOWN, ByVal 0&
                        Case MM_SendMessage
                            'Right Down SendMessage
                            SendMessage GetForegroundWindow, WM_RBUTTONDOWN, ByVal 0&, ByVal CLng(X + Y * &H10000)
                        Case MM_PostMessage
                            'Right Down PostMessage
                            PostMessage GetForegroundWindow, WM_RBUTTONDOWN, ByVal 0&, ByVal CLng(X + Y * &H10000)
                    End Select
                Case pUp
                    Select Case Method
                        Case MM_mouse_event
                            'Right Up mouse_event
                            mouse_event MOUSEEVENTF_RIGHTUP, X, Y, ByVal 0&, ByVal 0&
                        Case MM_keybd_event
                            'RightUp keybd_event
                            keybd_event vbKeyRButton, ByVal 0&, KEYEVENTF_KEYUP, ByVal 0&
                        Case MM_SendMessage
                            'Right Up SendMessage
                            SendMessage GetForegroundWindow, WM_RBUTTONUP, ByVal 0&, ByVal CLng(X + Y * &H10000)
                        Case MM_PostMessage
                            'Right Up PostMessage
                            PostMessage GetForegroundWindow, WM_RBUTTONUP, ByVal 0&, ByVal CLng(X + Y * &H10000)
                    End Select
            End Select
        Case pMiddle
            Select Case mEvent
                Case pDown
                    Select Case Method
                        Case MM_mouse_event
                            'Middle Down mouse_event
                            mouse_event MOUSEEVENTF_MIDDLEDOWN, X, Y, ByVal 0&, ByVal 0&
                        Case MM_keybd_event
                            'Middle Down keybd_event
                            keybd_event vbKeyMButton, ByVal 0&, KEYEVENTF_KEYDOWN, ByVal 0&
                        Case MM_SendMessage
                            'Middle Down SendMessage
                            SendMessage GetForegroundWindow, WM_MBUTTONDOWN, 0, ByVal CLng(X + Y * &H10000)
                        Case MM_PostMessage
                            'Middle Down PostMessage
                            PostMessage GetForegroundWindow, WM_MBUTTONDOWN, 0, ByVal CLng(X + Y * &H10000)
                    End Select
                Case pUp
                    Select Case Method
                        Case MM_mouse_event
                            'Middle Up mouse_event
                            mouse_event MOUSEEVENTF_MIDDLEUP, X, Y, 0, 0
                        Case MM_keybd_event
                            'Middle keybd_event
                            keybd_event vbKeyMButton, 0, KEYEVENTF_KEYUP, 0
                        Case MM_SendMessage
                            'Middle Up SendMessage
                            SendMessage GetForegroundWindow, WM_MBUTTONUP, 0, ByVal CLng(X + Y * &H10000)
                        Case MM_PostMessage
                            'Middle Up PostMessage
                            PostMessage GetForegroundWindow, WM_MBUTTONUP, 0, ByVal CLng(X + Y * &H10000)
                    End Select
            End Select
    End Select
    If blnIsClick Then
        Wait Settings.lngDelayDownUp, False
        MouseAction mKey, pUp, Method, X, Y
    End If
End Sub
