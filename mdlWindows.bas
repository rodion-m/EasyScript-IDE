Attribute VB_Name = "modWindows"
Option Explicit

Public Windows() As WindowMinAttribs
Dim WinIndex As Long, buff As String

Type RECT
   Left    As Long
   Top     As Long
   Right   As Long
   Bottom  As Long
End Type

Type WindowMinAttribs
    hwnd As Long
    Caption As String
    Position As RECT
End Type

Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Declare Function GetDesktopWindow Lib "user32" () As Long
Declare Function EnumChildWindows Lib "user32" (ByVal hWndParent As Long, ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Declare Function EnumWindows Lib "user32" (ByVal lpEnumFunc As Long, ByVal lParam As Long) As Boolean
Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long

Public Declare Function SetForegroundWindow Lib "user32.dll" (ByVal hwnd As Long) As Long

Declare Function APIDestroyWindow Lib "user32" Alias "DestroyWindow" (ByVal hwnd As Long) As Long
Declare Function IsWindowVisible Lib "user32" (ByVal hwnd As Long) As Long
Declare Function IsWindowEnabled Lib "user32" (ByVal hwnd As Long) As Long
Declare Function EnableWindow Lib "user32" (ByVal hwnd As Long, ByVal fEnable As Long) As Long
Declare Function APIShowWindow Lib "user32" Alias "ShowWindow" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccessas As Long, ByVal bInheritHandle As Long, ByVal dwProcId As Long) As Long
Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Public Const SW_RESTORE As Long = 9
Public Const STILL_ACTIVE As Long = &H103
Public Const SYNCHRONIZE As Long = &H100000
Public Const PROCESS_ALL_ACCESS As Long = &H1F0FFF

Public Const WM_CLOSE = &H10

Function GetWindows(Optional ByVal JustWithCaption As Long)
    ReDim Windows(1 To 1) ' чтобы начать нумерацию сначала
    EnumWindows AddressOf EnumWindowsProc, ByVal JustWithCaption
    ReDim Preserve Windows(1 To WinIndex - 1)
End Function

Function FindWindowByCaption(ByVal Caption As String) As Long
    Dim SplitedWindow() As String, c As Long
    GetWindows 1
    For c = 1 To UBound(Windows)
        If Compare(Left$(Windows(c).Caption, Len(Caption)), Caption) Then
            FindWindowByCaption = Windows(c).hwnd
            Exit Function
        End If
    Next c
    For c = 1 To UBound(Windows)
        If InStr(1, Windows(c).Caption, Caption, vbTextCompare) > 0 Then
            FindWindowByCaption = Windows(c).hwnd
            Exit Function
        End If
    Next c
    SplitedWindow = Split(Caption, " ") ' ЕСЛИ окно не найдено, то идет поиск по отдельным словам
    For i = 1 To UBound(SplitedWindow)
        For c = 1 To UBound(Windows)
            If (Len(SplitedWindow(i)) > 3 Or Len(Caption) < 7) And InStr(1, Windows(c).Caption, SplitedWindow(i), vbTextCompare) > 0 Then
                FindWindowByCaption = Windows(c).hwnd
                Exit Function
            End If
        Next c
    Next i
End Function

Public Function DestroyWindow(hwnd As Long, Optional bForced As Boolean)
    SendMessage hwnd, WM_CLOSE, 0&, 0&
    If bForced Then APIDestroyWindow hwnd
End Function

Sub ShowWindow(hwnd As Long)
    If Not IsWindowEnabled(hwnd) Then EnableWindow hwnd, 1
    SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE
    SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE
    If GetWindowStyle(hwnd) = vbMinimizedFocus Then
        APIShowWindow hwnd, SW_RESTORE
    End If
    SetForegroundWindow hwnd
End Sub

Function GetWindowStyle(hwnd As Long) As VbAppWinStyle
    Dim Width As Long, Height As Long
    If IsWindowVisible(hwnd) = 0 Then GetWindowStyle = vbHide: Exit Function
    GetWindowRect hwnd, Window
    Width = Window.Right - Window.Left
    Height = Window.Bottom - Window.Top
    If Width > (Screen_Width - 15) And Height > (Screen_Height - 30) Then
        GetWindowStyle = vbMaximizedFocus
    ElseIf Window.Left < -10000 Then
        GetWindowStyle = vbMinimizedFocus
    Else
        GetWindowStyle = vbNormalFocus
    End If
End Function

Function IsProcessExists(ByVal pID As Long) As Boolean
    Dim h As Long
    h = OpenProcess(SYNCHRONIZE, 0, pID)
    IsProcessExists = CBool(h)
    If IsProcessExists Then CloseHandle (h)
End Function

Function GetProcessHwnd(ByVal pID As Long) As Long
    If IsProcessExists(pID) = False Then Exit Function
    GetWindows
    For i = 1 To UBound(Windows)
        If GetWindowThreadProcessId(Windows(i).hwnd, 0) = pID Then
            GetProcessHwnd = Windows(i).hwnd
            Exit Function
        End If
    Next i
End Function

Function KillProcessByPID(ByVal pID As Long)
    Shell "cmd /c taskkill /f /pid " & pID, vbHide
End Function

Public Function EnumWindowsProc(ByVal hwnd As Long, ByVal JustWithCaption As Long) As Long
    buff = Space$(GetWindowTextLength(hwnd))
    GetWindowText hwnd, buff, Len(buff) + 1
    If JustWithCaption = 0 Or Len(buff) > 0 Then
        WinIndex = UBound(Windows)
        Windows(WinIndex).Caption = buff
        Windows(WinIndex).hwnd = hwnd
        ReDim Preserve Windows(1 To WinIndex + 1)
    End If
    EnumWindowsProc = 1&
End Function


'Public Function EnumChildProc(ByVal hWnd As Long, ByVal lParam As Long) As Long
'    Dim sSave As String
'    sSave = Space$(GetWindowTextLength(hWnd) + 1)
'    GetWindowText hWnd, sSave, Len(sSave)
'    'remove the last Chr$(0)
'    sSave = Left$(sSave, Len(sSave) - 1)
'    If sSave <> "" Then Form1.Print sSave
'    'continue enumeration
'    EnumChildProc = 1
'End Function
