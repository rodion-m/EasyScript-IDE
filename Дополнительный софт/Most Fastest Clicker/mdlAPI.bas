Attribute VB_Name = "mdlAPI"
Option Explicit

Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Public Const ES_NUMBER As Long = &H2000&
Public Const GWL_STYLE As Long = -16

Type POINTAPI
    x As Long
    y As Long
End Type

Declare Function GetCursorPosPT& Lib "user32" Alias "GetCursorPos" (lpPoint As POINTAPI)

Declare Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal y As Long) As Long
Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)

Public Const MOUSEEVENTF_LEFTDOWN = &H2
Public Const MOUSEEVENTF_LEFTUP = &H4

Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Long

Private Type InitCommonControlsEx
    dwSize As Long
    dwICC As Long
End Type

Private Declare Function InitCommonControlsEx Lib "comctl32.dll" (ByRef TLPINITCOMMONCONTROLSEX As InitCommonControlsEx) As Long

Public Function InitCommonControls() As Boolean
    On Error Resume Next
    Const ICC_USEREX_CLASSES = &H200&
    Dim udtICCEx As InitCommonControlsEx
    udtICCEx.dwSize = Len(udtICCEx)
    udtICCEx.dwICC = ICC_USEREX_CLASSES
    InitCommonControls = InitCommonControlsEx(udtICCEx)
End Function
