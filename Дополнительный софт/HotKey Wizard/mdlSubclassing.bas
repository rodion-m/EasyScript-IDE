Attribute VB_Name = "mdlSubclassing"
Option Explicit

Type POINTAPI
    X As Long
    Y As Long
End Type

Private Type MINMAXINFO
    ptReserved As POINTAPI
    ptMaxSize As POINTAPI
    ptMaxPosition As POINTAPI
    ptMinTrackSize As POINTAPI
    ptMaxTrackSize As POINTAPI
End Type

Private Declare Sub GetMem4 Lib "msvbvm60.dll" (pSrc As Any, pDst As Any)
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long


Public MainFormPrevProc As Long
Dim Info As MINMAXINFO, pt As POINTAPI

' ************************************************************** Subclassing ***********************************************************************
Public Const WM_CONTEXTMENU = &H7B
Public Const WM_KEYDOWN As Long = &H100
Public Const WM_KEYUP As Long = &H101
Public Const WM_CHAR As Long = &H102
Public Const WM_GETMINMAXINFO = &H24
Public Const GWL_WNDPROC = (-4)
' *****************************************************************************

Public Function wpMainForm(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    If uMsg = WM_GETMINMAXINFO Then
        SetMinSize Info, lParam, 650, 200
        wpMainForm = 1
    ElseIf uMsg = 0 And wParam = 0 And lParam = 0 Then
        frmMain.LoadCommandLine (ReadIni("Command", "null"))
        frmMain.ShowMe
        wpMainForm = 1
    Else
        wpMainForm = CallWindowProc(MainFormPrevProc, hwnd, uMsg, wParam, lParam)
    End If
End Function

'процедура для заполнения структуры
Private Sub SetMinSize(Info As MINMAXINFO, ByVal lParam As Long, X As Long, Y As Long)
    Call GetMem4(lParam, ByVal VarPtr(lParam) - 4)
    With Info
        .ptMinTrackSize.X = X
        .ptMinTrackSize.Y = Y
    End With
End Sub
