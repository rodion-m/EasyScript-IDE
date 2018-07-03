Attribute VB_Name = "mdlSubclassing"
Option Explicit

Private Type MINMAXINFO
    ptReserved As POINTAPI
    ptMaxSize As POINTAPI
    ptMaxPosition As POINTAPI
    ptMinTrackSize As POINTAPI
    ptMaxTrackSize As POINTAPI
End Type

Private Declare Sub GetMem4 Lib "msvbvm60.dll" (pSrc As Any, pDst As Any)
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)

Public MainFormPrevProc As Long, ScriptTextBoxPrevProc As Long
Dim Info As MINMAXINFO, pt As POINTAPI

' ************************************************************** Subclassing ***********************************************************************
Public Const WM_PASTE = 302&
Public Const WM_CONTEXTMENU = &H7B
Public Const WM_KEYDOWN As Long = &H100
Public Const WM_KEYUP As Long = &H101
Public Const WM_CHAR As Long = &H102
Public Const WM_GETMINMAXINFO = &H24
Public Const GWL_WNDPROC = (-4)
' *****************************************************************************

Public Function wpMainForm(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    If uMsg = WM_GETMINMAXINFO Then
        SetMinSize Info, lParam, 375, 300
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

Function wpScriptTextBox(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Select Case uMsg
        Case WM_PASTE
            frmMain.Paste
            wpScriptTextBox = 1
        Case WM_CONTEXTMENU
            frmMain.UpdateControlsState
            frmMain.PopupMenu frmMain.mnuEdit
            wpScriptTextBox = 1
'        Case WM_KEYDOWN
'            If Not frmMain.lstToolTip.Visible Then GoTo lDefault
'            Select Case lParam
'                Case 21495809 ' VK_UP
'                    If frmMain.lstToolTip.ListIndex > 0 Then frmMain.lstToolTip.ListIndex = frmMain.lstToolTip.ListIndex - 1
'                Case 22020097 'VK_DOWN
'                    If frmMain.lstToolTip.ListIndex < (frmMain.lstToolTip.ListCount - 1) Then frmMain.lstToolTip.ListIndex = frmMain.lstToolTip.ListIndex + 1
'                Case 1835009 'VK_RETURN
'                    frmMain.SelectFunc
'                Case Else
'                    GoTo lDefault
'            End Select
        Case Else
lDefault:
            wpScriptTextBox = CallWindowProc(ScriptTextBoxPrevProc, hwnd, uMsg, wParam, lParam)
    End Select
End Function
