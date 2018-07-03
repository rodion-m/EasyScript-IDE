Attribute VB_Name = "modManifest"
Option Explicit

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
