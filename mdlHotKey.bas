Attribute VB_Name = "mdlHotKey"
Option Explicit

Public HKSArray() As HotKey, Index As Long

Enum HKIndex
    hkExecuteScript = 1
    hkStartRec = 2
    hkSaveAs = 3
    hkInsClick = 4
    hkInsCursorPos = 5
End Enum

Enum HKType
    PublicHK = 0
    PrivateHK = 1
End Enum

Type HotKey
    vKeys() As KeyCodeConstants
    tType As HKType
    hwnd As Long
End Type

Function CheckHotKeys(Optional nIndex As HKIndex) As Boolean
    Dim TrueHotKey As Boolean
    If CommandLine.DisableHKs Then Exit Function
    For Index = 1 To UBound(HKSArray)
        If GetKeys(HKSArray(Index).vKeys) Then
            If HKSArray(Index).tType = PrivateHK Then 'проверка на локальность горячей клавиши
                If GetForegroundWindow <> HKSArray(Index).hwnd Then GoTo lNext
            End If
            If blnRec And Not blnExecuting And Index = hkExecuteScript Then Exit Function
            If blnExecuting And Not blnRec And Index = hkStartRec Then Exit Function
            If Index = hkExecuteScript And (Not IsValidScript(frmMain.txtMain)) Then Exit Function
            Call PreHotKeyHandler(Index)
            Sleep 300
            Call HotKeyHandler(Index)
            If Index = nIndex Then CheckHotKeys = True
        End If
lNext:
    Next Index
End Function

Function PreHotKeyHandler(Index As HKIndex, Optional vk As Long) As Boolean
    Select Case Index
        Case hkExecuteScript
            If Not blnExecuting And Not blnRec Then
                frmMain.SetExecState
                PreHotKeyHandler = True
            ElseIf blnExecuting Then
                frmMain.SetNormalState
                PreHotKeyHandler = True
            End If
        Case hkStartRec
            If Not blnRec And Not blnExecuting Then
                frmMain.SetRecState
                PreHotKeyHandler = True
            ElseIf blnRec Then
                frmMain.SetNormalState
                PreHotKeyHandler = True
            End If
    End Select
End Function

Function HotKeyHandler(Index As HKIndex, Optional vk As Long) As Boolean
    Select Case Index
        Case hkExecuteScript
            If Not blnRec And Not blnExecuting Then
                ExecuteScript frmMain.txtMain, frmMain.txtMain, frmMain.Arrow, frmMain.txtStatus
            ElseIf blnExecuting Then
                blnExecuting = False
                blnRec = False
            End If
        Case hkStartRec
            If Not blnRec And Not blnExecuting Then
                StartRec
            ElseIf blnRec Then
                blnExecuting = False
                blnRec = False
            End If
        Case hkSaveAs
            Call frmMain.mnuSaveAs_Click
        Case hkInsClick
            If blnMarkerShowed Then frmMain.InsertFunc "Клик(" & frmMarker.GetMarkerX & ", " & frmMarker.GetMarkerY & ")"
        Case hkInsCursorPos
            If blnMarkerShowed Then frmMain.InsertFunc "Переместить курсор(" & frmMarker.GetMarkerX & ", " & frmMarker.GetMarkerY & ")"
    End Select
End Function

Function CreateHotKey(Index As HKIndex, Key1 As KeyCodeConstants, _
                                        Optional Key2 As KeyCodeConstants, _
                                        Optional Key3 As KeyCodeConstants, _
                                        Optional tType As HKType = PublicHK, _
                                        Optional hwnd As Long) As Boolean
    If (Not Not HKSArray) <> 0 Then
        For i = LBound(HKSArray) To UBound(HKSArray)
            If (Not Not HKSArray(i).vKeys) <> 0 Then  'проверка, существует ли горячая клавиша. ибо если существует индекс, еще не факт, что есть клавиша
                If HKSArray(i).vKeys(1) = Key1 And HKSArray(i).vKeys(2) = Key2 And HKSArray(i).vKeys(3) = Key3 Then
                    If i = Index Then CreateHotKey = True
                    Exit Function
                End If
            End If
        Next i
    End If
    frmMain.tmrCheckHKs = False
    If (Not Not HKSArray) = 0 Then
        ReDim Preserve HKSArray(1 To Index)
    ElseIf UBound(HKSArray) < Index Then
        ReDim Preserve HKSArray(1 To Index)
    End If
    ReDim HKSArray(Index).vKeys(1 To 3)
    HKSArray(Index).vKeys(1) = Key1
    HKSArray(Index).vKeys(2) = Key2
    HKSArray(Index).vKeys(3) = Key3
    HKSArray(Index).tType = tType
    HKSArray(Index).hwnd = hwnd
    frmMain.tmrCheckHKs = True
    CreateHotKey = True
End Function
