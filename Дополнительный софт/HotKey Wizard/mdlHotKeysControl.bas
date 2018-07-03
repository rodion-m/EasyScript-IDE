Attribute VB_Name = "mdlHotKeysControl"
Option Explicit

Public HKSArray() As HotKey, Index As Long

Type HotKey
    vKeys() As KeyCodeConstants
    hProcesses() As Long
    CopysRuningCount As Long
    FileName As String
    IfRuning As AlreadyExistEvents
End Type

Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Long

Public Function GetKey(vKey As KeyCodeConstants, Optional AnyKeyPos As Boolean) As Boolean
    If AnyKeyPos Then
        If GetAsyncKeyState(vKey) Then If GetAsyncKeyState(vKey) Then GetKey = True
    Else
        If GetAsyncKeyState(vKey) And &H7FFF& Then GetKey = True
    End If
End Function

Public Function GetKeys(vKeys() As KeyCodeConstants) As Boolean
    If UBound(vKeys) = 0 Then
        'если всего одна клавиша
        GetKeys = GetKey(vKeys(0))
        Exit Function
    End If
    For i = 0 To UBound(vKeys)
        If Not GetKey(vKeys(i), True) Then Exit Function
    Next i
    GetKeys = True
End Function

Function CheckHotKeys() As Boolean
    If (Not Not HKSArray) = 0 Then Exit Function
    For Index = 1 To UBound(HKSArray)
        With HKSArray(Index)
            If (Not Not .vKeys) = 0 Then GoTo lContinue
            CheckProcesses (Index)
            If GetKeys(.vKeys) Then
                Call HotKeyHandler(Index)
                Sleep 500&
            End If
lContinue:
            DoEvents
        End With
    Next Index
    DoEvents
End Function

Sub CheckProcesses(Index As Long)
    With HKSArray(Index)
        If .CopysRuningCount = 0 Then Exit Sub
        If (Not Not .hProcesses) = 0 Then Exit Sub
        For i = 1 To .CopysRuningCount
            If i > UBound(.hProcesses) Then Exit Sub
            If Not IsProcessExists(.hProcesses(i)) Then
                .CopysRuningCount = .CopysRuningCount - 1
                .hProcesses(i) = 0
                DeleteZeroFromArray .hProcesses
                UpdateStatus
                frmMain.UpdateControls
            End If
        Next i
    End With
End Sub

Function HotKeyHandler(Index As Long) As Boolean
    Static antirecurse As Boolean
    If antirecurse Then Exit Function
    antirecurse = True
    frmMain.MousePointer = vbHourglass
    With HKSArray(Index)
        If .CopysRuningCount = 0 Then
            'Если ни один экземпляр не запущен
            ReDim .hProcesses(1 To 1)
            .hProcesses(1) = RunScript(.FileName)
            If .hProcesses(1) > 0 Then
                .CopysRuningCount = 1
            End If
        Else
            Select Case .IfRuning
                Case aeeJust_Kill: Call KillProcess(.hProcesses(1))
                Case aeeJust_Run: Call RunAnortherProcess(Index)
                Case aeeKill_Run ' убиваю текущий и рекурсивно создаю новый процесс
                    Call KillProcess(.hProcesses(1))
                    .CopysRuningCount = 0
                    antirecurse = False
                    Call HotKeyHandler(Index)
                Case aeeNothing  'ничего не делать
            End Select
        End If
    End With
    CheckProcesses (Index)
    UpdateStatus
    frmMain.UpdateControls
    frmMain.MousePointer = vbNormal
    antirecurse = False
End Function

Sub RunAnortherProcess(Index As Long)
    With HKSArray(Index)
        ReDim Preserve .hProcesses(1 To .CopysRuningCount + 1)
        .hProcesses(.CopysRuningCount + 1) = RunScript(.FileName)
        If .hProcesses(.CopysRuningCount + 1) > 0 Then .CopysRuningCount = .CopysRuningCount + 1
    End With
End Sub

Function CreateHotKey(Index As Long, vKeys() As Long) As Boolean
    On Error GoTo lBadKey
    If vKeys(0) = 0 Then CreateHotKey = True: Exit Function
    If (Not Not HKSArray) = 0 Then
        ReDim HKSArray(1 To Index)
    ElseIf UBound(HKSArray) < Index Then
        ReDim Preserve HKSArray(1 To Index)
    End If
    DeleteZeroFromArray vKeys
    If vKeys(0) = 0 Then GoTo lBadKey
    'If KeysAlreadyExists(vKeys) And Index = UBound(HKSArray) Then GoTo lBadKey
    With HKSArray(Index)
        .vKeys = vKeys
        UpdateSubItemsArray GetRealIndex(Index)
        .FileName = SubItemsArray(COLUMN_FILENAME)
        Select Case SubItemsArray(COLUMN_IFRUNING)
            Case STR_Just_Kill: .IfRuning = aeeJust_Kill
            Case STR_Nothing: .IfRuning = aeeNothing
            Case STR_Kill_Run: .IfRuning = aeeKill_Run
            Case STR_Just_Run: .IfRuning = aeeJust_Run
        End Select
    End With
    CreateHotKey = True
Exit Function
lBadKey:
    Err.Clear
    If Index > 1 Then ReDim Preserve HKSArray(1 To (Index - 1))
    Exit Function
End Function

Function KeysAlreadyExists(vKeys() As Long) As Boolean
    Dim lngConfidenceCounter As Long
    If UBound(HKSArray) = 1 Then Exit Function
    For i = 1 To UBound(HKSArray) - 1
        If (Not Not HKSArray(i).vKeys) = 0 Then GoTo lContinue
        lngConfidenceCounter = 0
        For n = 0 To UBound(HKSArray(i).vKeys)
            If n > UBound(vKeys()) Then Exit For
            If HKSArray(i).vKeys(n) = vKeys(n) Then lngConfidenceCounter = lngConfidenceCounter + 1
        Next n
        If lngConfidenceCounter >= (UBound(vKeys) + 1) Then
            KeysAlreadyExists = True
            Exit Function
        End If
lContinue:
    Next i
End Function
