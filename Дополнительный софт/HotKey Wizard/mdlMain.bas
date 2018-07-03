Attribute VB_Name = "mdlMain"
Option Explicit
Option Compare Text

Public Const COLUMN_HEADERS_COUNT = 7, COLUMN_NUMBER = 0, COLUMN_TITLE = 1, COLUMN_FILENAME = 2, _
                    COLUMN_HKS = 3, COLUMN_STATUS = 4, COLUMN_RUNCOUNT = 5, COLUMN_IFRUNING = 6

Public Const BAD_HOTKEY = "нет"
Public Const STATUS_RUNNING = "Запущен", STATUS_NOTRUNNING = "Остановлен", STATUS_BAD = "Ошибка"
Public Const ICON_STATUS_RUNNING = 2, ICON_STATUS_NOTRUNNING = 1, ICON_STATUS_BAD = 3
Public Const ICON_SYSTRAY_RUNNING = 5, ICON_SYSTRAY_NOTRUNNING = 4
Global i As Long, n As Long, Arr, List As ListView, SubItemsArray() As String, Editor_FN As String, _
ProgrammLoaded As Boolean, InitDir As String, ScriptsPath As String, ScriptsList_FN As String, blnRuning As Boolean

Public Const STR_Just_Kill = "Завершить скрипт", STR_Nothing = "Ничего не делать", _
                    STR_Kill_Run = "Завершить предыдущий и запустить", STR_Just_Run = "Запустить еще один экзмепляр"

Enum AlreadyExistEvents
    aeeJust_Kill
    aeeNothing
    aeeKill_Run
    aeeJust_Run
End Enum

Enum ChangeHotKeyModes
    chkmCreate = 1
    chkmChange = 2
End Enum

Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Any) As Long
'Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long

Private Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAcess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Const SYNCHRONIZE As Long = &H100000

Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function ShellExecuteEx Lib "shell32.dll" (SEI As SHELLEXECUTEINFO) As Long
Const SEE_MASK_INVOKEIDLIST = &HC
Const SEE_MASK_NOCLOSEPROCESS = &H40
Const SEE_MASK_FLAG_NO_UI = &H400
Private Type SHELLEXECUTEINFO
    cbSize As Long
    fMask As Long
    hwnd As Long
    lpVerb As String
    lpFile As String
    lpParameters As String
    lpDirectory As String
    nShow As Long
    hInstApp As Long
    lpIDList As Long
    lpClass As String
    hkeyClass As Long
    dwHotKey As Long
    hIcon As Long
    hProcess As Long
End Type

Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long

Type RECT
   Left    As Long
   Top     As Long
   Right   As Long
   Bottom  As Long
End Type

Declare Function GetForegroundWindow Lib "user32" () As Long
Declare Function SetForegroundWindow Lib "user32.dll" (ByVal hwnd As Long) As Long
Declare Function APIShowWindow Lib "user32" Alias "ShowWindow" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Sub SetWindowPos Lib "user32" _
    (ByVal hwnd As Long, ByVal hWndInsertAfter _
        As Long, ByVal X As Long, ByVal Y As Long, _
            ByVal cx As Long, ByVal cy As Long, _
                ByVal wFlags As Long)

Const SWP_NOSIZE = &H1&
Const SWP_NOMOVE = &H2&
Const SW_RESTORE As Long = 9
Const HWND_NOTOPMOST As Long = -2
Const HWND_TOPMOST As Long = -1
Const VK_STARTKEY = &H5B&

Sub ShowWindow(hwnd As Long)
    SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE
    SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE
    If GetWindowStyle(hwnd) = vbMinimizedFocus Then
        APIShowWindow hwnd, SW_RESTORE
    End If
    SetForegroundWindow hwnd
End Sub

Function GetWindowStyle(hwnd As Long) As VbAppWinStyle
    Dim Width As Long, Height As Long, Window As RECT, Screen_Width As Long, Screen_Height As Long
    GetWindowRect hwnd, Window
    Width = Window.Right - Window.Left
    Height = Window.Bottom - Window.Top
    Screen_Width = Screen.Width \ Screen.TwipsPerPixelX
    Screen_Height = Screen.Height \ Screen.TwipsPerPixelY
    If Width > (Screen_Width - 15) And Height > (Screen_Height - 30) Then
        GetWindowStyle = vbMaximizedFocus
    ElseIf Window.Left < -10000 Then
        GetWindowStyle = vbMinimizedFocus
    Else
        GetWindowStyle = vbNormalFocus
    End If
End Function

Public Sub AddScript(ByVal Title As String, ByVal FileName As String, ByVal HKs As String, ByVal IfRuning As String, Optional Index As Long)
    Dim RealIndex As Long
    If Index = 0 Then
        For i = 1 To List.ListItems.Count
            If GetRealIndex(i) = 0 Then
                Index = i
                Exit For
            End If
        Next i
        If Index = 0 Then Index = List.ListItems.Count + 1
        RealIndex = Index
    Else
        RealIndex = GetRealIndex(Index)
        If RealIndex <= List.ListItems.Count Then List.ListItems.Remove (RealIndex)
    End If
    List.ListItems.Add RealIndex, , CStr(Index), , ICON_STATUS_NOTRUNNING
    List.ListItems.Item(RealIndex).SubItems(COLUMN_TITLE) = Title
    List.ListItems.Item(RealIndex).SubItems(COLUMN_FILENAME) = FileName
    List.ListItems.Item(RealIndex).SubItems(COLUMN_HKS) = HKs
    List.ListItems.Item(RealIndex).SubItems(COLUMN_STATUS) = STATUS_NOTRUNNING
    List.ListItems.Item(RealIndex).SubItems(COLUMN_RUNCOUNT) = "0"
    List.ListItems.Item(RealIndex).SubItems(COLUMN_IFRUNING) = IfRuning
    AddHotKey Index, HKs
End Sub

Sub AddHotKey(Index As Long, ByVal HKs As String)
    Dim sKeys() As String, vKeys() As Long
    sKeys = Split(HKs, "+")
    ReDim vKeys(UBound(sKeys))
    For n = 0 To UBound(sKeys)
        vKeys(n) = GetVK(sKeys(n))
    Next n
    If Not CreateHotKey(Index, vKeys) Then
        List.ListItems.Item(GetRealIndex(Index)).SmallIcon = ICON_STATUS_BAD
    End If
    UpdateStatus
End Sub


Function AddKeysToComboBox(ParamArray VarList())
    Dim List As Variant
    For Each List In VarList
        List.Clear
        AddItemsToComboBox List, "Ctrl", "Alt", "Shift", "Win", "Tab", "Space"
        For n = Asc("A") To Asc("Z")
            List.AddItem Chr$(n)
        Next n
        For n = Asc("0") To Asc("9")
            List.AddItem Chr$(n)
        Next n
        For n = 1 To 12
            List.AddItem "F" & n
        Next n
        AddItemsToComboBox List, "=", "-", "/", "\", "[", "]", ";", ",", ".", "'", "`"
        List.AddItem "нет"
    Next
End Function

Sub AddItemsToComboBox(ByRef List, ParamArray Items())
    Dim Item
    For Each Item In Items
        List.AddItem Item
    Next
End Sub

Function GetVK(ByVal str As String) As KeyCodeConstants
    If Len(str) = 1 Then
        str = UCase(str)
        If IsNumeric(str) Then
            GetVK = Asc(str)
        ElseIf Asc(str) >= Asc("A") And Asc(str) <= Asc("Z") Then
            GetVK = Asc(str)
        Else
            Select Case str
                Case "=": GetVK = 187
                Case "-": GetVK = 189
                Case "/": GetVK = 191
                Case "\": GetVK = 220
                Case "[": GetVK = 219
                Case "]": GetVK = 221
                Case ";": GetVK = 186
                Case ",": GetVK = 188
                Case ".": GetVK = 190
                Case "'": GetVK = 222
                Case "`": GetVK = 192
                Case Else: GetVK = -1
            End Select
        End If
    ElseIf Left$(str, 1) = "F" And IsNumeric(Replace(str, "F", "")) Then
        If CLng(Replace(str, "F", "")) >= 1 And CLng(Replace(str, "F", "")) <= 12 Then
            GetVK = 111 + CLng(Replace(str, "F", ""))
        End If
    Else
        str = Replace(str, "{", "")
        str = Replace(str, "}", "")
        Select Case str
            Case "Ctrl": GetVK = vbKeyControl
            Case "Alt": GetVK = vbKeyMenu
            Case "Shift": GetVK = vbKeyShift
            Case "Win": GetVK = VK_STARTKEY
            Case "Tab": GetVK = vbKeyTab
            Case "Space": GetVK = vbKeySpace
            Case "нет": GetVK = 0
            Case Else: GetVK = -1
        End Select
    End If
End Function

Function Bool2Lng(Bool As Boolean) As Long
    If Bool Then Bool2Lng = 1 Else Bool2Lng = 0
End Function

Sub CopyArray(Source(), Dest() As KeyCodeConstants)
    Dim Index As Long
    Index = -1
    For Each Arr In Source
        If CLng(Arr) > 0 Then
            Index = Index + 1
            ReDim Preserve Dest(Index)
            Dest(Index) = CLng(Arr)
        End If
    Next
End Sub

Sub DeleteZeroFromArray(Arr() As Long)
    Dim blnZero As Boolean, i As Long
    For i = LBound(Arr) To UBound(Arr)
        If Arr(i) = 0& Then
            blnZero = True
            Arr(i) = UBound(Arr)
            Exit For
        End If
    Next i
    If blnZero Then
        If UBound(Arr) = LBound(Arr) Then Exit Sub
        ReDim Preserve Arr(LBound(Arr) To UBound(Arr) - 1)
        Call DeleteZeroFromArray(Arr()) ' рекурсивный проход, удаление каждого нуля
    End If
End Sub

Sub BubbleSortKeys(Arr() As KeyCodeConstants)
    Dim tmp As Long
    For i = LBound(Arr) To UBound(Arr) - 1
        If Arr(i) > Arr(i + 1) Then
            tmp = Arr(i)
            Arr(i) = Arr(i + 1)
            Arr(i + 1) = tmp
            i = LBound(Arr) - 1
        End If
    Next i
End Sub

Sub GetListSubItems(ByVal RealIndex As Long, ByRef GetTo() As String)
    Dim i As Long
    ReDim GetTo(0)
    GetTo(0) = List.ListItems.Item(RealIndex).Text
    For i = 1 To COLUMN_HEADERS_COUNT - 1
        ReDim Preserve GetTo(i)
        GetTo(i) = List.ListItems.Item(RealIndex).SubItems(i)
    Next i
End Sub

Sub UpdateSubItemsArray(Optional RealIndex As Long)
    If RealIndex = 0 Then RealIndex = GetSelectedItem
    GetListSubItems RealIndex, SubItemsArray()
End Sub

Function GetSelectedItem() As Long 'Optional List As ListView) As Long
    On Error GoTo EH
    GetSelectedItem = List.SelectedItem.Index
    If List.ListItems.Item(GetSelectedItem).Selected = False Then GetSelectedItem = 0
    Exit Function
EH:
    Err.Clear
    GetSelectedItem = 0
End Function

Function IsVirtKey(ByVal str As String) As Boolean
    If GetVK(str) <> -1 Then IsVirtKey = True
End Function

Function CChk(s) As Integer 'функция либо возвращает 0 либо 1
    If CBool(s) = True Then CChk = 1 Else CChk = 0
End Function

Function RunScript(ByVal FileName As String) As Long
    On Error Resume Next
    CheckScriptName (FileName)
    If Right$(FileName, 3) = ".es" Then
        If IsFileExists(Editor_FN) Then
            RunScript = APIShell(Editor_FN, "/hm /fn " & FileName)
        Else
            ShowNotFound (Editor_FN)
        End If
    Else 'тупо запускаю
        RunScript = APIShell(FileName, "", vbNormalFocus)
    End If
End Function

Sub UpdateStatus()
    Dim blnSetZero As Boolean, Index As Long
    blnRuning = False
    If (Not Not HKSArray) = 0 Then blnSetZero = True
    For i = 1 To List.ListItems.Count ' i - это RealIndex
        Index = GetIndex(i)
        If blnSetZero Then
            SetNotRunningStatus (i)
        ElseIf i <= UBound(HKSArray) Then
            If HKSArray(Index).CopysRuningCount > 0 Then
                blnRuning = True
                SetRunningStatus (i)
            Else
                SetNotRunningStatus (i)
            End If
        Else
            SetNotRunningStatus (i)
        End If
    Next i
End Sub

Sub SetBadStatus(ByVal RealIndex As Long)
    UpdateSubItemsArray RealIndex
    If SubItemsArray(COLUMN_STATUS) <> STATUS_BAD Then
        List.ListItems.Item(RealIndex).SubItems(COLUMN_STATUS) = STATUS_BAD
        List.ListItems.Item(RealIndex).SubItems(COLUMN_RUNCOUNT) = "0"
        List.ListItems(RealIndex).SmallIcon = ICON_STATUS_BAD
    End If
End Sub

Sub SetNotRunningStatus(ByVal RealIndex As Long)
    UpdateSubItemsArray RealIndex
    If SubItemsArray(COLUMN_HKS) = BAD_HOTKEY Then SetBadStatus (RealIndex): Exit Sub
    If Not IsFileExists(SubItemsArray(COLUMN_FILENAME)) Then SetBadStatus (RealIndex): Exit Sub
    If SubItemsArray(COLUMN_STATUS) <> STATUS_NOTRUNNING Then
        List.ListItems.Item(RealIndex).SubItems(COLUMN_STATUS) = STATUS_NOTRUNNING
        List.ListItems.Item(RealIndex).SubItems(COLUMN_RUNCOUNT) = "0"
        List.ListItems(RealIndex).SmallIcon = ICON_STATUS_NOTRUNNING
    End If
End Sub

Sub SetRunningStatus(ByVal RealIndex As Long)
    UpdateSubItemsArray RealIndex
    If SubItemsArray(COLUMN_RUNCOUNT) <> HKSArray(GetIndex(RealIndex)).CopysRuningCount Then
        List.ListItems.Item(RealIndex).SubItems(COLUMN_STATUS) = STATUS_RUNNING
        List.ListItems.Item(RealIndex).SubItems(COLUMN_RUNCOUNT) = HKSArray(GetIndex(RealIndex)).CopysRuningCount
        List.ListItems(RealIndex).SmallIcon = ICON_STATUS_RUNNING
    End If
End Sub

Function GetRealIndex(Index As Long) As Long
    Dim i As Long
    For i = 1 To List.ListItems.Count
        UpdateSubItemsArray i
        If SubItemsArray(COLUMN_NUMBER) = CStr(Index) Then GetRealIndex = i: Exit Function
    Next i
End Function

Function GetIndex(RealIndex As Long) As Long
    UpdateSubItemsArray RealIndex
    GetIndex = CLng(SubItemsArray(COLUMN_NUMBER))
End Function

Sub KillProcess(hProcess As Long)
    TerminateProcess hProcess, 0&
End Sub

Sub ShowNotFound(ByVal FileName As String)
    MsgBox "Файл " & FileName & " не найден. Проверьте правильность имени файла.", vbExclamation, "Файл не найден"
End Sub

Function APIShell(FileName As String, Optional Parameters As String, Optional Mode As VbAppWinStyle = vbHide) As Long
    Dim SEI As SHELLEXECUTEINFO
    With SEI
        .cbSize = Len(SEI)
        .fMask = SEE_MASK_NOCLOSEPROCESS 'Or SEE_MASK_INVOKEIDLIST Or SEE_MASK_FLAG_NO_UI
        .hwnd = frmMain.hwnd
        .lpVerb = "open"
        .lpFile = FileName
        .lpParameters = Parameters
        .nShow = Mode
    End With
    If ShellExecuteEx(SEI) = 1 Then APIShell = SEI.hProcess
End Function

Function IsProcessExists(ByVal hProcess As Long) As Boolean
    If hProcess = 0 Then Exit Function
    IsProcessExists = CBool(WaitForSingleObject(hProcess, 0)) 'если хендла не сущесуствует - возвращает 0
End Function

Sub KillProcesses(Index As Long)
    If (Not Not HKSArray) = 0 Then Exit Sub
    If Index > UBound(HKSArray) Then Exit Sub
    With HKSArray(Index)
        If .CopysRuningCount = 0 Then Exit Sub
        If (Not Not .hProcesses) = 0 Then Exit Sub
        For i = 1 To .CopysRuningCount
            KillProcess .hProcesses(i)
        Next i
        .CopysRuningCount = 0
    End With
    UpdateStatus
    frmMain.UpdateControls
End Sub

Function UpdateAllHotKeys() As Boolean
    Dim sKeys() As String, vKeys() As Long, Index As Long, i As Long
    ReDim HKSArray(1 To 1)
    UpdateAllHotKeys = True
    For i = 1 To List.ListItems.Count
        Index = GetIndex(i)
        UpdateSubItemsArray i
        sKeys = Split(SubItemsArray(COLUMN_HKS), "+")
        ReDim vKeys(UBound(sKeys))
        For n = 0 To UBound(sKeys)
            vKeys(n) = GetVK(sKeys(n))
        Next n
        If Not CreateHotKey(Index, vKeys) Then
            AddScript SubItemsArray(COLUMN_TITLE), SubItemsArray(COLUMN_FILENAME), BAD_HOTKEY, SubItemsArray(COLUMN_IFRUNING), Index
            UpdateStatus
            If Not ProgrammLoaded Then MsgBox "Не удается задать комбинацию горячих клавиш для скрипта №" & Index, vbExclamation, "Не удается задать клавиши"
        End If
    Next i
    UpdateStatus
End Function

Sub ShowFileInDir(ByVal FileName As String)
    CheckScriptName FileName
    Shell "explorer /select," & FileName, vbNormalFocus
End Sub

Sub LoadSettings()
    #If COMPILED Then
        Editor_FN = App.path & "\Editor.exe"
    #Else
        Editor_FN = "E:\Programming\vb\Easy Script\Editor.exe"
    #End If
    ScriptsPath = App.path & "\Scripts\"
    List.SmallIcons = frmMain.ImageList
    LoadFormSettings frmMain
    InitDir = ReadIni(InitDir, ScriptsPath)
    With frmMain
        .mnuAutoRun.Checked = Autorun_GetStatus
        .mnuEqvHotKeys.Checked = ReadIni("EqvHotKeys", 0, , dChk)
        .mnuTrayOn.Checked = ReadIni("TrayOn", 1, dChk)
    End With
    ScriptsList_FN = ReadIni("ScriptsList_FN", App.path & "\List.rshw")
    For i = 1 To List.ColumnHeaders.Count
        List.ColumnHeaders.Item(i).Width = ReadIni("Column" & i, List.ColumnHeaders.Item(i).Width, , dNumeric)
    Next i
End Sub

Sub LoadFormSettings(Form As Form)
    Dim wState As Single, strSection As String
    strSection = Form.LinkTopic
    wState = ReadIni("WindowState", 0, strSection, dNumeric)
    With Form
        .WindowState = wState
        If wState = 0 Then
            .Left = ReadIni("Left", Screen.Width / 2 - .Width / 2, strSection, dNumeric)
            .Top = ReadIni("Top", Screen.Height / 2 - .Height / 2, strSection, dNumeric)
            .Width = ReadIni("Width", .Width, strSection, dNumeric)
            .Height = ReadIni("Height", .Height, strSection, dNumeric)
        End If
    End With
End Sub

Sub SaveSettings()
    SaveFormSettings frmMain
    WriteIni "InitDir", InitDir
    WriteIni "ScriptsList_FN", ScriptsList_FN
    WriteIni "EqvHotKeys", frmMain.mnuEqvHotKeys.Checked
    WriteIni "TrayOn", frmMain.mnuTrayOn.Checked
    For i = 1 To List.ColumnHeaders.Count
        WriteIni "Column" & i, List.ColumnHeaders.Item(i).Width
    Next i
End Sub

Sub SaveFormSettings(Form As Form)
    Dim strSection As String
    strSection = Form.LinkTopic
    With Form
        WriteIni "WindowState", .WindowState, strSection
        WriteIni "Left", .Left, strSection
        WriteIni "Top", .Top, strSection
        WriteIni "Width", .Width, strSection
        WriteIni "Height", .Height, strSection
    End With
End Sub

Function IsHotKeyExists(ByVal sKeys As String) As Long
    ' Возвращает индекс существующей комбинации
    If sKeys = "" Then IsHotKeyExists = 1
    For i = 1 To List.ListItems.Count
        UpdateSubItemsArray i
        If CompareKeys(sKeys, SubItemsArray(COLUMN_HKS)) Then
            IsHotKeyExists = GetIndex(i)
            Exit Function
        End If
    Next i
End Function

Function CompareKeys(ByVal sKeys1 As String, ByVal sKeys2 As String) As Boolean ' 1 - новый , 2 - существующий
    Dim KeysArr1() As String, KeysArr2() As String, lngConfidenceCounter As Long, i As Long
    KeysArr1 = Split(sKeys1, "+")
    KeysArr2 = Split(sKeys2, "+")
    For i = 0 To UBound(KeysArr2)
        For n = 0 To UBound(KeysArr1)
            If (n > UBound(KeysArr2)) Then Exit For
            If KeysArr1(n) = KeysArr2(i) Then lngConfidenceCounter = lngConfidenceCounter + 1
        Next n
    Next i
    If lngConfidenceCounter >= (UBound(KeysArr1) + 1) Then
        CompareKeys = True
        Exit Function
    End If
End Function

Sub LoadList()
    Dim strBody As String, strLines() As String, strBuffer() As String
    List.ListItems.Clear
    strBody = ReadFile(ScriptsList_FN)
    If strBody = "" Then Exit Sub
    strLines() = Split(strBody, vbCrLf)
    For Each Arr In strLines
        If Arr = "" Then GoTo lNext
        strBuffer = Split(Arr, vbTab)
        If UBound(strBuffer) = 3 Then
            AddScript strBuffer(0), strBuffer(1), strBuffer(2), strBuffer(3)
        End If
lNext:
    Next
End Sub

Sub SaveList()
    Dim strLine As String, strBody As String
    For i = 1 To List.ListItems.Count
        UpdateSubItemsArray i
        strLine = SubItemsArray(COLUMN_TITLE) & vbTab & SubItemsArray(COLUMN_FILENAME) & vbTab & _
                        SubItemsArray(COLUMN_HKS) & vbTab & SubItemsArray(COLUMN_IFRUNING)
        AddLineToString strBody, strLine
    Next i
    CreateFile ScriptsList_FN, strBody
End Sub

Sub AddLineToString(ByRef strDest As String, ByVal strLine As String)
    If strDest = "" Then strDest = strLine: Exit Sub
    If Right$(strDest, 2) <> vbCrLf Then strLine = vbCrLf & strLine
    strDest = strDest & strLine
End Sub

Function AutoAdd(ByVal path As String, ByVal sKeys As String, ByVal IfRuning As String) As Boolean
    On Error GoTo EH
    Dim FileName As String
    If Not IsFileMaskExistsInDir(path) Then Exit Function
    If Right$(path, 1) <> "\" Then path = path & "\"
    sKeys = sKeys & "+"
    FileName = Dir(path, vbNormal)
    List.ListItems.Clear
    Do
        If Right$(FileName, 3) = ".es" Then
            sKeys = GetNextHotKey(sKeys)
            AddScript GetFileBaseTitle(FileName), GetFileBaseName(FileName), sKeys, IfRuning
        End If
        FileName = Dir
    Loop While FileName <> ""
    ScriptsList_FN = App.path & "\List.rshw"
    UpdateAllHotKeys
    SaveList
    AutoAdd = True
    Exit Function
EH:
    Err.Clear
    AutoAdd = False
End Function

Function IsFileMaskExistsInDir(ByVal path As String, Optional Mask As String = ".es") As Boolean
    On Error GoTo EH
    Dim FileName As String
    If Not IsValidePath(path) Then Exit Function
    If Right$(path, 1) <> "\" Then path = path & "\"
    FileName = Dir(path, vbNormal)
    Do While FileName <> ""
        If Right$(FileName, Len(Mask)) = Mask Then
            IsFileMaskExistsInDir = True
            Exit Function
        End If
        FileName = Dir
    Loop
    Exit Function
EH:
    Err.Clear
    IsFileMaskExistsInDir = False
End Function

Function IsValidePath(ByVal path As String) As Boolean
    On Error Resume Next
    If Len(path) < 4 Then Exit Function
    If Right$(Left$(path, 3), 2) <> ":\" Then Exit Function
    Call Dir(path)
    If Err.Number <> 0 Then Err.Clear: Exit Function
    IsValidePath = True
End Function

Function AddScriptFileToList(ByVal FileName As String, Optional ByVal Title As String, Optional ByVal sKeys As String, Optional ByVal IfRuning As String) As Boolean
    If Not IsFileExists(FileName) Then Call ShowNotFound(FileName): Exit Function
    If Title = "" Then Title = GetFileBaseTitle(FileName)
    If IfRuning = "" Then IfRuning = STR_Just_Kill
    SortKeys sKeys
    If IsHotKeyExists(sKeys) Then
        sKeys = GetFreeHotKey
    End If
    AddScript Title, FileName, sKeys, IfRuning
End Function

Function GetFreeHotKey() As String
    Dim sKeys As String
    ' назначаю клавишу по умолчаниюIf
    Do
        sKeys = GetNextHotKey(sKeys)
    Loop While (IsHotKeyExists(sKeys) > 0)
    GetFreeHotKey = sKeys
End Function

Sub SortKeys(ByRef sKeys As String)
    Dim sKeysArr() As String, Index As Long
    If Len(sKeys) < 4 Then Exit Sub
    If InStr(1, sKeys, "+") = 0 Then Exit Sub
    If Left$(Right$(sKeys, 2), 1) = "+" And GetCount(sKeys, "+") = 1 Then Exit Sub
    sKeysArr = Split(sKeys, "+")
    sKeys = ""
    For n = 0 To UBound(sKeysArr)
        Index = n
        For i = 0 To UBound(sKeysArr)
            If Len(sKeysArr(Index)) < Len(sKeysArr(i)) Or sKeysArr(i) = "Ctrl" Then Index = i
        Next i
        If sKeys = "" Then sKeys = sKeysArr(Index) Else sKeys = sKeys & "+" & sKeysArr(Index)
        sKeysArr(Index) = ""
    Next n
End Sub

Function GetCount(ByVal str As String, ByVal Find As String) As Long
    GetCount = UBound(Split(str, Find))
End Function

Sub ShowFormEx(ByRef Form As Form, ByVal Modal As Long, Owner As Form, ByRef FormToUnload As Form)
    Unload FormToUnload
    Form.Show Modal, Owner
End Sub

Function GetNextHotKey(ByVal sKeys As String) As String
    Dim CharNum As Long, sKeysArr() As String
    If sKeys = "" Then sKeys = "Ctrl+0"
    If InStr(1, sKeys, "+") = 0 Then sKeys = sKeys & "+0"
    CharNum = Asc(Right$(sKeys, 1))
    If ((CharNum >= vbKey0) And (CharNum < vbKey9)) Or ((CharNum >= vbKeyA) And (CharNum < vbKeyZ)) Then
        CharNum = CharNum + 1
    Else
        sKeysArr = Split(sKeys, "+")
        If sKeysArr(0) = "Ctrl" Then sKeysArr(1) = "Alt" Else sKeysArr(0) = "Ctrl"
        sKeysArr(UBound(sKeysArr)) = "1"
        GetNextHotKey = Join(sKeysArr, "+")
        Exit Function
    End If
    GetNextHotKey = Left$(sKeys, Len(sKeys) - 1) & Chr$(CharNum)
End Function

Sub CheckUpdate(Optional HideMode As Boolean = False, Optional ByVal hwnd As Long)
    On Error Resume Next
    Const Updater_FN = "UpdChecker.exe"
    Dim sParams As String
    If hwnd = 0 Then hwnd = frmMain.hwnd
    If Not HideMode Then
        If Dir(App.path & "\" & Updater_FN) = "" Then
            MsgBox "Файл с именем '" & App.path & "\" & Updater_FN & "' не найден. Повторная установка приложения может исправить ситуацию.", vbExclamation
            Exit Sub
        End If
        frmMain.MousePointer = vbHourglass
    End If
    sParams = "/au /pn " & App.ProductName & " /bn " & App.EXEName & " /hw " & hwnd
    If HideMode Then sParams = sParams & " /hm"
    Shell App.path & "\" & Updater_FN & " " & sParams, vbNormalFocus
    If Not HideMode Then Wait 3000: frmMain.MousePointer = vbNormal
End Sub

Sub ShowHelp(ByVal sSearchText As String, Optional ByRef Form As Form)
    If Dir$(App.path & "\Help.exe") = "" Then ShowNotFoundProgramm (App.path & "\Help.exe"): Exit Sub
    If Form Is Nothing Then Set Form = frmMain
    Form.MousePointer = vbHourglass
    Shell App.path & "\Help.exe " & sSearchText, vbNormalFocus
    Form.MousePointer = vbNormal
End Sub

Function StartProgramm(ByVal sProgrammName As String, Optional ByVal sParameters As String)
    If InStr(1, sProgrammName, ":\") = 0 Then sProgrammName = App.path & "\" & sProgrammName
    If IsFileExists(sProgrammName) Then
        Shell sProgrammName & " " & sParameters, vbNormalFocus
    Else
        ShowNotFoundProgramm sProgrammName
    End If
End Function

Sub ShowNotFoundProgramm(ByVal sFullName As String)
    MsgBox "Файл с именем '" & sFullName & "' не найден. Повторная установка приложения может решить данную проблему.", vbExclamation
End Sub

Function SendInfoToServer(Optional ByVal sInfo As String, Optional ByVal sInfoType As String, Optional ByVal hwnd As Long)
    Const Updater_FN = "UpdChecker.exe"
    Dim sParams As String
    If Dir(App.path & "\" & Updater_FN) = "" Then: ShowNotFoundProgramm App.path & "\" & Updater_FN: Exit Function
    If sInfo <> "" Then WriteIni "Send_Info", Replace(sInfo, vbCrLf, "[crlf]"), "Main", "Update.ini"
    If sInfoType <> "" Then WriteIni "Send_InfoType", sInfoType, "Main", "\Update.ini"
    sParams = "/si /hw " & hwnd & "/pn " & App.ProductName
    Shell App.path & "\" & Updater_FN & " " & sParams, vbNormalFocus
End Function
