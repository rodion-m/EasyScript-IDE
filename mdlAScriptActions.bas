Attribute VB_Name = "mdlFunctions"
'преффикс sf - Script Function
Option Explicit
Option Compare Text
Dim c As Long, MyLang As Long

Enum MouseEvent
    pDown = 1
    pUp = 2
    pClick = 3
    pDblClick = 4
End Enum

Enum MouseKey
    pLeft = 1
    pRight = 2
    pMiddle = 4
End Enum

Enum MouseModes
    MM_mouse_event
    MM_keybd_event
    MM_SendMessage
    MM_PostMessage
    MM_DirectInput
End Enum

Enum YesOrNo
    pYes = 1
    pNo = 0
End Enum

Function sfWait(ByVal Milliseconds As Long, ByVal RndRange As Long) As Boolean
    sfWait = Wait(Milliseconds + GenerateRnd(0, RndRange))
End Function

Function sfClick(ByVal X As Long, ByVal Y As Long, ByVal MouseEvent As MouseEvent, ByVal MouseKey As MouseKey, _
                        ByVal EventCount As Long, ByVal ReturnCursor As YesOrNo, _
                        ByVal WaitBefore As Long, ByVal WaitAfter As Long, ByVal xMaxPos As Long, ByVal yMaxPos As Long) As Boolean
    Dim PrevX As Long, PrevY As Long
    If Resolution.X > 0 And (Screen_Width <> Resolution.X Or Screen_Height <> Resolution.Y) Then
        ConvertResolution X, Y, Screen_Width, Screen_Height, Resolution.X, Resolution.Y
    End If
    If X = 0 And Y = 0 Then GetCursorPos X, Y
    X = X + SelectedWindow.Position.Left
    Y = Y + SelectedWindow.Position.Top
    If xMaxPos > 0 Then If X > xMaxPos Then sfClick = True: Exit Function
    If yMaxPos > 0 Then If Y > yMaxPos Then sfClick = True: Exit Function
    For c = 1 To EventCount
        Wait WaitBefore 'задержка перед действием
        If ReturnCursor = pYes Then GetCursorPos PrevX, PrevY
        ' специализированные настройки
        If Settings.bSpecial1 Then 'сдвинуть курсор перед кликом
            SetCursorPos X + GenerateRnd(5, 20), Y + GenerateRnd(5, 20)
            Wait 100
        End If
        ' ****************************
        If Settings.lngCoordsLimit > 0 Then
            X = RandomizeNumber(X, Settings.lngCoordsLimit)
            Y = RandomizeNumber(Y, Settings.lngCoordsLimit)
        End If
        If X > 0 Or Y > 0 Then SetCursorPos X, Y
        MouseAction MouseKey, MouseEvent, Settings.MouseMode, X, Y
        If ReturnCursor = pYes Then SetCursorPos PrevX, PrevY
        Wait WaitAfter ' задержка после
        DoEvents
        If Not blnExecuting Then sfClick = True: Exit Function
    Next c
    sfClick = True
End Function

Function sfPress(ByVal Chars As String, ByVal Count As Long, ByVal WaitBefore As Long, ByVal WaitAfter As Long) As Boolean
    'имеется солиднейший вариант написать парсер для апипресса
    Dim IsWinKey As Boolean
    If Left$(Chars, 1) = Chr$(34) And Right$(Chars, 1) = Chr$(34) And Len(Chars) >= 3 Then Chars = Mid$(Chars, 2, Len(Chars) - 2)
    ConvertSpecialChars Chars
    For c = 1 To Count
        Wait WaitBefore
        IsWinKey = InStr(1, Chars, "{Win}")
        If IsWinKey Then
            Chars = Replace(Chars, "{Win}+", "")
            Chars = Replace(Chars, "{Win}", "")
            keybd_event VK_STARTKEY, 0, KEYEVENTF_KEYDOWN, 0
            If IsVirtKey(Chars) Then
                keybd_event Str2VK(Chars), 0, KEYEVENTF_KEYDOWN, 0
                keybd_event Str2VK(Chars), 0, KEYEVENTF_KEYUP, 0
            Else
                Err.Raise ErrNum, , "Неверное использование клавиши Win"
            End If
            keybd_event VK_STARTKEY, 0, KEYEVENTF_KEYUP, 0
        Else
            CheckLang
            SendKeys Chars
            ActivateKeyboardLayout MyLang, 0
        End If
        Wait WaitAfter
        If Not blnExecuting Then sfPress = True: Exit Function
    Next c
    sfPress = True
End Function

Sub ConvertSpecialChars(Chars As String)
    Chars = Replace(Chars, "{Вниз}", "{Down}")
    Chars = Replace(Chars, "{Вверх}", "{Up}")
    Chars = Replace(Chars, "{Вправо}", "{Right}")
    Chars = Replace(Chars, "{Влево}", "{Left}")
    Chars = Replace(Chars, "{Ctrl}+", "^")
    Chars = Replace(Chars, "{Alt}+", "%")
    Chars = Replace(Chars, "{Shift}+", "+")
    Chars = Replace(Chars, "{}", "{{}{}}")
    If Len(Chars) = 2 Then Chars = Replace(Chars, "}{", "{}}{{}")
    If Len(Chars) = 1 Then
        For Each arr In SpecialCharsArray
            If Chars = arr Then Chars = "{" & Chars & "}": Exit Sub
        Next
    End If
    For Each arr In SpecialCharsArray
        If UBound(Split(Chars, "{")) = UBound(Split(Chars, "}")) Then GoTo lNext
        If Right(Chars, 1) = arr Then
            Chars = Left(Chars, Len(Chars) - 1)
            Chars = Chars & "{" & arr & "}"
        End If
lNext:
    Next
End Sub

Sub CheckLang()
    Dim NowLang As Long
    NowLang = GetLang
    MyLang = GetMyLang
    If NowLang <> GetMyLang Then
        If NowLang = Ru Then
            ActivateKeyboardLayout Ru, 0
        ElseIf NowLang = En Then
            ActivateKeyboardLayout En, 0
        End If
    End If
End Sub

Function sfSetCursorPos(ByVal X As Long, ByVal Y As Long, ByVal WaitBefore As Long, ByVal WaitAfter As Long) As Boolean
    If Resolution.X > 0 And (Screen_Width <> Resolution.X Or Screen_Height <> Resolution.Y) Then
        ConvertResolution X, Y, Resolution.X, Resolution.Y, Screen_Width, Screen_Height
    End If
    X = X + SelectedWindow.Position.Left
    Y = Y + SelectedWindow.Position.Top
    Wait WaitBefore
    SetCursorPos X, Y
    Wait WaitAfter
    sfSetCursorPos = True
End Function

Function sfLoop(ByVal LoopLimit As Long, ByVal JumpTo As Long, ByVal WaitBefore As Long, _
                        ByRef LineToExecute As Long, ByRef LoopRemain As Long) As Boolean
    If LoopLimit = pInfinitely Then
        LineToExecute = JumpTo - 1
        LoopRemain = pInfinitely
        sfLoop = True
        Exit Function
    End If
    LoopRemain = LoopLimit - LoopCounter ' для статуса
    Wait WaitBefore
    LoopCounter = LoopCounter + 1
    If LoopCounter <= LoopLimit Then LineToExecute = JumpTo - 1 Else LoopCounter = 0
    If LoopLimit = 0 Then LineToExecute = JumpTo - 1
    sfLoop = True
End Function

Function sfCreateFile(ByVal FileName As String, ByVal Content As String) As Boolean
    sfCreateFile = CreateFile(FileName, Content)
End Function

Function sfWriteFile(ByVal FileName As String, ByVal Content As String) As Boolean
    sfWriteFile = WriteFile(FileName, Content)
End Function

Function sfReadFile(ByVal FileName As String, ByRef GetContentTo As Variant, ByVal LineNumber As Long, ByVal LinesCount As Long) As Boolean
    Dim i As Long
    GetContentTo = ""
    If LineNumber = 0 And LinesCount = 0 Then
        GetContentTo = ReadFile(FileName)
    Else
        Dim Lines() As String
        ReadByLine FileName, Lines
        If LineNumber = 0 Then LineNumber = 1
        If LinesCount = 0 Then LinesCount = 1
        LineNumber = LineNumber - 1
        LinesCount = LinesCount - 1
        For i = LineNumber To LineNumber + LinesCount
            If i > UBound(Lines) Then Exit For
            If Len(GetContentTo) > 0 Then
                GetContentTo = GetContentTo & vbCrLf & Lines(i)
            Else
                GetContentTo = Lines(i)
            End If
        Next i
    End If
    sfReadFile = True
End Function

Function sfMsgBox(ByVal Text As String, ByVal Caption As String, ByVal MsgStyle As VbMsgBoxStyle, ByVal TOPMOST As YesOrNo) As Boolean
    If MsgStyle = vbInformation Or MsgStyle = vbCritical Or MsgStyle = vbExclamation Then
        If TOPMOST Then
            SetWindowPos msgform.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE
        Else
            SetWindowPos msgform.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE
        End If
        MessageBox msgform.hwnd, Text, Caption, MsgStyle
        sfMsgBox = True
    End If
End Function

Function sfInputBox(ByVal Prompt As String, ByVal Title As String, ByRef GetTo As Variant) As Boolean
    GetTo = InputBox(Prompt, Title)
    sfInputBox = True
End Function

Function sfShell(ByVal FileName As String, ByVal WaitForShow As YesOrNo, ByVal WaitWhileRunning As YesOrNo, ByVal WindowStyle As VbAppWinStyle) As Boolean
    Dim hwnd As Long, pID As Long
    pID = Shell(FileName, WindowStyle)
    If WaitForShow Then
        Wait 1000
'        Do While GetProcessHwnd(pID) = 0 ' дожидаемся пока окно появится в памяти
'            DoEvents
'            CheckHotKeys
'            If Not blnExecuting Then sfShell = True: Exit Function
'        Loop
'        hwnd = GetProcessHwnd(pID)
'        Do Until IsWindowVisible(hwnd) ' дожидаемся пока окно станет видимым
'            DoEvents
'            CheckHotKeys
'            If Not blnExecuting Then sfShell = True: Exit Function
'        Loop
    End If
    If WaitWhileRunning Then
        Do While IsProcessExists(pID) = True ' дожидаемся завершения процесса
            DoEvents
            CheckHotKeys
            If Not blnExecuting Then sfShell = True: Exit Function
        Loop
    End If
    sfShell = True
End Function

Function sfShowWindow(ByVal Caption As String, ByVal StopIfWindowNotFound As YesOrNo) As Boolean
    Dim hwnd As Long
    hwnd = FindWindowByCaption(Caption)
    ShowWindow hwnd
    If hwnd = 0 And StopIfWindowNotFound Then
        Err.Raise ErrNum, , "Окно не найдено"
    Else
        sfShowWindow = True
    End If
End Function

Function sfSetWindow(ByVal Caption As String, ByVal StopIfWindowNotFound As YesOrNo) As Boolean
    Dim hwnd As Long
    hwnd = FindWindowByCaption(Caption)
    If hwnd > 0 Then
        SelectedWindow.hwnd = hwnd
        GetWindowRect SelectedWindow.hwnd, SelectedWindow.Position
        sfSetWindow = True
    Else
        If StopIfWindowNotFound Then
            Err.Raise ErrNum, , "Окно не найдено"
        Else
            SelectedWindow.hwnd = WindowNotFound
            sfSetWindow = True
        End If
    End If
End Function

Function sfStartScript(ByVal ScriptName As String, ByVal Count As Long, ByVal WaitWhileExecuting As YesOrNo) As Boolean
    Dim pID As Long, MyFullName As String
    #If COMPILED Then
        MyFullName = App.Path & "\" & GetExeBaseName & ".exe"
    #Else
        MyFullName = App.Path & "\Editor.exe"
    #End If
    CheckScriptName ScriptName
    pID = Shell(MyFullName & " /hm dh 0" & " /co " & Count & " /fn " & ScriptName)
    If WaitWhileExecuting Then
        Do While IsProcessExists(pID)
            DoEvents
            CheckHotKeys
            If Not blnExecuting Then KillProcessByPID pID: sfStartScript = True: Exit Function
        Loop
    End If
    sfStartScript = True
End Function

Function sfChangeResolution(ByVal X As Long, ByVal Y As Long) As Boolean
    Resolution.X = JustPositive(X)
    Resolution.Y = JustPositive(Y)
    sfChangeResolution = True
End Function

Function sfDeleteFile(ByVal sFileName As String) As Boolean
    If IsDangerName(sFileName) Then Err.Raise 66, "Запрещенный файл", "Данная команда может нанести вред компьютеру": Exit Function
    If IsFileExists(sFileName) Then Kill sFileName
    sfDeleteFile = True
End Function

Function sfKillProcess(ByVal PocessName As String, ByVal bForcibly As Boolean) As Boolean
    Shell "cmd /c taskkill /im " & PocessName & IIf(bForcibly, " /f", "")
    sfKillProcess = True
End Function

Function sfDeleteDir(ByVal Path As String) As Boolean
    If IsDangerName(Path) Then Err.Raise 66, "Запрещенный путь", "Данная команда может нанести вред компьютеру": Exit Function
    Shell "cmd /c rmdir /q /s " & Path, vbHide
    sfDeleteDir = True
End Function

Function sfCloseWindow(ByVal Caption As String, ByVal bForcibly As Boolean, ByVal StopIfWindowNotFound As Boolean) As Boolean
    Dim hwnd As Long
    hwnd = FindWindowByCaption(Caption)
    If hwnd > 0 Then DestroyWindow hwnd, bForcibly
    If hwnd = 0 And StopIfWindowNotFound Then
        Err.Raise ErrNum, , "Окно не найдено"
    Else
        sfCloseWindow = True
    End If
End Function

Function sfDownloadFile(ByVal URL As String, ByVal FakeParam, ByVal Method As String, ByVal Headers As String, ByVal Content As String, ByVal Cookies As String, ByVal FakeParam2, ByVal EnabeRedirects As Long) As ResponseInfo
    If Headers <> "" Then Headers = Replace(Headers, "&&", vbCrLf)
    If Left$(URL, 7) <> "http://" Then URL = "http://" & URL
    bDownloaderOn = True
    sfDownloadFile = WinHTTP_EXEC(URL, Headers, Content, Cookies, Method, EnabeRedirects)
    bDownloaderOn = False
End Function

Function sfIf(ByVal Condition As String, ByVal TruePart As Variant, ByVal FalsePart As Variant) As Boolean
    Dim LeftPart As String, RightPart As String, ConditionWithoutQuotes As String, Operation As typeOperations
    Condition = ReplaceOutQuotesSTR(Condition, "==", "=")
    ConditionWithoutQuotes = CleanAllTextInQoutes(Condition)
    If InStr(1, ConditionWithoutQuotes, "!=") > 0 Then
        GetSides Condition, LeftPart, RightPart, "!="
        Operation = toNotEqv
    ElseIf InStr(1, ConditionWithoutQuotes, "<>") > 0 Then
        GetSides Condition, LeftPart, RightPart, "<>"
        Operation = toNotEqv
    ElseIf InStr(1, ConditionWithoutQuotes, "=>") > 0 Then
        GetSides Condition, LeftPart, RightPart, "=>"
        Operation = toMoreEqv
    ElseIf InStr(1, ConditionWithoutQuotes, "<=") > 0 Then
        GetSides Condition, LeftPart, RightPart, "<="
        Operation = toLessEqv
    ElseIf InStr(1, ConditionWithoutQuotes, "=") > 0 Then
        GetSides Condition, LeftPart, RightPart, "="
        Operation = toEqv
    ElseIf InStr(1, ConditionWithoutQuotes, ">") > 0 Then
        GetSides Condition, LeftPart, RightPart, ">"
        Operation = toMore
    ElseIf InStr(1, ConditionWithoutQuotes, "<") > 0 Then
        GetSides Condition, LeftPart, RightPart, "<"
        Operation = toLess
    End If
    If IsTextOnQuotes(LeftPart) Then LeftPart = CutQuotes(LeftPart) _
        Else If IsExpression(LeftPart) Then LeftPart = Calculate(LeftPart)
    If IsTextOnQuotes(RightPart) Then RightPart = CutQuotes(RightPart) _
        Else If IsExpression(RightPart) Then RightPart = Calculate(RightPart)
    Select Case Operation
        Case toEqv: If LeftPart = RightPart Then Condition = 1 Else Condition = 0
        Case toNotEqv: If LeftPart <> RightPart Then Condition = 1 Else Condition = 0
        Case toLess: If LeftPart < RightPart Then Condition = 1 Else Condition = 0
        Case toMore: If LeftPart > RightPart Then Condition = 1 Else Condition = 0
        Case toMoreEqv: If LeftPart >= RightPart Then Condition = 1 Else Condition = 0
        Case toLessEqv: If LeftPart <= RightPart Then Condition = 1 Else Condition = 0
    End Select
    If Condition = "0" Or Condition = "" Or Condition = "no" Or Condition = "false" And FalsePart <> "" Then 'falsepart
        If InStr(1, ConditionWithoutQuotes, "&&") > 0 Then FalsePart = Replace(FalsePart, "&&", vbCrLf)
        ConvertScript FalsePart
        ExecuteCleanScript FalsePart
    Else 'truepart
        If InStr(1, ConditionWithoutQuotes, "&&") > 0 Then TruePart = Replace(TruePart, "&&", vbCrLf)
        ConvertScript TruePart
        ExecuteCleanScript TruePart
    End If
    sfIf = True
End Function

Function sfVBScript(ByVal Code As String, ByRef GetResultTo As String, ByVal IsNewFunction As Boolean) As Boolean
    If IsNewFunction Then
        If InStr(1, Code, "&&") > 0 Then Code = Replace(Code, "&&", vbCrLf)
        VBScript.AddCode Code
    Else
        GetResultTo = VBScript.Eval(Code)
    End If
    sfVBScript = True
End Function

Function sfGoTo(ByVal LineNumber As Long, ByRef LineToExecute As Long) As Boolean
    'сделать проверку на валидность линии
    LineToExecute = LineNumber - 1
    sfGoTo = True
End Function

