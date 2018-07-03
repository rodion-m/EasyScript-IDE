Attribute VB_Name = "mdlAScriptIntepretator"
Option Explicit
Option Compare Text
Option Base 1

Public Const Delimiter = ","
Public Const ConstPrefix = "$"
Public Const WindowNotFound = 1&

Public LoopCounter As Long, sConsts As typeScriptConsts

Dim LinesCount As Long, c As Long, n As Long, LineToExecute As Long, _
NowLang As String, LoopRemain As Long       ' LineToExecut - общий счетчик линий

Dim bIsScript As Boolean ' переключается тогда, когда выполняется VBScript\JScript

Public Type typeScriptConsts
    Time As String
    Date As String
    Key As String
    Window As String
End Type

Public Enum typeOperations
    toEqv
    toNotEqv
    toLess ' меньше
    toMore 'больше
    toLessEqv
    toMoreEqv
End Enum
'--------------------------------------------------- Functions Consts
Public Enum Functions
    cClick = 1001
    cWait = 1002
    cPress = 1003
    cExit = 1004
    cStartScript = 1005
    cSetCursorPos = 1006
    cLoop = 1007
    cMsgBox = 1008
    cInputBox = 1009
    cPrint = 1010 'Debug Print
    cCreateFile = 1011
    cWriteFile = 1012
    cReadFile = 1013
    cSetWindow = 1014
    cShowWindow = 1015
    cShell = 1016
    cStop = 1017
    cSkipErrors = 1018
    cChangeResolution = 1019
    cDeleteFile = 1020
    cKillProcess = 1021
    cDeleteDir = 1022
    cCreateDir = 1023
    cCloseWindow = 1024
    cDownloadFile = 1025
    cIf = 1026
    cVBScript = 1027
    cGoTo = 1028
End Enum

'--------------------------------------------------- Parameters
Public Const pFullPress = 1
Public Const pInfinitely = -1
Public Const pSeconds = "sec", pMinutes = "min", pHours = "hor"

Function ExecuteScript(ByVal Script As Variant, Optional ScriptWindow As Object, Optional Arrow As Object, Optional Status As Object) As Boolean
    Dim RealLine() As String, RealLinesCount As Long, VisibleMode As Boolean, CleanScript As String, OrigStatusText As String
    If Script = "" Then Exit Function
    blnExecuting = True
    LogEvent "Запущено: " & LoadedFile.BaseTitle
    If Settings.bSkipErrors Then On Error GoTo lSkip Else On Error GoTo ErrHandler
    AddStage "Интеретация скрипта - главный модуль"
    Randomize
    CleanAll
    VisibleMode = (Not ScriptWindow Is Nothing) And (Not Arrow Is Nothing) And (Not Status Is Nothing)
    If VisibleMode Then
        frmMain.UpdateControlsState
        OrigStatusText = Status.Text
    End If
    GetLines Script, RealLine, RealLinesCount
    If Script = "" Then GoTo EndSub
    For LineToExecute = 1 To RealLinesCount
        AddStage "Интеретация скрипта - главный модуль"
        GlobalError.line = LineToExecute ' обновляю линию в ошибке
        CleanScript = RealLine(LineToExecute)
        ConvertScript CleanScript
        If Not IsValidScript(CleanScript) Then ' если команда не определена
            If CleanScript <> "" Then
                If Not Settings.bSkipErrors Then Script = CleanScript
                Err.Raise ErrNum, , "Команда не определена"
            End If
            GoTo lNext
        End If
        If VisibleMode Then
            Status.Text = "Выполняется строка №" & LineToExecute
            If LoopRemain <> 0 Then
                If LoopRemain = pInfinitely Then
                    Status.Text = Status.Text & vbCrLf & "Циклу осталось выполнятся бесконечно раз"
                Else
                    Status.Text = Status.Text & vbCrLf & "Циклу осталось выполнятся " & LoopRemain & " раз(а)"
                End If
            End If
            ScriptWindow.SelStart = GetSelStartFromLine(Script, LineToExecute)
            If LineToExecute = 1 Then
                Arrow.Top = 0
            ElseIf (LineToExecute - 1) < GetMaxLines Then
                Arrow.Top = GetTextHeight(ScriptWindow, String$(LineToExecute - 1, vbLf)) - 1.8
            Else
                Arrow.Top = GetTextHeight(ScriptWindow, String$(GetMaxLines - 1, vbLf))
            End If
            Arrow.Visible = True
        End If
        ExecuteCleanScript CleanScript
lNext:
        DoEvents
        If Not blnExecuting Then Exit For
    Next LineToExecute
EndSub:
    blnExecuting = False
    If ErrorTitle = "" Then
        LogEvent "Выполнено: " & LoadedFile.BaseTitle
        ExecuteScript = True
    Else
        LogEvent "Выполнено с ошибками: " & LoadedFile.BaseTitle & "; Ошибка: " & ErrorTitle
        ExecuteScript = False
        ErrorTitle = ""
    End If
    If VisibleMode Then
        Arrow.Visible = False
        frmMain.UpdateAll
        frmMain.SetNormalState
    End If
    LoopRemain = 0
    LoopCounter = 0
Exit Function
lSkip:
    ADL "неверная конструкция скрипта"
    Resume Next
Exit Function
ErrHandler:
    GlobalError.Description = "Ошибка на стадии парсинга скрипта"
    GlobalError.Expression = Script
    Call GlobalErrorHandler
End Function

Function IsValidScript(ByVal Script As Variant, Optional ByVal bConvert As Boolean) As Boolean
    If bConvert Then ConvertScript Script
    If GetActionType(Script, True) > 0 Then IsValidScript = True
End Function

Function GetActionType(ByVal line As String, Optional ByVal bBadScript As Boolean) As ActionType
    If bBadScript Then ConvertScript line
    If Len(line) < 1 Then Exit Function
    For Each arr In Funcs
        If InStr(1, line, arr) = 1 Then GetActionType = atFunc: Exit Function
    Next
    If Len(line) < 3 Then Exit Function
    For Each arr In Operations
        n = InStr(1, line, arr)
        If n > 1 Then
            If InStr(1, line, arr) > 0 Then GetActionType = atOperation: Exit Function
        End If
    Next
End Function

Sub ExecuteCleanScript(ByVal Script As Variant)
    Dim ScriptLine() As String, i As Long, ActionType As ActionType
    If Settings.bSkipErrors Then On Error GoTo lSkip Else On Error GoTo ErrHandler
    AddStage "Интеретация скрипта - исполнительный модуль"
    GetLines Script, ScriptLine(), LinesCount
    For i = 1 To LinesCount
        AddStage "Интеретация скрипта - исполнительный модуль"
        ActionType = GetActionType(ScriptLine(i))
        If ActionType = atFunc Then 'Function
            ExecuteFunc ScriptLine(i)
        ElseIf ActionType = atOperation Then 'Operation
            SetVarBySides ScriptLine(i)
        End If
        Pause
        If Not blnExecuting Then Exit For
    Next i
Exit Sub
lSkip:
    ADL "неверная конструкция скрипта"
    Resume Next
Exit Sub
ErrHandler:
    GlobalError.Description = "Ошибка при попытке произвести парсинг строки"
    GlobalError.Expression = Script
    Call GlobalErrorHandler
End Sub

'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

Sub ExecuteFunc(ByVal line As String)
    Dim Func As Functions, FuncLen As Long, ParamLine As String, Params() As Variant, FuncStatus As Boolean, buff As String, _
    VarIndex As Long 'номер переменное нужен, чтобы задавать им значения в ф-циях типа ReadFile
    If Settings.bSkipErrors Then On Error GoTo lSkip Else On Error GoTo ErrHandler
    AddStage "Интеретация команды"
    Func = GetFunc(line)
    If SelectedWindow.hwnd > 0 Then
        'провекра, в том ли окне происходят действия
        If IsWindowVisible(SelectedWindow.hwnd) = 0 Then
            If SelectedWindow.hwnd = WindowNotFound Then ' ЕСЛИ задана контстанта WNF, то ошибка пропускается и программа по сути дожидается окна
                If Func = cClick Or Func = cPress Then Exit Sub
            Else
                Err.Raise ErrNum, , "Невозможно найти заданное окно"
                Exit Sub
            End If
        End If
        If GetForegroundWindow <> SelectedWindow.hwnd Then
            If GetForegroundWindow <> frmMain.hwnd Then ShowWindow SelectedWindow.hwnd
        End If
    End If
    FuncLen = Len(Func)
    ParamLine = GetParamLine(line, FuncLen)
    ParseParameters ParamLine, Params()
    If Not blnExecuting Then Exit Sub
    Select Case Func
        Case cExit: End
        Case cClick
            ' click координата x=*, координата y=*, событие=клик, кнопка мыши=левая, сколько раз=1, вернуть курсор=да, задержка перед действием=0, задержка после действия=0
            ' click555,777,click,left,5,yes,100,100
            CheckParameters Params(), 0, 0, pClick, pLeft, 1, pYes, 0, 0, 0, 0
            AddStage "Выполнение команды Click"
            FuncStatus = sfClick(Params(1), Params(2), Params(3), Params(4), Params(5), Params(6), Params(7), Params(8), Params(9), Params(10))
        Case cWait
            ' wait время ожидания=1000, рандомизировать на=0
            ' wait100,50
            CheckParameters Params(), 1000, 0
            AddStage "Выполнение команды Wait"
            FuncStatus = sfWait(Params(1), Params(2))
        Case cPress
            ' press "клавиша", сколько раз=1, задержка перед действием=0, задержка после действия=0
            ' press"A",5,100,100
            CheckParameters Params(), "", 1, 0, 0
            AddStage "Выполнение команды Press"
            If Params(1) = "" Then Err.Raise ErrNum, , "Потерян обязательный параметр команды Press"
            FuncStatus = sfPress(Params(1), Params(2), Params(3), Params(4))
        Case cStartScript
            'StartScript имя скрипта, сколько раз=1, ждать завершения=да
            'StartScript"C:\Script.es",1,yes
            CheckParameters Params(), "", 1, pYes
            AddStage "Выполнение команды StartMacros"
            If Params(1) = "" Then Err.Raise ErrNum, , "Потерян обязательный параметр команды StartMacros"
            FuncStatus = sfStartScript(Params(1), Params(2), Params(3))
        Case cSetCursorPos
            ' SetCursoePos координата x=*, координата y=*, задержка до выполнения действия=0, задержка после выполнения действия=0
            ' setcursorpos 100,200,0,0
            AddStage "Выполнение команды SetCursorPos"
            CheckParameters Params(), 0, 0, 0, 0
            If Params(1) = 0 And Params(2) = 0 Then Err.Raise ErrNum, , "Потерян обязательный параметр команды SetCursorPos"
            FuncStatus = sfSetCursorPos(Params(1), Params(2), Params(3), Params(4))
        Case cLoop
            ' Loop сколько раз=бесконечно, начиная с какой линии=1, задержка перед прыжком=0
            ' loop5,1,0
            CheckParameters Params(), pInfinitely, 1, 0
            AddStage "Выполнение команды Loop"
            FuncStatus = sfLoop(Params(1), Params(2), Params(3), LineToExecute, LoopRemain)
        Case cMsgBox
            'MsgBox текст, заголовок=EasyScript, стиль сообщение=информация, поверх всех окон=да
            'MsgBox"сообщение", "заголовок", 48, да
            CheckParameters Params(), "", "EasyScript", vbInformation, pYes
            AddStage "Выполнение команды MsgBox"
            FuncStatus = sfMsgBox(Params(1), Params(2), Params(3), Params(4))
        Case cInputBox
            'InpuBox текст, заголовок=EasyScript, получить информацию в
            'InpuBox"сообщение", "заголовок", х
            CheckParameters Params(), "", "EasyScript", "#VAR#"
            AddStage "Выполнение команды InpuBox"
            FuncStatus = sfInputBox(Params(1), Params(2), buff)
            SetVarValue Params(3), buff
        Case cPrint
            AddStage "Выполнение команды Print"
            CheckParameters Params(), ""
            If Params(0) = "" Then Err.Raise ErrNum, , "Потерян обязательный параметр команды Print"
            ADL ParamLine, False
            If frmDebug.Visible = False Then frmDebug.Show , frmMain
            FuncStatus = True
        Case cCreateFile
            'CreateFile Имя файла, контент=""
            'CreateFile"C:\1.txt", "content"
            CheckParameters Params(), "", ""
            AddStage "Выполнение команды CreateFile"
            If Params(1) = "" Then Err.Raise ErrNum, , "Потерян обязательный параметр команды CreateFile"
            FuncStatus = sfCreateFile(Params(1), Params(2))
        Case cWriteFile
            'WriteFile Имя файла, контент=""
            'WriteFile"C:\1.txt", "content"
            CheckParameters Params(), "", ""
            AddStage "Выполнение команды WriteFile"
            If Params(1) = "" Then Err.Raise ErrNum, , "Потерян обязательный параметр команды WriteFile"
            FuncStatus = sfWriteFile(Params(1), Params(2))
        Case cReadFile
            'ReadFile имя файла, переменная, начиная с какой строки, сколько строк
            'ReadFile"C:\1.txt", x, 5, 5
            CheckParameters Params(), "", "#VAR#", 0, 0
            AddStage "Выполнение команды ReadFile"
            If Params(1) = "" Then Err.Raise ErrNum, , "Потерян обязательный параметр команды ReadFile"
            FuncStatus = sfReadFile(Params(1), buff, Params(3), Params(4))
            SetVarValue Params(2), buff
        Case cSetWindow
            ' SetWindow заголовок="", если окно не обнаружено показать ошибку=да
            ' SetWindow"any",yes
            CheckParameters Params(), "", pYes
            AddStage "Выполнение команды SetWindow"
            FuncStatus = sfSetWindow(Params(1), Params(2))
        Case cShowWindow
            ' SetWindow заголовок="", если окно не обнаружено показать ошибку=да
            ' SetWindow"any",yes
            CheckParameters Params(), "", pYes
            AddStage "Выполнение команды ShowWindow"
            FuncStatus = sfShowWindow(Params(1), Params(2))
        Case cShell
            ' Shell полное имя программы="", дожидаться появления окна=нет, дожидаться завершения процесса,стиль окна при запуске=нормально
            ' Shell"C:\1.exe",no,no,normal
            CheckParameters Params(), "", pNo, pNo, vbNormalFocus
            AddStage "Выполнение команды Shell"
            If Params(1) = "" Then Err.Raise ErrNum, , "Потерян обязательный параметр команды Shell"
            FuncStatus = sfShell(Params(1), Params(2), Params(3), Params(4))
        Case cStop
            ' Stop
            blnExecuting = False
            FuncStatus = True
        Case cSkipErrors
            ' SkipErrors yes
            CheckParameters Params(), pYes
            Settings.bSkipErrors = CBool(Params(1))
            FuncStatus = True
        Case cChangeResolution
            ' ChangeResolution ширина, высота
            ' ChangeResolution 800,600
            CheckParameters Params(), 0, 0
            AddStage "Выполнение команды ChangeResolution"
            If Params(1) = 0 Or Params(2) = 0 Then Err.Raise ErrNum, , "Потерян обязательный параметр команды ChangeResolution"
            FuncStatus = sfChangeResolution(Params(1), Params(2))
        Case cDeleteFile
            'DeleteFile Имя файла
            'DeleteFile"C:\1.txt"
            CheckParameters Params(), ""
            AddStage "Выполнение команды DeleteFile"
            If Params(1) = "" Then Err.Raise ErrNum, , "Потерян обязательный параметр команды DeleteFile"
            FuncStatus = sfDeleteFile(Params(1))
        Case cKillProcess
            'KillProcess Имя процесса
            'KillProcess"explorer.exe"
            CheckParameters Params(), "", pNo
            AddStage "Выполнение команды KillProcess"
            If Params(1) = "" Then Err.Raise ErrNum, , "Потерян обязательный параметр команды KillProcess"
            FuncStatus = sfKillProcess(Params(1), Params(2))
        Case cDeleteDir
            'DeleteDir Имя директории
            'DeleteFir"C:\dir\"
            CheckParameters Params(), ""
            AddStage "Выполнение команды DeleteDir"
            If Params(1) = "" Then Err.Raise ErrNum, , "Потерян обязательный параметр команды DeleteDir"
            FuncStatus = sfDeleteDir(Params(1))
        Case cCreateDir
            'CreateDir Имя папки
            'CreateDir"C:\dir\"
            CheckParameters Params(), ""
            AddStage "Выполнение команды CreateDir"
            If Params(1) = "" Then Err.Raise ErrNum, , "Потерян обязательный параметр команды CreateDir"
            MkDir (Params(1)): FuncStatus = True
        Case cCloseWindow
            'CloseWindow заголовок, усиленно
            'CloseWindow"блокнот",1
            CheckParameters Params(), "", pYes, pYes
            AddStage "Выполнение команды CloseWindow"
            If Params(1) = "" Then Err.Raise ErrNum, , "Потерян обязательный параметр команды CloseWindow"
            FuncStatus = sfCloseWindow(Params(1), Params(2), Params(3))
        Case cDownloadFile
            'DownloadFile адрес*, путь к файлу или имя переменной, метод , заголовки через |, POST-данные, Cookies, пееременная для получаения ответных Cookie
            'DownloadFile"asd.ru"
            CheckParameters Params(), "", "", "GET", "", "", "", "", pYes
            AddStage "Выполнение команды DownloadFile"
            If Params(1) = "" Then Err.Raise ErrNum, , "Потерян обязательный параметр команды DownloadFile"
            Downloader = sfDownloadFile(Params(1), Params(2), Params(3), Params(4), Params(5), Params(6), Params(7), Params(8))
            If Downloader.Headers <> "" Then
                If Params(7) <> "" Then ' ЕСЛИ переменная для получения куки введена
                    SetVarValue Params(7), Downloader.Cookie, True
                End If
                If InStr(1, Params(2), "\") > 0 Or InStr(1, Params(2), ".") > 0 Then
                    ' Введено имя файла
                    buff = Params(2)
                    CheckFileName buff
                    FuncStatus = CreateFile(buff, Downloader.Body)
                ElseIf Params(2) <> "" Then
                    ' Введена переменная
                    FuncStatus = SetVarValue(Params(2), Downloader.Body)
                Else
                    ' Ничего не введено, значит имя файла получаю из URL
                    buff = GetFileBaseName(Params(1), "/")
                    If buff = "" Then buff = "page.html"
                    CheckFileName buff
                    FuncStatus = CreateFile(buff, Downloader.Body)
                End If
            Else
                FuncStatus = False
            End If
        Case cIf ' Условие
            'If a=1, then Click, else End
            CheckParameters Params(), "#COND#", "", ""
            AddStage "Выполнение конструкции If"
            If Params(1) = "" Or Params(1) = "#COND#" Or Params(2) = "" Then Err.Raise ErrNum, , "Потерян обязательный параметр конструкции If"
            FuncStatus = sfIf(Params(1), Params(2), Params(3))
        Case cVBScript
            bIsScript = True
            CheckParameters Params(), "", "#VAR#", pNo
            AddStage "Выполнение кода VBScript"
            If Params(1) = "" Then Err.Raise ErrNum, , "Потерян обязательный параметр команды VBScript"
            FuncStatus = sfVBScript(Params(1), buff, Params(3))
            SetVarValue Params(2), buff
        Case cGoTo
            ' GoTo номер строки
            ' goto5
            CheckParameters Params(), 0
            AddStage "Выполнение команды GoTo"
            If Params(1) = 0 Then Err.Raise ErrNum, , "Потерян обязательный параметр команды GoTo"
            FuncStatus = sfGoTo(Params(1), LineToExecute)
    End Select
    If Err.Number <> 0 Or FuncStatus = False Then
        If Settings.bSkipErrors Then GoTo lSkip Else GoTo ErrHandler
    End If
Exit Sub
lSkip:
    ADL "команда не может быть выполнена"
Exit Sub
ErrHandler:
    GlobalError.Description = "Ошибка во время выполнения команды"
    GlobalError.Expression = line
    Call GlobalErrorHandler
End Sub

'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

Function SetVarValue(ByVal sVar As String, ByVal sValue As String, Optional bIsText As Boolean) As Long   'Возвращает индекс назначенной переменной
    Dim VarIndex As Long
   If sVar = "" Then Exit Function
    Select Case Right$(sValue, 1)
        Case "+": sValue = sVar & "+1"
        Case "-": sValue = sVar & "-1"
    End Select
    ReplaceVars sValue ' переменные заменяются в скрипте, и готовый вариант посылается в интерпретатор выражений
    VarIndex = GetVarIndex(sVar)
    If VarIndex = 0 Then
        VarsCount = VarsCount + 1
        VarIndex = VarsCount
        VarName(VarIndex) = sVar
    End If
    VarVal(VarIndex) = sValue
    If (Not bIsText) And (Not IsTextOnQuotes(sValue)) And IsExpression(VarVal(VarIndex)) Then
        VarVal(VarIndex) = Calculate(sValue)
    ElseIf Not IsTextOnQuotes(sValue) Then 'ЕСЛИ это текст, но он не в кавычках, убираю его в кавычки
        sValue = Chr$(34) & sValue & Chr$(34)
    End If
    SetVarValue = VarIndex
End Function

Function GetVarIndex(ByVal Name As String) As Long
    If VarsCount > 0 Then
        For i = 1 To VarsCount
            If VarName(i) = "" Then Exit Function
            If VarName(i) = Name Then GetVarIndex = i: Exit Function
        Next i
    End If
End Function

Function SetVarBySides(ByVal line As String) As Long  'Возвращает индекс назначенной переменной
    Dim LeftSide As String, RightSide As String
    If Settings.bSkipErrors Then On Error GoTo lSkip Else On Error GoTo ErrHandler
    AddStage "Назначение переменной"
    If InStr(1, line, Chr$(34)) = 0 Then line = LCase(line)
    If InStr(1, line, "=") Then
        GetSides line, LeftSide, RightSide, "="
        LeftSide = LCase(LeftSide)
    Else
        LeftSide = LCase(GetVar(line))
        RightSide = line
    End If
    If LeftSide = "" Then Err.Raise ErrNum, , "Переменная не введена"
    SetVarBySides = SetVarValue(LeftSide, RightSide)
Exit Function
lSkip:
    ADL "не получается задать переменную"
    'Resume Next
Exit Function
ErrHandler:
    GlobalError.Description = "Ошибка при попытке задать значение переменной"
    GlobalError.Expression = line
    Call GlobalErrorHandler
End Function

Function GetVar(line As String) As String
    Dim NextSymbol As String
    If Len(line) <= 3 Then GetVar = Left$(line, 1): Exit Function
    For c = 1 To Len(line)
        c = c + 1
        GetVar = Left$(line, c)
        NextSymbol = Mid$(line, c + 1, 1)
        If NextSymbol = "" Then Exit For
        If Asc(NextSymbol) < 65 Or _
            (Asc(NextSymbol) > 90 And Asc(NextSymbol) < 97) Or _
            (Asc(NextSymbol) > 122 And Asc(NextSymbol) < 127) Then Exit For
    Next c
End Function

Sub ReplaceVars(ByRef line As Variant, Optional ByVal bLeaveQuotes As Boolean)
    For c = 1 To VarsCount
        'If VarVal(c) = "" Then VarVal(c) = "0"
        'ReplaceOutQuotes line, VarName(c), VarVal(c), rtVar
        RealReplaceVar line, VarName(c), VarVal(c), bLeaveQuotes
    Next c
End Sub

Sub Pause(Optional Ms As Long = 10)
    If Settings.bFakeWait Then Wait Ms
    CheckHotKeys hkExecuteScript
End Sub
 
Function GetParamLine(line As String, FuncLen As Long) As String
    If Len(line) = FuncLen Then Exit Function
    GetParamLine = Mid$(line, FuncLen + 1, Len(line) - FuncLen)
    If Left$(GetParamLine, 1) = "(" And Right$(GetParamLine, 1) = ")" Then GetParamLine = Mid$(GetParamLine, 2, Len(GetParamLine) - 2)
End Function

Function RandomSign() As Long
    If Rnd > 0.5 Then RandomSign = 1& Else RandomSign = -1&
End Function

Sub ConvertNumber(str As Variant, Optional Default As Variant, Optional JustPositive As Boolean)
    Dim Unit As String, num As Long, cuted As String
    If Settings.bSkipErrors Then On Error GoTo lSkip Else On Error GoTo ErrHandler
    AddStage "Конвертация числа"
    If IsNumeric(str) Then Exit Sub
    'сек, мин, час
    If InStr(1, str, pHours) = 0 And InStr(1, str, pMinutes) = 0 And InStr(1, str, pSeconds) = 0 Then Exit Sub
    str = Replace(str, pHours, "*3600000+")
    str = Replace(str, pMinutes, "*60000+")
    str = Replace(str, pSeconds, "*1000+")
    If Right$(str, 1) = "+" Then str = Left(str, Len(str) - 1)
'    Unit = Right(str, 3)
'    If Unit = pSeconds Or Unit = pMinutes Or Unit = pHours Then
'        cuted = Left(str, Len(str) - 3)
'        If cuted = "" Then cuted = "1"
'    Else
'        cuted = str
'    End If
'    cuted = Calculate(cuted)
'    If (Not IsNumeric(cuted) And Len(str) < 3) Then
'        GoTo SetDefault
'    ElseIf IsNumeric(cuted) Then
'        If JustPositive And CLng(cuted) < 0 Then GoTo SetDefault
'    End If
'    If IsNumeric(cuted) Then num = CLng(cuted) Else num = 1
'    Select Case Unit
'        Case pSeconds
'            num = num * 1000
'        Case pMinutes
'            num = num * 1000 * 60
'        Case pHours
'            num = num * 1000 * 60 * 60
'        Case Else 'если непонятно что, то задается значение по умолчанию
'SetDefault:
'            Err.Raise ErrNum, , "неверное значение свойства"
'            num = CLng(Default)
'    End Select
'    str = CStr(num)
Exit Sub
lSkip:
    ADL "параметр не является числом"
    Resume Next
Exit Sub
ErrHandler:
    GlobalError.Description = "Параметр, который должен являтся числом, таковым не является"
    GlobalError.Expression = str
    Call GlobalErrorHandler
End Sub

Function DeleteLineComment(ByVal sInput As String, Optional ByVal sCommentSign = "'") As String
    Dim CommentStart As Long, CommentEnd As Long, QuotesStart As Long, QuotesEnd As Long, _
    LineStart As Long, LineEnd As Long, line As String, CommentPosInLine As Long, PrevCommentPosInLine As Long, _
    PrevLineStart As Long
    If InStr(1, sCommentSign, sCommentSign) = 0 Then DeleteLineComment = sInput: Exit Function
    '----------------- строковый коммент
    Do While InStr(CommentEnd + 1, sInput, sCommentSign)
        CommentStart = InStr(CommentEnd + 1, sInput, sCommentSign)
        LineStart = InStrRev(sInput, vbCrLf, CommentStart)
        If PrevLineStart <> LineStart Then PrevCommentPosInLine = 0
        PrevLineStart = LineStart
        CommentEnd = InStr(CommentStart, sInput, vbCrLf)
        If CommentEnd <= 0 Then CommentEnd = Len(sInput) + 1
        LineEnd = CommentEnd
        line = Mid$(sInput, LineStart + 1, LineEnd - LineStart)
        CommentPosInLine = InStr(PrevCommentPosInLine + 1, line, sCommentSign)
        PrevCommentPosInLine = CommentPosInLine
        QuotesStart = InStrRev(line, Chr$(34), CommentPosInLine)
        QuotesEnd = InStr(CommentPosInLine, line, Chr$(34))
        If QuotesStart > 0 And CommentPosInLine > QuotesStart And CommentPosInLine < QuotesEnd Then CommentEnd = CommentStart: GoTo lLoop
        sInput = Replace(sInput, Mid$(sInput, CommentStart, CommentEnd - CommentStart), "")
        CommentEnd = CommentStart
lLoop:
    Loop
    DeleteLineComment = sInput
End Function

Sub DeleteComments(Script As Variant)
    If Settings.bSkipErrors Then On Error GoTo lSkip Else On Error GoTo ErrHandler
    Dim CommentStart As Long, CommentEnd As Long, QuotesStart As Long, QuotesEnd As Long, _
    LineStart As Long, LineEnd As Long, line As String, CommentPosInLine As Long, PrevCommentPosInLine As Long, _
    PrevLineStart As Long
    '----------------- строковый коммент
    Script = DeleteLineComment(Script, "'")
    Script = DeleteLineComment(Script, "//")
    Script = DeleteLineComment(Script, ";")
    '----------------- потоковый коммент /* ... */
    If InStr(1, Script, "/*") = 0 Then Exit Sub
    Do While InStr(CommentEnd + 1, Script, "/*")
        CommentStart = InStr(CommentEnd + 1, Script, "/*")
        CommentEnd = InStr(CommentStart, Script, "*/") + 1
        If CommentEnd <= 2 Then CommentEnd = Len(Script)
        LineStart = InStrRev(Script, vbCrLf, CommentStart)
        If PrevLineStart <> LineStart Then PrevCommentPosInLine = 0
        PrevLineStart = LineStart
        LineEnd = InStr(CommentStart, Script, vbCrLf)
        If LineEnd <= 0 Then LineEnd = Len(Script) + 1
        line = Mid$(Script, LineStart + 1, LineEnd - LineStart)
        CommentPosInLine = InStr(PrevCommentPosInLine + 1, line, "/*")
        PrevCommentPosInLine = CommentPosInLine
        QuotesStart = InStrRev(line, Chr$(34), CommentPosInLine)
        QuotesEnd = InStr(CommentPosInLine, line, Chr$(34))
        If CommentPosInLine > QuotesStart And CommentPosInLine < QuotesEnd Then CommentEnd = CommentStart: GoTo lLoop2
        Script = Replace(Script, Mid$(Script, CommentStart, CommentEnd - CommentStart + 1), "")
        CommentEnd = CommentEnd - (CommentEnd - CommentStart + 1)
lLoop2:
    Loop
    '------------------------------------------
    Script = ReplaceCrLf(Script)
Exit Sub
lSkip:
    ADL "комментарий имеет неверную конструкцию"
    Resume Next
Exit Sub
ErrHandler:
    GlobalError.Description = "Ошибка на стадии удаления комментарий"
    GlobalError.Expression = Script
    Call GlobalErrorHandler
End Sub

Sub CheckParameters(Params() As Variant, ParamArray DefaultParams() As Variant)
    If Settings.bSkipErrors Then On Error GoTo lSkip Else On Error GoTo ErrHandler
    ' массив дефолтных параметров начинается с 0, а массив параметров с 1
    AddStage "Проверка параметров на валидность"
    ReDim Preserve Params(1 To UBound(DefaultParams) + 1)
    For n = 1 To UBound(DefaultParams) + 1
        If bIsScript And n = 1 Then
            ReplaceVars Params(n), True
            bIsScript = False
        Else
            If DefaultParams(n - 1) <> "#VAR#" Then ReplaceVars Params(n)
        End If
        If InStr(1, Params(n), ConstPrefix) > 0 Then ReplaceConsts Params(n)
        If Params(n) = "" Or Params(n) = Chr(34) & Chr(34) Then 'если пустые кавычки
            Params(n) = DefaultParams(n - 1)
        ElseIf InStr(1, Params(n), "~") > 0 And Left$(Params(n), 1) <> Chr$(34) And Len(Params(n)) < 30 Then ' ЕСЛИ используется рандомизатор ~
            Dim buff() As String
            buff = Split(Params(n), "~")
            If (Not IsNumeric(buff(0))) Or (Not IsNumeric(buff(1))) Then
                If (Not IsNumeric(buff(0))) And (Not IsNumeric(buff(1))) Then ' ЕСЛИ обе части не являются числом (1 мин 30 сек ~ 2 мин 50 сек)
                    ConvertNumber buff(0)
                    ConvertNumber buff(1)
                    buff(0) = Calculate(buff(0))
                    buff(1) = Calculate(buff(1))
                ElseIf IsNumeric(buff(0)) And (Not IsNumeric(buff(1))) Then   ' ЕСЛИ только правая часть не является числом (1~2 минуты)
                    buff(0) = buff(0) & Right$(buff(1), 3) ' Задаю левой части ту же систему измерения что и правой
                    ConvertNumber buff(0)
                    ConvertNumber buff(1)
                    buff(0) = Calculate(buff(0))
                    buff(1) = Calculate(buff(1))
                ElseIf (Not IsNumeric(buff(0))) And IsNumeric(buff(1)) Then   ' ЕСЛИ только левая часть не является числом (1 минута ~ 2)
                    buff(1) = buff(1) & Right$(buff(0), 3) ' Задаю правой части ту же систему измерения что и правой
                    ConvertNumber buff(0)
                    ConvertNumber buff(1)
                    buff(0) = Calculate(buff(0))
                    buff(1) = Calculate(buff(1))
                End If
            End If
            If (Not IsNumeric(buff(0))) Or (Not IsNumeric(buff(1))) Then Err.Raise ErrNum, "Неправильно использован рандомизатор ~"
            Params(n) = GenerateRnd(buff(0), buff(1), False)
        ElseIf IsNumeric(DefaultParams(n - 1)) And Not IsNumeric(Params(n)) Then
            'если по умолчанию должно быть числовое значение, а в параметре не числовое, то оно конвертируется
            ConvertNumber Params(n), DefaultParams(n - 1), True  ' На всякий случай попробовать конвертировать из систем измерения
            If IsExpression(Params(n)) Then ' ЕСЛИ выражение, то попробовать посчитать
                Params(n) = Calculate(Params(n))
            End If
        ElseIf IsExpression(Params(n)) And Not IsNumeric(DefaultParams(n - 1)) Then
            ' если написано выражение, то попробовать его высчитать, пропуская ошибки
            Params(n) = Calculate(Params(n), True)
        ElseIf InStr(1, Params(n), Chr$(34)) > 0 Then
            ' если есть кавычки, то убрать их
            If IsTextOnQuotes(Params(n)) And DefaultParams(n - 1) <> "#COND#" Then _
                Params(n) = Mid(Params(n), 2, Len(Params(n)) - 2)
        End If
        Params(n) = UnEscapeChars(Params(n)) 'Декодирую данные из экрана
    Next n
Exit Sub
lSkip:
    ADL "неверные параметры"
    Resume Next
Exit Sub
ErrHandler:
    GlobalError.Description = "Ошибка при попытке задать параметры по умолчанию"
    Call GlobalErrorHandler
End Sub

Sub ParseParameters(ByVal ParamLine As String, ByRef Params() As Variant)
    Dim RealCount As Long, NullParam() As String
    If Settings.bSkipErrors Then On Error GoTo lSkip Else On Error GoTo ErrHandler
    AddStage "Определение параметров"
    If ParamLine = "" Then Exit Sub
    ParamLine = ReplaceInQuotesSTR(ParamLine, Delimiter, vbLf & vbCr)
    NullParam() = Split(ParamLine, Delimiter)
    ReDim Params(1 To UBound(NullParam) + 1)
    For n = 1 To UBound(NullParam) + 1
        Params(n) = NullParam(n - 1)
        Params(n) = Replace(Params(n), vbLf & vbCr, Delimiter) ' реплейс ","
    Next n
Exit Sub
lSkip:
    ADL "неверные параметры"
    Resume Next
Exit Sub
ErrHandler:
    GlobalError.Description = "Ошибка при попытке получить параметры"
    GlobalError.Expression = ParamLine
    Call GlobalErrorHandler
End Sub

Function GetSelStartFromLine(ByVal str As String, LineNum As Long) As Long
    If LineNum = 1 Then GetSelStartFromLine = 0: Exit Function
    For n = 1 To LineNum - 1
        GetSelStartFromLine = InStr(GetSelStartFromLine + 1, str, vbCrLf)
    Next n
    GetSelStartFromLine = GetSelStartFromLine + 1
End Function

Function GetFunc(ByVal line As Variant, Optional SkipErrs As Boolean) As Functions
    ' поиск функции
    If SkipErrs Then
        On Error Resume Next
        ConvertScript line
    End If
    For c = 1 To Len(line)
        GetFunc = LCase(Mid$(line, 1, c))
        For Each arr In Funcs
            If GetFunc = arr Then Exit Function 'And (Mid$(line, Len(GetFunc) + 1, 1) = "(" Or Len(line) = Len(GetFunc)) Then Exit Function
        Next
    Next c
End Function

Function GetScript(ByVal line As Variant) As String
    ConvertScript line
    GetScript = CStr(line)
End Function

Function EscapeChars(ByVal sIn, Optional EscChar As String = "\") As Variant
    Dim c As Long, Chars(), EscedChar As String
    EscapeChars = sIn
    If InStr(1, EscapeChars, EscChar) = 0 Then Exit Function
    Chars = Array(Chr$(34), "{", "}", "'", "*", "+", "-", ":", "/", ",", "(", ")", "^", "~", "=", "$")
    For c = LBound(Chars) To UBound(Chars)
        If InStr(1, EscapeChars, EscChar & Chars(c)) > 0 Then
            EscedChar = "#" & GetEscedChar(c) & "#"
            EscapeChars = Replace(EscapeChars, EscChar & Chars(c), EscedChar)
        End If
    Next c
End Function

Function UnEscapeChars(ByVal sIn, Optional EscChar As String = "\") As Variant
    Dim c As Long, Chars(), EscedChar As String, EscSign As String
    UnEscapeChars = sIn
    If InStr(1, UnEscapeChars, "#") = 0 Then Exit Function
    Chars = Array(Chr$(34), "{", "}", "'", "*", "+", "-", ":", "/", ",", "(", ")", "^", "~", "=", "$")
    For c = LBound(Chars) To UBound(Chars)
        EscedChar = "#" & GetEscedChar(c) & "#"
        If InStr(1, UnEscapeChars, EscedChar) > 0 Then
            UnEscapeChars = Replace(UnEscapeChars, EscedChar, Chars(c))
        End If
    Next c
End Function

Function GetEscedChar(ByVal num As String) As String
    Dim sFirst As String, sSecond As String
    If Len(num) = 1 Then num = "0" & num
    sFirst = Left$(num, 1)
    sSecond = Right$(num, 1)
    GetEscedChar = Chr$(CLng(sFirst)) & Chr$(CLng(sSecond))
End Function

Sub ReplaceConsts(ByRef Data As Variant)
    With sConsts
        If ConstExists(Data, .Time) Then ReplaceConst Data, .Time, Format$(Time, "hh:mm:ss")
        If ConstExists(Data, .Date) Then ReplaceConst Data, .Date, Format$(Date, "yyyy.mm.dd")
        If ConstExists(Data, .Key) Then ReplaceConst Data, .Key, ScanKey
        If ConstExists(Data, .Window) Then ReplaceConst Data, .Window, GetForegroundWindowCaption
    End With
End Sub

Function ConstExists(ByRef Data As Variant, ByVal ConstName As String) As Boolean
    ConstExists = (InStr(1, Data, ConstPrefix & ConstName) > 0)
End Function

Sub ReplaceConst(ByRef Data As Variant, ByVal ConstName As String, ByVal ConstValue As String)
    If IsNumeric(ConstValue) Then
        RealReplaceVar Data, ConstPrefix & ConstName, ConstValue
    Else
        RealReplaceVar Data, ConstPrefix & ConstName, Chr$(34) & ConstValue & Chr$(34)
    End If
End Sub

Sub ConvertScript(ByRef Script As Variant)
    If Settings.bSkipErrors Then On Error GoTo lSkip Else On Error GoTo ErrHandler
    'AddStage "Конвертация скрипта"
    Script = EscapeChars(Script)
'    ' экранирую новые строки, если они в кавычках
'    ReplaceInQuotes Script, vbCrLf, ""
    DeleteComments Script
    If (Not Not ConstName) <> 0 Then 'если массив существует - реплейс констант
        For n = 1 To UBound(ConstName)
            ReplaceOutQuotes Script, ConstName(n), ConstVal(n)
        Next n
    End If
    If (Not Not FuncName) <> 0 Then 'если массив существует - реплейс функций
        For n = 1 To UBound(FuncName)
            ReplaceOutQuotes Script, FuncName(n), FuncVal(n)
        Next n
    End If
    ReplaceOutQuotes Script, ", тогда ", ","
    ReplaceOutQuotes Script, ", то ", ","
    ReplaceOutQuotes Script, ", иначе ", ","
    ReplaceOutQuotes Script, " тогда ", ","
    ReplaceOutQuotes Script, " то ", ","
    ReplaceOutQuotes Script, " иначе ", ","
    ReplaceOutQuotes Script, ", then ", ","
    ReplaceOutQuotes Script, ", else ", ","
    ReplaceOutQuotes Script, " then ", ","
    ReplaceOutQuotes Script, " else ", ","
    
    ReplaceOutQuotes Script, " ", ""
'-----------------------------------------------------------------------------------------------------------------
    ReplaceFunc Script, ClickV(), cClick
    ReplaceFunc Script, WaitV(), cWait
    ReplaceFunc Script, PressV(), cPress
    ReplaceFunc Script, ExitV(), cExit
    ReplaceFunc Script, StartScriptV(), cStartScript
    ReplaceFunc Script, SetCursorPosV(), cSetCursorPos
    ReplaceFunc Script, LoopV(), cLoop
    ReplaceFunc Script, MsgBoxV(), cMsgBox
    ReplaceFunc Script, InputBoxV(), cInputBox
    ReplaceFunc Script, PrintV(), cPrint
    ReplaceFunc Script, CreateFileV(), cCreateFile
    ReplaceFunc Script, WriteFileV(), cWriteFile
    ReplaceFunc Script, ReadFileV(), cReadFile
    ReplaceFunc Script, SetWindowV(), cSetWindow
    ReplaceFunc Script, ShowWindowV(), cShowWindow
    ReplaceFunc Script, ShellV(), cShell
    ReplaceFunc Script, StopV(), cStop
    ReplaceFunc Script, SkipErrorsV(), cSkipErrors
    ReplaceFunc Script, ChangeResolutionV(), cChangeResolution
    ReplaceFunc Script, DeleteFileV(), cDeleteFile
    ReplaceFunc Script, KillProcessV(), cKillProcess
    ReplaceFunc Script, DeleteDirV(), cDeleteDir
    ReplaceFunc Script, CreateDirV(), cCreateDir
    ReplaceFunc Script, CloseWindowV(), cCloseWindow
    ReplaceFunc Script, DownloadFileV(), cDownloadFile
    ReplaceFunc Script, IfV(), cIf
    ReplaceFunc Script, VBScriptV(), cVBScript
    ReplaceFunc Script, GoToV(), cGoTo
    ReplaceParam Script, "секунда секунды секунду секунд сек. сек", pSeconds, False, True
    ReplaceParam Script, "минута минуту минуты минут мин. мин", pMinutes, False, True
    ReplaceParam Script, "часов часа час. час", pHours, False, True
    ReplaceParam Script, "событие:клик клик", pClick
    ReplaceParam Script, "левая", pLeft
    ReplaceParam Script, "правая", pRight
    ReplaceParam Script, "средняя", pMiddle
    ReplaceParam Script, "вниз", pDown
    ReplaceParam Script, "вверх", pUp
    ReplaceParam Script, "да yes true", pYes
    ReplaceParam Script, "нет no false", pNo
    ReplaceParam Script, "раза раз мс", "", False
    ReplaceParam Script, "нормально", vbNormalFocus
    ReplaceParam Script, "свернуто", vbMinimizedFocus
    ReplaceParam Script, "развернуто", vbMaximizedFocus
    ReplaceParam Script, "свёрнуто", vbMinimizedFocus
    ReplaceParam Script, "развёрнуто", vbMaximizedFocus
    ReplaceParam Script, "скрыто", vbHide
    ReplaceParam Script, "информация", vbInformation
    ReplaceParam Script, "критическаяошибка", vbCritical
    ReplaceParam Script, "ошибка", vbExclamation
    ReplaceParam Script, "бесконечно", pInfinitely
    'Script = UnEscapeChars(Script) ' Данные вывожу из экрана только на момент парсинга параметров
Exit Sub
lSkip:
    ADL "скрипт сконвертирован не полностью"
    Resume Next
Exit Sub
ErrHandler:
    GlobalError.Description = "Ошибка при попытке привести скрипт в необходимый вид"
    GlobalError.Expression = Script
    Call GlobalErrorHandler
End Sub

Sub ReplaceFunc(Script As Variant, FuncVariantes() As Variant, nReplace As Functions)
    For Each arr In FuncVariantes
        ReplaceOutQuotes Script, Replace(arr, " ", ""), nReplace, rtFunc
    Next
End Sub

Sub ReplaceParam(Script As Variant, ParamVariantes As String, Replace As String, Optional ParamAlone As Boolean = True, Optional bReplaceAll As Boolean)
    For Each arr In Split(ParamVariantes, " ")
        If bReplaceAll Then
            ReplaceOutQuotes Script, arr, Replace
        Else
            ReplaceOutQuotes Script, arr, Replace, rtParam, ParamAlone
        End If
    Next
End Sub

Sub ReplaceOutQuotes(ByRef Script As Variant, ByVal Find As String, ByVal nReplace As String, Optional ByVal ReplaceType As ReplaceType, Optional ParamAlone As Boolean = True)
    Dim EmptyQuotesStringStart As Long, EmptyQuotesStringEnd As Long, EmptyQuotesString As String
    'EmptyQuotesString - значит пустая от "
    If Settings.bSkipErrors Then On Error GoTo lSkip Else On Error GoTo ErrHandler
    If InStr(1, Script, Find) = 0 Then Exit Sub
    If InStr(1, Script, Chr$(34)) = 0 Then
        'если кавычек нет, то реплейсить все
        Select Case ReplaceType
            Case rtFunc
                RealReplaceFunc Script, Find, nReplace
            Case rtParam
                RealReplaceParam Script, Find, nReplace, ParamAlone
            Case rtVar
                RealReplaceVar Script, Find, nReplace
            Case Else
                Script = Replace(Script, Find, nReplace)
        End Select
        Exit Sub
    End If
    Do While InStr(EmptyQuotesStringStart + 1, Script, Find) > 0
        EmptyQuotesStringEnd = InStr(EmptyQuotesStringStart + 1, Script, Chr$(34))
        If EmptyQuotesStringEnd = 0 Then EmptyQuotesStringEnd = Len(Script)
        EmptyQuotesString = Mid$(Script, EmptyQuotesStringStart + 1, EmptyQuotesStringEnd - EmptyQuotesStringStart)
        Select Case ReplaceType
            Case rtFunc
                RealReplaceFunc EmptyQuotesString, Find, nReplace
            Case rtParam
                RealReplaceParam EmptyQuotesString, Find, nReplace
            Case rtVar
                RealReplaceVar EmptyQuotesString, Find, nReplace
            Case Else
                EmptyQuotesString = Replace(EmptyQuotesString, Find, nReplace)
        End Select
        Script = AccurateReplace(Script, EmptyQuotesString, EmptyQuotesStringStart + 1, EmptyQuotesStringEnd)
        EmptyQuotesStringEnd = InStr(EmptyQuotesStringStart + 1, Script, Chr$(34)) ' Еще раз ищу кавычку, чтобы обновить ее позицию, ибо после замены она может измениться
        EmptyQuotesStringStart = InStr(EmptyQuotesStringEnd + 1, Script, Chr$(34))
        If EmptyQuotesStringEnd = 0 Or EmptyQuotesStringStart = 0 Then Exit Sub
    Loop
Exit Sub
lSkip:
    ADL "замена потерпела крах"
    Resume Next
Exit Sub
ErrHandler:
    GlobalError.Description = "Ошибка при попытке заменить данные в скрипте"
    GlobalError.Expression = Script
    Call GlobalErrorHandler
End Sub

'Sub ReplaceOutQuotes(ByRef Script As Variant, ByVal Find As String, ByVal nReplace As String, Optional ByVal ReplaceType As ReplaceType, Optional ParamAlone As Boolean = True)
'    Dim EmptyQuotesStringStart As Long, EmptyQuotesStringEnd As Long, EmptyQuotesString As String, SectionEnd As Long, TextBeforeLineEnd As String, LineStart As Long, LineEnd As Long
'    'EmptyQuotesString - значит пустая от "
'    'Line - значит фактическая линия от CrLf до CrLf
'    If Settings.bSkipErrors Then On Error GoTo lSkip Else On Error GoTo ErrHandler
'    If InStr(1, Script, Find) = 0 Then Exit Sub
'    If InStr(1, Script, Chr$(34)) = 0 Then
'        'если кавычек нет, то реплейсить все
'        Select Case ReplaceType
'            Case rtFunc
'                RealReplaceFunc Script, Find, nReplace
'            Case rtParam
'                RealReplaceParam Script, Find, nReplace, ParamAlone
'            Case rtVar
'                RealReplaceVar Script, Find, nReplace
'            Case Else
'                Script = Replace(Script, Find, nReplace)
'        End Select
'        Exit Sub
'    End If
'    Do While InStr(EmptyQuotesStringEnd + 1, Script, Find) > 0
'        EmptyQuotesStringEnd = InStr(EmptyQuotesStringStart + 1, Script, Chr$(34))
'        If EmptyQuotesStringEnd = 0 Then EmptyQuotesStringEnd = Len(Script)
'        LineStart = InStrRev(Script, vbCrLf, EmptyQuotesStringEnd) + 1
'        LineEnd = InStr(EmptyQuotesStringEnd, Script, vbCrLf)
'        If LineEnd = 0 Then LineEnd = Len(Script)
'        TextBeforeLineEnd = Mid$(Script, 1, LineEnd) 'текст до конца строки нужен, чтобы ограничить поиск конца секции, желательно потом написат спец функцию
'        SectionEnd = InStr(EmptyQuotesStringEnd, TextBeforeLineEnd, Delimiter)
'        If SectionEnd = 0 Then SectionEnd = Len(TextBeforeLineEnd)
'        EmptyQuotesString = Mid$(Script, EmptyQuotesStringStart + 1, EmptyQuotesStringEnd - EmptyQuotesStringStart)
'        Select Case ReplaceType
'            Case rtFunc
'                RealReplaceFunc EmptyQuotesString, Find, nReplace
'            Case rtParam
'                RealReplaceParam EmptyQuotesString, Find, nReplace
'            Case rtVar
'                RealReplaceVar EmptyQuotesString, Find, nReplace
'            Case Else
'                EmptyQuotesString = Replace(EmptyQuotesString, Find, nReplace)
'        End Select
'        Script = AccurateReplace(Script, EmptyQuotesString, EmptyQuotesStringStart + 1, EmptyQuotesStringEnd)
'        EmptyQuotesStringStart = InStrRev(Script, Chr$(34), SectionEnd)  'всегда увеличивается
'        'EmptyQuotesStringEnd = EmptyQuotesStringStart + Len(nReplace)
'        ' багнутая ф-ция
'    Loop
'Exit Sub
'lSkip:
'    ADL "замена потерпела крах"
'    Resume Next
'Exit Sub
'ErrHandler:
'    GlobalError.Description = "Ошибка при попытке заменить данные в скрипте"
'    GlobalError.Expression = Script
'    Call GlobalErrorHandler
'End Sub

Function ReplaceOutQuotesSTR(ByVal str As Variant, ByVal Find As String, ByVal nReplace As String, Optional ByVal ReplaceType As ReplaceType, Optional ParamAlone As Boolean = True) As String
    ReplaceOutQuotes str, Find, nReplace, ReplaceType, ParamAlone
    ReplaceOutQuotesSTR = CStr(str)
End Function

Function ReplaceInQuotesSTR(ByVal str As Variant, ByVal Find As String, ByVal nReplace As String, Optional ByVal ReplaceType As ReplaceType) As String
    ReplaceInQuotes str, Find, nReplace, ReplaceType
    ReplaceInQuotesSTR = CStr(str)
End Function

Sub ReplaceInQuotes(ByRef Script As Variant, ByVal Find As String, ByVal nReplace As String, Optional ByVal ReplaceType As ReplaceType)
    Dim QuotesStringStart As Long, QuotesStringEnd As Long, QuotesString As String, LineStart As Long, LineEnd As Long
    ' QuotesString - значит та строка, которая в кавычках от "
    ' Line - значит фактическая линия от CrLf до CrLf
    If Settings.bSkipErrors Then On Error GoTo lSkip Else On Error GoTo ErrHandler
    If InStr(1, Script, Find) = 0 Then Exit Sub
    If InStr(1, Script, Chr$(34)) = 0 Then Exit Sub
    Do While InStr(QuotesStringEnd + 1, Script, Find) > 0 And InStr(QuotesStringEnd + 1, Script, Chr$(34)) > 0
        QuotesStringStart = InStr(QuotesStringEnd + 1, Script, Chr$(34))
        If QuotesStringStart = 0 Then GoTo lLoop
        LineStart = InStrRev(Script, vbCrLf, QuotesStringStart) + 1
        LineEnd = InStr(QuotesStringStart, Script, vbCrLf)
        If LineEnd = 0 Then LineEnd = Len(Script)
        QuotesStringEnd = InStr(QuotesStringStart + 1, Script, Chr$(34))
        If QuotesStringEnd = 0 Then Exit Sub
        If QuotesStringEnd > LineEnd Then GoTo lLoop ' если закрывающаяся кавычка за пределами линии, то перейти к следующему реплейсу
        QuotesString = Mid$(Script, QuotesStringStart + 1, QuotesStringEnd - QuotesStringStart - 1)
        If InStr(1, QuotesString, Find) = 0 Then GoTo lLoop
        Select Case ReplaceType
            Case rtFunc
                RealReplaceFunc QuotesString, Find, nReplace
            Case rtParam
                RealReplaceParam QuotesString, Find, nReplace
            Case rtVar
                RealReplaceVar QuotesString, Find, nReplace
            Case Else
                QuotesString = Replace(QuotesString, Find, nReplace)
        End Select
        Script = AccurateReplace(Script, QuotesString, QuotesStringStart + 1, QuotesStringEnd - 1)
        QuotesStringEnd = InStr(QuotesStringStart + 1, Script, Chr$(34))
lLoop:
    Loop
Exit Sub
lSkip:
    ADL "замена потерпела крах"
    Resume Next
Exit Sub
ErrHandler:
    GlobalError.Description = "Ошибка при попытке заменить данные в скрипте"
    GlobalError.Expression = Script
    Call GlobalErrorHandler
End Sub

Sub RealReplaceVar(ByRef line As Variant, ByVal Find As String, ByVal Replace As String, Optional bLeaveQuotes As Boolean)
    Dim FindPos As Long, LeftSymbol As String, RightSymbol As String, LeftGood As Boolean, RightGood As Boolean, ExtentVal As Double
    Do While InStr(FindPos + 1, line, Find) > 0
        FindPos = InStr(FindPos + 1, line, Find)
        If FindPos = 1 And Len(line) = 1 Then GoTo lReplace
        If FindPos > 1 Then LeftSymbol = Mid$(line, FindPos - 1, 1)
        RightSymbol = Mid$(line, FindPos + Len(Find), 1)
        If LeftSymbol = "" Then
            LeftGood = True
        Else
            LeftGood = CBool(Asc(LeftSymbol) < 65 Or (Asc(LeftSymbol) > 90 And Asc(LeftSymbol) < 97) Or (Asc(LeftSymbol) > 122 And Asc(LeftSymbol) < 127))
            If LeftGood = False Then LeftGood = LeftGood Or FindCharsOnString(LeftSymbol, "<", ">", " ", "'", Chr$(34), "(")
        End If
        If RightSymbol = "" Then
            RightGood = True
        Else
            RightGood = CBool(Asc(RightSymbol) < 65 Or (Asc(RightSymbol) > 90 And Asc(RightSymbol) < 97) Or (Asc(RightSymbol) > 122 And Asc(RightSymbol) < 127))
            If RightGood = False Then RightGood = RightGood Or FindCharsOnString(RightSymbol, "<", ">", " ", "'", Chr$(34), ")")
        End If
        If LeftGood And RightGood Then
lReplace:
            If IsNumeric(Replace) Then
                'line = AccurateReplace(line, "(" & Replace & ")", FindPos, FindPos - 1 + Len(Find))
                line = AccurateReplace(line, Replace, FindPos, FindPos - 1 + Len(Find))
            Else
                If IsTextOnQuotes(Replace) And IsTextOnQuotes(line) And Not bLeaveQuotes Then ' Если bLeaveQuotes, то текстовые значения переменных будут заменяться в кавычках
                    ' Подрезаю кавычки в значении переменной, если текст, в котором она заменяется уже итак в кавычках
                    line = AccurateReplace(line, CutQuotes(Replace), FindPos, FindPos - 1 + Len(Find))
                Else
                    line = AccurateReplace(line, Replace, FindPos, FindPos - 1 + Len(Find))
                End If
            End If
            FindPos = FindPos + Len(Replace)
        End If
        DoEvents
    Loop
    'If IsNumeric(line) And Left$(line, 1) = "(" Then line = -line  ' ЕСЛИ с двух сторон скобки, то ВБ думает, что число отрицательное, хоть это и не так. В данной проверке меняю знак числа на положительный.
End Sub

Function CutQuotes(ByVal sInput As String) As String
    If IsTextOnQuotes(sInput) Then
        sInput = Mid$(sInput, 2, Len(sInput) - 2)
    End If
    CutQuotes = sInput
End Function

Function IsTextOnQuotes(ByVal sInput As String) As Boolean
    If (Left$(sInput, 1) = Chr$(34) And Right$(sInput, 1) = Chr$(34) And Len(sInput) >= 2) Then IsTextOnQuotes = True
End Function

Sub RealReplaceFunc(Script As Variant, Find As String, Replace As String)
    Dim LineStart As Long, LineEnd As Long, line As String, FindPos As Long, CrLfStart As Long
    If Left$(Find, 1) = Chr$(34) Or Right$(Find, 1) = Chr$(34) Then Exit Sub
    Do While InStr(FindPos + 1, Script, Find) > 0
        FindPos = InStr(FindPos + 1, Script, Find)
        CrLfStart = FindPos - 2
        If CrLfStart > 0 Then
            If Mid$(Script, CrLfStart, 2) = vbCrLf Then
                Script = AccurateReplace(Script, Replace, FindPos, FindPos - 1 + Len(Find))
            End If
        Else
            Script = AccurateReplace(Script, Replace, FindPos, FindPos - 1 + Len(Find))
        End If
    Loop
End Sub

Sub RealReplaceParam(Script As Variant, Find As String, Replace As String, Optional ParamAlone As Boolean = True)
    Dim LineStart As Long, LineEnd As Long, FindPos As Long, CrLfStart As Long, LeftSymbolValide As Boolean, RightSymbolValide As Boolean, Char As String
    Do While InStr(FindPos + 1, Script, Find) > 0
        FindPos = InStr(FindPos + 1, Script, Find)
        CrLfStart = FindPos - 2
        If FindPos > 1 Then
            Char = Mid$(Script, FindPos - 1, 1)
            LeftSymbolValide = CBool(Asc(Char) < 65 Or (Asc(Char) > 90 And Asc(Char) < 97) Or (Asc(Char) > 122 And Asc(Char) < 127)) Or FindCharsOnString(Char, "<", ">", " ", "'", Chr$(34), ",", ")", "(")
        Else
            LeftSymbolValide = False: RightSymbolValide = False
        End If
        If FindPos + Len(Find) < Len(Script) Then
            Char = Mid$(Script, FindPos + Len(Find), 1)
            RightSymbolValide = CBool(Asc(Char) < 65 Or (Asc(Char) > 90 And Asc(Char) < 97) Or (Asc(Char) > 122 And Asc(Char) < 127)) Or FindCharsOnString(Char, "<", ">", " ", "'", Chr$(34), ",", ")", "(")
        Else
            RightSymbolValide = True
        End If
        If FindPos = 1 Then LeftSymbolValide = True
        If LeftSymbolValide And RightSymbolValide Then
            If CrLfStart > 0 Then
                If Mid$(Script, CrLfStart, 2) <> vbCrLf Then
                    Script = AccurateReplace(Script, Replace, FindPos, FindPos - 1 + Len(Find))
                End If
            ElseIf (FindPos > 0 And CrLfStart < 1) Or Left(Script, 1) = "," Then
                Script = AccurateReplace(Script, Replace, FindPos, FindPos - 1 + Len(Find))
            End If
        End If
    Loop
End Sub

Function AccurateReplace(ByVal str As Variant, ByVal Replace As String, ByVal ReplaceFrom As Long, ByVal ReplaceTo As Long) As String
    AccurateReplace = Left$(str, ReplaceFrom - 1) & Replace & Right$(str, Len(str) - ReplaceTo)
End Function

Function IsExpression(ByVal strExpression As String) As Boolean
    If Len(strExpression) > 100 Then strExpression = False
    If InStr(1, strExpression, Chr$(0)) > 0 Then Exit Function
    If InStr(1, strExpression, Chr$(34)) > 0 Then Exit Function
    If InStr(1, strExpression, ":\") > 0 Then Exit Function
    If InStr(1, strExpression, "http://") > 0 Then Exit Function
    For Each arr In Operations
        If InStr(1, strExpression, arr) > 0 Then IsExpression = True: Exit Function
    Next
    If InStr(1, strExpression, "(") > 0 Then IsExpression = True
End Function

'Sub ParseLine(ByVal Script As Variant, ByRef ScriptLine() As String)
''    ScriptLine = Split(Script, vbCrLf)
''    Dim a As Long, b As Long, i As Long
''    Script = vbCrLf & Script & vbCrLf
''    Script = ReplaceCrLf(Script)
''    If Len(Script) < 5 Then Exit Sub
''    LinesCount = UBound(Split(Script, vbCrLf)) - 1
''    ReDim ScriptLine(LinesCount)
''    For i = 1 To LinesCount
''        a = InStr(a + 1, Script, vbCrLf)
''        b = InStr(a + 1, Script, vbCrLf)
''        If a < b Then ScriptLine(i) = Mid(Script, a + 2, (b - a) - 2)
''        If i Mod 10 = 0 Then DoEvents
''    Next i
'End Sub


'    If Len(Char) > 1 Then GoTo NotChar
'    Select Case Asc(Char)
'        ' ------------------ special
'
'        '---------------------------
'        Case 48 To 57
'            ScanCode = Asc(Char)
'        Case 65 To 90
'            UpCase = True
'            Lang = En
'            ScanCode = Asc(UCase(Char))
'        Case 97 To 122
'            UpCase = False
'            Lang = En
'            ScanCode = Asc(UCase(Char))
'        Case 192 To 223
'            UpCase = True
'            Lang = Ru
'            ScanCode = GetRuScanCode(Char)
'        Case 224 To 255
'            UpCase = False
'            Lang = Ru
'            ScanCode = GetRuScanCode(Char)
'    End Select
'    Select Case Char
'        Case "ё":  ScanCode = 192: UpCase = False: Lang = Ru
'        Case "Ё":  ScanCode = 192: UpCase = True: Lang = Ru
'        Case " ":  ScanCode = vbKeySpace
'        Case "!": ScanCode = vbKey1: UpCase = True
'        Case Chr(34):  ScanCode = 222: UpCase = True: Lang = En
'
'    End Select
'    ' если язык="" или регистр="", то оное не меняется
'NotChar:
