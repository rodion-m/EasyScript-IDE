Attribute VB_Name = "mdlMain"
Option Explicit
Option Compare Text

Public Const MARKER_SIZE = 24

Public Const FuncsCount = 28

' ******************** V - Variantes *******************************
Public ruFunctions(), enFunctions(), FuncsFirstTextParam(), _
ClickV(), WaitV(), PressV(), ExitV(), StartScriptV(), SetCursorPosV(), LoopV(), _
MsgBoxV(), InputBoxV(), PrintV(), CreateFileV(), WriteFileV(), ReadFileV(), _
SetWindowV(), ShowWindowV(), ShellV(), StopV(), SkipErrorsV(), ChangeResolutionV(), _
DeleteFileV(), KillProcessV(), DeleteDirV(), CreateDirV(), CloseWindowV(), DownloadFileV(), _
IfV(), VBScriptV(), GoToV()

Public blnExecuting As Boolean, NowHelp As String, blnMarkerShowed As Boolean, _
ConstName() As String, ConstVal() As String, FuncName(), FuncVal(), Funcs() As Variant, _
SpecialCharsArray() As String, arr As Variant, ClicksArray() As String, Operations() As String, _
VarName() As Variant, VarVal() As Variant, VarsCount As Long, tDebug As Object, blnCancel As Boolean, i&, n&, _
SelectedWindow As WindowMinAttribs, Window As RECT ' глобальная переменная Window, чтоб не объявлять по сто раз в модулях
Public MaxLines As Long, Settings As mySettings, Resolution As POINTAPI, History As typeHistory, LoadedFile As typeFileStruct, _
CommandLine As typeCommandLineSettings, ScriptsPath As String, ErrorTitle As String, Downloader As ResponseInfo

Public VBScript As New ScriptControl

Type typeCommandLineSettings
    HideMode As Boolean
    SkipErrors As Boolean
    Exit As Boolean
    DisableHKs As Boolean
    AutoStart As Boolean
    FileName As String
    Count As Long
End Type

Enum LoadSettingsMode
    lsmAll
    lsmPSettings
    lsmSSettings
End Enum

Type typeHistory
    Text() As String
    Step As Long
    SelStart() As Long
End Type

Type mySettings
    bRecCursor As Boolean
    bRecByWindow As Boolean
    bReturnCursor As Boolean
    lngUpdateMousePosInterval As Long
    bShowListTT As Boolean
    bShowTextTT As Boolean
    bShowDebug As Boolean
    bShowMarker As Boolean
    ' ---Script Settings ---
    sngSpeed As Single
    bSkipErrors As Boolean
    bFakeWait As Boolean
    lngIntervalLimit As Long
    lngCoordsLimit As Long
    lngDelayDownUp As Long
    MouseMode As MouseModes
    ' ******* Specials
    bSpecial1 As Boolean
End Type
 
 Enum ToolTipType
    tttFunc = 1
    tttParam = 2
 End Enum
 
  Enum ParamType
'    ptAny = 0
'    ptAnyNum = 1
    ptYesOrNo = 2
    ptMouseEvent = 3
    ptMouseKey = 4
    ptMsgBoxStyle = 5
    ptAppWinStyle = 6
    ptHTTPmethod = 7
    ptFunction = 10
 End Enum
 
 Enum ReplaceType
    rtFunc = 1
    rtParam = 2
    rtVar = 3
    rtExpression = 4 ' для матиматических выражений
 End Enum
 
 Enum ActionType
    atUnknown
    atFunc
    atOperation
 End Enum

' /////////////// Error Struct ////////////////////
Public GlobalError As MyError
Public Const ErrNum = 66
Type MyError
    Expression As Variant
    Stage() As String
    StageCount As Long
    Description As String
    line As Long 'LineNumber
End Type
' /////////////// Error Struct ////////////////////

Public Const DefaultColor = &HF1E4D9, RunColor = &H80000018, RecColor = &HC0&

Type POINTAPI
    x As Long
    Y As Long
End Type

Declare Function GetActiveWindow Lib "user32" () As Long
Declare Function GetForegroundWindow Lib "user32" () As Long

Declare Function GetTickCount Lib "kernel32.dll" () As Long

Declare Function GetUserDefaultLangID Lib "kernel32.dll" () As Integer '1049 - русский

Declare Function DrawText Lib "user32" Alias "DrawTextW" (ByVal hDC As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long

Public Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByRef lParam As Any) As Long
Public Declare Function PostMessage Lib "user32.dll" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Const WM_LBUTTONDBLCLK As Long = &H203
Public Const WM_LBUTTONDOWN As Long = &H201
Public Const WM_LBUTTONUP As Long = &H202
Public Const WM_RBUTTONDBLCLK As Long = &H206
Public Const WM_RBUTTONDOWN As Long = &H204
Public Const WM_RBUTTONUP As Long = &H205
Public Const WM_MBUTTONDBLCLK As Long = &H209
Public Const WM_MBUTTONDOWN As Long = &H207
Public Const WM_MBUTTONUP As Long = &H208
Public Const WM_SHOWWINDOW As Long = &H18

Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Public Const SM_CYVTHUMB = 9 'Height of scroll box on horizontal scroll bar
Public Const SM_CXHTHUMB = 10 ' Width of scroll box on horizontal scroll bar
Public Const SM_CYCAPTION = 4 'Height of windows caption

Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Declare Function GetCaretPos Lib "user32.dll" (ByRef lpPoint As POINTAPI) As Long

Declare Function HideCaret Lib "user32" (ByVal hwnd As Long) As Long
Declare Function ShowCaret Lib "user32" (ByVal hwnd As Long) As Long

'Public Declare Function IsUserAnAdmin Lib "shell32" () As Long
'
Public Declare Function MessageBox Lib "user32" Alias "MessageBoxA" (ByVal hwnd As Long, ByVal lpText As String, ByVal lpCaption As String, ByVal wType As Long) As Long
'MessageBox 0, "Текст сообщения", "Заголовок сообщения", 48 (type)

Public Declare Function WinExec Lib "kernel32" (ByVal lpCmdLine As String, ByVal nCmdShow As Long) As Long
''WinExec "cmd /c C:\Windows\system32\calc.exe", 0
'
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
''для runas = ShellExecute 0, "runas", "C:\Windows\system32\cmd.exe", "", "C:\Windows\system32", 1

Function StartProgramm(ByVal sProgrammName As String, Optional ByVal sParameters As String)
    If InStr(1, sProgrammName, ":\") = 0 Then sProgrammName = App.Path & "\" & sProgrammName
    If IsFileExists(sProgrammName) Then
        Shell sProgrammName & " " & sParameters, vbNormalFocus
    Else
        ShowNotFoundProgramm sProgrammName
    End If
End Function

Sub ShowNotFoundProgramm(ByVal sFullName As String)
    MsgBox "Файл с именем '" & sFullName & "' не найден. Повторная установка приложения может решить данную проблему.", vbExclamation, "Файл не найден"
End Sub

Function GetCount(ByVal str As String, ByVal Find As String) As Long
    If str = "" Then Exit Function
    GetCount = UBound(Split(str, Find))
End Function

Function GetLinesCount(ByVal str As String) As Long
    str = vbCrLf & str
    GetLinesCount = UBound(Split(str, vbCrLf))
End Function

Function GetLineNumberByCharNumber(ByVal str As String, ByVal CharNum As Long) As Long
    Dim PrevCrLf As Long
    If CharNum = 0 Then GetLineNumberByCharNumber = 1: Exit Function
    If InStrRev(str, vbCrLf, CharNum) = 0 Then GetLineNumberByCharNumber = 1: Exit Function
    If InStr(CharNum, str, vbCrLf) = 0 Then GetLineNumberByCharNumber = GetLinesCount(str): Exit Function
    For n = 1 To GetLinesCount(str)
        GetLineNumberByCharNumber = InStr(GetLineNumberByCharNumber + 1, str, vbCrLf)
        If CharNum > PrevCrLf And CharNum < GetLineNumberByCharNumber Then
            GetLineNumberByCharNumber = n
            Exit For
        End If
        PrevCrLf = GetLineNumberByCharNumber
    Next n
End Function

Function JustPositive(ByVal num As Long) As Long
    If num < 0 Then JustPositive = 0 Else JustPositive = num
End Function

Sub ConvertResolution(ByRef x As Long, ByRef Y As Long, ByVal RecordResX As Long, ByVal RecordResY As Long, ByVal CurrentResX As Long, ByVal CurrentResY As Long)
    x = x / (RecordResX / CurrentResX)
    Y = Y / (RecordResY / CurrentResY)
End Sub

Sub GetLines(ByVal str As String, ByRef GetTo() As String, Optional GetLinesCountTo As Long)
    str = vbCrLf & str 'чтобы начать с 1, а не с 0
    GetTo = Split(str, vbCrLf)
    GetLinesCountTo = UBound(GetTo)
End Sub

Function ReplaceCrLf(ByVal str As String)
    ReplaceCrLf = str
    While InStr(1, ReplaceCrLf, vbCrLf & vbCrLf)
        ReplaceCrLf = Replace(ReplaceCrLf, vbCrLf & vbCrLf, vbCrLf)
    Wend
    If Right(ReplaceCrLf, 2) = vbCrLf Then ReplaceCrLf = Left(ReplaceCrLf, Len(ReplaceCrLf) - 2)
    If Left(ReplaceCrLf, 2) = vbCrLf Then ReplaceCrLf = Right(ReplaceCrLf, Len(ReplaceCrLf) - 2)
End Function

Sub GetSides(ByVal str As String, Side1, Side2, ByVal sDelimiter As String)
    Dim DelimiterPos As Long
    DelimiterPos = InStr(1, str, sDelimiter)
    Side1 = Left(str, DelimiterPos - 1)
    Side2 = Right(str, Len(str) - DelimiterPos - (Len(sDelimiter) - 1))
End Sub

Sub SetCustomConsts(str As String)
    Dim line() As String, LinesCount As Long, n As Long
    str = ReplaceCrLf(str)
    GetLines str, line, LinesCount
    ReDim ConstName(1 To LinesCount)
    ReDim ConstVal(1 To LinesCount)
    For n = 1 To LinesCount
        GetSides line(n), ConstName(n), ConstVal(n), "="
    Next n
End Sub

Function GetLineByLineNumber(ByVal str As String, ByVal LineNumber As Long, Optional GetLineStartTo As Long, Optional GetLineEndTo As Long) As String
    Dim Lines() As String
    GetLines str, Lines()
    GetLineByLineNumber = Lines(LineNumber)
    For n = 1 To LineNumber - 1
        GetLineStartTo = InStr(GetLineStartTo + 1, str, vbCrLf)
    Next n
    If GetLineStartTo = 0 Then GetLineStartTo = 1
    If GetLineStartTo > 1 Then GetLineStartTo = GetLineStartTo + 2
    GetLineEndTo = InStr(GetLineStartTo + 1, str, vbCrLf)
    If GetLineEndTo = 0 Then GetLineEndTo = Len(str)
End Function

Function GetLineByCharNumber(ByVal str As String, ByVal CharNumber As Long, Optional GetLineStartTo As Long, Optional GetLineEndTo As Long) As String
    GetLineByCharNumber = GetLineByLineNumber(str, GetLineNumberByCharNumber(str, CharNumber), GetLineStartTo, GetLineEndTo)
End Function

Sub SetCustomFunctions(str As String)
    Dim n As Long, EquallyPos As Long, FuncStart As Long, FuncEnd As Long, BlockStart As Long, BlockEnd As Long, Block As String, BlocksCount As Long
    str = ReplaceCrLf(str)
    ReplaceOutQuotes str, "[", "{"
    ReplaceOutQuotes str, "]", "}"
    BlocksCount = UBound(Split(str, "="))
    ReDim FuncName(1 To BlocksCount)
    ReDim FuncVal(1 To BlocksCount)
    For n = 1 To BlocksCount
        EquallyPos = InStr(BlockEnd + 1, str, "=") ' EquallyPos - Позиция знака =
        BlockStart = InStrRev(str, vbCrLf, EquallyPos)
        BlockEnd = InStr(BlockStart + 1, str, "}")
        Block = Mid(str, BlockStart + 1, BlockEnd - BlockStart)
        GetSides Block, FuncName(n), FuncVal(n), "="
        ReplaceOutQuotes FuncVal(n), "}", ""
        ReplaceOutQuotes FuncVal(n), "{", ""
        FuncVal(n) = ReplaceCrLf(FuncVal(n))
    Next n
End Sub

Function IsFormLoaded(ByVal sName As String) As Boolean
    Dim arr
    For Each arr In Forms
        If arr.Name = sName Then IsFormLoaded = True: Exit Function
    Next
End Function

Sub ShowMarker(Optional ByVal FuncInLine As Long, Optional ByVal bCallByUser As Boolean)
    On Error Resume Next
    Const lngTimeOfAppearance As Long = 200, lngUpdateInterval As Long = 10 ' интервал обновления маркера - каждые 10 мс
    Dim Point As POINTAPI, Params(), Left As Single, Top As Single, FormSize As Single, _
    lngEndTime As Long, lngTransparency As Long, lngDifference As Long, lngPrevDifference As Long
    If FuncInLine = cClick Or FuncInLine = cSetCursorPos Then Point = GetPoint
    If (Point.x = 0 And Point.Y = 0) Then
        If bCallByUser Then GetCursorPosPT Point Else GoTo lHide
    End If
    Left = Point.x * Screen.TwipsPerPixelX
    Top = Point.Y * Screen.TwipsPerPixelY
    If Not IsFormLoaded("frmMarker") Then
        Load frmMarker
        frmMarker.Show , frmMain
        Wait 1, False
        frmMain.Show
    End If
    With frmMarker
        If CompareValues(Point.x, .GetMarkerX, 5) And CompareValues(Point.Y, .GetMarkerY, 5) And blnMarkerShowed Then Exit Sub
        blnMarkerShowed = True
        .Move Left, Top
        Alpha .hwnd, 5
        lngPrevDifference = -1
        lngEndTime = GetPerformanceTime + lngTimeOfAppearance
        Do While lngEndTime > GetPerformanceTime
            lngDifference = lngTimeOfAppearance - (lngEndTime - GetPerformanceTime)
            If CompareValues(lngDifference, lngPrevDifference, lngUpdateInterval) Then GoTo lContinue
            lngTransparency = GetSmartValue(lngDifference, lngTimeOfAppearance, 5, 220) ' от 5 до 220
            FormSize = GetSmartValue(lngDifference, lngTimeOfAppearance, 1, MARKER_SIZE)   ' размер формы в пикселях
            DrawForm .hwnd, CLng(FormSize)
            Alpha .hwnd, lngTransparency
            .Left = Left - ((FormSize * Screen.TwipsPerPixelX) / 2)
            .Top = Top - ((FormSize * Screen.TwipsPerPixelY) / 2)
            lngPrevDifference = lngDifference
lContinue:
            DoEvents
        Loop
    End With
Exit Sub
lHide:
    HideMarker
End Sub

Function CompareValues(ByVal Val1 As Double, ByVal Val2 As Double, ByVal Imprecision As Double) As Boolean
    If Abs(Val1 - Val2) <= Imprecision Then CompareValues = True
End Function

Function HideMarker()
    If blnMarkerShowed Then
        If frmMarker.mnuAlwaysOnTop.Checked = False Then
            blnMarkerShowed = False
            Alpha frmMarker.hwnd, 5
        End If
    End If
End Function

Function GetSmartValue(ByVal nVal As Double, ByVal nMaxVal As Double, ByVal Minimum As Double, ByVal Maximum As Double) As Double
    GetSmartValue = Maximum * (nVal / nMaxVal)
    If GetSmartValue < Minimum Then GetSmartValue = Minimum
End Function

Function GetPoint() As POINTAPI
    Dim Params()
    ParseParameters GetParamLine(GetScript(frmMain.GetScriptLine), 4), Params()
    If (Not Not Params) = 0 Then Exit Function
    If IsNumeric(Params(1)) Then GetPoint.x = Params(1)
    If UBound(Params) > 1 Then
        If IsNumeric(Params(2)) Then GetPoint.Y = Params(2)
    End If
    If GetPoint.x > Screen_Width Then GetPoint.x = Screen_Width
    If GetPoint.Y > Screen_Height Then GetPoint.Y = Screen_Height
End Function

Function Any2Str(num As Variant) As String
    Dim i&, d&, a&
    If Settings.bSkipErrors Then
        On Error GoTo lSkip
    Else
        On Error GoTo ErrHandler
    End If
    a = InStr(1, num, "+")
    If a Then
        Any2Str = Left(num, a - 2)
        d = Right(num, Len(CStr(num)) - a)
        For i = 1 To d
                Any2Str = IIf(Abs(Any2Str) < 100000000000000#, Any2Str * 10, Any2Str & 0)
        Next i
    Else
        a = InStr(2, num, "-")
        If a Then
            Any2Str = Left(num, a - 2)
            d = Right(num, Len(CStr(num)) - a)
            For i = 1 To d
                Any2Str = IIf(InStr(1, Any2Str, ",") = 0 Or Abs(Any2Str) > 1, Any2Str / 10, Replace(Any2Str, ",", ",0"))
            Next i
        End If
    End If
Ends:
    If Any2Str = "" Then Any2Str = num
Exit Function
lSkip:
    ADL "не получается преобразовать число в текстовый вид"
Exit Function
ErrHandler:
    GlobalError.Description = "Ошибка при попытке преобразовать число в текстовый вид"
    Call GlobalErrorHandler
End Function

Sub GlobalErrorHandler()
    ErrorTitle = Err.Description
    If Not blnExecuting Then Err.Clear: Exit Sub
    'ADL "Критическая ошибка", , True
    frmErrorBox.ShowError
    Err.Clear
    blnExecuting = False
    frmMain.SetNormalState
    frmMain.UpdateAll
End Sub

Sub LogEvent(ByVal sEvent As String)
    Const LogName = "Log.txt"
    If CommandLine.HideMode Then Exit Sub
    If frmPSettings.chkLog.Value = 1 Then
        WriteFileRev App.Path & "\" & LogName, EventTime & sEvent & vbCrLf
    End If
End Sub

Function fDate()
    fDate = Format(Date, "dd.mm.yy")
End Function

Function fTime()
    fTime = Format(Time, "hh:mm:ss")
End Function

Function EventTime() As String
    EventTime = "[" & fDate & "]" & "[" & fTime & "] "
End Function

Sub ADL(line As String, Optional IsError As Boolean = True, Optional SaveErr As Boolean)  ' Add Debug Line
    ErrorTitle = Err.Description
    If Not Settings.bShowDebug Then Exit Sub
    If CommandLine.HideMode Then Exit Sub
    If IsError Then
        line = "Ошибка на строке " & GlobalError.line & ": " & line
        If Not SaveErr Then Err.Clear
    End If
    If Right(tDebug, 2) = vbCrLf Then
        tDebug = tDebug & line
    ElseIf tDebug = "" Then
        tDebug = line
    Else
        tDebug = tDebug & vbCrLf & line
    End If
    If frmDebug.Left > (Screen.TwipsPerPixelX * (Screen_Width - 66)) Or frmDebug.Left < 0 Then frmDebug.Left = Screen.Width / 2 - frmDebug.Width / 2
    If frmDebug.Top > (Screen.TwipsPerPixelY * (Screen_Height - 66)) Or frmDebug.Top < 0 Then frmDebug.Top = Screen.Height / 2 - frmDebug.Height / 2
    If frmDebug.Visible = False Then frmDebug.Show , frmMain
    tDebug.SelStart = Len(tDebug)
End Sub

Sub CleanAll()
    Dim EmptyError As MyError, EmptyWindow As WindowMinAttribs
    Resolution.x = 0
    Resolution.Y = 0
    GlobalError = EmptyError
    SelectedWindow = EmptyWindow
    tDebug = ""
    ReDim VarName(1000)
    ReDim VarVal(1000)
    VarsCount = 0
End Sub

Sub AddStage(StageName As String)
    Dim ExistSubNum As Long
    GlobalError.StageCount = GlobalError.StageCount + 1
    ReDim Preserve GlobalError.Stage(1 To GlobalError.StageCount)
    For n = 1 To GlobalError.StageCount 'специальная проверка на то, существует ли функция уже в массиве, и если да, то стереть все последующие за ней функции
        If Len(GlobalError.Stage(n)) > 8 Then If Right$(GlobalError.Stage(n), Len(GlobalError.Stage(n)) - 8) = StageName Then ExistSubNum = n
        If ExistSubNum > 0 And n > ExistSubNum Then
            GlobalError.Stage(n) = ""
        End If
    Next n
    If ExistSubNum = 0 Then
        GlobalError.Stage(GlobalError.StageCount) = "Этап " & GlobalError.StageCount & ": " & StageName
    Else
        GlobalError.StageCount = ExistSubNum
    End If
    'ADL GlobalError.Stage(GlobalError.StageCount)
End Sub

Function GetLargestNum(Numbers() As Double, Optional ByRef GetPosTo As Long) As Variant  'пока работает только для чисел > 0
    For n = LBound(Numbers) To UBound(Numbers)
            If GetLargestNum < CDbl(Numbers(n)) Then GetLargestNum = Numbers(n): GetPosTo = n
    Next n
End Function

Function GetLargestLen(Strings() As String) As Variant  'пока работает только для чисел > 0
    For n = LBound(Strings) To UBound(Strings)
            If GetLargestLen < Len(Strings(n)) Then GetLargestLen = Len(Strings(n))
    Next n
End Function

Sub StopAll()
    blnRec = False
    blnExecuting = False
    DoEvents
End Sub

Sub SaveSettings(Optional blnSaveScriptSettings As Boolean)
    Dim strSection As String
    'If Dir(App.Path & "\" & SETTINGS_FILENAME) <> "" Then Kill App.Path & "\" & SETTINGS_FILENAME
    strSection = "Main"
    With frmMain
        WriteIni "WindowState", .WindowState, strSection
        WriteIni "Left", .Left, strSection
        WriteIni "Top", .Top, strSection
        WriteIni "Width", .Width, strSection
        WriteIni "Height", .Height, strSection
    End With
    strSection = "File"
    With LoadedFile
        WriteIni "FileName", .FileName, strSection
        WriteIni "Path", .Path, strSection
        WriteIni "BaseName", .BaseName, strSection
        WriteIni "BaseTitle", .BaseTitle, strSection
    End With
    strSection = "Debug"
    With frmDebug
        WriteIni "Left", .Left, strSection
        WriteIni "Top", .Top, strSection
        WriteIni "Width", .Width, strSection
        WriteIni "Height", .Height, strSection
    End With
    strSection = "PSettings"
    With frmPSettings
        WriteIni "RecCursor", .chkRecCursor, strSection
        WriteIni "UpdateMousePosInterval", .txtUpdateMousePosInterval, strSection
        WriteIni "ReturnCursor", .chkReturnCursor, strSection
        WriteIni "RecByWindow", .chkRecByWindow, strSection
        WriteIni "ShowDebug", .chkShowDebug, strSection
        WriteIni "ShowListTT", .chkShowListTT, strSection
        WriteIni "ShowTextTT", .chkShowTextTT, strSection
        WriteIni "ShowMarker", .chkShowMarker, strSection
        WriteIni "execKey1", .execKey(1), strSection
        WriteIni "execKey2", .execKey(2), strSection
        WriteIni "execKey3", .execKey(3), strSection
        WriteIni "recKey1", .recKey(1), strSection
        WriteIni "recKey2", .recKey(2), strSection
        WriteIni "recKey3", .recKey(3), strSection
        WriteIni "Minimize", .chkMinimize, strSection
        WriteIni "Maximize", .chkMaximize, strSection
        WriteIni "ShowButtons", .chkShowButtons, strSection
        WriteIni "MinimizeOnRecord", .chkMinimizeOnRecord, strSection
        WriteIni "Log", .chkLog, strSection
    End With
    CacheProgrammSettings
    If blnSaveScriptSettings Then
        strSection = "SSettings"
        With frmSSettings
            WriteIni "Speed", .txtSpeed, strSection
            WriteIni "SkipErrors", .chkSkipErrors, strSection
            WriteIni "FakeWait", .chkFakeWait, strSection
            WriteIni "RandomizeInterval", .txtIntervalLimit, strSection
            WriteIni "RandomizeCoords", .txtCoordsLimit, strSection
            WriteIni "IntervalLimit", .txtIntervalLimit, strSection
            WriteIni "CoordsLimit", .txtCoordsLimit, strSection
            WriteIni "DelayDownUp", .chkDelayDownUp, strSection
            WriteIni "MouseMode", .cmbMouseMode.ListIndex, strSection
            WriteIni "Special1", .chkSpecial1, strSection
        End With
        CacheScriptSettings
    End If
    UpdateHotKeys
End Sub

Sub LoadSettings(Optional Mode As LoadSettingsMode)
    Select Case Mode
        Case lsmAll
            If Dir(App.Path & "\Scripts", vbDirectory) <> "" Then
                ScriptsPath = App.Path & "\Scripts\"
            Else
                ScriptsPath = App.Path & "\"
            End If
            LoadFormSettings frmMain
            LoadFormSettings frmDebug
            LoadFormSettings frmPSettings
            LoadFileSettings
        Case lsmPSettings
            LoadFormSettings frmPSettings
        Case lsmSSettings
            LoadFormSettings frmSSettings
    End Select
    CacheProgrammSettings
End Sub

Sub LoadFileSettings()
    With LoadedFile
        .FileName = ReadIni("FileName", "", "File")
        If IsFileExists(.FileName) Then
            .Path = ReadIni("Path", "", "File")
            .BaseName = ReadIni("BaseName", "", "File")
            .BaseTitle = ReadIni("BaseTitle", "", "File")
            LoadScript .FileName
        Else
            ' ЕСЛИ файла нет, анулировать данные
            .FileName = ""
        End If
    End With
End Sub

Sub LoadFormSettings(Form As Object)
    Dim strSection As String, obj As Object, buff As Variant
    strSection = Form.LinkTopic
    With Form
    Select Case Form.LinkTopic
        Case frmMain.LinkTopic
            Dim wState As Single
            wState = CSng(ReadIni("WindowState", 0, strSection, dNumeric))
            If wState >= 0 And wState <= 2 Then .WindowState = wState
            If wState = 0 Then
                .Left = CSng(ReadIni("Left", (Screen.Width \ 2 - .Width \ 2), strSection, dNumeric))
                .Top = CSng(ReadIni("Top", (Screen.Height \ 2 - .Height \ 2), strSection, dNumeric))
                .Width = CSng(ReadIni("Width", .Width, strSection, dNumeric))
                .Height = CSng(ReadIni("Height", .Height, strSection, dNumeric))
            End If
        Case frmDebug.LinkTopic
            .Left = CSng(ReadIni("Left", .Left, strSection, dNumeric))
            .Top = CSng(ReadIni("Top", .Top, strSection, dNumeric))
            .Width = CSng(ReadIni("Width", .Width, strSection, dNumeric))
            .Height = CSng(ReadIni("Height", .Height, strSection, dNumeric))
        Case frmPSettings.LinkTopic
            .chkRecCursor = CLng(ReadIni("RecCursor", 1, strSection, dChk))
            .chkReturnCursor = CLng(ReadIni("ReturnCursor", 1, strSection, dChk))
            .txtUpdateMousePosInterval = CLng(ReadIni("UpdateMousePosInterval", 100, strSection, dNumeric))
            .chkRecByWindow = CLng(ReadIni("RecByWindow", 0, strSection, dChk))
            .chkShowDebug = CLng(ReadIni("ShowDebug", 1, strSection, dChk))
            .chkShowListTT = CLng(ReadIni("ShowListTT", 1, strSection, dChk))
            .chkShowTextTT = CLng(ReadIni("ShowTextTT", 1, strSection, dChk))
            .chkShowMarker = CLng(ReadIni("ShowMarker", 1, strSection, dChk))
            .execKey(1) = ReadIni("execKey1", "Ctrl", strSection, dHotKey)
            .execKey(2) = ReadIni("execKey2", "E", strSection, dHotKey)
            .execKey(3) = ReadIni("execKey3", "нет", strSection, dHotKey)
            .recKey(1) = ReadIni("recKey1", "Ctrl", strSection, dHotKey)
            .recKey(2) = ReadIni("recKey2", "R", strSection, dHotKey)
            .recKey(3) = ReadIni("recKey3", "нет", strSection, dHotKey)
            If .execKey(1) = "нет" Then .execKey(1) = "Ctrl"
            If .recKey(1) = "нет" Then .recKey(1) = "Ctrl"
            .chkMinimize = ReadIni("Minimize", 1, strSection, dChk)
            .chkMaximize = ReadIni("Maximize", 1, strSection, dChk)
            .chkShowButtons = ReadIni("ShowButtons", 1, strSection, dChk)
            .chkMinimizeOnRecord = ReadIni("MinimizeOnRecord", 1, strSection, dChk)
            .chkLog = ReadIni("Log", 0, strSection, dChk)
'        Case frmSSettings.LinkTopic
'            .SetSpeed ReadIni("Speed", "1,0", strSection)
'            .chkSkipErrors = ReadIni("SkipErrors", 0, strSection, dChk)
'            .chkFakeWait = ReadIni("FakeWait", 1, strSection, dChk)
'            .txtIntervalLimit = ReadIni("IntervalLimit", 100, strSection)
'            .txtCoordsLimit = ReadIni("CoordsLimit", 0, strSection)
'            .chkDelayDownUp = ReadIni("DelayDownUp", 0, strSection, dChk)
'            If ReadIni("MouseMode", 1, strSection, dNumeric) < 4 Then .cmbMouseMode.ListIndex = ReadIni("MouseMode", 0, strSection, dNumeric)
'            .chkSpecial1 = ReadIni("Special1", 0, strSection, dChk)
    End Select
    End With
End Sub

Sub LoadDefaultScriptSettings()
    Dim strSection As String
    strSection = "SSettings"
    With Settings
        .sngSpeed = ReadIni("Speed", "1,0", strSection)
        .bSkipErrors = ReadIni("SkipErrors", 0, strSection, dChk)
        .bFakeWait = ReadIni("FakeWait", 1, strSection, dChk)
        .lngIntervalLimit = ReadIni("IntervalLimit", 10, strSection)
        .lngCoordsLimit = ReadIni("CoordsLimit", 0, strSection)
        .lngDelayDownUp = ReadIni("DelayDownUp", 0, strSection, dChk)
        If ReadIni("MouseMode", 1, strSection, dNumeric) < 4 Then .MouseMode = ReadIni("MouseMode", 0, strSection, dNumeric)
        .bSpecial1 = ReadIni("Special1", 0, strSection, dChk)
    End With
End Sub

Function GetDataType(obj As Object) As TypeOfData
    If TypeOf obj Is CheckBox Then GetDataType = dChk
    If TypeOf obj Is TextBox Then If IsNumeric(obj.Text) Then GetDataType = dNumeric
    If obj.Name = "execKey" Or obj.Name = "recKey" Then GetDataType = dHotKey
End Function

Sub LoadScript(Optional ByVal FileName As String, Optional JustUpdateInfo As Boolean, Optional JustLoadSettings As Boolean, Optional ByVal Content As String) ' Content значит загрузить скрипт без файла
    Dim ScriptBody As String, ParamsBody As String, Params() As String
    If JustLoadSettings Then GoTo lLoadSettings
    If Not IsFileExists(FileName) And Content = "" Then Exit Sub
    If Not JustUpdateInfo And Not JustLoadSettings Then frmMain.MousePointer = vbHourglass
    DoEvents
    With LoadedFile
        .FileName = FileName
        .Path = Left$(.FileName, InStrRev(.FileName, "\"))
        .BaseName = Right$(.FileName, Len(.FileName) - InStrRev(.FileName, "\"))
        .BaseTitle = Left$(.BaseName, IIf(InStrRev(.BaseName, ".") > 0, InStrRev(.BaseName, ".") - 1, Len(.BaseName)))
        If Content <> "" Then .Content = Content Else .Content = ReadFile(FileName)
        n = InStrRev(.Content, vbCrLf & "Settings::")
        If n > 0 Then
            If n > 1 Then .ScriptBody = Left$(.Content, n - 1)
            .SettingsBody = Right$(.Content, Len(.Content) - n - 1)
        Else
            .ScriptBody = LoadedFile.Content
        End If
    End With
    If JustUpdateInfo Then Exit Sub
lLoadSettings:
    With Settings
         .sngSpeed = GetScriptSettings("Speed", "1,0", 100)
         .bSkipErrors = GetScriptSettings("SkipErrors", 0)
         .bFakeWait = GetScriptSettings("FakeWait", 1)
         .lngIntervalLimit = GetScriptSettings("IntervalLimit", 5)
         .lngCoordsLimit = GetScriptSettings("CoordsLimit", 0)
         If CChk(GetScriptSettings("DelayDownUp", 0)) Then
            .lngDelayDownUp = 50
        Else
            .lngDelayDownUp = 1
         End If
         .MouseMode = GetScriptSettings("MouseMode", 0, 3, , True)
         .bSpecial1 = GetScriptSettings("Special1", 0)
    End With
'    With frmSSettings
'         .SetSpeed GetScriptSettings("Speed", "1,0", 100)
'         .chkSkipErrors = GetScriptSettings("SkipErrors", 0)
'         .chkFakeWait = GetScriptSettings("FakeWait", 1)
'         .txtIntervalLimit = GetScriptSettings("IntervalLimit", 10)
'         .txtCoordsLimit = GetScriptSettings("CoordsLimit", 0)
'         .chkDelayDownUp = CChk(GetScriptSettings("DelayDownUp", 0))
'         .cmbMouseMode.ListIndex = GetScriptSettings("MouseMode", 0, 3)
'         .chkSpecial1 = GetScriptSettings("Special1", 0)
'    End With
    If JustLoadSettings Then Exit Sub
    frmMain.txtMain = LoadedFile.ScriptBody
    ReDim History.Text(0)
    ReDim History.SelStart(0)
    History.Step = 0
    History.Text(0) = frmMain.txtMain
    frmMain.MousePointer = vbNormal
End Sub

Function CChk(s) As Integer 'функция либо возвращает 0 либо 1
    If CBool(s) = True Then CChk = 1 Else CChk = 0
End Function


Function SaveScript(Optional ByVal FileName As String, Optional JustSetNewParams As Boolean) As Boolean
    Dim OldFile As typeFileStruct
     If Not JustSetNewParams Then frmMain.MousePointer = vbHourglass
    OldFile = LoadedFile
    LoadedFile.SettingsBody = "Settings::"
    With Settings
        AddScriptSettings "Speed", .sngSpeed
        AddScriptSettings "SkipErrors", .bSkipErrors
        AddScriptSettings "FakeWait", .bFakeWait
        AddScriptSettings "RandomizeInterval", .lngIntervalLimit
        AddScriptSettings "RandomizeCoords", .lngCoordsLimit
        AddScriptSettings "IntervalLimit", .lngIntervalLimit
        AddScriptSettings "CoordsLimit", .lngCoordsLimit
        AddScriptSettings "DelayDownUp", (.lngDelayDownUp > 1)
        AddScriptSettings "MouseMode", .MouseMode
        AddScriptSettings "Special1", .bSpecial1
    End With
    If JustSetNewParams Then Exit Function
    LoadedFile.Content = frmMain.txtMain & vbCrLf & LoadedFile.SettingsBody
    If FileName = "" Then FileName = LoadedFile.FileName
    If CreateFile(FileName, LoadedFile.Content) Then
        LoadScript FileName, True
        SaveScript = True
    ElseIf LoadedFile.FileName <> "" Then
        LoadedFile = OldFile
    End If
     frmMain.MousePointer = vbNormal
End Function

Function AddScriptSettings(ByVal strName As String, ByVal strValue As String)
    If strValue = "True" Then strValue = "1"
    If strValue = "False" Then strValue = "0"
    
    LoadedFile.SettingsBody = LoadedFile.SettingsBody & strName & "=" & strValue & ";"
End Function

Function GetScriptSettings(ByVal strValue As String, ByVal strDefault, Optional ByVal Limit As Double = -100000, Optional IsCheckBox As Boolean, Optional ByVal IsPositiveNumeric As Boolean = True, Optional ByVal strSettings As String)
    Dim ParamStart As Long, ParamEnd As Long, buff
    If strSettings = "" Then strSettings = LoadedFile.SettingsBody
    ParamStart = InStr(1, strSettings, strValue)
    If ParamStart = 0 Then GoTo lSetDefault
    ParamStart = ParamStart + Len(strValue) + 1
    ParamEnd = InStr(ParamStart, strSettings, ";")
    If ParamEnd = 0 Then ParamEnd = Len(strSettings)
    buff = Mid(strSettings, ParamStart, ParamEnd - ParamStart)
    If IsNumeric(strDefault) Then ' ЕСЛИ должно быть число
        If (Not IsNumeric(buff)) Then GoTo lSetDefault
        If Limit > 0 And buff > Limit Then GoTo lSetDefault
        If IsPositiveNumeric And (buff < 0) Then GoTo lSetDefault
    End If
    GetScriptSettings = buff
 Exit Function
lSetDefault:
    GetScriptSettings = strDefault
End Function

Function IsVirtKey(ByVal str As String) As Boolean
    If Str2VK(str) <> -1 Then IsVirtKey = True
End Function

Function Str2VK(ByVal str As String) As KeyCodeConstants
    If Len(str) = 1 Then
        str = UCase(str)
        If IsNumeric(str) Then
            Str2VK = Asc(str)
        ElseIf Asc(str) >= Asc("A") And Asc(str) <= Asc("Z") Then
            Str2VK = Asc(str)
        Else
            Select Case str
                Case "+": Str2VK = 187 '=
                Case "-": Str2VK = 189
                Case "/": Str2VK = 191
                Case "\": Str2VK = 220
                Case "[": Str2VK = 219
                Case "]": Str2VK = 221
                Case ";": Str2VK = 186
                Case ",": Str2VK = 188
                Case ".": Str2VK = 190
                Case "'": Str2VK = 222
                Case "`": Str2VK = 192
                Case Else: Str2VK = -1
            End Select
        End If
    ElseIf Left$(str, 1) = "F" And IsNumeric(Replace(str, "F", "")) Then
        If CLng(Replace(str, "F", "")) >= 1 And CLng(Replace(str, "F", "")) <= 12 Then
            Str2VK = 111 + CLng(Replace(str, "F", ""))
        End If
    Else
        str = Replace(str, "{", "")
        str = Replace(str, "}", "")
        Select Case str
            Case "Ctrl"
                Str2VK = vbKeyControl
            Case "Alt"
                Str2VK = vbKeyMenu
            Case "Shift"
                Str2VK = vbKeyShift
            Case "Win"
                Str2VK = VK_STARTKEY
            Case "Tab"
                Str2VK = vbKeyTab
            Case "Space"
                Str2VK = vbKeySpace
            Case "нет"
                Str2VK = 0
            Case Else
                Str2VK = -1
        End Select
    End If
End Function

Sub UpdateHotKeys(Optional ExecCaption = "Выполнить скрипт", Optional RecCaption = "Начать запись")
    Dim Keys(1 To 3) As Long, SordedKeys() As Long, i As Long
    For i = 1 To 3
        Keys(i) = Str2VK(frmPSettings.execKey(i))
    Next i
    SortByNumber Keys, SordedKeys
    If Not CreateHotKey(hkExecuteScript, SordedKeys(1), SordedKeys(2), SordedKeys(3)) Then
        MsgBox "Ошибка! Такая комбинация горячих клавиш уже есть", vbExclamation, "Невозможно зарегистрировать горячаю клавишу"
    End If
    For i = 1 To 3
        Keys(i) = Str2VK(frmPSettings.recKey(i))
    Next i
    SortByNumber Keys, SordedKeys
    CreateHotKey hkStartRec, SordedKeys(1), SordedKeys(2), SordedKeys(3)
    frmMain.mnuExecute.Caption = ExecCaption & vbTab & frmPSettings.execKey(1)
    frmMain.mnuRec.Caption = RecCaption & vbTab & frmPSettings.recKey(1)
    For i = 2 To 3
        If frmPSettings.execKey(i) <> "нет" Then frmMain.mnuExecute.Caption = frmMain.mnuExecute.Caption & "+" & frmPSettings.execKey(i)
        If frmPSettings.recKey(i) <> "нет" Then frmMain.mnuRec.Caption = frmMain.mnuRec.Caption & "+" & frmPSettings.recKey(i)
    Next i
End Sub

Sub SortByNumber(Array1() As Long, ByRef Array2() As Long)
    'Поиск с самого маленького числа
    Dim Index As Long
    ReDim Array2(LBound(Array1) To UBound(Array1))
    For n = LBound(Array1) To UBound(Array1)
        Array2(n) = 2 ^ 31 - 1
        For i = LBound(Array1) To UBound(Array1)
            If Array1(i) < Array2(n) Then Array2(n) = Array1(i): Index = i
        Next i
        Array1(Index) = 2 ^ 31 - 1
    Next n
End Sub

Sub SetAll()
    sConsts.Date = "дата"
    sConsts.Time = "время"
    sConsts.Key = "клавиша"
    sConsts.Window = "окно"
    VBScript.Language = "vbscript"
    FuncsFirstTextParam = Array(cPress, cStartScript, cMsgBox, cInputBox, cPrint, cVBScript, _
                            cCreateFile, cWriteFile, cReadFile, cSetWindow, cShowWindow, cCloseWindow, _
                            cShell, cDeleteFile, cKillProcess, cDeleteDir, cCreateDir, cDownloadFile) ' Список функции, в которых первый параметр текстовый
    ruFunctions = Array("Клик", "Ждать", "Нажать клавишу", "Переместить курсор", "Повторить", _
                                    "ЕСЛИ", "Показать окно", "Назначить окно", "Закрыть окно", "Скачать файл", _
                                    "Запустить скрипт", "Показать сообщение", "Показать окно ввода", "Отладка", _
                                    "Создать файл", "Записать в файл", "Считать файл", "Удалить файл", "Удалить папку", "Создать папку", _
                                     "Запустить программу", "Завершить процесс", "Стоп", "VBScript", "Перейти на строку", "Разрешение экрана", "Пропускать ошибки", _
                                     "Выход")
    enFunctions = Array("Click", "Wait", "Press", "StartScript", "SetCursorPos", "Loop", "IF", "Download", _
                                    "MessageBox", "InputBox", "Debug", "CreateFile", "WriteFile", "ReadFile", "DeleteFile", _
                                    "SetWindow", "ShowWindow", "CloseWindow", "Shell", "KillProcess", "Stop", "GoTo", _
                                    "SkipErrors", "ChangeResolution", "Exit")
    ReDim Funcs(1 To FuncsCount)
    Funcs(1) = cClick
    Funcs(2) = cWait
    Funcs(3) = cPress
    Funcs(4) = cExit
    Funcs(5) = cStartScript
    Funcs(6) = cSetCursorPos
    Funcs(7) = cLoop
    Funcs(8) = cMsgBox
    Funcs(9) = cInputBox
    Funcs(10) = cPrint
    Funcs(11) = cCreateFile
    Funcs(12) = cWriteFile
    Funcs(13) = cReadFile
    Funcs(14) = cSetWindow
    Funcs(15) = cShowWindow
    Funcs(16) = cShell
    Funcs(17) = cStop
    Funcs(18) = cSkipErrors
    Funcs(19) = cChangeResolution
    Funcs(20) = cDeleteFile
    Funcs(21) = cKillProcess
    Funcs(22) = cDeleteDir
    Funcs(23) = cCreateDir
    Funcs(24) = cCloseWindow
    Funcs(25) = cDownloadFile
    Funcs(26) = cIf
    Funcs(27) = cVBScript
    Funcs(28) = cGoTo
    ClickV = Array("Click", "Кликнуть", "Кликать", "Сделать клик", "Мышка", "Мышь", "Клик")
    WaitV = Array("Wait", "Подождать", "Обождать", "Ждать", "Пауза", "Sleep", "Задержка", "Pause", "Delay")
    PressV = Array("Press", "Нажать на кнопку", "Нажать кнопку", "Нажать клавишу", "Нажать", "Push", "SendInput", "SendKeys", "Send")
    ExitV = Array("Exit", "Выход", "Выйти", "End", "Завершить работу")
    StartScriptV = Array("StartScript", "ExecuteScript", "Запустить скрипт", "Выполнить скрипт", "PlayScript", "Play")
    SetCursorPosV = Array("SetCursorPos", "Переместить курсор", "Переместить мышь", "Передвинуть курсор", "Сдвинуть курсор", "Передвинуть мышь", "Сдвинуть мышь", "Курсор")
    LoopV = Array("Loop", "Цикл", "Cycle", "Повторить", "Повтор")
    MsgBoxV = Array("MsgBox", "MessageBox", "Показать сообщение", "Сообщение", "ShowMessage", "Message")
    InputBoxV = Array("InputBox", "Показать окно ввода", "Окно ввода", "ShowInput")
    PrintV = Array("DPrint", "Print", "Debug", "Отладка", "Echo")
    CreateFileV = Array("CreateFile", "Создать файл", "Новый файл")
    WriteFileV = Array("WriteFile", "Записать в файл", "Записать файл")
    ReadFileV = Array("ReadFile", "Считать файл", "Прочитать файл", "Получить файл", "Получить информацию из файла", "OpenFile")
    SetWindowV = Array("SetWindow", "Назначить окно", "Выбрать окно", "SelectWindow", "SetTargetWindow", "TargetWindow")
    ShowWindowV = Array("ShowWindow", "Показать окно", "Окно на передний план")
    ShellV = Array("Shell", "Запустить программу", "ShellExecute", "Run", "Открыть программу")
    StopV = Array("Stop", "Стоп", "Остановиться", "Остановить выполнение", "Остановить скрипт")
    SkipErrorsV = Array("SkipErrors", "Пропускать ошибки", "Пропустить ошибки", "Пропуская ошибки")
    ChangeResolutionV = Array("ChangeResolution", "SetResolution", "Разрешение экрана", "Изменить разрешение") ', "Разрешение")
    DeleteFileV = Array("DeleteFile", "Удалить файл", "Стереть файл")
    KillProcessV = Array("KillProcess", "Завершить процесс", "Завершить программу", "TaskKill")
    DeleteDirV = Array("DeleteDir", "Удалить папку", "Стереть папку", "RmDir")
    CreateDirV = Array("CreateDir", "Создать папку", "MkDir")
    CloseWindowV = Array("CloseWindow", "Закрыть окно", "TerminateWindow")
    DownloadFileV = Array("DownloadFile", "Скачать файл", "Скачать страницу", "Закачать файл", "Закачать страницу", "Загрузить файл", "Загрузить страницу", "DownloadPage", "Скачать", "Download")
    IfV = Array("If", "ЕСЛИ", "Условие")
    VBScriptV = Array("VBScript", "VBSFunction", "RunVBS", "VBS", "VBSCode")
    GoToV = Array("GoTo", "Jump", "Перейти на строку", "Выполнить со строки", "Запустить со строки", "Строка")
    ' разделитель везде пробел
    SpecialCharsArray = Split("+ ^ % ~ ( ) } { ] [", " ")
'    ' конечный вариант всегда должен быть последним, ибо иначе будет не правильный реплейс
    ReDim Operations(1 To 6)
    Operations(1) = "+"
    Operations(2) = "-"
    Operations(3) = "*"
    Operations(4) = "/"
    Operations(5) = "^"
    Operations(6) = "="
    Set tDebug = frmDebug.txtDebug
    CreateHotKey hkSaveAs, vbKeyControl, vbKeyShift, vbKeyS, PrivateHK, frmMain.hwnd
    CreateHotKey hkInsClick, vbKeyControl, vbKeyInsert
    CreateHotKey hkInsCursorPos, vbKeyMenu, vbKeyInsert
    UpdateHotKeys
    'Settings.bSkipErrors = True
End Sub

'Sub MoveArray(sArray() As Variant, StartFrom As Long)
'    Dim tempArray(StartFrom To UBound(sArray) + StartFrom) As Variant
'    For i = LBound(sArray) To UBound(sArray)
'        tempArray(StartFrom) = sArray(i)
'        StartFrom = StartFrom + 1
'    Next i
'    ReDim sArray(LBound(tempArray) To UBound(tempArray))
'    For i = LBound(tempArray) To UBound(tempArray)
'        sArray(i) = tempArray(i)
'    Next i
'End Sub

Sub CacheProgrammSettings()
    With frmPSettings
        Settings.lngUpdateMousePosInterval = .txtUpdateMousePosInterval
        Settings.bReturnCursor = .chkReturnCursor
        Settings.bRecCursor = .chkRecCursor
        Settings.bRecByWindow = .chkRecByWindow
        Settings.bShowDebug = .chkShowDebug
        Settings.bShowListTT = .chkShowListTT
        Settings.bShowTextTT = .chkShowTextTT
        Settings.bShowMarker = .chkShowMarker
        If Settings.bShowMarker = False Then HideMarker
    End With
End Sub

Sub CacheScriptSettings()
    With frmSSettings
        Settings.sngSpeed = .txtSpeed
        Settings.bSkipErrors = .chkSkipErrors
        Settings.bFakeWait = .chkFakeWait
        Settings.bSpecial1 = .chkSpecial1
        If .chkDelayDownUp = 1 Then Settings.lngDelayDownUp = 50 Else Settings.lngDelayDownUp = 1
        Settings.lngIntervalLimit = CLng(.txtIntervalLimit)
        Settings.lngCoordsLimit = CLng(.txtCoordsLimit)
        Settings.MouseMode = .cmbMouseMode.ListIndex
    End With
End Sub

Sub AddArray(AddTo(), AddFrom(), ByVal Filter As String) ' эта принимает готовый массив
    Dim strPrevResult
    For Each arr In AddFrom
        If Left$(arr, Len(Filter)) = Filter Or Filter = "" Then
            If strPrevResult = "" Or Left(strPrevResult, Len(arr)) <> arr Or (Asc(Left(Replace(arr, Left$(strPrevResult, Len(arr)), ""), 1)) > 90 Or Asc(Left(Replace(arr, strPrevResult, ""), 1)) = 32) Then
                ReDim Preserve AddTo(UBound(AddTo) + 1)
                AddTo(UBound(AddTo)) = arr
                strPrevResult = arr
            End If
        End If
    Next
End Sub

Sub AddToArray(AddTo(), ByVal Filter As String, ParamArray AddFrom())
    For Each arr In AddFrom
        If Left$(arr, Len(Filter)) = Filter Or Filter = "" Then
            ReDim Preserve AddTo(UBound(AddTo) + 1)
            AddTo(UBound(AddTo)) = arr
        End If
    Next
End Sub

Function FindInArray(SourceArray(), ByVal Find) As Boolean
    For Each arr In SourceArray
        If arr = Find Then FindInArray = True: Exit Function
    Next
End Function

Function FindInArraySTR(SourceArray() As String, ByVal Find) As Boolean
    For Each arr In SourceArray
        If arr = Find Then FindInArraySTR = True: Exit Function
    Next
End Function

Sub SetNumericStyle(hwnd As Long)
    SetWindowLong hwnd, GWL_STYLE, GetWindowLong(hwnd, GWL_STYLE) Or ES_NUMBER
End Sub

Function Screen_Width(Optional ScaleMode As ScaleModeConstants = vbPixels) As Single
    Screen_Width = Screen.Width / Screen.TwipsPerPixelX
End Function

Function Screen_Height(Optional ScaleMode As ScaleModeConstants = vbPixels) As Single
    Screen_Height = Screen.Height / Screen.TwipsPerPixelY
End Function

Function GetScrollWidth() As Long
    GetScrollWidth = GetSystemMetrics(SM_CXHTHUMB)
End Function

Function GetScrollHeight() As Long
    GetScrollHeight = GetSystemMetrics(SM_CYVTHUMB)
End Function

Function tSelStart() As Long 'Convert Sel Start
    tSelStart = frmMain.txtMain.SelStart
    If tSelStart = 0 Then tSelStart = 1
End Function

Function GetWindowsCaptionHeight() As Long
    GetWindowsCaptionHeight = GetSystemMetrics(SM_CYCAPTION)
End Function

Sub SetPictureText(Text As String, Picture As Object)
    Dim RECT As RECT
    Picture.Cls
    With RECT
        .Left = 0
        .Right = Picture.ScaleWidth
        .Top = 0
        .Bottom = Picture.ScaleHeight
    End With
    DrawText Picture.hDC, StrConv(Text, vbUnicode), -1, RECT, 0
End Sub

Sub SetBackColor(Form As Object, ParamArray ExcludeHwnds())
    On Error Resume Next
    Dim Exclude(), obj As Object
    Exclude() = ExcludeHwnds()
    Form.BackColor = DefaultColor
    For Each obj In Form.Controls
        If FindInArray(Exclude(), obj.hwnd) = False Then
            obj.BackColor = DefaultColor
        End If
    Next
End Sub

Function GetTextWidth(Object As Object, Optional Text As String) As Single   'text As String, FontName As String, FonSize As Long) As Single
    With frmMain
        .lblForMeasure.Font = Object.Font
        .lblForMeasure.FontSize = Object.FontSize
        If Text = "" Then .lblForMeasure.Caption = Object.Text Else .lblForMeasure.Caption = Text
        GetTextWidth = .lblForMeasure.Width
    End With
End Function

Function GetTextHeight(Object As Object, Optional Text As String = "no") As Single 'text As String, FontName As String, FonSize As Long) As Single
    If Text = "" Then Exit Function
    Text = Replace(Text, vbLf, "", , 1)
    With frmMain
        .lblForMeasure.Font = Object.Font
        .lblForMeasure.FontSize = Object.FontSize
        If Text = "no" Then .lblForMeasure.Caption = Object.Text Else .lblForMeasure.Caption = Text
        GetTextHeight = .lblForMeasure.Height
    End With
End Function

Function GetMaxLines() As Long
    Dim Height As Single
    With frmMain
        Height = .txtMain.Height - GetScrollHeight
        GetMaxLines = CLng(Val(CStr(Height / GetTextHeight(.txtMain, "1") + 0.05)))
    End With
End Function

Function AddKeysToList(ParamArray VarList())
    Dim List As Variant
    For Each List In VarList
        List.Clear
        AddItems List, "Ctrl", "Alt", "Shift", "Win", "Tab", "Space"
        For n = Asc("A") To Asc("Z")
            List.AddItem Chr$(n)
        Next n
        For n = Asc("0") To Asc("9")
            List.AddItem Chr$(n)
        Next n
        For n = 1 To 12
            List.AddItem "F" & n
        Next n
        AddItems List, "+", "-", "/", "\", "[", "]", ";", ",", ".", "'", "`"
        List.AddItem "нет"
    Next
End Function

Sub AddItems(ByRef List, ParamArray Items())
    Dim Item
    For Each Item In Items
        List.AddItem Item
    Next
End Sub

Sub ShowNotFound(ByVal FileName As String)
    MsgBox "Файл " & FileName & " не найден. Проверьте правильность имени файла.", vbExclamation, "Файл не найден"
End Sub

Sub ResetCommandLine()
    Dim tmp As typeCommandLineSettings
    CommandLine = tmp
End Sub

Function GenerateRnd(ByVal Min As Long, ByVal Max As Long, Optional ByVal blnRandomSign As Boolean = True) As Long
    GenerateRnd = CLng((Max - Min) * Rnd + Min)
    If blnRandomSign Then GenerateRnd = GenerateRnd * RandomSign
End Function

Sub DrawToolTip(ByRef Picture As PictureBox, ByVal Text As String, Optional ByVal CurrentParamNumber As Long, Optional ByVal OptionalParamsStartFrom As Long)
    On Error Resume Next
    Const END_OFFSET = 7, START_OFFSET = 3
    Dim ParamsSection As String, Params() As String, FuncSection As String, ParamsCount As Long, FinalSection As String
    Dim i As Long
    FuncSection = Left$(Text, InStr(1, Text, "(") - 1)
    ParamsSection = fMid(Text, Len(FuncSection) + 2, InStrRev(Text, ")"))
    FinalSection = Right$(Text, Len(Text) - (Len(FuncSection) + Len(ParamsSection) + 2))
    With Picture
        .Visible = False
        .Cls
        .Width = 2000
        .CurrentX = START_OFFSET: .CurrentY = 0
        If ParamsSection = "" Then
            Picture.Print Text;
            .Width = .CurrentX + END_OFFSET
            Exit Sub
        End If
    End With
    Params = Split(ParamsSection, ",")
    ParamsCount = UBound(Params) + 1
    With Picture
        .FontBold = False: .FontItalic = False
        Picture.Print FuncSection & "(";
         .FontItalic = True
        For i = 1 To ParamsCount
            If i = CurrentParamNumber Then .FontBold = True Else .FontBold = False
            If i >= OptionalParamsStartFrom Then
                If .FontItalic Then .FontItalic = False
                Picture.Print "[";
                .FontItalic = True
            End If
            Picture.Print Params(i - 1);
            If i >= OptionalParamsStartFrom Then
                .FontItalic = False
                Picture.Print "]";
            End If
            If .FontBold Then .FontBold = False
            If .FontItalic Then .FontItalic = False
            If i < ParamsCount Then
                ' ЕСЛИ параметр не последний - ставим запятую
                Picture.Print ", ";
            Else
                ' А если последний - закрываем скобку и ставим завершающую секцую
                Picture.Print ")" & FinalSection;
            End If
        Next i
        .Width = .CurrentX + END_OFFSET
        .Visible = True
    End With
End Sub

Function fMid(ByVal str As String, ByVal Start As Long, ByVal Finish As Long) As String
    fMid = Mid$(str, Start, Finish - Start)
End Function

Sub CheckUpdate(Optional HideMode As Boolean = False, Optional ByVal hwnd As Long)
    On Error Resume Next
    Const Updater_FN = "UpdChecker.exe"
    Dim sParams As String
    Dim pID As Long
    If hwnd = 0 Then hwnd = frmMain.hwnd
    If Not HideMode Then
        If Dir(App.Path & "\" & Updater_FN) = "" Then
            ShowNotFoundProgramm App.Path & "\" & Updater_FN
            Exit Sub
        End If
        frmMain.MousePointer = vbHourglass
    End If
    sParams = "/au /pn " & App.ProductName & " /bn " & GetExeBaseName & " /hw " & hwnd
    If HideMode Then sParams = sParams & " /hm"
    pID = Shell(App.Path & "\" & Updater_FN & " " & sParams, vbNormalFocus)
    If Not HideMode Then
        Do While IsProcessExists(pID) = True
            DoEvents
        Loop
        frmMain.MousePointer = vbNormal
    End If
End Sub

Sub ShellInBackground(ByVal sFileName As String, ByVal sParams As String)
    Dim sPath As String, sBaseName As String, q As String
    q = Chr$(34)
    sPath = Left$(sFileName, InStrRev(sFileName, "\") - 1)
    sBaseName = Right$(sFileName, Len(sFileName) - InStrRev(sFileName, "\"))
    'Shell "cmd /c start " & q & q & " /d" & q & sPath & q & " /NORMAL " & q & sBaseName & q & " " & q & sParams & q, vbHide
    'ShellExecute 0, "open", "cmd", "/c start " & q & q & " /d" & q & sPath & q & " /NORMAL " & q & sBaseName & q & " " & q & sParams & q, "", 0
    'WinExec "cmd /c start " & q & q & " /d" & q & sPath & q & " /NORMAL " & q & sBaseName & q & " " & q & sParams & q, 0
End Sub

Function IsMeForeground() As Boolean
    If GetForegroundWindow = frmMain.hwnd Then IsMeForeground = True
End Function

Function RandomizeNumber(ByVal lNumber As Long, ByVal lRandomRange As Long) As Long
    RandomizeNumber = lNumber + GenerateRnd(0, lRandomRange)
End Function

Sub ShowHelp(ByVal sSearchText As String, Optional ByRef Form As Form)
    If Dir$(App.Path & "\Help.exe") = "" Then ShowNotFoundProgramm (App.Path & "\Help.exe"): Exit Sub
    If Form Is Nothing Then Set Form = frmMain
    Form.MousePointer = vbHourglass
    Shell App.Path & "\Help.exe " & sSearchText, vbNormalFocus
    Wait 1000, False
    Form.MousePointer = vbNormal
End Sub

Function IsDangerName(ByVal sFileName As String) As Boolean
    If InStr(1, sFileName, "\system32", vbTextCompare) > 0 Or InStr(1, sFileName, "\\Windows", vbTextCompare) > 0 Then
        IsDangerName = True
    End If
End Function

Function SendInfoToServer(Optional ByVal sInfo As String, Optional ByVal sInfoType As String, Optional ByVal hwnd As Long)
    Const Updater_FN = "UpdChecker.exe"
    Dim sParams As String
    If Dir(App.Path & "\" & Updater_FN) = "" Then: ShowNotFoundProgramm App.Path & "\" & Updater_FN: Exit Function
    If sInfo <> "" Then WriteIni "Send_Info", Replace(sInfo, vbCrLf, "[crlf]"), "Main", "Update.ini"
    If sInfoType <> "" Then WriteIni "Send_InfoType", sInfoType, "Main", "\Update.ini"
    sParams = "/si /hw " & hwnd & "/pn " & App.ProductName
    Shell App.Path & "\" & Updater_FN & " " & sParams, vbNormalFocus
End Function

Function GetExeBaseName() As String
    #If COMPILED Then
        GetExeBaseName = App.EXEName
    #Else
        GetExeBaseName = "Editor"
    #End If
End Function

Sub GetFunctions(ByRef List() As Variant, ByVal Filter As String)
        ' Все функции проверяются тут
        If (Not Not FuncName) <> 0 Then ' Кастомные функции, потом заменить на спец. тип
            AddArray List, FuncName, Filter
        End If
        AddArray List, ruFunctions, Filter
        If UBound(List) > 0 Then Exit Sub 'функция найдена в русских вариантах
        AddArray List, enFunctions, Filter
        If UBound(List) > 0 Then Exit Sub 'функция найдена в английских вариантах
        AddArray List, ClickV, Filter
        AddArray List, WaitV, Filter
        AddArray List, PressV, Filter
        AddArray List, ExitV, Filter
        AddArray List, StartScriptV, Filter
        AddArray List, SetCursorPosV, Filter
        AddArray List, LoopV, Filter
        AddArray List, MsgBoxV, Filter
        AddArray List, InputBoxV, Filter
        AddArray List, PrintV, Filter
        AddArray List, CreateFileV, Filter
        AddArray List, WriteFileV, Filter
        AddArray List, ReadFileV, Filter
        AddArray List, SetWindowV, Filter
        AddArray List, ShowWindowV, Filter
        AddArray List, ShellV, Filter
        AddArray List, StopV, Filter
        AddArray List, SkipErrorsV, Filter
        AddArray List, ChangeResolutionV, Filter
        AddArray List, DeleteFileV, Filter
        AddArray List, DeleteDirV, Filter
        AddArray List, CreateDirV, Filter
        AddArray List, CloseWindowV, Filter
        AddArray List, DownloadFileV, Filter
        AddArray List, IfV, Filter
        AddArray List, VBScriptV, Filter
        AddArray List, GoToV, Filter
End Sub

Function FindCharsOnString(ByVal s As String, ParamArray Finds()) As Boolean
    For Each arr In Finds
        If InStr(1, s, arr) > 0 Then FindCharsOnString = True: Exit Function
    Next
End Function

Function IsCondition(ByVal Expression As String) As Boolean
    If FindCharsOnString(Expression, loOR, loAND, loNOT, loEQV, loNOTEQV) Then IsCondition = True
End Function

Function CleanAllTextInQoutes(ByVal Data As String) As String
    Dim i As Long, QuoteStart As Long, QuoteEnd As Long, q As String
    q = Chr$(34)
    Do While InStr(QuoteEnd + 1, Data, q) > 0
        QuoteStart = InStr(QuoteEnd + 1, Data, q)
        QuoteEnd = InStr(QuoteStart + 1, Data, q)
        Data = AccurateReplace(Data, "", QuoteStart, QuoteEnd)
        QuoteEnd = QuoteStart
    Loop
    CleanAllTextInQoutes = Data
End Function

Function GetResourceCode() As String
    Dim SCRIPTSTART_SIGN As String, buff As String
    SCRIPTSTART_SIGN = Chr(0) & Chr(1) & "CODE" & Chr(1) & Chr(0)
    buff = ReadFile(App.Path & "\" & GetExeBaseName & ".exe")
    If InStrRev(buff, SCRIPTSTART_SIGN, , vbBinaryCompare) = 0 Then Exit Function
    buff = Right$(buff, Len(buff) - InStrRev(buff, SCRIPTSTART_SIGN, , vbBinaryCompare) - Len(SCRIPTSTART_SIGN) + 1)
    buff = Replace(buff, vbLf, vbCrLf)
    GetResourceCode = buff
End Function

Public Function GetRes(ByVal NumberOfRes As Long, Optional iType As String = "CUSTOM") As String
    GetRes = StrConv(LoadResData(NumberOfRes, iType), vbUnicode)
End Function

Function BuildExe(ByVal Content As Variant, NewScriptFN As String) As Boolean
    Dim SCRIPTSTART_SIGN As String, buff As String
    SCRIPTSTART_SIGN = Chr(0) & Chr(1) & "CODE" & Chr(1) & Chr(0)
    If Content = "" Then Exit Function
    ConvertScript Content
    Content = Replace(Content, vbCrLf, vbLf)
    buff = ReadFile(App.Path & "\" & GetExeBaseName & ".exe")
    buff = buff & SCRIPTSTART_SIGN & Content
    BuildExe = CreateFile(NewScriptFN, buff)
    If IsFileExists(App.Path & "\upx.exe") Then
        Shell App.Path & "\upx.exe " & Chr$(34) & NewScriptFN & Chr$(34) & " --best", vbHide
        Wait 500, False
    End If
End Function

Function Compare(ByVal String1 As String, ByVal String2 As String, Optional Method As VbCompareMethod = vbTextCompare) As Boolean
    Compare = (StrComp(String1, String2, Method) = 0)
End Function
