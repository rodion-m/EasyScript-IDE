Attribute VB_Name = "mdlRecord"
Option Explicit
Public blnRec As Boolean
Dim X As Long, Y As Long, RecTo As Object, SmallCounter As Long, blnClick As Boolean, _
    PrevCounter As Long, PrevX As Long, PrevY As Long, n As Long, k As String 'k - Key

Dim WindowCaption As String, lStartTime As Long, lTime As Long

Sub StartRec()
    Dim vKey As Long, FuncCounter As Long, Char As String
    Static antirecurse As Boolean
    If antirecurse Then Exit Sub
    antirecurse = True
    Set RecTo = frmMain.txtMain
    blnRec = True
    frmMain.UpdateControlsState
    frmMain.txtStatus = "Активирован режим записи..."
    Resolution.X = 0
    Resolution.X = 0
    lStartTime = GetPerformanceTime
    Do While blnRec
        If CheckHotKeys(hkStartRec) Then Exit Do
        'LeftMouseKey
        If GetKey(vbKeyLButton, True) Then
            If FuncCounter = 0 Then: FuncCounter = FuncCounter + 1: GoTo lLoop
            GetCursorPos X, Y
            PrevX = X
            PrevY = Y
            CheckRecordMod X, Y 'проверка, не идет ли запись относительно окна
            CheckResolution
            AddLine "Клик(" & X & ", " & Y & ", " & "Вниз" & ", " & "Левая , 1 раз, " & ReturnCursor & ", " & Ms2Str(GetCounter) & ")"
            WaitForKeyUp vbKeyLButton, True, "Левая"
            GetCursorPos X, Y
            CheckRecordMod X, Y 'проверка, не идет ли запись относительно окна
            If Not blnClick Then AddLine "Клик(" & X & ", " & Y & ", " & "Вверх" & ", " & "Левая, 1 раз, " & ReturnCursor & ", " & Ms2Str(GetCounter) & ")"
            FuncCounter = FuncCounter + 1
        End If
        'RightMouseKey
        If GetKey(vbKeyRButton, True) Then
            If FuncCounter = 0 Then: FuncCounter = FuncCounter + 1: GoTo lLoop
            GetCursorPos X, Y
            PrevX = X
            PrevY = Y
            CheckRecordMod X, Y 'проверка, не идет ли запись относительно окна
            CheckResolution
            AddLine "Клик(" & X & ", " & Y & ", " & "Вниз" & ", " & "Правая , 1 раз, " & ReturnCursor & ", " & Ms2Str(GetCounter) & ")"
            WaitForKeyUp vbKeyRButton, True, "Правая"
            GetCursorPos X, Y
            CheckRecordMod X, Y 'проверка, не идет ли запись относительно окна
            If Not blnClick Then AddLine "Клик(" & X & ", " & Y & ", " & "Вверх" & ", " & "Правая, 1 раз, " & ReturnCursor & ", " & Ms2Str(GetCounter) & ")"
            FuncCounter = FuncCounter + 1
        End If
        'MiddleMouseKey
        If GetKey(vbKeyMButton, True) Then
            If FuncCounter = 0 Then: FuncCounter = FuncCounter + 1: GoTo lLoop
            GetCursorPos X, Y
            PrevX = X
            PrevY = Y
            CheckRecordMod X, Y 'проверка, не идет ли запись относительно окна
            CheckResolution
            AddLine "Клик(" & X & ", " & Y & ", " & "Вниз" & ", " & "Средняя , 1 раз, " & ReturnCursor & ", " & Ms2Str(GetCounter) & ")"
            WaitForKeyUp vbKeyMButton, True, "Средняя"
            GetCursorPos X, Y
            CheckRecordMod X, Y 'проверка, не идет ли запись относительно окна
            If Not blnClick Then AddLine "Клик(" & X & ", " & Y & ", " & "Вверх" & ", " & "Средняя, 1 раз, " & ReturnCursor & ", " & Ms2Str(GetCounter) & ")"
            FuncCounter = FuncCounter + 1
        End If
        'Keys
        Char = ScanKey
        If Char <> "" Then
            For Each arr In SpecialCharsArray
                If Char = arr Then Char = "{" & Char & "}"
            Next
            CheckRecordMod 0, 0 'проверка, не идет ли запись относительно окна
            If Not CheckHotKeys(hkStartRec) Then AddLine "Нажать клавишу(" & Chr$(34) & GetSpecialKeys & Char & Chr$(34) & ", 1 раз, " & Ms2Str(GetCounter) & ")"
        End If
lLoop:
        DoEvents
    Loop
    frmMain.UpdateAll
    antirecurse = False
End Sub

Function ReturnCursor() As String
    If Settings.bReturnCursor Then ReturnCursor = "Да" Else ReturnCursor = "Нет"
End Function

Sub CheckResolution()
    If (GetResolution.X <> Resolution.X) Or (GetResolution.Y <> Resolution.Y) Then
        Resolution = GetResolution
        AddLine "Разрешение экрана(" & Resolution.X & ", " & Resolution.Y & ")"
    End If
End Sub

Function GetResolution() As POINTAPI
    GetResolution.X = Screen_Width
    GetResolution.Y = Screen_Height
End Function

Sub CheckRecordMod(ByRef X As Long, ByRef Y As Long)
    Dim ForegroundWindowCaption As String
    If Settings.bRecByWindow Then
        ForegroundWindowCaption = GetForegroundWindowCaption
        If WindowCaption <> ForegroundWindowCaption And ForegroundWindowCaption <> frmMain.Caption Then
            If ForegroundWindowCaption = "" Then GoTo lSetW
            GetWindowRect GetForegroundWindow, Window
            X = JustPositive(X - Window.Left)
            Y = JustPositive(Y - Window.Top)
lSetW:
            WindowCaption = ForegroundWindowCaption
            AddLine "Назначить окно(" & Chr$(34) & WindowCaption & Chr$(34) & ")"
        End If
    End If
End Sub

Function GetSpecialKeys() As String
    If GetKey(vbKeyControl, True) Then GetSpecialKeys = "{Ctrl}+"
    If GetKey(vbKeyMenu, True) Then GetSpecialKeys = GetSpecialKeys & "{Alt}+"
    If GetKey(VK_STARTKEY, True) Then GetSpecialKeys = "{Win}+"
'    If GetKey(vbKeyShift, True) Then GetSpecialKeys = GetSpecialKeys & "{Shift}+"
End Function

Sub WaitForKeyUp(vKey As Long, Optional blnIsMouse As Boolean = False, Optional MouseKey As String)
    SmallCounter = 0
    blnClick = True
    Do While GetKey(vKey, True)
        If (GetCounter(False) >= Settings.lngUpdateMousePosInterval) And blnIsMouse Then
            GetCursorPos X, Y
            If X <> PrevX Or Y <> PrevY Then
                blnClick = False
                AddLine "Передвинуть курсор(" & X & ", " & Y & ", " & Ms2Str(GetCounter) & ")"
                PrevX = X
                PrevY = Y
            End If
        End If
        DoEvents
        If CheckHotKeys(hkStartRec) Then Exit Do
        If blnRec = False Then Exit Do
    Loop
    If GetCounter(False) > 100 Then blnClick = False
    If blnClick Then
        RecTo.Text = Left(RecTo.Text, Len(RecTo.Text) - Len("Клик(" & X & ", " & Y & ", " & "Вниз" & ", " & MouseKey & " , 1 раз, " & ReturnCursor & ", " & Ms2Str(PrevCounter) & ")"))
        AddLine "Клик(" & X & ", " & Y & ", " & "Клик" & ", " & MouseKey & ", 1 раз, " & ReturnCursor & ", " & Ms2Str(PrevCounter + GetCounter) & ")"
    End If
End Sub

Sub AddLine(line As String)
    k = ""
    If RecTo.Text = "" Then RecTo.Text = line: Exit Sub
    If Right(RecTo.Text, 2) = vbCrLf Then RecTo.Text = RecTo.Text & line Else RecTo.Text = RecTo.Text & vbCrLf & line
    RecTo.SelStart = Len(RecTo.Text)
End Sub

Function ScanKey() As String
    k = ""
    If Not CheckMouse Then Exit Function
    If GetLang = En Then
        'Пишим англ. клавиши:
        For n = 65 To 90 '128
            If GetKey(n) Then
                k = LCase(Chr(n))
                GoTo SetKey
            End If
        Next n
        'Спец символы Англ
        If GetKey(186) Then If GetShift Then k = ":": GoTo SetKey Else k = ";": GoTo SetKey
        If GetKey(188) Then If GetShift Then k = "<": GoTo SetKey Else k = ",": GoTo SetKey
        If GetKey(190) Then If GetShift Then k = ">": GoTo SetKey Else k = ".": GoTo SetKey
        If GetKey(191) Then If GetShift Then k = "?": GoTo SetKey Else k = "/": GoTo SetKey
        If GetKey(192) Then If GetShift Then k = "~": GoTo SetKey Else k = "`": GoTo SetKey
        If GetKey(219) Then If GetShift Then k = "{": GoTo SetKey Else k = "[": GoTo SetKey
        If GetKey(220) Then If GetShift Then k = "|": GoTo SetKey Else k = "\": GoTo SetKey
        If GetKey(221) Then If GetShift Then k = "}": GoTo SetKey Else k = "]": GoTo SetKey
        If GetKey(222) Then If GetShift Then k = Chr(34): GoTo SetKey Else k = "'": GoTo SetKey
        If GetKey(vbKey2) Then If GetShift Then k = "@": GoTo SetKey Else k = "2": GoTo SetKey
        If GetKey(vbKey3) Then If GetShift Then k = "#": GoTo SetKey Else k = "3": GoTo SetKey
        If GetKey(vbKey4) Then If GetShift Then k = "$": GoTo SetKey Else k = "4": GoTo SetKey
        If GetKey(vbKey6) Then If GetShift Then k = "^": GoTo SetKey Else k = "6": GoTo SetKey
        If GetKey(vbKey7) Then If GetShift Then k = "&": GoTo SetKey Else k = "7": GoTo SetKey
    Else
        'Русские клавиши
        For n = 65 To 90 '128
            If GetKey(CLng(n)) Then
            k = LCase(Chr(n))
            Select Case k
                Case "q": k = "й": Case "w": k = "ц": Case "e": k = "у": Case "r": k = "к": Case "t": k = "е": Case "y": k = "н": Case "u": k = "г": Case "i": k = "ш": Case "o": k = "щ": Case "p": k = "з"
                Case "a": k = "ф": Case "s": k = "ы": Case "d": k = "в": Case "f": k = "а": Case "g": k = "п": Case "h": k = "р": Case "j": k = "о": Case "k": k = "л": Case "l": k = "д"
                Case "z": k = "я": Case "x": k = "ч": Case "c": k = "с": Case "v": k = "м": Case "b": k = "и": Case "n": k = "т": Case "m": k = "ь"
            End Select
            GoTo SetKey
            End If
        Next n
        'Спец символы Рус
        If GetKey(186) Then '*** ж
            k = "ж"
            GoTo SetKey
        End If
        If GetKey(188) Then '*** б
            k = "б"
            GoTo SetKey
        End If
        If GetKey(190) Then '*** ю
            k = "ю"
            GoTo SetKey
        End If
        If GetKey(192) Then '*** ё
            k = "ё"
            GoTo SetKey
        End If
        If GetKey(219) Then '*** х
            k = "х"
            GoTo SetKey
        End If
        If GetKey(222) Then '*** э
            k = "э"
            GoTo SetKey
        End If
        If GetKey(221) Then '*** ъ
            k = "ъ"
            GoTo SetKey
        End If
        If GetKey(191) Then If GetShift Then k = ",": GoTo SetKey Else k = ".": GoTo SetKey
        If GetKey(220) Then If GetShift Then k = "/": GoTo SetKey Else k = "\": GoTo SetKey
        If GetKey(vbKey2) Then If GetShift Then k = Chr(34): GoTo SetKey Else k = "2": GoTo SetKey
        If GetKey(vbKey3) Then If GetShift Then k = "№": GoTo SetKey Else k = "3": GoTo SetKey
        If GetKey(vbKey4) Then If GetShift Then k = ";": GoTo SetKey Else k = "4": GoTo SetKey
        If GetKey(vbKey6) Then If GetShift Then k = ":": GoTo SetKey Else k = "6": GoTo SetKey
        If GetKey(vbKey7) Then If GetShift Then k = "?": GoTo SetKey Else k = "7": GoTo SetKey
    End If
    'Действия, не зависимые от проверок
    If GetKey(187) Then If GetShift Then k = "+": GoTo SetKey Else k = "=": GoTo SetKey
    If GetKey(189) Then If GetShift Then k = "_": GoTo SetKey Else k = "-": GoTo SetKey
    If GetKey(32) Then k = " ": GoTo SetKey
    If GetKey(vbKeyMultiply) Then k = "*": GoTo SetKey
    If GetKey(vbKeyAdd) Then k = "+": GoTo SetKey
    If GetKey(vbKeySubtract) Then k = "-": GoTo SetKey
    If GetKey(vbKeyDecimal) Then If GetLang = En Then k = ".": GoTo SetKey Else k = ",": GoTo SetKey
    If GetKey(vbKeyDivide) Then k = "/": GoTo SetKey
    If GetKey(vbKey0) Then If GetShift Then k = ")": GoTo SetKey Else k = "0": GoTo SetKey
    If GetKey(vbKey1) Then If GetShift Then k = "!": GoTo SetKey Else k = "1": GoTo SetKey
    If GetKey(vbKey5) Then If GetShift Then k = "%": GoTo SetKey Else k = "5": GoTo SetKey
    If GetKey(vbKey8) Then If GetShift Then k = "*": GoTo SetKey Else k = "8": GoTo SetKey
    If GetKey(vbKey9) Then If GetShift Then k = "(": GoTo SetKey Else k = "9": GoTo SetKey
    If GetKey(vbKeyReturn) Then k = "{Enter}": GoTo SetKey
    If GetKey(vbKeyEscape) Then k = "{Esc}": GoTo SetKey
    If GetKey(vbKeyBack) Then k = "{BackSpace}": GoTo SetKey
    'SpecialKeys
'    If GetKey(vbKeyControl) Then k = "{Ctrl}": GoTo SetKey
'    If GetKey(vbKeyMenu) Then k = "{Alt}": GoTo SetKey
'    If GetKey(vbKeyCapital) Then k = "{Caps Lock}": GoTo SetKey
'    If GetKey(vbKeyShift) Then k = "{Shift}": GoTo SetKey
    If GetKey(vbKeyTab) Then k = "{Tab}": GoTo SetKey
    If GetKey(vbKeyLeft) Then k = "{Влево}": GoTo SetKey
    If GetKey(vbKeyUp) Then k = "{Вверх}": GoTo SetKey
    If GetKey(vbKeyRight) Then k = "{Вправо}": GoTo SetKey
    If GetKey(vbKeyDown) Then k = "{Вниз}": GoTo SetKey
    If GetKey(vbKeyDelete) Then k = "{Del}": GoTo SetKey
    If GetKey(vbKeyF1) Then k = "{F1}": GoTo SetKey
    If GetKey(vbKeyF2) Then k = "{F2}": GoTo SetKey
    If GetKey(vbKeyF3) Then k = "{F3}": GoTo SetKey
    If GetKey(vbKeyF4) Then k = "{F4}": GoTo SetKey
    If GetKey(vbKeyF5) Then k = "{F5}": GoTo SetKey
    If GetKey(vbKeyF6) Then k = "{F6}": GoTo SetKey
    If GetKey(vbKeyF7) Then k = "{F7}": GoTo SetKey
    If GetKey(vbKeyF8) Then k = "{F8}": GoTo SetKey
    If GetKey(vbKeyF9) Then k = "{F9}": GoTo SetKey
    If GetKey(vbKeyF10) Then k = "{F10}": GoTo SetKey
    If GetKey(vbKeyF11) Then k = "{F11}": GoTo SetKey
    If GetKey(vbKeyF12) Then k = "{F12}": GoTo SetKey
    'NUMPAD Независимо от языка
    'делать идеальное определение нумпада - чересчур, также есть большая вероятность, что при зажатом шифте, код клавиши меняется..
    If GetNumLock Then
        If GetKey(vbKeyNumpad0) And Not GetShift Then k = "0": GoTo SetKey
        If GetKey(vbKeyNumpad1) And Not GetShift Then k = "1": GoTo SetKey
        If GetKey(vbKeyNumpad2) And Not GetShift Then k = "2": GoTo SetKey
        If GetKey(vbKeyNumpad3) And Not GetShift Then k = "3": GoTo SetKey
        If GetKey(vbKeyNumpad4) And Not GetShift Then k = "4": GoTo SetKey
        If GetKey(vbKeyNumpad5) And Not GetShift Then k = "5": GoTo SetKey
        If GetKey(vbKeyNumpad6) And Not GetShift Then k = "6": GoTo SetKey
        If GetKey(vbKeyNumpad7) And Not GetShift Then k = "7": GoTo SetKey
        If GetKey(vbKeyNumpad8) And Not GetShift Then k = "8": GoTo SetKey
        If GetKey(vbKeyNumpad9) And Not GetShift Then k = "9": GoTo SetKey
    End If
SetKey:
    If GetUpCase And Len(k) = 1 Then k = UCase(k)
    ScanKey = k
End Function

Function Ms2Str(ByVal Ms As Long) As String
    Select Case Ms
    Case 10000 To 600000
        Ms2Str = CStr(Ms \ 1000) & " сек"
    Case Is > 600000
        Ms2Str = CStr(Ms \ 60000) & " мин"
    Case Else
        Ms2Str = CStr(Ms) & " мс"
    End Select
End Function

Function GetCounter(Optional bSetToZero As Boolean = True) As Long
    lTime = GetPerformanceTime
    GetCounter = lTime - lStartTime
    If bSetToZero Then lStartTime = lTime: PrevCounter = GetCounter
End Function

