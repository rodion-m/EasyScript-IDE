Attribute VB_Name = "mdlFileDialog"
Option Explicit

Public Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (ByRef pOpenfilename As OPENFILENAME) As Long
Public Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (ByRef pOpenfilename As OPENFILENAME) As Long
'Public Declare Function FindText Lib "comdlg32.dll" Alias "FindTextA" (ByRef pFindreplace As FINDREPLACE) As Long
'Public Declare Function ReplaceText Lib "comdlg32.dll" Alias "ReplaceTextA" (ByRef pFindreplace As FINDREPLACE) As Long
'Public Declare Function ChooseColor Lib "comdlg32.dll" Alias "ChooseColorA" (ByRef pChoosecolor As ChooseColor) As Long
'Public Declare Function ChooseFont Lib "comdlg32.dll" Alias "ChooseFontA" (ByRef pChoosefont As ChooseFont) As Long
'Public Declare Function PrintDlg Lib "comdlg32.dll" Alias "PrintDlgA" (ByRef pPrintdlg As PrintDlg) As Long
'Public Declare Function PrintDlgEx Lib "comdlg32.dll" Alias "PrintDlgExA" (ByRef TLPPRINTDLGEXA As PRINTDLGEXA) As Long

Type OPENFILENAME
    lStructSize         As Long
    hwndOwner           As Long
    hInstance           As Long
    lpstrFilter         As String
    lpstrCustomFilter   As Long
    nMaxCustrFilter     As Long
    nFilterIndex        As Long
    lpstrFile           As String
    nMaxFile            As Long
    lpstrFileTitle      As String
    nMaxFileTitle       As Long
    lpstrInitialDir     As String
    lpstrTitle          As String
    Flags               As Long
    nFileOffset         As Integer
    nFileExtension      As Integer
    lpstrDefExt         As String
    lCustrData          As Long
    lpfnHook            As Long
    lpTemplateName      As Long
End Type

Public Type typeFileStruct
    Path               As String
    FileName            As String
    BaseName    As String ' имя с расширением
    BaseTitle       As String ' имя без расширения
    Content         As String
    ScriptBody    As String
    SettingsBody As String
End Type

Public Const OFN_ALLOWMULTISELECT = &H200       'выбор нескольких файлов
Public Const OFN_CREATEPROMPT = &H2000          'Уст., что диалог. окно сообщ. пользователю об отсутствии файла и предлагает его создать
Public Const OFN_EXPLORER = &H80000
Public Const OFN_FILEMUSTEXIST = &H1000         'Уст., что польз. может выбрать только сущ. файл или увидит сообщение "Файл не существует. Проверьте правильность имени файла".
Public Const OFN_HIDEREADONLY = &H4             'Прячет галочку "Только чтение".
Public Const OFN_NOCHANGEDIR = &H8              'Открывать диалог в том же каталоге, в котором он был открыть первый раз
Public Const OFN_NODEREFERENCELINKS = &H100000
Public Const OFN_NONETWORKBUTTON = &H20000
Public Const OFN_NOREADONLYRETURN = &H8000
Public Const OFN_NOVALIDATE = &H100             'Можно использовать любые символы при указании имени файла.
Public Const OFN_OVERWRITEPROMPT = &H2
Public Const OFN_PATHMUSTEXIST = &H800
Public Const OFN_READONLY = &H1
Public Const OFN_SHOWHELP = &H10                'Показывает кнопку справки.

Function OpenFile(ByVal Title As String, ByVal InitDir As String, ByVal FileName As String, ByVal Filter As String, ByVal hwnd As Long, Optional ByVal FilterIndex As Long = 1) As typeFileStruct
    Dim FileDialog As OPENFILENAME
    If Filter = vbNullString Then Filter = "Все файлы (*.*)" & Chr$(0) & "*.*"
    With FileDialog
        .hwndOwner = hwnd
        .lpstrFilter = Filter
        .nFilterIndex = FilterIndex
        .lpstrFile = String$(1024, 0)
        .nMaxFile = 1023
        .lpstrFileTitle = String$(1024, 0)
        .nMaxFileTitle = 1023
        .lpstrTitle = Title
        .lpstrInitialDir = InitDir
        .lpstrDefExt = "mde"
        .Flags = OFN_CREATEPROMPT + OFN_PATHMUSTEXIST + OFN_EXPLORER + OFN_HIDEREADONLY
        .lStructSize = Len(FileDialog)
    End With
    If GetOpenFileName(FileDialog) Then
        With OpenFile
            .FileName = GetStringBeforeNull(FileDialog.lpstrFile)
            If Not IsFileExists(.FileName) Then .FileName = "": Exit Function
            .BaseName = GetStringBeforeNull(FileDialog.lpstrFileTitle)
            .Path = GetPath(.FileName)
            .BaseTitle = GetFileBaseTitle(.FileName)
        End With
    End If
End Function

Function SaveFile(ByVal Title As String, ByVal InitDir As String, ByVal FileName As String, ByVal Filter As String, ByVal hwnd As Long, Optional ByVal FilterIndex As Long = 1) As typeFileStruct
    Dim FileDialog As OPENFILENAME
    If Filter = vbNullString Then Filter = "Все файлы (*.*)" & Chr$(0) & "*.*"
    With FileDialog
        .hwndOwner = hwnd
        .lpstrFilter = Filter
        .nFilterIndex = FilterIndex
        .lpstrFile = FileName & String$(1024 - Len(FileName), 0)
        .nMaxFile = 1023
        .lpstrFileTitle = String$(1024, 0)
        .nMaxFileTitle = 1023
        .lpstrTitle = Title
        .lpstrInitialDir = InitDir
        .lpstrDefExt = "mde"
        .Flags = OFN_CREATEPROMPT + OFN_PATHMUSTEXIST + OFN_EXPLORER + OFN_HIDEREADONLY + OFN_OVERWRITEPROMPT
        .lStructSize = Len(FileDialog)
    End With
    If GetSaveFileName(FileDialog) Then
        With SaveFile
            .FileName = GetStringBeforeNull(FileDialog.lpstrFile)
            If .FileName = "" Then Exit Function
            .BaseName = GetStringBeforeNull(FileDialog.lpstrFileTitle)
            .Path = GetPath(.FileName)
            .BaseTitle = GetFileBaseTitle(.FileName)
        End With
    End If
End Function

Function GetStringBeforeNull(ByVal str As String) As String
    GetStringBeforeNull = Left$(str, InStr(1, str, vbNullChar) - 1)
    If GetStringBeforeNull = "" Then GetStringBeforeNull = str
End Function

Function GetFileBaseTitle(ByVal FileName As String) As String
    Dim BaseName As String, lngDotPos As Long
    BaseName = GetFileBaseName(FileName)
    lngDotPos = InStrRev(BaseName, ".") - 1
    If lngDotPos = -1 Then lngDotPos = Len(BaseName)
    GetFileBaseTitle = Left$(BaseName, lngDotPos)
End Function

Function GetFileBaseName(ByVal FileName As String) As String
    Dim lngSlashPos As Long
    lngSlashPos = InStrRev(FileName, "\")
    GetFileBaseName = Right$(FileName, Len(FileName) - lngSlashPos)
End Function

Function GetPath(ByVal FileName As String) As String
    Dim lngSlashPos As Long
    lngSlashPos = InStrRev(FileName, "\")
    If lngSlashPos > 0 Then GetPath = Left$(FileName, lngSlashPos)
End Function

Function IsFileExists(ByVal FileName As String) As Boolean
    On Error Resume Next
    Dim fn As Integer
    CheckScriptName FileName
    fn = FreeFile
    Open FileName For Input As fn
    If Err.Number = 52 Or Err.Number = 53 Or Err.Number = 76 Or Err.Number = 75 Then
       Err.Clear
       Exit Function
    End If
    Close fn
    IsFileExists = True
End Function

Sub CheckScriptName(ByRef FileName As String)
    If InStr(1, FileName, ".") = 0 Then
        FileName = FileName & ".es"
    End If
    If InStr(1, FileName, "\") = 0 Then
        FileName = ScriptsPath & FileName
    End If
End Sub

Public Function CreateFile(ByVal FileName As String, ByVal strContent As String) As Boolean
    On Error GoTo ErrHandler
    Dim FileNumber As Integer
    If Dir(FileName) <> "" Then Kill (FileName)
    FileNumber = FreeFile
    Open FileName For Binary As FileNumber
        Put FileNumber, , strContent
    Close FileNumber
    CreateFile = True
Exit Function
ErrHandler:
    CreateFile = False
    Err.Clear
End Function

Public Function ReadFile(ByVal FileName As String) As String
    On Error GoTo ErrHandler
    Dim FileNumber As Integer
    FileNumber = FreeFile
    If Dir(FileName) = "" Then Exit Function
    Open FileName For Binary As FileNumber
        ReadFile = Space(LOF(FileNumber))
        Get FileNumber, , ReadFile
    Close FileNumber
Exit Function
ErrHandler:
    ReadFile = ""
    Err.Clear
End Function
