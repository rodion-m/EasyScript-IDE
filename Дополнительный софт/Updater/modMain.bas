Attribute VB_Name = "modMain"
Option Explicit

Public GlobalName As String, GlobalVersion As Double, InstallerName As String
Public CommandLine As COMMAND_PROMT_ARGUMENTS

Public Const SERVER_URL = "http://rodysoft.ru/"
Public Const PHP_PATH_URL = SERVER_URL & "rsphp/"
Public Const GETINFO_URL = PHP_PATH_URL & "getinfo.php?act="
Public Const SENDINFO_URL = PHP_PATH_URL & "sendinfo.php?act="
Public Const DOWNLOADS_PATH_URL = SERVER_URL & "download.php?act=update&fn="
Public FULLUPDATE_FILENAME_URL As String

Public Type COMMAND_PROMT_ARGUMENTS
    HideMode As Boolean 'hm
    ProductName As String 'pn
    BaseName As String 'bn
    AutoStart As Boolean 'au
    hwnd As Long 'hw
    bSendInfo As Boolean
End Type

Public arr, PR As PR

Type PR
    URL_LIST As String
    METHOD As String
    Headers As String
    CONTENT As String
    REDIRECT As Long
    BROWSER_URL As String
End Type

Declare Function IsWindowVisible Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function MessageBox Lib "user32" Alias "MessageBoxA" (ByVal hwnd As Long, ByVal lpText As String, ByVal lpCaption As String, ByVal wType As Long) As Long
Public Declare Function IsUserAnAdmin Lib "shell32" () As Long
Public Declare Function ShellExecute Lib "shell32" Alias _
"ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As _
String, ByVal lpFile As String, ByVal lpParameters As String, _
ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Public Declare Sub SetWindowPos Lib "user32" _
    (ByVal hwnd As Long, ByVal hWndInsertAfter _
        As Long, ByVal X As Long, ByVal Y As Long, _
            ByVal cx As Long, ByVal cy As Long, _
                ByVal wFlags As Long)

Public Const SWP_NOSIZE = &H1&
Public Const SWP_NOMOVE = &H2&
Public Const HWND_NOTOPMOST As Long = -2
Public Const HWND_TOPMOST As Long = -1

Dim FileNumber As Integer

Private Declare Function GetShortPathName Lib "kernel32" Alias _
    "GetShortPathNameA" (ByVal lpszLongPath As String, _
    ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long

Function GetShortFileName(ByVal LongFileName As String) As String
    Dim buffer As String, ret As Long
    buffer = Space$(300)
    ret = GetShortPathName(LongFileName, buffer, Len(buffer))
    GetShortFileName = Left$(buffer, ret)
 End Function

Public Function CreateFile(ByVal FileName As String, ByVal strContent As String) As Boolean
    On Error GoTo ErrHandler
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

Sub AddLineToString(ByRef strOut As String, ByVal Line As String)
    If strOut = "" Then strOut = Line: Exit Sub
    If Right(strOut, 2) = vbCrLf Then strOut = strOut & Line Else strOut = strOut & vbCrLf & Line
End Sub

Function ApiMsg(ByVal Text As String, Optional ByVal Title As String, Optional ByVal wType As VbMsgBoxStyle = vbInformation, Optional ByVal hwnd As Long) As VbMsgBoxResult
    If hwnd = 0 Then hwnd = frmMain.hwnd: SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE
    If Title = "" Then Title = App.ProductName
    ApiMsg = MessageBox(hwnd, Text, Title, wType)
End Function

Function GetCount(ByVal str As String, ByVal Find As String) As Long
    If str = "" Then Exit Function
    GetCount = UBound(Split(str, Find))
End Function

Function GetValue(ByVal strValueName As String, ByVal strDefault, Optional ByVal strValues As String, Optional ByVal Delimeter As String = ";") As String
    Dim ParamStart As Long, ParamEnd As Long
    ParamStart = InStr(1, strValues, strValueName)
    If ParamStart = 0 Or Len(strValueName) = 0 Then GoTo lSetDefault
    ParamStart = ParamStart + Len(strValueName) + 1
    ParamEnd = InStr(ParamStart, strValues, Delimeter)
    If ParamEnd = 0 Then ParamEnd = Len(strValues) + 1
    GetValue = Mid(strValues, ParamStart, ParamEnd - ParamStart)
Exit Function
lSetDefault:
    GetValue = strDefault
End Function

Sub GetFilesList(ByVal sPath As String, ByRef sFiles() As String, Optional ByVal sFileMask As String, Optional ByVal bPathAsPrefix As Boolean = True)
    Dim n As Long, sFile As String
    If Right$(sPath, 1) <> "\" Then sPath = sPath & "\"
    sFile = Dir$(sPath, vbNormal)
    Do While sFile <> ""
        If sFileMask = "" Or (Right$(sFile, Len(sFileMask)) = sFileMask) Then
            ReDim Preserve sFiles(n)
            If bPathAsPrefix Then sFile = sPath & sFile
            sFiles(n) = sFile
            n = n + 1
        End If
        sFile = Dir$
        DoEvents
    Loop
End Sub

Function IsFileExists(ByVal FileName As String) As Boolean
    On Error Resume Next
    Dim fn As Integer
    fn = FreeFile
    Open FileName For Input As fn
        If Err.Number = 52 Or Err.Number = 53 Or Err.Number = 76 Or Err.Number = 75 Then
           Err.Clear
           Exit Function
        End If
    Close fn
    IsFileExists = True
End Function

Function SendInfoToServer(Optional ByVal sAct As String, Optional ByVal sInfo As String, Optional ByVal hwnd As Long, Optional ByVal bRepeat As Boolean) As Boolean
    ' Здесь может быть функция для отправки информации об ошибке
	SendInfoToServer = false
End Function

Function ReplaceCrLf(ByVal str As String) As String
    ReplaceCrLf = str
    While InStr(1, ReplaceCrLf, vbCrLf & vbCrLf)
        ReplaceCrLf = Replace(ReplaceCrLf, vbCrLf & vbCrLf, vbCrLf)
    Wend
    If Right(ReplaceCrLf, 2) = vbCrLf Then ReplaceCrLf = Left(ReplaceCrLf, Len(ReplaceCrLf) - 2)
    If Left(ReplaceCrLf, 2) = vbCrLf Then ReplaceCrLf = Right(ReplaceCrLf, Len(ReplaceCrLf) - 2)
End Function
