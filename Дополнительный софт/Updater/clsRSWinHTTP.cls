VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRSWinHTTP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'RaiseEvent OnResponseStart[(arg1, arg2, ... , argn)]
Event OnResponseStart(ByVal Status As Long, ByVal ContentType As String)
Event Complete(ByVal Status As Long, ByVal StatusText As String)
Event Progress(ByVal Progress As Long, ByVal ProgressMax As Long)
Event OnError(ByVal ErrorNumber As Long, ByVal ErrorDescription As String)

Private WithEvents WHR As WinHttpRequest
Attribute WHR.VB_VarHelpID = -1
Dim sOutPUBLIC As Object, lProgressCounter As Long, lContentSize As Long, bUpdateProgress As Boolean
Public Data As String
Attribute Data.VB_VarUserMemId = 0
Public ResponseText As String
Public ResponseBody As String
Public ResponseHeaders As String
Public Status As Long
Public StatusText As String

Public Sub Abort()
    On Error Resume Next
    WHR.Abort
    Set WHR = Nothing
End Sub

Public Function DownloadToString(ByVal sURL As String, Optional ByRef sOut As String, Optional ByVal bSetProgress As Boolean = True, Optional ByVal METHOD = "POST") As Boolean
    On Error GoTo ErrHandler
    Dim b() As Byte, sHeads As String
    Set WHR = New WinHttpRequest
    Data = ""
    lProgressCounter = 0
    ResponseText = ""
    ResponseBody = ""
    Status = 0
    StatusText = ""
    If bSetProgress Then lContentSize = GetContentSize(sURL)
    bUpdateProgress = bSetProgress
    RaiseEvent OnResponseStart(0, "") ' Закачка началась...
    WHR.Open METHOD, sURL, WINHTTP_FLAG_ASYNC
    sHeads = _
    "Cache-Control: no-chache" & vbCrLf & _
    "User-Agent: " & USERAGENT & vbCrLf & _
    "Accept: text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8" & vbCrLf & _
    "Referer: http://rodysoft.ru/updater" & vbCrLf & _
    "Accept-Language: ru-RU" & vbCrLf & _
    "Accept-Charset: utf-8"
    SetWinHTTP_Headers WHR, sHeads
    WHR.Send
    WHR.WaitForResponse
    b = WHR.ResponseBody
    sOut = Utf8ToUcs2(b)
    With PR
        .URL_LIST = ParseData(sOut, "pr url", True)
        .METHOD = ParseData(sOut, "pr method", True)
        .Headers = ParseData(sOut, "pr headers", True)
        .CONTENT = ParseData(sOut, "pr content", True)
        .REDIRECT = CLng(ParseData(sOut, "pr redirect", True, 0))
        .BROWSER_URL = ParseData(sOut, "pr browser_url", True)
    End With
    sOut = ParseData(sOut)
    If WHR.Status = 200 Then DownloadToString = True
    Exit Function
ErrHandler:
    Err.Clear
End Function

Private Function ParseData(ByVal sData As String, Optional sDataTeg As String = "rodysoft data", Optional bIfNotExistSetNull As Boolean, Optional ByVal sDefault As String) As String
    Dim sTegStart As String, sTegEnd As String
    sTegStart = "<" & sDataTeg & ">"
    sTegEnd = "</" & sDataTeg & ">"
    ParseData = GetStringByStartEndSign(sData, sTegStart, sTegEnd)
    If ParseData = "" And Not bIfNotExistSetNull Then ParseData = sData
    If ParseData = "" Then ParseData = sDefault
End Function

Public Function DownloadToFile(ByVal sURL As String, Optional ByVal sFileName As String) As Boolean
    DownloadToFile = DownloadToString(sURL)
    If DownloadToFile Then DownloadToFile = CreateFile(sFileName, Data)
End Function

Private Sub WHR_OnResponseDataAvailable(Data() As Byte)
    If Not bUpdateProgress Then Exit Sub
    lProgressCounter = lProgressCounter + UBound(Data)
    RaiseEvent Progress(lProgressCounter, lContentSize)
End Sub

Private Sub WHR_OnResponseFinished()
    Dim b() As Byte
    If Len(WHR.ResponseBody) > 0 Then
        b = WHR.ResponseBody
        Data = StrConv(b, vbUnicode)
    End If
    ParseData Data
    ResponseText = WHR.ResponseText
    ResponseBody = WHR.ResponseBody
    Status = WHR.Status
    StatusText = WHR.StatusText
    RaiseEvent Complete(WHR.Status, WHR.StatusText)
End Sub

