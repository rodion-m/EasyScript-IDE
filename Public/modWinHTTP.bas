Attribute VB_Name = "modWinHTTP"
Option Explicit

Public sHeads As String, WHR As New WinHttpRequest, bDownloaderOn As Boolean

Public Const HTTPREQUEST_PROXYSETTING_DEFAULT = 0
Public Const HTTPREQUEST_PROXYSETTING_PRECONFIG = 0
Public Const HTTPREQUEST_PROXYSETTING_DIRECT = 1
Public Const HTTPREQUEST_PROXYSETTING_PROXY = 2
Public Const WINHTTP_FLAG_ASYNC = &H10000000
Public Const USERAGENT = "RodySoft EasyScript"

Type ResponseInfo
    Text As String
    Status As Long
    Headers As String
    Cookie As String
    Body As String
End Type

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMs As Long)

Function GetContentSize(ByVal sURL As String) As Long
    On Error Resume Next
    Dim WHR As New WinHttpRequest, buff As String
    WHR.Open "HEAD", sURL, WINHTTP_FLAG_ASYNC
    WHR.SetTimeouts 5000, 5000, 5000, 5000
    WHR.Option(WinHttpRequestOption_EnableRedirects) = 0
    WHR.Send
    WHR.WaitForResponse
    buff = WHR.GetResponseHeader("Content-Length")
    If Not IsNumeric(buff) Then Exit Function
    GetContentSize = CLng(buff)
End Function

Function GetHeaderName(ByVal sHeader As String) As String
    Dim i As Long
    For i = 1 To Len(sHeader)
        If VBA.Mid$(sHeader, i, 2) = ": " Then
            GetHeaderName = VBA.Left$(sHeader, i - 1)
        End If
    Next i
End Function

Function GetCookieByHeaders(ByVal sHeads As String) As String
     Dim sBuffer As String, lPos As Long
     Do
        sBuffer = GetStringByStartEndSign(sHeads, "Set-Cookie: ", ";", lPos)
        If lPos = 0 Then Exit Do
        If GetCookieByHeaders = "" Then
            GetCookieByHeaders = sBuffer
        Else
            GetCookieByHeaders = GetCookieByHeaders & "; " & sBuffer
        End If
    Loop
End Function

Function GetStringByStartEndSign(ByVal strSource As String, ByVal strStartSignature As String, Optional ByVal strEndSignature As String, _
                                                        Optional ByRef lngSectionStartPos As Long, Optional ByRef lngSectionEndPos As Long, _
                                                        Optional ByVal lngStartOffset As Long, Optional ByVal lngEndOffset As Long, Optional ByVal lngNumber As Long = 1, _
                                                        Optional ByVal bContinueIfEndPosIsZero As Boolean) As String
    Dim n As Long
    For n = 1 To lngNumber
        lngSectionStartPos = InStr(lngSectionStartPos + 1, strSource, strStartSignature)
    Next n
    If lngSectionStartPos = 0 Then Exit Function
    lngSectionStartPos = lngSectionStartPos + lngStartOffset + Len(strStartSignature)
    lngSectionEndPos = InStr(lngSectionStartPos + 1, strSource, strEndSignature)
    If lngSectionEndPos = 0 Or strEndSignature = "" Then If bContinueIfEndPosIsZero Then lngSectionEndPos = Len(strSource) + 1 Else Exit Function
    lngSectionEndPos = lngSectionEndPos - lngEndOffset
    GetStringByStartEndSign = VBA.Mid$(strSource, lngSectionStartPos, lngSectionEndPos - lngSectionStartPos)
End Function

Function GetHeaderValue(ByVal sHeaders As String, ByVal sHeaderName As String) As String
    sHeaders = Replace(sHeaders, vbCrLf, vbLf)
    GetHeaderValue = GetStringByStartEndSign(sHeaders, sHeaderName & ": ", vbLf, , , , , , True)
End Function

Function WinHTTP_EXEC(ByVal sURL As String, Optional ByVal sHeads As String, Optional ByVal sContent As String, Optional ByVal sCookie As String, Optional Method = "GET", Optional EnableRedirects As Long = 0) As ResponseInfo
    'On Error GoTo ErrHandler
    Dim sHeadsArray, sHeader As String, sValue As String, arr
    Dim b() As Byte
    WHR.Open Method, sURL, WINHTTP_FLAG_ASYNC
    WHR.SetTimeouts 10000, 10000, 10000, 10000
    WHR.Option(WinHttpRequestOption_EnableRedirects) = EnableRedirects
    If sHeads = "" Then
    sHeads = _
        "Cache-Control: no-chache" & vbCrLf & _
        "User-Agent: " & USERAGENT & vbCrLf & _
        "Accept: text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8" & vbCrLf & _
        "Accept-Language: ru-RU" & vbCrLf & _
        "Accept-Charset: utf-8" & vbCrLf & _
        "Accept-Encoding: gzip"
    End If
    sHeadsArray = Split(sHeads, vbCrLf)
    For Each arr In sHeadsArray
        sHeader = GetHeaderName(arr)
        sValue = Right$(arr, Len(arr) - Len(sHeader) - 2)
        WHR.SetRequestHeader sHeader, sValue
    Next
    If sCookie <> "" Then WHR.SetRequestHeader "Cookie", sCookie
    If Compare(Method, "POST") Then WHR.SetRequestHeader "Content-Type", "application/x-www-form-urlencoded"
    WHR.Send sContent
    WHR.WaitForResponse
    With WinHTTP_EXEC
        .Headers = WHR.GetAllResponseHeaders
        .Status = WHR.Status
        .Cookie = GetCookieByHeaders(WinHTTP_EXEC.Headers)
        .Body = WHR.ResponseBody
        If GetHeaderValue(.Headers, "Content-Encoding") = "gzip" Then .Body = DecompressGZip(.Body)
        If IsUnicode(.Body) Then .Body = StrConv(.Body, vbUnicode)
        If IsUTF8(.Body) Then .Body = Utf8AsIsToUcs2(.Body)
    End With
Exit Function
ErrHandler:
    Err.Clear
End Function
