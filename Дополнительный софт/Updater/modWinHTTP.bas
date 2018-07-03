Attribute VB_Name = "modWinHTTP"
Option Explicit

Public Const WINHTTP_FLAG_ASYNC = &H10000000
Public Const USERAGENT = "Mozilla/5.0 (compatible; MSIE 9.0; Windows NT 6.1; Trident/5.0)"

Type ResponseInfo
    Text As String
    Body As String
    Status As Long
    Headers As String
    Cookie As String
End Type

Private Declare Function MultiByteToWideChar Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, lpMultiByteStr As Any, ByVal cchMultiByte As Long, lpWideCharStr As Any, ByVal cchWideChar As Long) As Long
Private Const CP_UTF8 = 65001

Public Function Utf8ToUcs2(ByRef sInputUtf8Str() As Byte) As String
    Dim ret As Long
    Dim nInpLength As Long: nInpLength = ByteArySize(sInputUtf8Str)
    If nInpLength = 0 Then Exit Function Else Utf8ToUcs2 = Space(nInpLength)
    ret = MultiByteToWideChar(CP_UTF8, 0, _
                              sInputUtf8Str(LBound(sInputUtf8Str)), nInpLength, _
                              ByVal StrPtr(Utf8ToUcs2), nInpLength)
    If ret = 0 Then Error 51 Else Utf8ToUcs2 = Left$(Utf8ToUcs2, ret)
End Function

Private Function ByteArySize(ByRef b() As Byte) As Long
    On Error Resume Next
    ByteArySize = UBound(b) - LBound(b) + 1
End Function

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

Function WinHTTP_EXEC(ByVal sURL As String, Optional ByVal sHeads As String, Optional ByVal sContent As String, Optional ByVal sCookie As String, Optional ByVal METHOD = "GET", Optional ByVal EnableRedirects As Long = 0) As ResponseInfo
    On Error GoTo ErrHandler
    Dim WHR As New WinHttpRequest
    Dim b() As Byte
    WHR.Open METHOD, sURL, WINHTTP_FLAG_ASYNC
    WHR.SetTimeouts 10000, 10000, 10000, 10000
    WHR.Option(WinHttpRequestOption_EnableRedirects) = EnableRedirects
    SetWinHTTP_Headers WHR, sHeads
    If sCookie <> "" Then WHR.SetRequestHeader "Cookie", sCookie
    WHR.Send sContent
    WHR.WaitForResponse
    WinHTTP_EXEC.Headers = WHR.GetAllResponseHeaders
    WinHTTP_EXEC.Status = WHR.Status
    WinHTTP_EXEC.Cookie = GetCookieByHeaders(WinHTTP_EXEC.Headers)
    WinHTTP_EXEC.Text = WHR.ResponseText
    b = WHR.ResponseBody
    WinHTTP_EXEC.Body = Utf8ToUcs2(b)
    Exit Function
ErrHandler:
    Err.Clear
End Function

Sub SetWinHTTP_Headers(ByRef WHR As WinHttpRequest, ByVal sHeads As String)
    Dim sHeadsArray() As String, sHeader As String, sValue As String, arr
    If Len(sHeads) < 2 Then Exit Sub
    sHeadsArray = Split(sHeads, vbCrLf)
    For Each arr In sHeadsArray
        sHeader = GetHeaderName(arr)
        sValue = Right$(arr, Len(arr) - Len(sHeader) - 2)
        WHR.SetRequestHeader sHeader, sValue
    Next
End Sub

Private Function GetHeaderName(ByVal sHeader As String) As String
    Dim i As Long
    For i = 1 To Len(sHeader)
        If Mid$(sHeader, i, 2) = ": " Then
            GetHeaderName = Left$(sHeader, i - 1)
        End If
    Next i
End Function

Private Function GetCookieByHeaders(ByVal sHeads As String) As String
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

Private Function GetStringByStartEndSign(ByVal strSource As String, ByVal strStartSignature As String, Optional ByVal strEndSignature As String, _
                                                        Optional ByRef lngSectionStartPos As Long, Optional ByRef lngSectionEndPos As Long, _
                                                        Optional ByVal lngStartOffset As Long, Optional ByVal lngEndOffset As Long, Optional ByVal lngNumber As Long = 1) As String
    Dim n As Long
    For n = 1 To lngNumber
        lngSectionStartPos = InStr(lngSectionStartPos + 1, strSource, strStartSignature)
    Next n
    If lngSectionStartPos = 0 Then Exit Function
    lngSectionStartPos = lngSectionStartPos + lngStartOffset + Len(strStartSignature)
    lngSectionEndPos = InStr(lngSectionStartPos + 1, strSource, strEndSignature)
    If lngSectionEndPos = 0 Then Exit Function
    lngSectionEndPos = lngSectionEndPos - lngEndOffset
    GetStringByStartEndSign = Mid$(strSource, lngSectionStartPos, lngSectionEndPos - lngSectionStartPos)
End Function

Function GetHeaderValue(ByVal sHeaders As String, ByVal sHeaderName As String) As String
    sHeaders = Replace(sHeaders, vbCrLf, vbLf)
    GetHeaderValue = GetStringByStartEndSign(sHeaders, sHeaderName & ": ", vbLf)
End Function
