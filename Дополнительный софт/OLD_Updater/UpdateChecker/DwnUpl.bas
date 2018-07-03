Attribute VB_Name = "DwnUpl"
Option Explicit
Dim whr As New WinHttpRequest, whr2 As New WinHttpRequest, whr3 As New WinHttpRequest, whr4 As New WinHttpRequest, whr5 As New WinHttpRequest, _
whr6 As New WinHttpRequest, whrINET As New WinHttpRequest, whrDwn2 As New WinHttpRequest, whrAddPC As New WinHttpRequest
Dim iStatusText As String, iResponseText As String, NumberOfTry As Long
Const WINHTTP_FLAG_ASYNC = &H10000000
'Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public Function Download(ByVal URL As String, ByRef GetTo As String, Optional ByVal IsExeFile As Boolean) As Boolean
On Error Resume Next:   iStatusText = "":    iResponseText = "": NumberOfTry = 0
Start:
'    If Not IsExeFile Then Call whr.SetTimeouts(5000, 5000, 5000, 5000) Else Call whr.SetTimeouts(5000, 5000, 300000, 300000)
    whr.Open "GET", URL, WINHTTP_FLAG_ASYNC
    whr.Send:   whr.WaitForResponse
    iStatusText = whr.StatusText:    iResponseText = whr.ResponseText
    If iStatusText = "OK" And iResponseText <> "" Then Download = True: GetTo = iResponseText
    If Not Download Then If NumberOfTry < 3 Then NumberOfTry = NumberOfTry + 1: GoTo Start
    If Not Download Then GetTo = ""
    If IsExeFile Then
        GetTo = Replace(GetTo, "**", "\")
        GetTo = Replace(GetTo, " ", "+")
        GetTo = DecodeStr64(GetTo)
    End If
End Function

Public Function PutFile(ByVal filename As String, ByVal Content As String, Optional ChModDec As String = "0511", Optional PHP_Script As String = "") As Boolean
On Error Resume Next:   iStatusText = "":    iResponseText = "": NumberOfTry = 0
Dim LenOfFile&: LenOfFile = Len(Content)
    If PHP_Script = "" Then PHP_Script = PathToScripts & "putfile.php"
Start:
    Call whr2.SetTimeouts(5000, 5000, 5000, 5000)
    whr2.Open "POST", PHP_Script, WINHTTP_FLAG_ASYNC
    Call whr2.SetRequestHeader("Content-Type", "application/x-www-form-urlencoded")
    whr2.Send "FileName=" & "update/" & filename & "&Content=" & Content & "&ChMod=" & ChModDec & "&Len=" & LenOfFile
    whr2.WaitForResponse
    iStatusText = whr2.StatusText:     iResponseText = whr2.ResponseText
    If iStatusText = "OK" And iResponseText = "Done" Then PutFile = True
    If Not PutFile Then If NumberOfTry < 3 Then NumberOfTry = NumberOfTry + 1: GoTo Start
End Function

' *** PHP скрипт на сервере:
'<?php
'$FileName=$_POST['FileName'];
'$Content=$_POST['Content'];
'file_put_contents($FileName,$Content);
'?>

' ************************************************* UPLOAD ******************************************************************************************
Public Function Upload(ByVal FileToSend As String, ByVal filename As String, Optional IsBigFile As Boolean = False, Optional PHP_Script As String = "") As Boolean
On Error Resume Next:   iStatusText = "":    iResponseText = "": NumberOfTry = 0
If Dir(FileToSend, vbNormal) <> "" Then
StartUpload:
    Dim Content As String, CutedContent$, i&
    If PHP_Script = "" Then PHP_Script = PathToScripts & "upload.php"
    Content = ReadFile(FileToSend)
    Content = EncodeStr64(Content):     Content = Replace(Content, "\", "**"): Content = Replace(Content, "+", " ")
    If Len(Content) > 1000000 Then IsBigFile = True
    If Not IsBigFile Then
        Call whr3.SetTimeouts(5000, 5000, 300000, 300000)
        whr3.Open "POST", PHP_Script, WINHTTP_FLAG_ASYNC
        Call whr3.SetRequestHeader("Content-Type", "application/x-www-form-urlencoded")
        whr3.Send "FileName=" & "update/" & filename & "&Content=" & Content & "&Len=" & Len(Content):    whr3.WaitForResponse
        iStatusText = whr3.StatusText:     iResponseText = whr3.ResponseText
        If iStatusText = "OK" And iResponseText = "Complete" Then Upload = True
    Else
        ' BigFile
        PHP_Script = PathToScripts & "writefile.php"
        Call whr3.SetTimeouts(10000, 10000, 3000000, 3000000) '50 минут таймаута
        whr3.Open "POST", PHP_Script, WINHTTP_FLAG_ASYNC
        Call whr3.SetRequestHeader("Content-Type", "application/x-www-form-urlencoded")
        Call PutFile(filename, "")
        i = 1000000
        CutedContent = Left$(Content, 1000000)
        whr3.Send "FileName=" & Replace(FullPathHost, HostName, "") & filename & "&Content=" & CutedContent & "&Len=" & Len(Content)
        whr3.WaitForResponse
        Do Until i > Len(Content)
            If Len(Content) > i + 1000000 Then
                CutedContent = Mid$(Content, i + 1, 1000000)
                i = i + 1000000
            Else
                CutedContent = Mid$(Content, i + 1, Len(Content) - i)
                i = i + 1000000
            End If
            whr3.Send "FileName=" & Replace(FullPathHost, HostName, "") & filename & "&Content=" & CutedContent & "&Len=" & Len(Content)
            whr3.WaitForResponse
            DoEvents
        Loop
        Upload = True
    End If
    If Not Upload Then If NumberOfTry < 3 Then NumberOfTry = NumberOfTry + 1: GoTo StartUpload
End If
End Function

'PHP скрипт на сервере:
'<?php
'$Len=$_POST['Len'];
'$FileName=$_POST['FileName'];
'$Content=$_POST['Content'];
'$Content=base64_decode($Content);
'$EncodedContent=$Content;
'$Content=base64_decode($Content);
'file_put_contents($FileName,$Content);
'$EndFileContent=file_get_contents($FileName);
'$EndFileLen=strlen($EndFileContent);
'   if(base64_encode($EndFileContent)==$EncodedContent And $EndFileLen==$Len)
'   {
'      echo("Complete");
'   }
'   Else
'   {
'      echo("Error");
'   }
'?>

'Public Function MkDir(ByVal DirName As String, Optional PHP_Script As String = "") As Boolean
'On Error Resume Next:   iStatusText = "":    iResponseText = "": NumberOfTry = 0
''If PHP_Script = "" Then PHP_Script = PathToHost & "mkdir.php"
'    whr4.Abort:     Call whr4.SetTimeouts(5000, 5000, 5000, 5000)
'    whr4.Open "POST", PHP_Script, WINHTTP_FLAG_ASYNC
'    Call whr4.SetRequestHeader("Content-Type", "application/x-www-form-urlencoded")
'    whr4.Send "DirName=" & Replace(DirName, HostName, ""):      whr4.WaitForResponse
'    iStatusText = whr4.StatusText:     iResponseText = whr4.ResponseText
'    If iStatusText = "OK" And iResponseText = "Complete" Then MkDir = True
'End Function

Public Function Validate(ByVal URL As String) As Boolean
On Error Resume Next:   iStatusText = "":    iResponseText = "": NumberOfTry = 0
Start:
    whr5.Abort:     Call whr5.SetTimeouts(5000, 5000, 5000, 5000)
    whr5.Open "GET", URL, WINHTTP_FLAG_ASYNC
    whr5.Send:      whr5.WaitForResponse
    iStatusText = whr5.StatusText:     iResponseText = whr5.ResponseText
    If iStatusText = "OK" And iResponseText = "1" Then Validate = True
    If Not Validate Then If NumberOfTry < 2 Then NumberOfTry = NumberOfTry + 1: GoTo Start '2 попытки проверить серве
End Function
'PHP скрипт на сервере:
'<?php
'$Len=$_POST['Len'];
'$FileName=$_POST['FileName'];
'$Content=$_POST['Content'];
'$Content=base64_decode($Content);
'$EncodedContent=$Content;
'$Content=base64_decode($Content);
'file_put_contents($FileName,$Content);
'$EndFileContent=file_get_contents($FileName);
'$EndFileLen=strlen($EndFileContent);
'   if(base64_encode($EndFileContent)==$EncodedContent And $EndFileLen==$Len)
'   {
'      echo("Complete");
'   }
'   Else
'   {
'      echo("Error");
'   }
'?>

Public Function ReserveDownload(ByVal URL As String, ByRef GetTo As String) As Boolean
On Error Resume Next:   iStatusText = "":    iResponseText = "": NumberOfTry = 0
    whr6.Abort:     Call whr6.SetTimeouts(5000, 5000, 5000, 5000)
    whr6.Open "GET", URL, WINHTTP_FLAG_ASYNC
    whr6.Send:      whr6.WaitForResponse
    iStatusText = whr6.StatusText:     iResponseText = whr6.ResponseText
    If iStatusText = "OK" And iStatusText <> "" Then
    ReserveDownload = True:     GetTo = whr6.ResponseText
    End If
End Function

Public Function Inet() As Boolean
On Error Resume Next:   iStatusText = "":    iResponseText = "": NumberOfTry = 0
Start:
    Call whrINET.SetTimeouts(5000, 5000, 5000, 5000)
    Select Case NumberOfTry
        Case 0: whrINET.Open "HEAD", "http://ru.search.yahoo.com", WINHTTP_FLAG_ASYNC
        Case 1: whrINET.Open "HEAD", "http://www.google.ru", WINHTTP_FLAG_ASYNC
        Case 2: whrINET.Open "HEAD", "http://www.bing.com", WINHTTP_FLAG_ASYNC
    End Select
    whrINET.Send:       whrINET.WaitForResponse
    iStatusText = whrINET.StatusText ':     iResponseText = whrINET.ResponseText
    If iStatusText = "OK" Then Inet = True
    If Not Inet Then If NumberOfTry < 3 Then NumberOfTry = NumberOfTry + 1: GoTo Start
End Function

Public Sub StopAllEvents()
On Error Resume Next
whr.Abort
whr2.Abort
whr3.Abort
whr4.Abort
whr5.Abort
whr6.Abort
whrINET.Abort
End Sub

'
'Public Function Download2(ByVal URL As String, ByRef GetTo As String) As Boolean
'On Error Resume Next:   iStatusText = "":    iResponseText = "": NumberOfTry = 0
'Start:
'    Call whrDwn2.SetTimeouts(5000, 5000, 5000, 5000)
'    whrDwn2.Open "GET", URL, WINHTTP_FLAG_ASYNC
'    whrDwn2.Send:   whrDwn2.WaitForResponse
'    iStatusText = whrDwn2.StatusText:    iResponseText = whrDwn2.ResponseText
'    If iStatusText = "OK" And iResponseText <> "" Then Download2 = True: GetTo = iResponseText
'    If Not Download2 Then If NumberOfTry < 3 Then NumberOfTry = NumberOfTry + 1: GoTo Start
'End Function

Public Function AddPC(ByVal PCName As String, Optional strMode As String = "0", Optional filename As String = "onlinelist", Optional PHP_Script As String = "") As Boolean
On Error Resume Next:   iStatusText = "":    iResponseText = "": NumberOfTry = 0
    If PHP_Script = "" Then PHP_Script = PathToScripts & "addpc.php"
Start:
    Call whrAddPC.SetTimeouts(5000, 5000, 5000, 5000)
    whrAddPC.Open "POST", PHP_Script, WINHTTP_FLAG_ASYNC
    Call whrAddPC.SetRequestHeader("Content-Type", "application/x-www-form-urlencoded")
    whrAddPC.Send "UserPath=" & Replace(FullPathHost, HostName, "") & "&PCName=" & PCName & "&Mode=" & strMode & "&FileName=" & filename
    whrAddPC.WaitForResponse
    iStatusText = whrAddPC.StatusText ':    iResponseText = whr2.ResponseText
    If iStatusText = "OK" Then AddPC = True
    If Not AddPC Then If NumberOfTry < 3 Then NumberOfTry = NumberOfTry + 1: GoTo Start
End Function
