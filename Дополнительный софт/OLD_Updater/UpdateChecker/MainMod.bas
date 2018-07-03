Attribute VB_Name = "MainMod"
Public Declare Function ShellExecute Lib "shell32.dll" Alias _
"ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As _
String, ByVal lpFile As String, ByVal lpParameters As String, _
ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
'для runas = ShellExecute 0, "runas", "C:\Windows\system32\cmd.exe", "", "C:\Windows\system32", 1

Public Declare Sub SetWindowPos Lib "USER32" _
    (ByVal hwnd As Long, ByVal hWndInsertAfter _
        As Long, ByVal x As Long, ByVal y As Long, _
            ByVal cx As Long, ByVal cy As Long, _
                ByVal wFlags As Long)

Public Const HWND_TOPMOST = -1
Public Const SWP_NOSIZE = &H1
Public Const SWP_NOMOVE = &H2

Public Function Text2Hex(ByVal iText As String) As String
On Error Resume Next
Dim iChar$, OutText As String, HexChar$
For n = 1 To Len(iText)
iChar = Mid$(iText, n, 1)
If Len(Hex(Asc(iChar))) > 1 Then HexChar = Hex(Val(Asc(iChar))) Else HexChar = "0" & Hex((Asc(iChar)))
OutText = OutText & HexChar
DoEvents
Next n
Text2Hex = OutText
End Function

Public Function Hex2Text(ByVal iTextInHex As String) As String
On Error Resume Next
Dim iChar$, OutText As String, OutChar$
For n = 1 To Len(iTextInHex) Step 2
iChar = Mid$(iTextInHex, n, 2)
OutChar = Chr$(CLng("&H" & iChar))
OutText = OutText & OutChar
DoEvents
Next n
Hex2Text = OutText
End Function

