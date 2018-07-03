Attribute VB_Name = "mm"

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
