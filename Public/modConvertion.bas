Attribute VB_Name = "modConvertion"
Option Explicit

Private Declare Function MultiByteToWideChar Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpMultiByteStr As String, ByVal cchMultiByte As Long, ByVal lpWideCharStr As String, ByVal cchWideChar As Long) As Long

Private Declare Function WideCharToMultiByte Lib "kernel32" ( _
    ByVal CodePage As Long, ByVal dwFlags As Long, _
    lpWideCharStr As Integer, _
    ByVal cchWideChar As Long, _
    lpMultiByteStr As Byte, _
    ByVal cchMultiByte As Long, _
    ByVal lpDefaultChar As String, _
    ByVal lpUsedDefaultChar As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)

Private Declare Function MultiByteToWideChar_Xakep Lib "kernel32" Alias "MultiByteToWideChar" (ByVal CodePage As Long, ByVal dwFlags As Long, lpMultiByteStr As Any, ByVal cchMultiByte As Long, lpWideCharStr As Any, ByVal cchWideChar As Long) As Long
Private Const CP_UTF8 As Long = 65001

Public Function Utf8AsIsToUcs2(ByRef sAsIsBytes As String) As String
    Dim ret As Long
    Dim nInpLength As Long: nInpLength = LenB(sAsIsBytes)
    If nInpLength = 0 Then Exit Function Else Utf8AsIsToUcs2 = Space(nInpLength)
    ret = MultiByteToWideChar_Xakep(CP_UTF8, 0, _
                              ByVal StrPtr(sAsIsBytes), nInpLength, _
                              ByVal StrPtr(Utf8AsIsToUcs2), nInpLength)
    If ret = 0 Then Error 51 Else Utf8AsIsToUcs2 = Left$(Utf8AsIsToUcs2, ret)
End Function

Public Function Utf8DirtyToUcs2(ByVal sDirtyString As String) As String
    Dim ret As Long
    Dim nInpLength As Long: nInpLength = Len(sDirtyString)
    If nInpLength = 0 Then Exit Function Else Utf8DirtyToUcs2 = Space(nInpLength)
    sDirtyString = StrConv(sDirtyString, vbFromUnicode)
    ret = MultiByteToWideChar_Xakep(CP_UTF8, 0, _
                              ByVal StrPtr(sDirtyString), nInpLength, _
                              ByVal StrPtr(Utf8DirtyToUcs2), nInpLength)
    If ret = 0 Then Error 51 Else Utf8DirtyToUcs2 = Left$(Utf8DirtyToUcs2, ret)
End Function

Public Function Utf8ToUcs2(ByRef sInputUtf8Str() As Byte) As String
    Dim ret As Long
    Dim nInpLength As Long: nInpLength = ByteArySize(sInputUtf8Str)
    If nInpLength = 0 Then Exit Function Else Utf8ToUcs2 = Space(nInpLength)
    ret = MultiByteToWideChar_Xakep(CP_UTF8, 0, _
                              sInputUtf8Str(LBound(sInputUtf8Str)), nInpLength, _
                              ByVal StrPtr(Utf8ToUcs2), nInpLength)
    If ret = 0 Then Error 51 Else Utf8ToUcs2 = Left$(Utf8ToUcs2, ret)
End Function

Private Function ByteArySize(ByRef b() As Byte) As Long
    On Error Resume Next
    ByteArySize = UBound(b) - LBound(b) + 1
End Function

Public Function UTF8ToWin(ByVal inString As String) As String
    Dim iStrSize    As Long
    Dim Uc As String
    If Len(inString) = 0 Then Exit Function
    
    iStrSize& = MultiByteToWideChar(CP_UTF8, 0&, inString, &HFFFF, Uc, 0)
    If Len(iStrSize) Then
         Uc = String$(iStrSize * 2, 0&)
         Call MultiByteToWideChar(CP_UTF8, 0&, inString, &HFFFF, Uc, iStrSize)
         UTF8ToWin = StrConv(Uc, vbFromUnicode)
    End If
End Function

Public Function WinToUTF8(ByVal inString As String) As String
    Dim res As Long
    Dim UTF8Buffer() As Byte
    Dim UTF16Buffer() As Integer
    Dim Length As Long
    Dim Uc$
    
    Length = Len(inString)
    If Length = 0 Then Exit Function
    
    ReDim UTF16Buffer(0 To Length + 1) As Integer
    CopyMemory UTF16Buffer(0), ByVal StrPtr(inString), Length * 2
    '
    
    ' First, find out how big a UTF8 buffer we need:
    res = WideCharToMultiByte(CP_UTF8, 0, UTF16Buffer(LBound(UTF16Buffer)), Length, 0, 0, vbNullString, 0)
    If res = 0 Then Exit Function
    
    ' OK, now do the conversion:
    ReDim UTF8Buffer(0 To res - 1)
    res = WideCharToMultiByte(CP_UTF8, 0, UTF16Buffer(LBound(UTF16Buffer)), Length, UTF8Buffer(0), res, vbNullString, 0)
    If res = 0 Then Exit Function
    
    ' I never got this far in the function to see if this works correctly...
    WinToUTF8 = StrConv(UTF8Buffer, vbUnicode)
End Function

Function GenerateUTF8Data(ParamArray Bytes()) As String
    Dim i As Long, b() As Byte
    ReDim b(UBound(Bytes))
    For i = 0 To UBound(Bytes)
        b(i) = CByte(Bytes(i))
    Next i
    GenerateUTF8Data = StrConv(b, vbUnicode)
End Function

Function IsUTF8(ByVal Data As String) As Boolean
    Dim utf8Signs() As String, c As Long
    If Len(Data) < 2 Then Exit Function
    GetUTF8Signs utf8Signs()
    For c = LBound(utf8Signs) To UBound(utf8Signs)
        If InStrB(1, Data, utf8Signs(c), vbBinaryCompare) > 0 Then IsUTF8 = True: Exit Function
    Next c
End Function

Function GetUTF8Signs(ByRef sOut() As String)
    Dim c As Long, i As Long
    ReDim sOut(65)
    For c = 144 To 191 'À-ï
        sOut(i) = ChrB(&HD0) & ChrB(c)
        i = i + 1
    Next c
    For c = 128 To 145 'ï-¸
        sOut(i) = ChrB(&HD1) & ChrB(c)
        i = i + 1
    Next c
End Function

Function IsUnicode(ByVal Data As String) As Boolean
    If InStrB(1, Data, Chr$(0), vbBinaryCompare) > 0 Then IsUnicode = True
End Function
