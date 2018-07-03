Attribute VB_Name = "modFileInfo"
Option Explicit
Option Compare Text

Type FileInfo
    BaseName As String
    ProductName As String
    ProductVersion As String
End Type

Sub GetFileInfo(ByVal sFileName As String, ByRef File As FileInfo)
    Dim sContent As String
    sContent = ReadFile(sFileName)
    File.BaseName = Right$(sFileName, Len(sFileName) - InStrRev(sFileName, "\"))
    File.ProductName = GetFileInfoValue(sContent, "ProductName")
    File.ProductVersion = GetFileInfoValue(sContent, "ProductVersion")
End Sub

Function GetFileInfoValue(ByVal sContent As String, ByVal sValueName As String) As String
    Dim sValueNamePrefix As String, sValueNamePostfix As String
    Dim sValuePrefix As String, sValuePostfix As String
    sValueNamePrefix = Chr$(0) & Chr$(1) & Chr$(0)
    Select Case sValueName
        Case "CompanyName", "ProductName", "FileVersion", "FileDescription"
            sValueNamePostfix = String$(5 - 1, 0) 'после каждой буквы в уникоде подставляет ноль, поэтому один ноль в постфиксу вычитается
        Case Else
            sValueNamePostfix = String$(3 - 1, 0)
    End Select
    sValuePostfix = String$(2, 0)
    GetFileInfoValue = GetStringByStartEndSign(sContent, _
                                                                sValueNamePrefix & StrConv(sValueName, vbUnicode) & sValueNamePostfix, _
                                                                sValuePostfix)
    GetFileInfoValue = StrConv(GetFileInfoValue & vbNullChar, vbFromUnicode)
End Function

Function GetStringByStartEndSign(ByVal strSource As String, ByVal strStartSignature As String, Optional ByVal strEndSignature As String, _
                                                        Optional ByRef lngSectionStartPos As Long, Optional ByRef lngSectionEndPos As Long, _
                                                        Optional ByVal lngStartOffset As Long, Optional ByVal lngEndOffset As Long, Optional ByVal lngNumber As Long = 1) As String
    Dim n As Long
    For n = 1 To lngNumber
        lngSectionStartPos = InStr(lngSectionStartPos + 1, strSource, strStartSignature)
    Next n
    If lngSectionStartPos = 0 Then Exit Function
    lngSectionStartPos = lngSectionStartPos + lngStartOffset + Len(strStartSignature)
    lngSectionEndPos = InStr(lngSectionStartPos + 1, strSource, strEndSignature)
    If lngSectionEndPos = 0 Then Exit Function 'lngSectionEndPos = Len(strSource) + 1
    lngSectionEndPos = lngSectionEndPos - lngEndOffset
    GetStringByStartEndSign = Mid$(strSource, lngSectionStartPos, lngSectionEndPos - lngSectionStartPos)
End Function

