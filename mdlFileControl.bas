Attribute VB_Name = "mdlFileControl"
Option Explicit
Dim FileNumber As Integer

Public Function CreateFile(ByVal FileName As String, ByVal strContent As String) As Boolean
    On Error GoTo ErrHandler
    If IsFileExists(FileName) Then Kill (FileName)
    FileNumber = FreeFile
    Open FileName For Binary As FileNumber
        Put FileNumber, , strContent
    Close FileNumber
    CreateFile = True
Exit Function
ErrHandler:
    CreateFile = False
End Function

Public Function ReadFile(ByVal FileName As String) As String
    On Error GoTo ErrHandler
    FileNumber = FreeFile
    If IsFileExists(FileName) = False Then Exit Function
    Open FileName For Binary As FileNumber
        ReadFile = Space(LOF(FileNumber))
        Get FileNumber, , ReadFile
    Close FileNumber
Exit Function
ErrHandler:
    ReadFile = ""
End Function

Public Function WriteFile(ByVal FileName As String, ByVal strContent As String) As Boolean
    On Error GoTo ErrHandler
    FileNumber = FreeFile
    Open FileName For Binary As FileNumber
        Put FileNumber, LOF(FileNumber) + 1, strContent
    Close FileNumber
    WriteFile = True
Exit Function
ErrHandler:
    WriteFile = False
End Function

Public Function WriteFileRev(ByVal FileName As String, ByVal strContent As String) As Boolean
    On Error GoTo ErrHandler
    Dim sOld As String
    FileNumber = FreeFile
    sOld = ReadFile(FileName)
    strContent = strContent & sOld
    WriteFileRev = CreateFile(FileName, strContent)
Exit Function
ErrHandler:
    WriteFileRev = False
End Function

Public Sub ReadByLine(ByVal FileName As String, ByRef ReadTo() As String)
    On Error Resume Next
    Dim Index As Long
    If IsFileExists(FileName) = False Then Exit Sub
    FileNumber = FreeFile
    Open FileName For Input As FileNumber
        Do Until EOF(FileNumber)
            ReDim Preserve ReadTo(Index)
            Line Input #FileNumber, ReadTo(Index)
            Index = Index + 1
        Loop
    Close FileNumber
End Sub
