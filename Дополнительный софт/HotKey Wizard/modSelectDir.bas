Attribute VB_Name = "modSelectDir"
Option Explicit

Declare Function SHBrowseForFolder Lib "shell32.dll" _
    Alias "SHBrowseForFolderA" (lpBrowseInfo As BROWSEINFO) As Long
Declare Function SHGetPathFromIDList Lib "shell32.dll" _
    Alias "SHGetPathFromIDListA" (ByVal pidl As Long, _
    ByVal pszPath As String) As Long

Type BROWSEINFO
    hOwner As Long
    pidlRoot As Long
    pszDisplayName As String
    lpszTitle As String
    ulFlags As Long
    lpfn As Long
    lParam As Long
    iImage As Long
End Type

Public Const BIF_RETURNONLYFSDIRS = &H1

Dim udtBrowseInfo As BROWSEINFO, pid As Long
Dim path As String

Function SelectDir(Optional ByVal sDefault As String) As String
    With udtBrowseInfo
        .hOwner = frmMain.hwnd
        .lpszTitle = "Выберите папку"
        .ulFlags = BIF_RETURNONLYFSDIRS
    End With
    pid = SHBrowseForFolder(udtBrowseInfo)
    path = Space(260)
    SHGetPathFromIDList ByVal pid, ByVal path
    SelectDir = Left$(path, InStr(1, path, Chr(0)) - 1)
    If Len(SelectDir) < 3 Then SelectDir = sDefault
End Function

