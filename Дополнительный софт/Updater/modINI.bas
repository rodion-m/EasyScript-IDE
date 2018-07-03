Attribute VB_Name = "modINI"
Option Explicit

Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" ( _
        ByVal lpApplicationName As String, _
        ByVal lpKeyName As String, _
        ByVal lpString As String, _
        ByVal lpFileName As String) As Long

Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" ( _
        ByVal lpApplicationName As String, _
        ByVal lpKeyName As String, _
        ByVal lpDefault As String, _
        ByVal lpReturnedString As String, _
        ByVal nSize As Long, _
        ByVal lpFileName As String) As Long

Enum TypeOfData
    dNumeric
    dHotKey
    dChk
End Enum

Public Function WriteIni(ByVal strKeyName As String, ByVal strValue As String, Optional iniSection As String = "Main", Optional iniFile As String = "Update.ini") As Long
    iniFile = App.Path & "\" & iniFile
    WriteIni = WritePrivateProfileString(iniSection, strKeyName, strValue, iniFile)
End Function

Public Function ReadIni(ByVal strKeyName As String, Optional ByVal strDefault As String, Optional iniSection As String = "Main", Optional TypeOfData As TypeOfData = -1, Optional iniFile As String = "Update.ini") As String
    Dim strData As String * 1024
    iniFile = App.Path & "\" & iniFile
    If GetPrivateProfileString(iniSection, strKeyName, strDefault, strData, Len(strData), iniFile) > 0 Then
        ReadIni = Left$(strData, InStr(strData$, Chr$(0)) - 1)
        Select Case TypeOfData
            Case dNumeric: If Not IsNumeric(ReadIni) Then ReadIni = strDefault
        End Select
    Else
        ReadIni = strDefault
    End If
End Function
