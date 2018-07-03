Attribute VB_Name = "mdlRegistry"
Option Explicit

Private Declare Function RegOpenKey Lib "advapi32" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long
Private Declare Function RegDeleteValue Lib "advapi32" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
Private Declare Function RegSetValueEx Lib "advapi32" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpData As String, ByVal cbData As Long) As Long
Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" _
                (ByVal hKey As Long, ByVal lpValueName As String, _
                 ByVal lpReserved As Long, lpType As Long, ByVal lpData As String, lpcbData As Long) As Long


Enum REGISTRY_KEYS
    HKEY_CLASSES_ROOT = &H80000000
    HKEY_CURRENT_USER = &H80000001
    HKEY_LOCAL_MACHINE = &H80000002
    HKEY_USERS = &H80000003
End Enum

Enum REGISTRY_TYPES
    REG_BINARY = 3&
    REG_DWORD = 4&
    REG_EXPAND_SZ = 2&
    REG_LINK = 6&
    REG_MULTI_SZ = 7&
    REG_SZ = 1&
    REG_QWORD = 11&
End Enum

Private Const ERROR_SUCCESS As Long = 0&
Private Const KEY_ALL_ACCESS = &HF003F
Private Const AUTORUN_KEY = "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Run"

Dim hKey As Long

Function RegAdd(ByVal strKey, ByVal strValueName, ByVal strValue, Optional dwType As REGISTRY_TYPES = REG_SZ) As Boolean
    OpenRegistry strKey
    If RegSetValueEx(hKey, strValueName, 0&, REG_SZ, strValue, Len(strValue)) = ERROR_SUCCESS Then
        RegAdd = True
    End If
    CloseRegistry
End Function

Function RegDel(ByVal strKey, ByVal strValueName) As Boolean
    If RegGet(strKey, strValueName) = False Then RegDel = True: Exit Function
    If RegDel Then Exit Function
    OpenRegistry strKey
    If RegDeleteValue(hKey, strValueName) = ERROR_SUCCESS Then
        RegDel = True
    End If
    CloseRegistry
End Function

Function RegGet(ByVal strKey As String, ByVal strValueName As String)
    Dim lngValueLen As Long, strData As String
    strData = String$(256, Chr$(0))
    lngValueLen = 256
    OpenRegistry strKey
    If RegQueryValueEx(hKey, strValueName, 0&, 0&, strData, lngValueLen) <> ERROR_SUCCESS Then RegGet = False: Exit Function
    If lngValueLen > 1 And Left$(strData, 1) <> Chr$(0) Then
        RegGet = Left(strData, lngValueLen - 1)
    Else
        If lngValueLen = 2 Then RegGet = "" Else RegGet = False
    End If
    CloseRegistry
End Function

Function OpenRegistry(ByVal strKey As String, Optional ByRef strMainKey As String, Optional ByRef strSubKey As String) As Boolean
    strMainKey = Left$(strKey, InStr(1, strKey, "\") - 1)
    strSubKey = Right$(strKey, Len(strKey) - Len(strMainKey) - 1)
    If RegOpenKey(Get_hKey(strMainKey), strSubKey, hKey) = ERROR_SUCCESS Then
        OpenRegistry = True
    End If
End Function

Function Get_hKey(ByVal strMainKey As String) As REGISTRY_KEYS
    Select Case strMainKey
        Case "HKEY_CLASSES_ROOT", "HKCR":       Get_hKey = HKEY_CLASSES_ROOT
        Case "HKEY_LOCAL_MACHINE", "HKLM":   Get_hKey = HKEY_LOCAL_MACHINE
        Case "HKEY_CURRENT_USER", "HKCU":      Get_hKey = HKEY_CURRENT_USER
        Case "HKEY_USERS", "HKU":                           Get_hKey = HKEY_USERS
    End Select
End Function

Sub CloseRegistry()
    If hKey <> 0 Then RegCloseKey hKey
End Sub

Function Autorun_Add(Optional ByVal Title As String, Optional ByVal FileName As String) As Boolean
    If Title = "" Then Title = App.ProductName
    If FileName = "" Then FileName = App.Path & "\" & App.EXEName & ".exe /autorun"
    Autorun_Add = RegAdd(AUTORUN_KEY, Title, FileName)
End Function

Function Autorun_Delete(Optional ByVal Title As String) As Boolean
    If Title = "" Then Title = App.ProductName
    Autorun_Delete = RegDel(AUTORUN_KEY, Title)
End Function

Function Autorun_GetStatus(Optional ByVal Title As String, Optional ByVal FileName As String) As Boolean
    Dim RegVal As String
    If Title = "" Then Title = App.ProductName
    If FileName = "" Then FileName = App.Path & "\" & App.EXEName & ".exe /autorun"
    RegVal = RegGet(AUTORUN_KEY, Title)
    If FileName <> "" Then
        Autorun_GetStatus = (RegVal = FileName)
    ElseIf RegVal <> "" Then
        Autorun_GetStatus = True
    End If
End Function
