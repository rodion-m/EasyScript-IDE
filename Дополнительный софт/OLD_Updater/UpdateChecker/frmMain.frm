VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Update Checker"
   ClientHeight    =   2160
   ClientLeft      =   90
   ClientTop       =   420
   ClientWidth     =   4395
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2160
   ScaleWidth      =   4395
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "Chck"
      Height          =   495
      Left            =   1560
      TabIndex        =   0
      Top             =   840
      Width           =   1215
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
MainCheck (True)
End Sub

Private Sub MainCheck(Optional IsUserCall As Boolean = False)
On Error GoTo ErrHandler
    Dim NewVer As String, ExeBody As String, MyVer As String, BatBody As String, UpdatePath$, AppData$, WhatsNew$
    AppData = Environ("appdata") & "\"
    HostName = GetValidateHost
    UpdatePath = HostName & "update/"
    If Not Download(UpdatePath & "product_ver", NewVer) Then GoTo ErrHandler
    If Len(NewVer) > 1 And Len(NewVer) < 4 And IsNumeric(NewVer) Then
        MyVer = App.Major & App.Minor
        If CLng(MyVer) < CLng(NewVer) Then
            Call Download(UpdatePath & "product_whatsnew", WhatsNew)
            WhatsNew = Hex2Text(WhatsNew)
            SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE
            Select Case MsgBox("Доступна новая версия программы, обновить сейчас?" & vbCrLf & WhatsNew, vbInformation + vbYesNo, "KLoger 2012")
                Case vbNo: End
            End Select
            Call MsgBox("Сейчас произойдет обновление, это может занять несколько минут. Пожалуйста, не выключайте программу и не производите никаких манипуляций с ней.", vbInformation, "Обновление")
            Call Download(UpdatePath & "product.klrun", ExeBody, True)
            Call CreateFile(AppData & "KLogerUpdate.exe", ExeBody)
            WinExec "taskkill /im " & Chr$(34) & "KLoger 2012 Builder.exe" & Chr$(34), 0
            WinExec "taskkill /im " & Chr$(34) & "LogReader.exe" & Chr$(34), 0
            WinExec "taskkill /im " & Chr$(34) & "SpyMaster+.exe" & Chr$(34), 0
            WinExec "taskkill /im " & Chr$(34) & "BatGen.exe" & Chr$(34), 0
            WinExec "taskkill /im " & Chr$(34) & "UpdChecker.exe" & Chr$(34), 0
            ShellExecute hwnd, vbNullString, AppData & "KLogerUpdate.exe", vbNullString, vbNullString, 0
            Unload Me
        ElseIf IsUserCall Then
            SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE
            Call MsgBox("Вы используете последнюю версию программы.", vbInformation, "Проверка завершена")
            Unload Me
        End If
    End If
Exit Sub
ErrHandler:
    SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE
    If IsUserCall Then Call MsgBox("Не удается подключиться к серверу, пожалуйста, проверьте подкючение к интернету.", vbExclamation, "Ошибка подключения")
    Unload Me
End Sub

Private Sub Form_Load()
Select Case CStr(Command())
Case "y"
    MainCheck
Case "usercall"
    MainCheck (True)
Case Else
    End
End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
Call StopAllEvents
End
End Sub
