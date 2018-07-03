Attribute VB_Name = "mdlWait"
Option Explicit

Private Declare Function QueryPerformanceCounter Lib "kernel32.dll" (ByRef lpPerformanceCount As Currency) As Long
Private Declare Function QueryPerformanceFrequency Lib "kernel32.dll" (ByRef lpFrequency As Currency) As Long
Dim tCounter As Currency, tFrequency As Currency

Function Wait(ByVal lngInterval As Long, Optional ByVal blnCallInScript As Boolean = True) As Boolean
    On Error Resume Next
    If lngInterval < 0 Then Wait = False: Exit Function
    If blnCallInScript Then
        lngInterval = RandomizeNumber(lngInterval, Settings.lngIntervalLimit) ' Рандомизация интервала
        If Settings.sngSpeed > 0 Then lngInterval = CLng(lngInterval / Settings.sngSpeed)
    End If
    If lngInterval < 1 Then Wait = True: Exit Function
    Select Case lngInterval
        Case Is <= 100: PerformanceWait lngInterval
        Case Is > 20000: MySleep lngInterval, 100, blnCallInScript ' ЕСЛИ более 20-ти секунд, то частота обновления 100 мс.
        Case Is > 100: MySleep lngInterval, 20, blnCallInScript
    End Select
    Wait = True
End Function

Sub MySleep(ByVal lngInterval As Long, Optional ByVal lFrequency As Long = 20, Optional ByVal blnCallInScript As Boolean)
    Dim lCounter As Long
    Do While lCounter < lngInterval
        Sleep lFrequency
        lCounter = lCounter + lFrequency
        DoEvents
        If blnCallInScript Then CheckHotKeys: If Not blnExecuting Then Exit Do
    Loop
End Sub

Private Sub PerformanceWait(ByVal lngInterval As Long)
    On Error Resume Next
    Dim lngEndTime As Currency, lngTime As Long
    lngTime = GetPerformanceTime
    lngEndTime = lngTime + lngInterval
    Do While (GetPerformanceTime < lngEndTime)
        DoEvents
    Loop
End Sub

Function GetPerformanceTime() As Long ' Возвращает милисекунды
    If tFrequency = 0 Then QueryPerformanceFrequency tFrequency
    QueryPerformanceCounter tCounter
    GetPerformanceTime = CLng(tCounter / (tFrequency / 1000))
End Function
