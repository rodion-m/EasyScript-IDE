Attribute VB_Name = "modWait"
Option Explicit

Public Declare Function GetTickCount Lib "kernel32.dll" () As Long
Private Declare Function QueryPerformanceCounter Lib "kernel32.dll" (ByRef lpPerformanceCount As Currency) As Long
Private Declare Function QueryPerformanceFrequency Lib "kernel32.dll" (ByRef lpFrequency As Currency) As Long
Dim tCounter As Currency, tFrequency As Currency

Function Wait(ByVal lngInterval As Long) As Boolean
    On Error Resume Next
    Dim lngEndTime As Currency, lngTime As Long
    If lngInterval < 1 Then Wait = True: Exit Function
    lngTime = GetPerformanceTime
    lngEndTime = lngTime + lngInterval
    Do While (GetPerformanceTime < lngEndTime)
        'тело ожидания
        DoEvents
    Loop
    Wait = True
End Function

Function GetPerformanceTime() As Long
    If tFrequency = 0 Then QueryPerformanceFrequency tFrequency
    QueryPerformanceCounter tCounter
    GetPerformanceTime = CLng(tCounter / (tFrequency / 1000))
End Function
