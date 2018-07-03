VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cBalloonSysTray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
   
Private Declare Function Shell_NotifyIconW Lib "shell32.dll" _
   (ByVal dwMessage As Long, lpData As NOTIFYICONDATAW) As Long
Private Declare Function GetForegroundWindow Lib "user32" () As Long


Private Const NIF_ICON = &H2
Private Const NIF_MESSAGE = &H1
Private Const NIF_TIP = &H4
Private Const NIF_STATE = &H8
Private Const NIF_INFO = &H10

Private Const NIM_ADD = &H0
Private Const NIM_MODIFY = &H1
Private Const NIM_DELETE = &H2
Private Const NIM_SETFOCUS = &H3
Private Const NIM_SETVERSION = &H4

Private Const NOTIFYICON_VERSION = 3

Private Type NOTIFYICONDATAW
   cbSize As Long             ' 4
   hwnd As Long               ' 8
   uID As Long                ' 12
   uFlags As Long             ' 16
   uCallbackMessage As Long   ' 20
   hIcon As Long              ' 24
   szTip(0 To 255) As Byte    ' 280
   dwState As Long            ' 284
   dwStateMask As Long        ' 288
   szInfo(0 To 511) As Byte   ' 800
   uTimeOutOrVersion As Long  ' 804
   szInfoTitle(0 To 127) As Byte ' 932
   dwInfoFlags As Long        ' 936
   guidItem As Long           ' 940
End Type

Private nfIconDataW As NOTIFYICONDATAW
Private Const NOTIFYICONDATAA_V2_SIZE_U = 936

Private Const WM_MOUSEMOVE = &H200
Private Const WM_LBUTTONDBLCLK = &H203
Private Const WM_LBUTTONDOWN = &H201
Private Const WM_LBUTTONUP = &H202
Private Const WM_RBUTTONDBLCLK = &H206
Private Const WM_RBUTTONDOWN = &H204
Private Const WM_RBUTTONUP = &H205
Private Const WM_MOUSELEAVE = &H2A3

Private Const WM_USER = &H400

Private Const NIN_SELECT = WM_USER
Private Const NINF_KEY = &H1
Private Const NIN_KEYSELECT = (NIN_SELECT Or NINF_KEY)
Private Const NIN_BALLOONSHOW = (WM_USER + 2)
Private Const NIN_BALLOONHIDE = (WM_USER + 3)
Private Const NIN_BALLOONTIMEOUT = (WM_USER + 4)
Private Const NIN_BALLOONUSERCLICK = (WM_USER + 5)

Public Event SysTrayMouseDown(ByVal eButton As MouseButtonConstants)
Public Event SysTrayMouseUp(ByVal eButton As MouseButtonConstants)
Public Event SysTrayMouseMove()
Public Event SysTrayDoubleClick(ByVal eButton As MouseButtonConstants)
Public Event BalloonShow()
Public Event BalloonHide()
Public Event BalloonTimeOut()
Public Event BalloonClicked()

Public Enum EBalloonIconTypes
   NIIF_NONE = 0
   NIIF_INFO = 1
   NIIF_WARNING = 2
   NIIF_ERROR = 3
   NIIF_NOSOUND = &H10
End Enum

Private m_Owner As Form
Private m_hWnd As Long
Private m_hIcon As Long

Private m_ToolTip As String

Private m_bAddedMenuItem As Boolean

Private m_bUseUnicode As Boolean
Private m_bSupportsNewVersion As Boolean

Private m_MenuObject As Object
Private m_MenuDefault As Object

Private blnIconAdded As Boolean

Public Sub ShowTrayIcon()
    With nfIconDataW
       .hwnd = m_hWnd
       .uID = 0&
       .uFlags = NIF_ICON Or NIF_MESSAGE Or NIF_TIP
       .uCallbackMessage = WM_MOUSEMOVE
       .hIcon = m_hIcon
       stringToArray m_ToolTip, .szTip, 256
      .uTimeOutOrVersion = NOTIFYICON_VERSION
       .cbSize = nfStructureSize
    End With
    Shell_NotifyIconW NIM_ADD, nfIconDataW
    Shell_NotifyIconW NIM_SETVERSION, nfIconDataW
    blnIconAdded = True
End Sub

Public Sub HideTrayIcon()
    Class_Terminate
End Sub

Public Sub ShowBalloonTip(ByVal sMessage As String, Optional ByVal sTitle As String, Optional ByVal eIcon As EBalloonIconTypes)
    If Not blnIconAdded Then Exit Sub
    stringToArray sMessage, nfIconDataW.szInfo, 512
    stringToArray sTitle, nfIconDataW.szInfoTitle, 128
    nfIconDataW.dwInfoFlags = eIcon
    nfIconDataW.uFlags = NIF_INFO
    Shell_NotifyIconW NIM_MODIFY, nfIconDataW
End Sub

Public Function ShowMenu()
   If Not (m_MenuDefault Is Nothing) Then
      m_Owner.PopupMenu m_MenuObject, , , , m_MenuDefault
   Else
      m_Owner.PopupMenu m_MenuObject
   End If
End Function

Private Sub stringToArray(ByVal sString As String, bArray() As Byte, ByVal lMaxSize As Long)
Dim b() As Byte
Dim i As Long
Dim j As Long
   If Len(sString) > 0 Then
      b = sString
      For i = LBound(b) To UBound(b)
         bArray(i) = b(i)
         If (i = (lMaxSize - 2)) Then
            Exit For
         End If
      Next i
      For j = i To lMaxSize - 1
         bArray(j) = 0
      Next j
   End If
End Sub

Private Property Get nfStructureSize() As Long
    nfStructureSize = NOTIFYICONDATAA_V2_SIZE_U
End Property

Public Sub CallEvents(X As Single)
    If (m_Owner Is Nothing) Then Exit Sub
     Dim uMsg As Long
     If GetForegroundWindow = m_Owner.hwnd Then Exit Sub
    uMsg = m_Owner.ScaleX(X, m_Owner.ScaleMode, vbPixels)
    Select Case uMsg
    Case WM_MOUSEMOVE
       RaiseEvent SysTrayMouseMove
    Case WM_LBUTTONDOWN
       RaiseEvent SysTrayMouseDown(vbLeftButton)
    Case WM_LBUTTONUP
       RaiseEvent SysTrayMouseUp(vbLeftButton)
    Case WM_LBUTTONDBLCLK
       RaiseEvent SysTrayDoubleClick(vbLeftButton)
    Case WM_RBUTTONDOWN
       RaiseEvent SysTrayMouseDown(vbRightButton)
    Case WM_RBUTTONUP
       RaiseEvent SysTrayMouseUp(vbRightButton)
    Case WM_RBUTTONDBLCLK
       RaiseEvent SysTrayDoubleClick(vbRightButton)
    Case NIN_BALLOONSHOW
       RaiseEvent BalloonShow
    Case NIN_BALLOONHIDE
       RaiseEvent BalloonHide
    Case NIN_BALLOONTIMEOUT
       RaiseEvent BalloonTimeOut
    Case NIN_BALLOONUSERCLICK
       RaiseEvent BalloonClicked
    End Select
End Sub

Public Property Get OwnerForm() As Form
    Set OwnerForm = m_Owner
End Property

Public Property Let OwnerForm(ByVal fmOwner As Form)
    Set m_Owner = fmOwner
    m_hWnd = fmOwner.hwnd
End Property

Public Property Get IconHandle() As Long
    IconHandle = nfIconDataW.hIcon
End Property

Public Property Let IconHandle(ByVal hIcon As Long)
    If (hIcon <> nfIconDataW.hIcon) And blnIconAdded Then
       m_hIcon = hIcon
       nfIconDataW.hIcon = hIcon
       nfIconDataW.uFlags = NIF_ICON
       Shell_NotifyIconW NIM_MODIFY, nfIconDataW
    End If
End Property

Public Property Get ToolTip() As String
    ToolTip = m_ToolTip
End Property

Public Property Let ToolTip(ByVal sTip As String)
    If Not blnIconAdded Then Exit Property
    m_ToolTip = sTip
    stringToArray sTip, nfIconDataW.szTip, 256
    nfIconDataW.uFlags = NIF_TIP
    Shell_NotifyIconW NIM_MODIFY, nfIconDataW
End Property

Public Property Let DefaultMenuObject(ByVal objMenuDef As Object)
   Set m_MenuDefault = objMenuDef
End Property

Public Property Let OwnerMenuObject(ByVal objMenuObject As Object)
   Set m_MenuObject = objMenuObject
End Property

Private Sub Class_Terminate()
      Shell_NotifyIconW NIM_DELETE, nfIconDataW
End Sub
