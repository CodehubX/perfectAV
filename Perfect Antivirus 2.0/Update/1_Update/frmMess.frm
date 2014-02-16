VERSION 5.00
Object = "{6DC1AB90-DCD1-47A1-AB36-924FFD67ADBF}#1.0#0"; "UniControls_v2.0.ocx"
Begin VB.Form frmMess 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Loi nhan tu trang chu"
   ClientHeight    =   3105
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5655
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMess.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3105
   ScaleWidth      =   5655
   StartUpPosition =   2  'CenterScreen
   Begin UniControls.UniLabel UniLabel2 
      Height          =   255
      Left            =   120
      Top             =   2880
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   450
      BackStyle       =   0
      Caption         =   "Tin nha81n d9u7o75c gu73i tu72: dinhquangtrung90@yahoo.com"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin UniControls.UniTextBox txtMess 
      Height          =   2175
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   3836
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   0
      Text            =   ""
      MultiLine       =   -1  'True
      Locked          =   -1  'True
      BorderStyle     =   2
      Scrollbar       =   2
   End
   Begin UniControls.UniLabel UniLabel1 
      Height          =   495
      Left            =   600
      Top             =   120
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   873
      Alignment       =   1
      Caption         =   "Lo72i nha81n tu72 trang chu3"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmMess"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
frmMain.Timer1.Enabled = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
If frmMain.FileExists(AppPath & "UpdateInfo.txt") = True Then
Kill AppPath & "UpdateInfo.txt"
Kill AppPath & "Mess.txt"
End If

Unload frmMain

End Sub
