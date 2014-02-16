VERSION 5.00
Object = "{E8FDD05C-3067-4198-8AEC-1A013A46ABDD}#1.0#0"; "FVUnicodeControl.ocx"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Object = "{6DC1AB90-DCD1-47A1-AB36-924FFD67ADBF}#1.0#0"; "UniControls_v2.0.ocx"
Begin VB.Form frmMain 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Thac mac - gop y ve chuong trinh PAV2009"
   ClientHeight    =   7650
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8910
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7650
   ScaleWidth      =   8910
   StartUpPosition =   2  'CenterScreen
   Begin UniControls.UniLabel LBL1 
      Height          =   615
      Left            =   2520
      Top             =   4440
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   1085
      Alignment       =   1
      Caption         =   "D9ang ta3i ho65p thu7 go1p y1"
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
   Begin SHDocVwCtl.WebBrowser Web1 
      Height          =   4815
      Left            =   120
      TabIndex        =   6
      Top             =   2760
      Width           =   8655
      ExtentX         =   15266
      ExtentY         =   8493
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin FVUnicodeControl.FVistaUniButton cmdSend 
      Height          =   375
      Left            =   6600
      TabIndex        =   5
      Top             =   1080
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      BackColor       =   14737632
      ButtonStyle     =   3
      Caption         =   "D9a1nh gia1"
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
   Begin FVUnicodeControl.FVistaUniLabel Label1 
      Height          =   375
      Left            =   480
      TabIndex        =   4
      Top             =   1920
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   661
      Alignment       =   1
      BackStyle       =   0
      Caption         =   "50"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   8454016
   End
   Begin FVUnicodeControl.FVistaSlider barDanhGia 
      Height          =   270
      Left            =   480
      TabIndex        =   3
      Top             =   1560
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   476
      BackColor       =   16777215
      Value           =   50
   End
   Begin FVUnicodeControl.FVistaUniLabel FVistaUniLabel3 
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   1200
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   450
      Caption         =   "Ve62 cha61t lu7o75ng, ba5n d9a1nh gia1 PAV 2009 phie6n ba3n II d9u7o75c bao nhie6u d9ie63m?"
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
   Begin FVUnicodeControl.FVistaUniLabel FVistaUniLabel2 
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   661
      Alignment       =   1
      Caption         =   "Mo5i y1 kie61n cu3a ba5n se4 la2m co7 so7 d9e63 chu7o7ng tri2nh nga2y ca2ng pha1t trie63n ho7n!"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16711680
   End
   Begin FVUnicodeControl.FVistaUniLabel FVistaUniLabel1 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   661
      Alignment       =   1
      Caption         =   "Ha4y gu73i nhu74ng tha81c ma81c, go1p y1 cu3a ba5n ve62 chu7o7ng tri2nh d9e61n ta1c gia3"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16711680
   End
   Begin VB.Image Image1 
      Height          =   360
      Left            =   480
      Picture         =   "frmMain.frx":628A
      Stretch         =   -1  'True
      Top             =   1920
      Width           =   8100
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub barDanhGia_ValueChanged()
Label1.Caption = barDanhGia.Value
End Sub

Private Sub cmdSend_Click()
cmdSend.Enabled = False
Me.barDanhGia.Enabled = False

SendMail "dinhquangtrung90@yahoo.com", barDanhGia.Value
cmdSend.Enabled = True
Me.barDanhGia.Enabled = True

UniMsgBox "D9a4 gu73i d9e61n ta1c gia3! Ra61t ca3m o7n ba5n d9a4 go1p y1 cho chu7o7ng tri2nh!", vbOKOnly, "OK!", Me.hWnd
End Sub

Private Sub Form_Load()
Web1.Navigate "http://www2.shoutmix.com/?qtsoft"
End Sub

Private Sub Web1_DocumentComplete(ByVal pDisp As Object, URL As Variant)
LBL1.Visible = False
End Sub
