VERSION 5.00
Object = "{D1ECD258-D329-4388-AB83-DEC261A66B86}#1.0#0"; "UniControls_v2.0.ocx"
Begin VB.Form frmMain 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Perfect Antivirus 2009"
   ClientHeight    =   5625
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8115
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5625
   ScaleWidth      =   8115
   StartUpPosition =   2  'CenterScreen
   Begin UniControls.UniLabel UniLabel1 
      Height          =   615
      Left            =   4080
      Top             =   480
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   1085
      Caption         =   "Nha61n va2o nu1t ""Ca2i d9a85t Perfect Antivirus 2009"" d9e63 ba81t d9a62u ca2i d9a85t."
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16777215
      Link            =   ""
   End
   Begin UniControls.UniButton cmdSetup 
      Height          =   495
      Left            =   3840
      TabIndex        =   0
      Top             =   1320
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   873
      Icon            =   "frmMain.frx":57E4
      Style           =   2
      Caption         =   "Ca2i d9a85t Perfect Antivirus 2009"
      IconAlign       =   3
      iNonThemeStyle  =   2
      MaskColor       =   16711935
      FontColor       =   255
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      RoundedBordersByTheme=   0   'False
      ShowFocusRectangle=   0   'False
   End
   Begin UniControls.UniButton cmdReadMe 
      Height          =   495
      Left            =   3840
      TabIndex        =   1
      Top             =   1920
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   873
      Icon            =   "frmMain.frx":5800
      Style           =   2
      Caption         =   "Xem gio71i thie65u"
      IconAlign       =   3
      iNonThemeStyle  =   2
      MaskColor       =   16711935
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      RoundedBordersByTheme=   0   'False
      ShowFocusRectangle=   0   'False
   End
   Begin UniControls.UniButton cmdSource 
      Height          =   495
      Left            =   3840
      TabIndex        =   2
      Top             =   2520
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   873
      Icon            =   "frmMain.frx":581C
      Style           =   2
      Caption         =   "Xem ma4 nguo62n cu3a chu7o7ng tri2nh"
      IconAlign       =   3
      iNonThemeStyle  =   2
      MaskColor       =   16711935
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      RoundedBordersByTheme=   0   'False
      ShowFocusRectangle=   0   'False
   End
   Begin UniControls.UniButton cmdData 
      Height          =   495
      Left            =   3840
      TabIndex        =   3
      Top             =   3120
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   873
      Icon            =   "frmMain.frx":5838
      Style           =   2
      Caption         =   "Bo65 CSDL 700.000 ma4 nha65n da5ng Virus"
      IconAlign       =   3
      iNonThemeStyle  =   2
      MaskColor       =   16711935
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      RoundedBordersByTheme=   0   'False
      ShowFocusRectangle=   0   'False
   End
   Begin UniControls.UniButton cmdVB6 
      Height          =   495
      Left            =   3840
      TabIndex        =   4
      Top             =   3720
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   873
      Icon            =   "frmMain.frx":5854
      Style           =   2
      Caption         =   "Ca2i d9a85t ngo6n ngu74 Visual Basic 6.0"
      IconAlign       =   3
      iNonThemeStyle  =   2
      MaskColor       =   16711935
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      RoundedBordersByTheme=   0   'False
      ShowFocusRectangle=   0   'False
   End
   Begin UniControls.UniButton UniButton1 
      Height          =   495
      Left            =   3840
      TabIndex        =   5
      Top             =   4320
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   873
      Icon            =   "frmMain.frx":5870
      Style           =   2
      Caption         =   "Ca2i d9a85t co65ng cu5 gia3i ne1n WinRAR"
      IconAlign       =   3
      iNonThemeStyle  =   2
      MaskColor       =   16711935
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      RoundedBordersByTheme=   0   'False
      ShowFocusRectangle=   0   'False
   End
   Begin VB.Image Image1 
      Height          =   5625
      Left            =   0
      Picture         =   "frmMain.frx":588C
      Top             =   0
      Width           =   3750
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdData_Click()
On Error Resume Next
Shell "explorer " & AppPath & "BoCoSoDuLieu", vbNormalFocus
End Sub

Private Sub cmdReadMe_Click()
On Error Resume Next
Shell ChrW(34) & "C:\Program Files\Microsoft Office\OFFICE11\WINWORD.EXE" & ChrW(34) & " " & ChrW(34) & AppPath & "GioiThieu\ReadMe.doc" & ChrW(34), vbNormalFocus
End Sub

Private Sub cmdSetup_Click()
On Error Resume Next
Shell AppPath & "PAVSetup.exe", vbNormalFocus
End Sub

Private Sub cmdSource_Click()
On Error Resume Next
Shell "explorer " & AppPath & "MaNguon", vbNormalFocus
End Sub

Private Sub cmdVB6_Click()
On Error Resume Next
Shell "explorer " & AppPath & "NgonNguLapTrinh", vbNormalFocus
End Sub

Private Sub UniButton1_Click()
On Error Resume Next
Shell "explorer " & AppPath & "CongCuGiaiNen", vbNormalFocus
End Sub
