VERSION 5.00
Object = "{E8FDD05C-3067-4198-8AEC-1A013A46ABDD}#1.0#0"; "FVUnicodeControl.ocx"
Object = "{6DC1AB90-DCD1-47A1-AB36-924FFD67ADBF}#1.0#0"; "UniControls_v2.0.ocx"
Begin VB.Form frmTangToc 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PAV 2009 - Tang Toc may Tinh"
   ClientHeight    =   4815
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7320
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmTangToc.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4815
   ScaleWidth      =   7320
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin FVUnicodeControl.FVistaUniButton cmdOK 
      Height          =   375
      Left            =   6120
      TabIndex        =   13
      Top             =   4320
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      BackColor       =   14737632
      ButtonStyle     =   3
      Caption         =   "OK!"
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
   Begin UniControls.UniLabel UniLabel2 
      Height          =   255
      Left            =   0
      Top             =   600
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   450
      Alignment       =   1
      Caption         =   "Ta61t ca3 ca1c thao ta1c d9u7o71i d9a6y d9e62u co1 the63 giu1p ma1y ti1nh cu3a ba5n cha5y nhanh ho7n"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16711680
   End
   Begin FVUnicodeControl.FVistaUniButton cmdSaveSetting 
      Height          =   375
      Left            =   240
      TabIndex        =   12
      Top             =   4200
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   661
      BackColor       =   14737632
      ButtonStyle     =   3
      Caption         =   "Thie61t la65p"
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
   Begin FVUnicodeControl.FVistaUniButton cmdDelTemp 
      Height          =   375
      Left            =   5040
      TabIndex        =   11
      Top             =   1080
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   661
      BackColor       =   14737632
      ButtonStyle     =   3
      Caption         =   "Xo1a ca1c file thu72a"
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
   Begin FVUnicodeControl.FVistaUniCheckbox CHKTANGTOC 
      Height          =   195
      Index           =   10
      Left            =   240
      TabIndex        =   10
      Top             =   3840
      Width           =   2205
      _ExtentX        =   3889
      _ExtentY        =   344
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Value           =   -1  'True
      Caption         =   "Ta8ng to61c truy ca65p d9i4a me62m"
      ForeColor       =   0
      ShowFocusRectangle=   0   'False
   End
   Begin FVUnicodeControl.FVistaUniCheckbox CHKTANGTOC 
      Height          =   195
      Index           =   9
      Left            =   240
      TabIndex        =   9
      Top             =   3600
      Width           =   4680
      _ExtentX        =   8255
      _ExtentY        =   344
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Value           =   -1  'True
      Caption         =   "Ta8ng to61c hoa5t d9o65ng cu3a o63 d9i4a quang (CDROM-DVD-CD RW…)"
      ForeColor       =   0
      ShowFocusRectangle=   0   'False
   End
   Begin FVUnicodeControl.FVistaUniCheckbox CHKTANGTOC 
      Height          =   195
      Index           =   8
      Left            =   240
      TabIndex        =   8
      Top             =   3360
      Width           =   3435
      _ExtentX        =   6059
      _ExtentY        =   344
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Value           =   -1  'True
      Caption         =   "Ta8ng to61c kho73i d9o65ng tu72 5-8 gia6y tu2y ca61u hi2nh"
      ForeColor       =   0
      ShowFocusRectangle=   0   'False
   End
   Begin FVUnicodeControl.FVistaUniCheckbox CHKTANGTOC 
      Height          =   195
      Index           =   7
      Left            =   240
      TabIndex        =   7
      Top             =   3120
      Width           =   4320
      _ExtentX        =   7620
      _ExtentY        =   344
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Value           =   -1  'True
      Caption         =   "Ta8ng to61c ta81t ma1y xuo61ng co2n 5-15 gia6y tu2y ca61u hi2nh ma1y"
      ForeColor       =   0
      ShowFocusRectangle=   0   'False
   End
   Begin FVUnicodeControl.FVistaUniCheckbox CHKTANGTOC 
      Height          =   195
      Index           =   6
      Left            =   240
      TabIndex        =   6
      Top             =   2880
      Width           =   2070
      _ExtentX        =   3651
      _ExtentY        =   344
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Value           =   -1  'True
      Caption         =   "Ta8ng to61c na5p sa84n du74 lie65u"
      ForeColor       =   0
      ShowFocusRectangle=   0   'False
   End
   Begin FVUnicodeControl.FVistaUniCheckbox CHKTANGTOC 
      Height          =   195
      Index           =   5
      Left            =   240
      TabIndex        =   5
      Top             =   2640
      Width           =   2805
      _ExtentX        =   4948
      _ExtentY        =   344
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Value           =   -1  'True
      Caption         =   "Ta8ng su71c hoa5t d9o65ng cu3a bo65 nho71 a3o"
      ForeColor       =   0
      ShowFocusRectangle=   0   'False
   End
   Begin FVUnicodeControl.FVistaUniCheckbox CHKTANGTOC 
      Height          =   195
      Index           =   4
      Left            =   240
      TabIndex        =   4
      Top             =   2400
      Width           =   3840
      _ExtentX        =   6773
      _ExtentY        =   344
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Value           =   -1  'True
      Caption         =   "Thie61t la65p ca1c tho6ng so61 ta8ng to61c khi truy ca65p ma5ng"
      ForeColor       =   0
      ShowFocusRectangle=   0   'False
   End
   Begin FVUnicodeControl.FVistaUniCheckbox CHKTANGTOC 
      Height          =   555
      Index           =   3
      Left            =   240
      TabIndex        =   3
      Top             =   1800
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   979
      AutoSize        =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Value           =   -1  'True
      Caption         =   "Thie61t la65p cho Windows go74 bo3 hoa2n toa2n ca1c DLL ra kho3i bo65 nho71 khi thoa1t chu7o7ng tri2nh lie6n quan"
      ForeColor       =   0
      ShowFocusRectangle=   0   'False
   End
   Begin FVUnicodeControl.FVistaUniCheckbox CHKTANGTOC 
      Height          =   195
      Index           =   2
      Left            =   240
      TabIndex        =   2
      Top             =   1560
      Width           =   4320
      _ExtentX        =   7620
      _ExtentY        =   344
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Value           =   -1  'True
      Caption         =   "Ta81t che61 d9o65 ca65p nha65t tho72i gian truy ca65p File cu4a Windows"
      ForeColor       =   0
      ShowFocusRectangle=   0   'False
   End
   Begin FVUnicodeControl.FVistaUniCheckbox CHKTANGTOC 
      Height          =   195
      Index           =   1
      Left            =   240
      TabIndex        =   1
      Top             =   1320
      Width           =   3765
      _ExtentX        =   6641
      _ExtentY        =   344
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Value           =   -1  'True
      Caption         =   "Kho6ng na5p ca1c thu7 vie65n he65 tho61ng va2o Bo65 nho71 a3o"
      ForeColor       =   0
      ShowFocusRectangle=   0   'False
   End
   Begin FVUnicodeControl.FVistaUniCheckbox CHKTANGTOC 
      Height          =   195
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   1080
      Width           =   2310
      _ExtentX        =   4075
      _ExtentY        =   344
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Value           =   -1  'True
      Caption         =   "Ta8ng to61c d9o65 truy xua61t Menu"
      ForeColor       =   0
      ShowFocusRectangle=   0   'False
   End
   Begin UniControls.UniLabel UniLabel1 
      Height          =   375
      Left            =   480
      Top             =   120
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   661
      Alignment       =   1
      Caption         =   "To63ng ho75p ca1c thu3 thua65t ta8ng to61c ma1y ti1nh"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmTangToc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdDelRecycle_Click()
EmpRecBin
UniMsgBox "D9a4 la2m sa5ch thu2ng ra1c!", vbOKOnly, "OK!", Me.hwnd
End Sub

Private Sub cmdDelTemp_Click()
ClearJunkFile
UniMsgBox "D9a4 xo1a xong!", vbOKOnly, "OK!", Me.hwnd
End Sub

Private Sub cmdOK_Click()
Unload Me
End Sub

Private Sub cmdSaveSetting_Click()
Dim ocxDir$
Dim bytResourceData() As Byte

tXoaFile AppPath & "TangToc1.REG"
ocxDir = AppPath & "TangToc1.REG"
bytResourceData = LoadResData(101, "TANGTOC")
Open ocxDir For Binary Shared As #1
Put #1, 1, bytResourceData
Close #1
Shell "REG IMPORT " & AppPath & "TangToc1.REG", vbHide

tXoaFile AppPath & "TangToc2.REG"
ocxDir = AppPath & "TangToc2.REG"
bytResourceData = LoadResData(102, "TANGTOC")
Open ocxDir For Binary Shared As #1
Put #1, 1, bytResourceData
Close #1
Shell "REG IMPORT " & AppPath & "TangToc2.REG", vbHide

UniMsgBox "OK! Ma1y ti1nh cu3a ba5n d9a4 d9u7o75c ta8ng to61c le6n ra61t nhie62u!", vbOKOnly, "OK!", Me.hwnd
tXoaFile AppPath & "TangToc2.REG"
tXoaFile AppPath & "TangToc1.REG"
End Sub

