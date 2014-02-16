VERSION 5.00
Object = "{E8FDD05C-3067-4198-8AEC-1A013A46ABDD}#1.0#0"; "FVUnicodeControl.ocx"
Object = "{6DC1AB90-DCD1-47A1-AB36-924FFD67ADBF}#1.0#0"; "UniControls_v2.0.ocx"
Begin VB.Form frmScan 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Scan Virus"
   ClientHeight    =   7545
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10650
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmScan.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "frmScan.frx":058A
   ScaleHeight     =   7545
   ScaleWidth      =   10650
   StartUpPosition =   2  'CenterScreen
   Begin FVUnicodeControl.FVistaUniCheckbox OffCom 
      Height          =   195
      Left            =   120
      TabIndex        =   28
      Top             =   3600
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
      Caption         =   "Ta81t ma1y sau khi que1t xong"
      BackStyle       =   0
      ForeColor       =   0
   End
   Begin VB.Timer Timer6 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1080
      Top             =   3120
   End
   Begin FVUnicodeControl.FVistaUniButton cmdHide 
      Height          =   495
      Left            =   120
      TabIndex        =   27
      Top             =   3960
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   873
      BackColor       =   -2147483633
      ButtonStyle     =   3
      Caption         =   "A63n xuo61ng Traybar"
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
   Begin UniControls.UniLabel LBL 
      Height          =   255
      Left            =   6000
      Top             =   4560
      Width           =   10335
      _ExtentX        =   18230
      _ExtentY        =   450
      BackStyle       =   0
      Caption         =   "Nha61n nu1t Start d9e63 que1t."
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16776960
   End
   Begin UniControls.UniLabel lblChuanBi3 
      Height          =   255
      Left            =   7680
      Top             =   6960
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   450
      BackStyle       =   0
      Caption         =   "UniLabel2"
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
   Begin UniControls.UniLabel lblChuanBi4 
      Height          =   615
      Left            =   8640
      Top             =   6720
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1085
      BackStyle       =   0
      Caption         =   $"frmScan.frx":1024C
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   65535
   End
   Begin UniControls.UniLabel lblChuanBi2 
      Height          =   255
      Left            =   7920
      Top             =   6720
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   450
      BackStyle       =   0
      Caption         =   "Xin ha4y kie63m tra la5i nhu74ng tho6ng tin ma2 ba5n d9a4 ca2i d9a85t 1 la62n nu74a."
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
   Begin UniControls.UniLabel lblChuanBi1 
      Height          =   375
      Left            =   7440
      Top             =   6360
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      BackStyle       =   0
      Caption         =   "Chua63n bi5 que1t"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   255
   End
   Begin VB.Timer Timer5 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   3600
      Top             =   5400
   End
   Begin UniControls.UniLabel lblPhatHien 
      Height          =   255
      Left            =   3600
      Top             =   5520
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   450
      BackStyle       =   0
      Caption         =   "Pha1t Hie65n:"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   12648384
   End
   Begin UniControls.UniLabel lbl_Da_Quet 
      Height          =   255
      Left            =   3600
      Top             =   5280
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   450
      BackStyle       =   0
      Caption         =   "D9a4 que1t:"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   12648384
   End
   Begin UniControls.UniLabel lbl_Tong_So_File 
      Height          =   255
      Left            =   3600
      Top             =   5040
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   450
      BackStyle       =   0
      Caption         =   "To63ng so61 File:"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   12648384
   End
   Begin VB.Timer Timer4 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   6120
      Top             =   360
   End
   Begin FVUnicodeControl.FVistaUniProgressbar Proc 
      Height          =   225
      Left            =   3600
      Top             =   5040
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   397
      Max             =   100
      Value           =   0
      TStyle          =   3
      Min             =   0
      Style           =   2
      Text            =   "D9ang chua63n bi5..."
      Align           =   1
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   5
      Left            =   7680
      Top             =   0
   End
   Begin FVUnicodeControl.FVistaUniButton cmdBack 
      Height          =   615
      Left            =   2400
      TabIndex        =   21
      Top             =   5400
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1085
      BackColor       =   -2147483633
      ButtonStyle     =   3
      Caption         =   "Thie61t la65p la5i"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PicAlign        =   8
      PicNormal       =   "frmScan.frx":10318
      PicSizeH        =   16
      PicSizeW        =   16
   End
   Begin FVUnicodeControl.FVistaUniButton cmdCachLy 
      Height          =   375
      Left            =   4680
      TabIndex        =   20
      Top             =   4440
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      BackColor       =   16761024
      ButtonStyle     =   3
      Caption         =   "Ca1ch ly Virus"
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      State           =   3
   End
   Begin FVUnicodeControl.FVistaUniButton cmdDeleteVirus 
      Height          =   375
      Left            =   3480
      TabIndex        =   19
      Top             =   4440
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      BackColor       =   16761024
      ButtonStyle     =   3
      Caption         =   "Xo1a Virus"
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      State           =   3
   End
   Begin FVUnicodeControl.FVistaUniButton cmdStartStopScan 
      Height          =   855
      Left            =   2400
      TabIndex        =   18
      ToolTipText     =   "Start/Stop"
      Top             =   4440
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1508
      BackColor       =   -2147483633
      ButtonStyle     =   3
      Caption         =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PicAlign        =   0
      PicDown         =   "frmScan.frx":10D2A
      PicNormal       =   "frmScan.frx":1651C
      PicSizeH        =   48
      PicSizeW        =   48
   End
   Begin UniControls.UniLabel lblHayNhanVao 
      Height          =   255
      Left            =   2520
      Top             =   840
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   450
      BackStyle       =   0
      Caption         =   "Ha4y nha61n va2o nu1t mu4i te6n ma2u xanh d9e63 chu7o7ng tri2nh ba81t d9a62u que1t."
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
   Begin UniControls.UniListView LV1 
      Height          =   3135
      Left            =   2400
      TabIndex        =   17
      Top             =   1200
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   5530
      AutoUnicode     =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   192
      View            =   1
      LabelEdit       =   0   'False
      GridLines       =   -1  'True
      FullRowSelect   =   -1  'True
      AutoArrange     =   0   'False
      BorderStyle     =   2
      CheckBoxes      =   -1  'True
      HeaderButtons   =   0   'False
      HeaderTrackSelect=   0   'False
      HideSelection   =   0   'False
      InfoTips        =   0   'False
   End
   Begin UniControls.UniLabel lbl_HoangTat 
      Height          =   375
      Left            =   2400
      Top             =   480
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      BackStyle       =   0
      Caption         =   "D9a4 hoa2n ta61t!"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   255
   End
   Begin UniControls.UniLabel lblPro 
      Height          =   255
      Index           =   3
      Left            =   600
      Top             =   1920
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   450
      BackStyle       =   0
      Caption         =   "Ba81t d9a62u que1t"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   12632256
   End
   Begin UniControls.UniLabel lblPro 
      Height          =   255
      Index           =   2
      Left            =   600
      Top             =   1560
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   450
      BackStyle       =   0
      Caption         =   "Chua63n bi5 que1t"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   12632256
   End
   Begin FVUnicodeControl.FVistaUniLabel lbl_KhongCanThietLap 
      Height          =   255
      Left            =   3000
      TabIndex        =   16
      Top             =   6480
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   450
      BackStyle       =   0
      Caption         =   "Ne61u ba5n kho6ng ca62n chi3nh gi2 the6m, ha4y nha61n nu1t ""Tie61p Tu5c"""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   12632319
   End
   Begin UniControls.UniLabel UniLabel5 
      Height          =   255
      Left            =   240
      Top             =   480
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   450
      BackStyle       =   0
      Caption         =   "Co6ng vie65c ca62n la2m"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   12632319
   End
   Begin FVUnicodeControl.FVistaUniCheckbox chk 
      Height          =   195
      Index           =   4
      Left            =   960
      TabIndex        =   15
      Top             =   7200
      Width           =   4080
      _ExtentX        =   7197
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
      Caption         =   "Phu5c ho62i la5i ca1c kho1a Registry bi5 hu7 ho3ng sau khi die65t"
      BackStyle       =   0
      ForeColor       =   8454143
      ShowFocusRectangle=   0   'False
   End
   Begin FVUnicodeControl.FVistaUniCheckbox chk 
      Height          =   195
      Index           =   3
      Left            =   1200
      TabIndex        =   14
      Top             =   6960
      Width           =   2685
      _ExtentX        =   4736
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
      Caption         =   "Ta5o file ba1o ca1o sau khi die65t Virus"
      BackStyle       =   0
      ForeColor       =   8454143
      ShowFocusRectangle=   0   'False
   End
   Begin FVUnicodeControl.FVistaUniCheckbox chk 
      Height          =   195
      Index           =   2
      Left            =   1080
      TabIndex        =   13
      Top             =   6480
      Width           =   3630
      _ExtentX        =   6403
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
      Caption         =   "Bo3 qua nhu74ng file co1 dung lu7o75ng lo71n ho7n 10MB"
      BackStyle       =   0
      ForeColor       =   8454143
      ShowFocusRectangle=   0   'False
   End
   Begin FVUnicodeControl.FVistaUniCheckbox chk 
      Height          =   195
      Index           =   1
      Left            =   480
      TabIndex        =   12
      Top             =   7320
      Width           =   3405
      _ExtentX        =   6006
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
      Caption         =   "Que1t ca1c tie61n tri2nh d9ang cha5y trong bo65 nho71"
      BackStyle       =   0
      ForeColor       =   8454143
      ShowFocusRectangle=   0   'False
   End
   Begin FVUnicodeControl.FVistaUniCheckbox chk 
      Height          =   195
      Index           =   0
      Left            =   1440
      TabIndex        =   11
      Top             =   6600
      Width           =   3045
      _ExtentX        =   5371
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
      Caption         =   "Que1t ca1c File kho73i d9o65ng cu2ng he65 tho61ng"
      BackStyle       =   0
      ForeColor       =   8454143
      ShowFocusRectangle=   0   'False
   End
   Begin UniControls.UniLabel lbl_CacThiet 
      Height          =   255
      Left            =   1800
      Top             =   7080
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   450
      BackStyle       =   0
      Caption         =   "Ca1c thie61t la65p khi que1t"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   0
   End
   Begin VB.ComboBox cboEXT 
      Height          =   315
      Left            =   2040
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   6360
      Width           =   2895
   End
   Begin FVUnicodeControl.FVistaUniOption optEXT 
      Height          =   195
      Index           =   1
      Left            =   3480
      TabIndex        =   9
      Top             =   7080
      Width           =   1980
      _ExtentX        =   3493
      _ExtentY        =   344
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Value           =   -1  'True
      ShowFocusRectangle=   0   'False
      Caption         =   "Chi3 que1t file co1 d9uo6i:"
      BackStyle       =   0
      ForeColor       =   0
   End
   Begin FVUnicodeControl.FVistaUniOption optEXT 
      Height          =   195
      Index           =   0
      Left            =   3720
      TabIndex        =   8
      Top             =   6840
      Width           =   1950
      _ExtentX        =   3440
      _ExtentY        =   344
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ShowFocusRectangle=   0   'False
      Caption         =   "Ta61t ca3 ca1c loa5i (*.*)"
      BackStyle       =   0
      ForeColor       =   0
   End
   Begin UniControls.UniLabel lbl_DuoiFile 
      Height          =   255
      Left            =   2640
      Top             =   6720
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   450
      BackStyle       =   0
      Caption         =   "D9uo6i file se4 que1t"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   0
   End
   Begin UniControls.UniLabel lbl_HayCaiDat 
      Height          =   375
      Left            =   1320
      Top             =   6720
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   661
      BackStyle       =   0
      Caption         =   "Ha4y ca2i d9a85t ca61u hi2nh que1t cho phu2 ho75p vo71i mu5c d9i1ch que1t cu3a ba5n."
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
   Begin FVUnicodeControl.FVistaUniLabel lbl1 
      Height          =   255
      Left            =   600
      TabIndex        =   6
      Top             =   6720
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   450
      BackStyle       =   0
      Caption         =   "Bu7o71c 1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   255
   End
   Begin FVUnicodeControl.FVistaUniButton cmdDelPath 
      Height          =   225
      Left            =   840
      TabIndex        =   5
      Top             =   6480
      Width           =   225
      _ExtentX        =   397
      _ExtentY        =   397
      BackColor       =   16761024
      ButtonStyle     =   3
      Caption         =   "Xóa bo3"
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
   Begin FVUnicodeControl.FVistaUniButton cmdAddPath 
      Height          =   225
      Left            =   840
      TabIndex        =   4
      Top             =   6240
      Width           =   225
      _ExtentX        =   397
      _ExtentY        =   397
      BackColor       =   16761024
      ButtonStyle     =   3
      Caption         =   "Thêm vào"
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
   Begin UniControls.UniLabel lblPro 
      Height          =   255
      Index           =   1
      Left            =   600
      Top             =   1200
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   450
      BackStyle       =   0
      Caption         =   "Ca2i d9a85t ca61u hi2nh"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   12632256
   End
   Begin UniControls.UniLabel lblPro 
      Height          =   255
      Index           =   0
      Left            =   600
      Top             =   840
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   450
      BackStyle       =   0
      Caption         =   "Cho5n thu7 mu5c"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   12632256
   End
   Begin UniControls.UniLabel lbl3 
      Height          =   255
      Left            =   1080
      Top             =   6720
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   450
      BackStyle       =   0
      Caption         =   "Tie61p Tu5c"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   255
   End
   Begin FVUnicodeControl.FVistaUniButton cmdNext 
      Height          =   345
      Left            =   5760
      TabIndex        =   3
      Top             =   7200
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   609
      ButtonStyle     =   3
      Caption         =   "Tie61p tu5c"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaskColor       =   16777215
      PicNormal       =   "frmScan.frx":1BD0E
      PicSizeH        =   20
      PicSizeW        =   20
   End
   Begin VB.Timer Timer2 
      Left            =   8160
      Top             =   -120
   End
   Begin VB.Timer Timer1 
      Left            =   8520
      Top             =   -120
   End
   Begin UniControls.UniListBox lstPath 
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   6240
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   661
      IconMaskColor   =   16711935
      AutoUnicode     =   0   'False
      Picture         =   "frmScan.frx":1C348
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   64
      SelForeColor    =   0
      FullRowSelect   =   0   'False
      RowHeight       =   19
      AutoHideScrollBars=   -1  'True
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   8640
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   735
   End
   Begin FVUnicodeControl.FVistaUniButton cmdAllDrive 
      Height          =   225
      Left            =   1080
      TabIndex        =   0
      Top             =   6240
      Width           =   225
      _ExtentX        =   397
      _ExtentY        =   397
      BackColor       =   16761024
      ButtonStyle     =   3
      Caption         =   "Que1t ta61t ca3 o63 d9i4a"
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
   Begin UniControls.UniLabel lbl2 
      Height          =   255
      Left            =   120
      Top             =   6720
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   450
      BackStyle       =   0
      Caption         =   "Ha4y cho5n 1 hoa85c nhie62u khu vu75c d9e63 que1t sau d9o1 nha61n"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16711680
      Link            =   ""
   End
   Begin UniControls.UniLabel UniLabel1 
      Height          =   255
      Left            =   2760
      Top             =   240
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   450
      BackStyle       =   0
      Caption         =   "Que1t Virus"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   8388608
   End
   Begin FVUnicodeControl.FVistaUniLabel lbl_Buoc2 
      Height          =   375
      Left            =   2520
      TabIndex        =   7
      Top             =   6720
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      BackStyle       =   0
      Caption         =   "Bu7o71c 2"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   255
   End
   Begin FVUnicodeControl.FVistaUniButton cmdStopScan 
      Height          =   855
      Left            =   2400
      TabIndex        =   24
      Top             =   4440
      Visible         =   0   'False
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1508
      BackColor       =   -2147483633
      ButtonStyle     =   3
      Caption         =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PicAlign        =   0
      PicNormal       =   "frmScan.frx":2241B
      PicSizeH        =   48
      PicSizeW        =   48
   End
   Begin UniControls.UniLabel lblxTime 
      Height          =   255
      Left            =   3600
      Top             =   5760
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   450
      BackStyle       =   0
      Caption         =   "Tho72i gian:"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   12648384
   End
   Begin UniControls.UniLabel lblPro 
      Height          =   255
      Index           =   4
      Left            =   600
      Top             =   2280
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   450
      BackStyle       =   0
      Caption         =   "Que1t xong!"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   12632256
   End
   Begin UniControls.UniLabel lblPro 
      Height          =   255
      Index           =   5
      Left            =   600
      Top             =   2640
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   450
      BackStyle       =   0
      Caption         =   "Su73a lo64i he65 tho61ng"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   12632256
   End
   Begin UniControls.Resizer Resizer1 
      Left            =   9600
      Top             =   0
      _ExtentX        =   820
      _ExtentY        =   794
      ControlsSetupString=   $"frmScan.frx":27C0D
      MinFormWidth    =   718
      MinFormHeight   =   537
   End
   Begin VB.Image PicIcon 
      Height          =   255
      Index           =   5
      Left            =   360
      Top             =   2640
      Width           =   255
   End
   Begin VB.Image PicIcon 
      Height          =   255
      Index           =   4
      Left            =   360
      Top             =   2280
      Width           =   255
   End
   Begin VB.Label lblTime 
      BackStyle       =   0  'Transparent
      Caption         =   "00:00:00"
      ForeColor       =   &H008080FF&
      Height          =   255
      Left            =   4680
      TabIndex        =   26
      Top             =   5760
      Width           =   1095
   End
   Begin VB.Label lblVirus 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   4800
      TabIndex        =   25
      Top             =   5520
      Width           =   1215
   End
   Begin VB.Label lblTotalFile 
      BackStyle       =   0  'Transparent
      Caption         =   "123456"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H008080FF&
      Height          =   255
      Left            =   4920
      TabIndex        =   23
      Top             =   5040
      Width           =   855
   End
   Begin VB.Label lblProcess 
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "0 %"
      ForeColor       =   &H008080FF&
      Height          =   255
      Left            =   4560
      TabIndex        =   22
      Top             =   5280
      Width           =   1815
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H0080C0FF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   1095
      Left            =   3480
      Shape           =   4  'Rounded Rectangle
      Top             =   4920
      Width           =   2415
   End
   Begin VB.Image PicIcon 
      Height          =   255
      Index           =   3
      Left            =   360
      Top             =   1920
      Width           =   255
   End
   Begin VB.Image PicIcon 
      Height          =   255
      Index           =   2
      Left            =   360
      Top             =   1560
      Width           =   255
   End
   Begin VB.Image PicIcon 
      Height          =   255
      Index           =   1
      Left            =   360
      Top             =   1200
      Width           =   255
   End
   Begin VB.Image PicIcon 
      Height          =   255
      Index           =   0
      Left            =   360
      Top             =   840
      Width           =   255
   End
   Begin VB.Image picDang 
      Height          =   240
      Left            =   840
      Picture         =   "frmScan.frx":27F8B
      Top             =   120
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image picOK 
      Height          =   240
      Left            =   1680
      Picture         =   "frmScan.frx":283A7
      Top             =   120
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Shape Vong1 
      BorderColor     =   &H000080FF&
      Height          =   255
      Left            =   3840
      Shape           =   4  'Rounded Rectangle
      Top             =   6840
      Width           =   1815
   End
   Begin VB.Shape Vong2 
      BorderColor     =   &H000080FF&
      Height          =   255
      Left            =   3360
      Shape           =   4  'Rounded Rectangle
      Top             =   6840
      Width           =   3135
   End
   Begin VB.Line Line5 
      BorderColor     =   &H000000FF&
      X1              =   2280
      X2              =   9360
      Y1              =   6120
      Y2              =   6120
   End
   Begin VB.Line Line4 
      BorderColor     =   &H000000FF&
      X1              =   9360
      X2              =   9360
      Y1              =   360
      Y2              =   6120
   End
   Begin VB.Line Line3 
      BorderColor     =   &H000000FF&
      X1              =   3960
      X2              =   9360
      Y1              =   360
      Y2              =   360
   End
   Begin VB.Line Line2 
      BorderColor     =   &H000000FF&
      X1              =   2280
      X2              =   2640
      Y1              =   360
      Y2              =   360
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000FF&
      X1              =   2280
      X2              =   2280
      Y1              =   360
      Y2              =   6120
   End
   Begin VB.Image Image1 
      Height          =   1920
      Left            =   120
      Picture         =   "frmScan.frx":287D5
      Top             =   4320
      Width           =   1920
   End
End
Attribute VB_Name = "frmScan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long

Private Const GWL_EXSTYLE = (-20)
Private Const LWA_ALPHA = &H2
Private Const WS_EX_LAYERED = &H80000

Dim m_lAlpha
Dim m_Folder
Dim m_Handle As Long
Dim m_Folder2
Dim m_Handle2 As Long
Dim xProcess As Integer
Public xTongSoFile As Long
Public xStopScan As Boolean
Dim xScanNow As Boolean
Dim xTime As Long
Public xNowReport As String
Private Sub cmdBack_Click()
Timer3.Enabled = True
End Sub

Private Sub cmdCachLy_Click()
If UniMsgBox("Ba5n co1 cha81c cha81n muo61n ca1ch ly ta61t ca3 ca1c Virus d9a4 d9a1nh da61u kho6ng?", vbYesNo + vbQuestion, "Xo1a Virus", Me.hwnd) = vbYes Then
    'MsgBox "Delete"
LBL.Caption = "D9ang ca1ch ly..."
BaTdAuXoAlAi:
    Dim Jj As Long
    For Jj = 1 To LV1.ListItems.Count
    If LV1.ListItems(Jj).Checked = True Then
        CachLyFile LV1.ListItems(Jj).SubItems(1).Caption
        If modMain.FileExists(xFile) = False Then
            LV1.ListItems.Remove (Jj)
            GoTo BaTdAuXoAlAi
        End If
    End If
    Next Jj
    LBL.Caption = "D9a4 ca1ch ly xong!"
End If
End Sub

Private Sub cmdDeleteVirus_Click()
If GetSetting("PAV2009", "CauHinhVirus", "OffCom", False) = False Then
    If UniMsgBox("Ba5n co1 cha81c cha81n muo61n xo1a ta61t ca3 ca1c Virus d9a4 d9a1nh da61u kho6ng?", vbYesNo + vbQuestion, "Xo1a Virus", Me.hwnd) = vbYes Then
        'MsgBox "Delete"
    LBL.Caption = "D9ang xo1a..."
    GoTo BaTdAuXoAlAi
    End If
Else

        LBL.Caption = "D9ang xo1a..."
BaTdAuXoAlAi:
        Dim Jj As Long
        For Jj = 1 To LV1.ListItems.Count
        If LV1.ListItems(Jj).Checked = True Then
            tXoaFile LV1.ListItems(Jj).SubItems(1).Caption
            If modMain.FileExists(xFile) = False Then
                LV1.ListItems.Remove (Jj)
                GoTo BaTdAuXoAlAi
            End If
        End If
        Next Jj
        LBL.Caption = "D9a4 xo1a xong!"
        
        If GetSetting("PAV2009", "CauHinhVirus", "OffCom", False) = True Then
        frmOffComputer.Show
        End If
        
    End If

    'frmOffComputer.Show
End Sub

Private Sub cmdHide_Click()
frmMain.xScanning = True
frmScan.Hide
If xScanNow = True Then Timer6.Enabled = True
'frmMain.Tray.BalloonTip "Nha61n d9u1p chuo65t va2o d9a6 d9e63 hie65n chu7o7ng tri2nh.", btsInfo, "Tho6ng ba1o"
End Sub

Private Sub cmdNext_Click()
If lstPath.List(0) = ToUnicode("[Chu7a co1 thu7 mu5c na2o ca62n que1t!]") Then
    UniMsgBox "Ba5n ha4y cho5n no7i d9e63 que1t"
    Exit Sub
End If

xProcess = xProcess + 1

If xProcess = 1 Then
'\\\\\\\\\\ - Buoc thu nhat - \\\\\\\\\\\
lstPath.Visible = False
cmdAddPath.Visible = False
cmdDelPath.Visible = False
cmdAllDrive.Visible = False
lbl1.Visible = False
lbl2.Visible = False
lbl3.Visible = False

lblPro(0).ForeColor = &HC0C0C0
PicIcon(0).Picture = picOK.Picture
lblPro(1).ForeColor = &H8080FF
PicIcon(1).Picture = picDang.Picture
'=============
Me.lbl_Buoc2.Visible = True
Me.lbl_HayCaiDat.Visible = True
Me.lbl_CacThiet.Visible = True
Me.lbl_KhongCanThietLap.Visible = True
Me.lbl_DuoiFile.Visible = True
Me.chk(0).Visible = True
Me.chk(1).Visible = True
Me.chk(2).Visible = True
Me.chk(3).Visible = True
Me.chk(4).Visible = True
Me.optEXT(0).Visible = True
Me.optEXT(1).Visible = True
Me.cboEXT.Visible = True
Me.Vong1.Visible = True
Me.Vong2.Visible = True

Me.lbl_Buoc2.Left = 2400
Me.lbl_Buoc2.Top = 600

Me.lbl_HayCaiDat.Left = 2520
Me.lbl_HayCaiDat.Top = 1080

Me.lbl_CacThiet.Left = 2520
Me.lbl_CacThiet.Top = 2760


Me.lbl_KhongCanThietLap.Left = 2520
Me.lbl_KhongCanThietLap.Top = 5160

Me.lbl_DuoiFile.Left = 2520
Me.lbl_DuoiFile.Top = 1320

Me.chk(0).Left = 2640
Me.chk(0).Top = 3240
Me.chk(1).Left = 2640
Me.chk(1).Top = 3600
Me.chk(2).Left = 2640
Me.chk(2).Top = 3960
Me.chk(3).Left = 2640
Me.chk(3).Top = 4320
Me.chk(4).Left = 2640
Me.chk(4).Top = 4680
Me.optEXT(0).Left = 2640
Me.optEXT(0).Top = 1680
Me.optEXT(1).Left = 2640
Me.optEXT(1).Top = 1920

Me.cboEXT.Left = 2640
Me.cboEXT.Top = 2160

Me.Vong1.Left = 2520
Me.Vong1.Top = 1560
Me.Vong2.Left = 2520
Me.Vong2.Top = 3000
Me.Vong1.Height = 1095
Me.Vong1.Width = 4695
Me.Vong2.Height = 2055
Me.Vong2.Width = 6360
'\\\\\\\\\\\\\\\\\\\\\\\
End If



If xProcess = 2 Then
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
lblPro(1).ForeColor = &HC0C0C0
PicIcon(1).Picture = picOK.Picture
lblPro(2).ForeColor = &H8080FF
PicIcon(2).Picture = picDang.Picture
Me.lbl_Buoc2.Visible = False
Me.lbl_HayCaiDat.Visible = False
Me.lbl_CacThiet.Visible = False
Me.lbl_KhongCanThietLap.Visible = False
Me.lbl_DuoiFile.Visible = False
Me.chk(0).Visible = False
Me.chk(1).Visible = False
Me.chk(2).Visible = False
Me.chk(3).Visible = False
Me.chk(4).Visible = False
Me.optEXT(0).Visible = False
Me.optEXT(1).Visible = False
Me.cboEXT.Visible = False
Me.Vong1.Visible = False
Me.Vong2.Visible = False


'===================

Me.lblChuanBi1.Visible = True
Me.lblChuanBi2.Visible = True
Me.lblChuanBi3.Visible = True
Me.lblChuanBi4.Visible = True


    cmdBack.Visible = True
    
    lblChuanBi1.Left = 2400
    lblChuanBi1.Top = 480
    lblChuanBi1.AutoSize = True
    lblChuanBi2.Left = 2520
    lblChuanBi2.Top = 840
    lblChuanBi2.AutoSize = True
    
    lblChuanBi4.Top = cmdBack.Top - lblChuanBi4.Height - 120
    lblChuanBi4.Left = 2400
    lblChuanBi4.Height = 855
    lblChuanBi4.Width = Line4.x1 - lblChuanBi4.Left - 120
    
    lblChuanBi3.Left = 2400
    lblChuanBi3.Top = 1200
    lblChuanBi3.Width = Line4.x1 - lblChuanBi3.Left - 120
    lblChuanBi3.Height = lblChuanBi4.Top - lblChuanBi3.Top - 120
    lblChuanBi3.AutoUnicode = False
    lblChuanBi3.Caption = LamBaiChuanBi
    
'\\\\\\\\\\\\\\\\\\\\\
End If


If xProcess = 3 Then
lblPro(2).ForeColor = &HC0C0C0
PicIcon(2).Picture = picOK.Picture
lblPro(3).ForeColor = &H8080FF
PicIcon(3).Picture = picDang.Picture

    lblChuanBi1.Visible = False
    lblChuanBi2.Visible = False
    lblChuanBi3.Visible = False
    lblChuanBi4.Visible = False
    cmdNext.Visible = False

LBL.Visible = True
Me.lbl_HoangTat.Visible = True
Me.lblHayNhanVao.Visible = True
LV1.Visible = True
cmdStartStopScan.Visible = True
cmdStopScan.Visible = False
cmdDeleteVirus.Visible = True
cmdCachLy.Visible = True
cmdBack.Visible = True
Shape1.Visible = True

'\\\\\\\\\\\\\\\\\\\\\\\\
End If
End Sub

Private Sub cmdStartStopScan_Click()
If LV1.ListItems.Count > 1 Then
    If UniMsgBox("Trong sanh sa1ch hie65n d9ang to62n ta5i ke61t qua3 que1t Virus!" & vbCrLf & "Ne61u ba5n que1t virus ba6y gio72 thi2 danh sa1ch se4 bi5 xo1a sa5ch!" & vbCrLf & vbCrLf & " Ba5n co1 muo61n xo1a sa5ch ke61t qua3 va2 que1t la5i tu72 d9a62u kho6ng?", vbYesNo + vbCritical, "Tho6ng ba1o", Me.hwnd) = vbNo Then
        Exit Sub
    End If
End If

'SaveSetting "PAV2009", "CauHinhVirus", "OffCom", OffCom.Value


cmdStartStopScan.Visible = False
cmdStopScan.Visible = True

    Dim xTotal As Long
    Dim xFileEXT As String
    Dim Jh As Integer
    Dim JhX As Integer
    
    
    '============
    LBL.Caption = "Chua63n bi5 que1t...": DoEvents
    LV1.ListItems.Clear: DoEvents
    LV1.Enabled = False
    xScanNow = True
    xStopScan = False
    nFile = 0
    nVirus = 0
    lblVirus.Caption = 0
    lblProcess.Visible = False
    Me.lbl_Da_Quet.Visible = False
    Me.lblPhatHien.Visible = False
    Me.lblVirus.Visible = False
    Me.lblTime.Visible = False
    Me.lblxTime.Visible = False
    Me.lblTotalFile.Visible = False
    Me.lbl_Tong_So_File.Visible = False
    cmdStopScan.Enabled = False
    cmdDeleteVirus.Enabled = False
    cmdCachLy.Enabled = False
    cmdBack.Enabled = False
    Sleep 500: DoEvents
    LBL.Caption = "La61y tho6ng tin ca61u hi2nh que1t..."
    xTongSoFile = 0
    If frmScan.optEXT(0).Value = True Then
        xFileEXT = "*.*"
    ElseIf frmScan.optEXT(1).Value = True Then
        xFileEXT = Left(cboEXT.Text, 5)
    End If
    'MsgBox xFileEXT
    xTime = 0
    Me.lblTime.Caption = "00:00:01"
    Proc.Visible = True
    Timer4.Enabled = True
    Shape1.FillStyle = 0
    Sleep 354: DoEvents
    
    LBL.Caption = "D9ang d9e61m file se4 que1t..."
    Sleep 411: DoEvents
    'For Jh = 0 To lstPath.ListCount - 1
    '    'MsgBox lstPath.List(Jh)
    '    TongSoFile lstPath.List(Jh), xFileEXT
    'Next Jh
    
    Timer5.Enabled = True
    lblPro(3).ForeColor = &HC0C0C0
    PicIcon(3).Picture = picOK.Picture
    lblPro(4).ForeColor = &H8080FF
    PicIcon(4).Picture = picDang.Picture
    
    lblTotalFile.Caption = xTongSoFile: DoEvents
    Proc.Visible = False
    Timer4.Enabled = False
    Shape1.FillStyle = 1
    Proc.Visible = False
    lblProcess.Visible = True
    Me.lbl_Da_Quet.Visible = True
    Me.lblPhatHien.Visible = True
    Me.lblVirus.Visible = True
    Me.lblTime.Visible = True
    Me.lblxTime.Visible = True

    Me.lblTotalFile.Visible = True
    Me.lblTotalFile.Caption = "---"
    Me.lbl_Tong_So_File.Visible = True

    cmdStopScan.Enabled = True
    LBL.Caption = "Ba81t d9a62u que1t..."
    Sleep 200: DoEvents
    
    If chk(1).Value = True Then
        ScanProcess
    End If
    LBL.Caption = "D9ang que1t tie61n tri2nh"
    Sleep 200: DoEvents
    For JhX = 0 To lstPath.ListCount - 1
        LBL.Caption = "D9ang que1t ta5i: " & lstPath.List(JhX)
        scanvirus lstPath.List(JhX), xFileEXT, xTongSoFile
    Next JhX
    
    If chk(4).Value = True Then
        LBL.Caption = "D9ang phu5c ho62i Registry..."
        RegistryClean
        Sleep 200
    End If
    LBL.Caption = "D9a4 que1t xong!"
    ScanDone

End Sub

Private Sub cmdStopScan_Click()
    xStopScan = True
'ScanDone

End Sub

Private Sub Form_Load()

'&H008080FF&
    Dim lStyle As Long
    lStyle = GetWindowLong(Me.hwnd, GWL_EXSTYLE)
    lStyle = lStyle Or WS_EX_LAYERED
    SetWindowLong Me.hwnd, GWL_EXSTYLE, lStyle
    SetLayeredWindowAttributes Me.hwnd, 0, 0, LWA_ALPHA
    Timer1.Interval = 5
    Timer2.Interval = 5
    Timer2.Enabled = False
    Timer1.Enabled = True
    
    RefereshListPath

xScanNow = False
xProcess = 0
ConnectDB
'\\\\\\\\\\\\\\\\\\
lbl1.Left = 2520
lbl1.Top = 600
lbl2.Left = 3240
lbl2.Top = 960
lbl3.Left = 8280
lbl3.Top = 960
lbl1.AutoSize = True
lbl2.AutoSize = True
lbl3.AutoSize = True


lstPath.Left = 2400
lstPath.Top = 1320
cmdAddPath.Left = 2400
cmdDelPath.Left = 3960
cmdAllDrive.Left = 5520
cmdNext.Left = 7800
cmdAddPath.Width = 1455
cmdAddPath.Height = 375
cmdDelPath.Width = 1455
cmdDelPath.Height = 375
cmdAllDrive.Width = 1455
cmdAllDrive.Height = 375
cmdNext.Width = 1455
cmdNext.Height = 375
'\\\\\\\\\\\\\\\

'\\\\\\\\\\\\\\\
cboEXT.AddItem "*.EXE - Application"
cboEXT.AddItem "*.BAT - MS-DOS Batch File"
cboEXT.AddItem "*.CMD - Windows NT Command Script"
cboEXT.AddItem "*.COM - MS-DOS Application"
cboEXT.AddItem "*.DLL - Application Extension"
cboEXT.AddItem "*.OCX - ActiveX Control"
cboEXT.AddItem "*.PIF - Shortcut to MS-DOS Program"
cboEXT.AddItem "*.SCR - Screen Saver"

optEXT(0).Value = GetSetting("PAV2009", "CauHinhVirus", "OPTEXT0", False)
optEXT(1).Value = GetSetting("PAV2009", "CauHinhVirus", "OPTEXT1", True)
cboEXT.ListIndex = GetSetting("PAV2009", "CauHinhVirus", "IndexEXT", "0")
cboEXT.Enabled = GetSetting("PAV2009", "CauHinhVirus", "CboEn", True)
chk(0).Value = GetSetting("PAV2009", "CauHinhVirus", "Check0", True)
chk(1).Value = GetSetting("PAV2009", "CauHinhVirus", "Check1", True)
chk(2).Value = GetSetting("PAV2009", "CauHinhVirus", "Check2", True)
chk(3).Value = GetSetting("PAV2009", "CauHinhVirus", "Check3", True)
chk(4).Value = GetSetting("PAV2009", "CauHinhVirus", "Check4", True)

'\\\\\\\\\\\\\\\\

'\\\\\\\\\\\\\\
PicIcon(0).Picture = picDang.Picture
PicIcon(1).Picture = picDang.Picture
PicIcon(2).Picture = picDang.Picture
PicIcon(3).Picture = picDang.Picture
PicIcon(4).Picture = picDang.Picture
PicIcon(5).Picture = picDang.Picture

lblPro(0).ForeColor = &H8080FF
PicIcon(0).Picture = picDang.Picture
'\\\\\\\\\\\\\\

'\\\\\\\\\\\\\\
Me.lbl_Buoc2.Visible = False
Me.lbl_HayCaiDat.Visible = False
Me.lbl_CacThiet.Visible = False
Me.lbl_KhongCanThietLap.Visible = False
Me.lbl_DuoiFile.Visible = False
Me.chk(0).Visible = False
Me.chk(1).Visible = False
Me.chk(2).Visible = False
Me.chk(3).Visible = False
Me.chk(4).Visible = False
Me.optEXT(0).Visible = False
Me.optEXT(1).Visible = False
Me.cboEXT.Visible = False
Me.Vong1.Visible = False
Me.Vong2.Visible = False
'\\\\\\\\\\\\\\\\\

'\\\\\\\\\\\\\\\
Me.lbl_HoangTat.Visible = False
Me.lblHayNhanVao.Visible = False
LV1.Visible = False
cmdStartStopScan.Visible = False
cmdStopScan.Visible = False
cmdDeleteVirus.Visible = False
cmdCachLy.Visible = False
cmdBack.Visible = False
Shape1.Visible = False
LBL.Visible = False

lblProcess.Visible = False
Me.lbl_Da_Quet.Visible = False
Me.lblPhatHien.Visible = False
Me.lblVirus.Visible = False
Me.lblTime.Visible = False
Me.lblxTime.Visible = False

Me.lblTotalFile.Visible = False
Me.lbl_Tong_So_File.Visible = False
Proc.Visible = False
'\\\\\\\\\\\\\\\\\\\\\\\\

'\\\\\\\\\\\
Me.lblChuanBi1.Visible = False
Me.lblChuanBi2.Visible = False
Me.lblChuanBi3.Visible = False
Me.lblChuanBi4.Visible = False
'\\\\\\\\\\\



'\\\\\\\\\\\\\
LV1.AutoUnicode = False
LV1.FullRowSelect = True
LV1.View = eViewDetails
LV1.HeaderButtons = False
LV1.GridLines = True

LV1.Columns.add , , ToUnicode("Te6n Virus")
LV1.Columns.add , , ToUnicode("D9u7o72ng da64n"), , 4000
'\\\\\\\\\\\\\\\

'\\\\\\\\\\\\
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Unload frmMenu
If UnloadMode = 2 Then End
If xScanNow = True Then
    Cancel = 1
    Exit Sub
End If
    If UnloadMode <> vbFormCode Then
        Cancel = True
        Timer2.Enabled = True
    End If
End Sub

Private Sub Image2_Click()

End Sub




Private Sub LV1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
    PopupMenu frmMenu.scanvi
End If
End Sub

Private Sub OffCom_Click()
SaveSetting "PAV2009", "CauHinhVirus", "OffCom", OffCom.Value
End Sub

Private Sub optEXT_Click(Index As Integer)
If Index = 1 Then cboEXT.Enabled = True Else cboEXT.Enabled = False
End Sub

Private Sub Timer1_Timer()
    m_lAlpha = m_lAlpha + 15
    If (m_lAlpha > 255) Then
        m_lAlpha = 255
        Timer1.Enabled = False
    Else
        SetLayeredWindowAttributes Me.hwnd, 0, m_lAlpha, LWA_ALPHA
    End If
End Sub

Private Sub Timer2_Timer()
    m_lAlpha = m_lAlpha - 15
    If (m_lAlpha < 0) Then
        m_lAlpha = 0
        If KhanCap = False Then frmMain.Show
        frmMain.xScanning = False
        frmMenu.Tray.ToolTipText = "Perfect Antivirus 2009"
        Unload Me
    Else
        SetLayeredWindowAttributes Me.hwnd, 0, m_lAlpha, LWA_ALPHA
    End If
End Sub




Private Sub cmdAddPath_Click()

m_Folder = 0
Dim xPa As String
Dim xPi As Integer

xPa = ChonThuMuc(frmScan)
If xPa <> "" Then

For xPi = 0 To lstPath.ListCount - 1
    If xPa = lstPath.List(xPi) Then
        UniMsgBox "D9a4 to62n ta5i thu7 mu5c na2y ro62i!"
        RefereshListPath
        RefereshListPath
        Exit Sub
    End If
Next xPi

lstPath.AddItem xPa
RefereshListPath
RefereshListPath


End If


End Sub

Private Sub cmdAllDrive_Click()
lstPath.Clear
Drive1.Refresh
Dim xH As Integer
For xH = 0 To Drive1.ListCount - 1
    lstPath.AddItem UCase(Left(Drive1.List(xH), 1) & ":\")
Next xH
RefereshListPath
End Sub

Private Sub cmdDelPath_Click()
If lstPath.ListIndex <> -1 Then
    lstPath.Remove lstPath.ListIndex
    RefereshListPath
End If
End Sub

Sub RefereshListPath()
If lstPath.ListCount = 0 Then
    lstPath.AddItem ToUnicode("[Chu7a co1 thu7 mu5c na2o ca62n que1t!]")
    cmdNext.Enabled = False
Else
    If lstPath.List(0) = ToUnicode("[Chu7a co1 thu7 mu5c na2o ca62n que1t!]") Then lstPath.Remove (0)
    cmdNext.Enabled = True
End If
lstPath.Refresh
End Sub

Private Sub Form_Resize()
If Me.WindowState <> 1 Then
Image1.Top = Me.Height - Image1.Height - 480
cmdHide.Top = Me.Height - Image1.Height - cmdHide.Height - 480

'LV1.Columns(1).Width = 4000


Line5.X2 = Me.Width - 240
Line3.X2 = Me.Width - 240
Line4.x1 = Me.Width - 240
Line4.X2 = Me.Width - 240
Line1.Y2 = Me.Height - 700
Line4.Y2 = Me.Height - 700
Line5.Y1 = Me.Height - 700
Line5.Y2 = Me.Height - 700

cmdNext.Top = Line5.Y1 - cmdNext.Height - 240
cmdAddPath.Top = Line5.Y1 - cmdAddPath.Height - 240
cmdDelPath.Top = Line5.Y1 - cmdDelPath.Height - 240
cmdAllDrive.Top = Line5.Y1 - cmdAllDrive.Height - 240

cmdNext.Left = Line4.x1 - cmdNext.Width - 480

lstPath.Width = Line4.x1 - lstPath.Left - 120
lstPath.Height = cmdNext.Top - lstPath.Top - 120


cmdBack.Top = Line5.Y1 - cmdBack.Height - 120
cmdStartStopScan.Top = cmdBack.Top - cmdStartStopScan.Height - 120
cmdStopScan.Left = cmdStartStopScan.Left
cmdStopScan.Top = cmdStartStopScan.Top



LV1.Width = Line4.x1 - LV1.Left - 120
LV1.Height = cmdStartStopScan.Top - LV1.Top - 120
LV1.Columns(1).Width = LV1.Width / 4
LV1.Columns(2).Width = LV1.Width - LV1.Width / 4 - 480

LBL.Top = LV1.Top + LV1.Height + 120

cmdDeleteVirus.Top = LV1.Top + LV1.Height + 120
cmdCachLy.Top = LV1.Top + LV1.Height + 120

Shape1.Top = cmdDeleteVirus.Top + cmdDeleteVirus.Height + 120
Shape1.Height = Line5.Y1 - Shape1.Top - 120
Shape1.Width = cmdCachLy.Left + cmdCachLy.Width - Shape1.Left
Proc.Top = Shape1.Top + Shape1.Height / 2 - Proc.Height
Proc.Width = cmdCachLy.Left + cmdCachLy.Width - 240 - Proc.Left
Me.lbl_Tong_So_File.Top = Shape1.Top
Me.lblTotalFile.Top = Me.lbl_Tong_So_File.Top
Me.lbl_Da_Quet.Top = Me.lblTotalFile.Top + Me.lbl_Da_Quet.Height '+ 120
Me.lblProcess.Top = Me.lbl_Da_Quet.Top
Me.lblPhatHien.Top = Me.lbl_Da_Quet.Top + Me.lblPhatHien.Height '+ 120
Me.lblVirus.Top = Me.lblPhatHien.Top
Me.lblxTime.Top = Me.lblPhatHien.Top + Me.lblxTime.Height
Me.lblTime.Top = Me.lblxTime.Top




lblChuanBi1.Left = 2400
lblChuanBi1.Top = 480
lblChuanBi1.AutoSize = True
lblChuanBi2.Left = 2520
lblChuanBi2.Top = 840
lblChuanBi2.AutoSize = True
    
lblChuanBi4.Top = cmdBack.Top - lblChuanBi4.Height - 120
lblChuanBi4.Left = 2400
lblChuanBi4.Height = 855
lblChuanBi4.Width = Line4.x1 - lblChuanBi4.Left - 120
    
lblChuanBi3.Left = 2400
lblChuanBi3.Top = 1200
lblChuanBi3.Width = Line4.x1 - lblChuanBi3.Left - 120
lblChuanBi3.Height = lblChuanBi4.Top - lblChuanBi3.Top - 120





If Me.WindowState <> 2 Then
Me.Left = (Screen.Width - Me.Width) / 2
Me.Top = (Screen.Height - Me.Height) / 2
End If
End If
End Sub




Private Sub Timer3_Timer()
    m_lAlpha = m_lAlpha - 15
    If (m_lAlpha < 0) Then
        m_lAlpha = 0
        frmMain.Scan4Virus
    Else
        SetLayeredWindowAttributes Me.hwnd, 0, m_lAlpha, LWA_ALPHA
    End If
End Sub


Private Sub Timer4_Timer()
Proc.Value = Proc.Value + 1
If Proc.Value > 99 Then Proc.Value = 0
End Sub

Private Sub Timer5_Timer()
xTime = xTime + 1
Me.lblTime.Caption = CoTime(xTime)
End Sub

Public Function LamBaiChuanBi()
LamBaiChuanBi = ToUnicode("Tho6ng tin ca2i d9a85t:" & vbCrLf & vbCrLf _
& "-  Ca1c thu7 mu5c se4 que1t:" & vbCrLf)
Dim X As Integer
For X = 0 To lstPath.ListCount - 1
    LamBaiChuanBi = LamBaiChuanBi & "   ]" & lstPath.List(X) & vbCrLf
Next X
LamBaiChuanBi = LamBaiChuanBi & ToUnicode(vbCrLf & "-  Ca1c tho6ng so61 ca61u hi2nh:" & vbCrLf _
& "  - D9uo6i file se4 que1t: " & IIf(Me.optEXT(0).Value, "Ta61t ca3 ca1c loa5i file (*.*)", cboEXT.Text) & vbCrLf _
& "  - Que1t file kho73i d9o65ng cu2ng he65 tho61ng: " & IIf(chk(0).Value, "Co1", "Kho6ng") & vbCrLf _
& "  - Que1t ca1c tie61n tri2nh d9ang cha5y trong bo65 nho71: " & IIf(chk(1).Value, "Co1", "Kho6ng") & vbCrLf _
& "  - Bo3 qua nhu74ng file co1 dung lu7o75ng lo71n ho7n 10MB: " & IIf(chk(2).Value, "Co1", "Kho6ng") & vbCrLf _
& "  - Ta5o file ba1o ca1o sau khi que1t Virus: " & IIf(chk(3).Value, "Co1", "Kho6ng") & vbCrLf _
& "  - Phu5c ho62i la5i ca1c kho1a Registry bi5 ho3ng: " & IIf(chk(4).Value, "Co1", "Kho6ng") & vbCrLf)
End Function

Public Sub CheckAll()
'If LV1.ListItems.Count > 0 Then
Dim xAx As String
xAx = LBL.Caption
LBL.Caption = "D9ang d9a1nh da61u/ bo3 d9a1nh da61u..."
Dim U As Integer
For U = 1 To LV1.ListItems.Count
    DoEvents
    LV1.ListItems(U).Checked = True
    DoEvents
Next U
LBL.Caption = xAx
'End If
End Sub

Public Sub ScanDone()
    lblPro(4).ForeColor = &HC0C0C0
    PicIcon(4).Picture = picOK.Picture
    lblPro(5).ForeColor = &H8080FF
    PicIcon(5).Picture = picDang.Picture
    Timer5.Enabled = False
    LBL.Caption = "D9ang loa5i bo3 file tin tu7o73ng..."
    LoaiBoTinTuong
    CheckAll
    LBL.Caption = "D9ang la2m ba1o ca1o..."
    MakeReport
    frmMenu.Tray.ToolTipText = "Perfect Antivirus 2009"
    xScanNow = False
    LV1.Enabled = True
    cmdBack.Enabled = True
    cmdDeleteVirus.Enabled = True
    cmdCachLy.Enabled = True
    
cmdStartStopScan.Visible = True
cmdStopScan.Visible = False
'SaveSetting "PAV2009", "CauHinhVirus", "OffCom", OffCom.Value
If GetSetting("PAV2009", "CauHinhVirus", "OffCom", False) = True Then
    cmdDeleteVirus_Click
End If
End Sub

Private Sub Timer6_Timer()
frmMenu.Tray.ToolTipText = "D9ang que1t... " & frmScan.lblProcess.Caption
End Sub
