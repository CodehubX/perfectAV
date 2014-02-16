VERSION 5.00
Object = "{D1ECD258-D329-4388-AB83-DEC261A66B86}#1.0#0"; "UniControls_v2.0.ocx"
Begin VB.Form frmMain 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Perfect Antivirus 2009"
   ClientHeight    =   8175
   ClientLeft      =   2625
   ClientTop       =   1080
   ClientWidth     =   10230
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
   ScaleHeight     =   8175
   ScaleWidth      =   10230
   StartUpPosition =   2  'CenterScreen
   Begin UniControls.UniFrame fm 
      Height          =   6735
      Index           =   7
      Left            =   3000
      Top             =   1320
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   11880
      MaskColor       =   16711935
      FrameColor      =   -2147483629
      Style           =   0
      Caption         =   "Qua3n ly1 Tie61n Tri2nh"
      TextColor       =   13579779
      Alignment       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.ListBox lstPro 
         Height          =   255
         Left            =   6360
         TabIndex        =   44
         Top             =   240
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Timer tmrPro 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   120
         Top             =   240
      End
      Begin UniControls.UniLabel UniLabel33 
         Height          =   255
         Left            =   1440
         Top             =   600
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   450
         Alignment       =   1
         Caption         =   "Co1 the63 xem d9u7o75c ca1c Process a63n"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   33023
      End
      Begin UniControls.UniListView LVPro 
         Height          =   5295
         Left            =   240
         TabIndex        =   43
         Top             =   1200
         Width           =   6615
         _ExtentX        =   11668
         _ExtentY        =   9340
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MultiSelect     =   -1  'True
         LabelEdit       =   0   'False
         FullRowSelect   =   -1  'True
         AutoArrange     =   0   'False
         HeaderButtons   =   0   'False
         HeaderTrackSelect=   0   'False
         HideSelection   =   0   'False
         InfoTips        =   0   'False
      End
      Begin UniControls.UniLabel UniLabel31 
         Height          =   255
         Left            =   240
         Top             =   360
         Width           =   6615
         _ExtentX        =   11668
         _ExtentY        =   450
         Alignment       =   1
         Caption         =   "Theo do4i va2 qua3n ly1 ca1c tie61n tri2nh d9ang cha5y"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   255
      End
   End
   Begin UniControls.UniFrame fm 
      Height          =   6735
      Index           =   6
      Left            =   3000
      Top             =   1320
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   11880
      MaskColor       =   16711935
      FrameColor      =   -2147483629
      Style           =   0
      Caption         =   "Tu75 d9o65ng xo1a Autorun"
      TextColor       =   13579779
      Alignment       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin UniControls.UniLabel UniLabel36 
         Height          =   495
         Left            =   360
         Top             =   2040
         Width           =   6375
         _ExtentX        =   11245
         _ExtentY        =   873
         Caption         =   $"frmMain.frx":57E2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   255
      End
      Begin UniControls.UniLabel atplblStaAutorun 
         Height          =   255
         Left            =   5760
         Top             =   360
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   450
         Caption         =   "D9ang mo73"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   255
      End
      Begin VB.Timer atptmrAutorun 
         Enabled         =   0   'False
         Interval        =   5000
         Left            =   5400
         Top             =   1080
      End
      Begin UniControls.UniButton atpcmdAutorun 
         Height          =   735
         Left            =   2280
         TabIndex        =   41
         Top             =   3840
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   1296
         Icon            =   "frmMain.frx":588F
         Style           =   2
         Caption         =   "Ta81t chu71c na8ng na2y"
         IconAlign       =   3
         iNonThemeStyle  =   2
         BackColor       =   -2147483643
         MaskColor       =   16711935
         FontColor       =   16711680
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
      Begin UniControls.UniLabel UniLabel35 
         Height          =   495
         Left            =   240
         Top             =   720
         Width           =   6495
         _ExtentX        =   11456
         _ExtentY        =   873
         Caption         =   $"frmMain.frx":58AB
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   33023
      End
      Begin UniControls.UniLabel UniLabel30 
         Height          =   375
         Left            =   120
         Top             =   360
         Width           =   6855
         _ExtentX        =   12091
         _ExtentY        =   661
         Alignment       =   1
         Caption         =   "Tu75 d9o65ng xo1a Autorun"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   255
      End
   End
   Begin UniControls.UniFrame fm 
      Height          =   6735
      Index           =   5
      Left            =   3000
      Top             =   1320
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   11880
      MaskColor       =   16711935
      FrameColor      =   -2147483629
      Style           =   0
      Caption         =   "Ba3o ve65 Registry"
      TextColor       =   13579779
      Alignment       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.Timer atptmrREG 
         Interval        =   5000
         Left            =   5160
         Top             =   600
      End
      Begin UniControls.UniButton cmdDelREG 
         Height          =   375
         Left            =   2280
         TabIndex        =   40
         Top             =   6240
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   661
         Icon            =   "frmMain.frx":5974
         Style           =   2
         Caption         =   "Xo1a chu71c na8ng d9a4 d9a1nh da61u"
         IconAlign       =   3
         iNonThemeStyle  =   2
         BackColor       =   -2147483643
         MaskColor       =   16711935
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         RoundedBordersByTheme=   0   'False
         ShowFocusRectangle=   0   'False
      End
      Begin UniControls.UniLabel UniLabel34 
         Height          =   255
         Left            =   120
         Top             =   1800
         Width           =   6855
         _ExtentX        =   12091
         _ExtentY        =   450
         Caption         =   "Ca1c chu71c na8ng d9ang d9u7o75c ba3o ve65. (Nha61n nu1t The6m Va2o d9e63 the6m chu71c na8ng ca62n ba3o ve65)."
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   16711680
      End
      Begin UniControls.UniButton atpREGAdd 
         Height          =   375
         Left            =   120
         TabIndex        =   39
         Top             =   6240
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   661
         Icon            =   "frmMain.frx":5F0E
         Style           =   2
         Caption         =   "The6m va2o"
         IconAlign       =   3
         iNonThemeStyle  =   2
         BackColor       =   -2147483643
         MaskColor       =   16711935
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         RoundedBordersByTheme=   0   'False
         ShowFocusRectangle=   0   'False
      End
      Begin UniControls.UniListView atpLVREG 
         Height          =   3975
         Left            =   120
         TabIndex        =   38
         Top             =   2160
         Width           =   6855
         _ExtentX        =   12091
         _ExtentY        =   7011
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   12648447
         MultiSelect     =   -1  'True
         LabelEdit       =   0   'False
         AutoArrange     =   0   'False
         HeaderButtons   =   0   'False
         HeaderTrackSelect=   0   'False
         HideSelection   =   0   'False
         InfoTips        =   0   'False
      End
      Begin UniControls.UniLabel UniLabel29 
         Height          =   255
         Left            =   120
         Top             =   240
         Width           =   6855
         _ExtentX        =   12091
         _ExtentY        =   450
         Alignment       =   1
         BackStyle       =   0
         Caption         =   "Tu75 d9o65ng ba3o ve65 Registry"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   255
      End
      Begin UniControls.UniLabel UniLabel32 
         Height          =   495
         Left            =   240
         Top             =   600
         Width           =   6615
         _ExtentX        =   11668
         _ExtentY        =   873
         BackStyle       =   0
         Caption         =   $"frmMain.frx":64A8
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   33023
      End
      Begin UniControls.UniButton atpcmdREG 
         Height          =   375
         Left            =   2400
         TabIndex        =   37
         Top             =   1080
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   661
         Style           =   2
         Caption         =   "Ta81t chu71c na8ng na2y"
         IconAlign       =   3
         iNonThemeStyle  =   2
         BackColor       =   -2147483643
         MaskColor       =   16711935
         FontColor       =   16711680
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
      Begin UniControls.UniLabel atplbREG 
         Height          =   255
         Left            =   5640
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         Caption         =   "D9ang mo73"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   255
      End
   End
   Begin UniControls.UniFrame fm 
      Height          =   6735
      Index           =   4
      Left            =   3000
      Top             =   1320
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   11880
      MaskColor       =   16711935
      FrameColor      =   -2147483629
      Style           =   0
      Caption         =   "Nha65t Ky1 Virus"
      TextColor       =   13579779
      Alignment       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin UniControls.UniButton cmdDelSelected 
         Height          =   375
         Left            =   1680
         TabIndex        =   36
         Top             =   6240
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         Icon            =   "frmMain.frx":6586
         Style           =   2
         Caption         =   "Xo1a mu5c d9a4 cho5n"
         IconAlign       =   3
         iNonThemeStyle  =   2
         BackColor       =   -2147483643
         MaskColor       =   16711935
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         RoundedBordersByTheme=   0   'False
         ShowFocusRectangle=   0   'False
      End
      Begin UniControls.UniButton cmdEventsVirusKillAll 
         Height          =   375
         Left            =   120
         TabIndex        =   35
         Top             =   6240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         Icon            =   "frmMain.frx":6B20
         Style           =   2
         Caption         =   "Xo1a he61t"
         IconAlign       =   3
         iNonThemeStyle  =   2
         BackColor       =   -2147483643
         MaskColor       =   16711935
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         RoundedBordersByTheme=   0   'False
         ShowFocusRectangle=   0   'False
      End
      Begin UniControls.UniLabel UniLabel28 
         Height          =   375
         Left            =   240
         Top             =   720
         Width           =   6735
         _ExtentX        =   11880
         _ExtentY        =   661
         Caption         =   "Ghi che1p ca1c la62n que1t Virus cu3a ba5n."
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   33023
      End
      Begin UniControls.UniListView LVVirusEvents 
         Height          =   4935
         Left            =   120
         TabIndex        =   34
         Top             =   1200
         Width           =   6855
         _ExtentX        =   12091
         _ExtentY        =   8705
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   12648447
         MultiSelect     =   -1  'True
         LabelEdit       =   0   'False
         AutoArrange     =   0   'False
         CheckBoxes      =   -1  'True
         HeaderButtons   =   0   'False
         HeaderTrackSelect=   0   'False
         HideSelection   =   0   'False
         InfoTips        =   0   'False
      End
      Begin UniControls.UniLabel UniLabel27 
         Height          =   375
         Left            =   360
         Top             =   360
         Width           =   6375
         _ExtentX        =   11245
         _ExtentY        =   661
         Alignment       =   1
         Caption         =   "Nha65t ky1 ca1c la62n que1t Virus"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   255
      End
   End
   Begin UniControls.UniFrame fm 
      Height          =   6735
      Index           =   3
      Left            =   3000
      Top             =   1320
      Visible         =   0   'False
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   11880
      MaskColor       =   16711935
      FrameColor      =   -2147483629
      Style           =   0
      Caption         =   "Ca61u hi2nh que1t"
      TextColor       =   13579779
      Alignment       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin UniControls.UniLabel UniLabel17 
         Height          =   375
         Left            =   240
         Top             =   600
         Width           =   6615
         _ExtentX        =   11668
         _ExtentY        =   661
         Caption         =   "- Ha4y ca2i d9a85t ca61u hi2nh que1t cho phu2 ho75p d9e63 co1 d9u7o75c ke61t qua3 que1t to61t nha61t."
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   33023
      End
      Begin UniControls.UniButton cmdFullScan 
         Height          =   375
         Left            =   360
         TabIndex        =   17
         Top             =   4080
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   661
         Icon            =   "frmMain.frx":70BA
         Style           =   2
         Caption         =   "Que1t Virus Toa2n Bo65"
         IconAlign       =   3
         iNonThemeStyle  =   2
         BackColor       =   -2147483643
         MaskColor       =   16711935
         FontColor       =   255
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         RoundedBordersByTheme=   0   'False
         ShowFocusRectangle=   0   'False
      End
      Begin UniControls.UniCheckBox VSdll 
         Height          =   195
         Left            =   2040
         TabIndex        =   15
         Top             =   1680
         Width           =   705
         _ExtentX        =   1244
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
         Caption         =   "*.DLL"
         ForeColor       =   0
      End
      Begin UniControls.UniLabel UniLabel11 
         Height          =   255
         Left            =   240
         Top             =   2160
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   450
         Caption         =   "Chu71c Na8ng Que1t:"
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
      Begin UniControls.NumbericUpDown VSLimitSize 
         Height          =   330
         Left            =   3480
         TabIndex        =   14
         Top             =   3120
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   582
         Locked          =   -1  'True
         Text            =   "1"
         Max             =   10
         Min             =   1
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
      Begin UniControls.UniLabel UniLabel10 
         Height          =   255
         Left            =   4200
         Top             =   3240
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   450
         Caption         =   "MB"
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
      Begin UniControls.UniCheckBox VSDontScanSize 
         Height          =   195
         Left            =   360
         TabIndex        =   13
         Top             =   3240
         Width           =   3060
         _ExtentX        =   5398
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
         Caption         =   "Bo3 qua ca1c File co1 dung lu7o75ng lo71n ho7n:"
         ForeColor       =   0
      End
      Begin UniControls.UniCheckBox VSScanProcess 
         Height          =   195
         Left            =   360
         TabIndex        =   12
         Top             =   2880
         Width           =   3510
         _ExtentX        =   6191
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
         Caption         =   "Que1t ca1c u71ng du5ng d9ang cha5y trong bo65 nho71."
         ForeColor       =   0
      End
      Begin UniControls.UniCheckBox VSScanStartUp 
         Height          =   195
         Left            =   360
         TabIndex        =   11
         Top             =   2520
         Width           =   3780
         _ExtentX        =   6668
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
         Caption         =   "Que1t ca1c chu7o7ng tri2nh kho73i d9o65ng cu2ng he65 tho61ng."
         ForeColor       =   0
      End
      Begin UniControls.UniLabel UniLabel7 
         Height          =   255
         Left            =   120
         Top             =   240
         Width           =   6855
         _ExtentX        =   12091
         _ExtentY        =   450
         Alignment       =   1
         Caption         =   "Ca2i D9a85t Ca61u Hi2nh Que1t"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   255
      End
      Begin UniControls.UniLabel UniLabel9 
         Height          =   255
         Left            =   240
         Top             =   1080
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   450
         Caption         =   "Kie63u File Que1t:"
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
      Begin UniControls.UniCheckBox VSexe 
         Height          =   195
         Left            =   360
         TabIndex        =   6
         Top             =   1440
         Width           =   720
         _ExtentX        =   1270
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
         Caption         =   "*.EXE"
         Enabled         =   0   'False
         ForeColor       =   0
      End
      Begin UniControls.UniCheckBox VSbat 
         Height          =   195
         Left            =   360
         TabIndex        =   7
         Top             =   1680
         Width           =   735
         _ExtentX        =   1296
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
         Caption         =   "*.BAT"
         ForeColor       =   0
      End
      Begin UniControls.UniCheckBox VScom 
         Height          =   195
         Left            =   1200
         TabIndex        =   8
         Top             =   1440
         Width           =   795
         _ExtentX        =   1402
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
         Caption         =   "*.COM"
         ForeColor       =   0
      End
      Begin UniControls.UniCheckBox VScmd 
         Height          =   195
         Left            =   1200
         TabIndex        =   9
         Top             =   1680
         Width           =   780
         _ExtentX        =   1376
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
         Caption         =   "*.CMD"
         ForeColor       =   0
      End
      Begin UniControls.UniCheckBox VSscr 
         Height          =   195
         Left            =   2040
         TabIndex        =   10
         Top             =   1440
         Width           =   750
         _ExtentX        =   1323
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
         Caption         =   "*.SCR"
         ForeColor       =   0
      End
   End
   Begin UniControls.UniFrame fm 
      Height          =   6735
      Index           =   2
      Left            =   3000
      Top             =   1320
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   11880
      MaskColor       =   16711935
      FrameColor      =   -2147483629
      Style           =   0
      Caption         =   "Tu75 d9o65ng que1t"
      TextColor       =   13579779
      Alignment       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin UniControls.UniLabel UniLabel26 
         Height          =   255
         Left            =   480
         Top             =   4920
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   450
         Caption         =   "[Tu75 d9o65ng que1t va2 kie63m tra Keylogger mo64i 5 gia6y]"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   33023
      End
      Begin UniControls.UniCheckBox atsScanKeylogger 
         Height          =   195
         Left            =   240
         TabIndex        =   33
         Top             =   4680
         Width           =   5295
         _ExtentX        =   9340
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
         Caption         =   "Tu75 d9o65ng pha1t hie65n va2 ca3nh ba1o Virus cho nhu74ng tie61n tri2nh d9ang cha5y."
         ForeColor       =   0
         ShowFocusRectangle=   0   'False
      End
      Begin UniControls.UniLabel UniLabel25 
         Height          =   255
         Left            =   480
         Top             =   4320
         Width           =   6495
         _ExtentX        =   11456
         _ExtentY        =   450
         Caption         =   "[Ba5n ne6n su73 du5ng chu71c na8ng na2y d9e63 d9a3m ba3o cho USB cu3a ba5n tra1nh kho3i Virus.]"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   33023
      End
      Begin UniControls.UniCheckBox atsAutoScanUSB 
         Height          =   195
         Left            =   240
         TabIndex        =   20
         Top             =   4080
         Width           =   5145
         _ExtentX        =   9075
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
         Caption         =   "Tu75 d9o65ng que1t Virus cho USB khi pha1t hie65n USB ke61t no61i va2o ma1y ti1nh."
         ForeColor       =   0
         ShowFocusRectangle=   0   'False
      End
      Begin UniControls.UniLabel UniLabel19 
         Height          =   495
         Left            =   480
         Top             =   3600
         Width           =   6495
         _ExtentX        =   11456
         _ExtentY        =   873
         Caption         =   $"frmMain.frx":7ACC
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   33023
      End
      Begin UniControls.UniLabel UniLabel18 
         Height          =   855
         Left            =   480
         Top             =   2400
         Width           =   6375
         _ExtentX        =   11245
         _ExtentY        =   1508
         Caption         =   $"frmMain.frx":7B53
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   33023
      End
      Begin UniControls.UniCheckBox atsAlwaysScanFolder 
         Height          =   195
         Left            =   240
         TabIndex        =   19
         Top             =   3360
         Width           =   4815
         _ExtentX        =   8493
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
         Caption         =   "Tu75 d9o65ng que1t va2 ca3nh ba1o Virus trong Thu7 mu5c d9ang d9u7o75c mo73."
         ForeColor       =   0
         ShowFocusRectangle=   0   'False
      End
      Begin UniControls.UniCheckBox atsScanEXE 
         Height          =   195
         Left            =   240
         TabIndex        =   18
         Top             =   2160
         Width           =   4710
         _ExtentX        =   8308
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
         Caption         =   "Tu75 d9o65ng que1t va2 ca3nh ba1o Virus cho nhu74ng File sa81p d9u7o75c mo73."
         ForeColor       =   0
         ShowFocusRectangle=   0   'False
      End
      Begin UniControls.UniLabel UniLabel16 
         Height          =   735
         Left            =   240
         Top             =   960
         Width           =   6615
         _ExtentX        =   11668
         _ExtentY        =   1296
         Caption         =   $"frmMain.frx":7CCF
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   33023
      End
      Begin UniControls.UniLabel UniLabel15 
         Height          =   375
         Left            =   120
         Top             =   480
         Width           =   6855
         _ExtentX        =   12091
         _ExtentY        =   661
         Alignment       =   1
         Caption         =   "Ca2i d9a85t chu71c na8ng tu75 d9o65ng que1t"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   255
      End
   End
   Begin UniControls.UniTrayIcon Tray1 
      Left            =   7800
      Top             =   480
      _ExtentX        =   529
      _ExtentY        =   529
      TooltipText     =   "Perfect Antivirus 2009 - Ma1y ti1nh cu3a ba5n d9ang o73 ti2nh tra5ng to61t nha61t!"
      Icon            =   "frmMain.frx":7DFE
   End
   Begin UniControls.UniFrame fm 
      Height          =   6735
      Index           =   1
      Left            =   3000
      Top             =   1320
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   11880
      MaskColor       =   16711935
      FrameColor      =   -2147483629
      Style           =   0
      Caption         =   "Que1t tu2y cho5n"
      TextColor       =   13579779
      Alignment       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin UniControls.UniButton csBack 
         Height          =   375
         Left            =   6480
         TabIndex        =   32
         Top             =   6000
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         Icon            =   "frmMain.frx":8398
         Style           =   2
         IconAlign       =   2
         iNonThemeStyle  =   2
         Enabled         =   0   'False
         BackColor       =   15398133
         MaskColor       =   16711935
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         RoundedBordersByTheme=   0   'False
         ShowFocusRectangle=   0   'False
      End
      Begin UniControls.UniButton csKill 
         Height          =   375
         Left            =   4560
         TabIndex        =   31
         Top             =   6000
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         Icon            =   "frmMain.frx":8DAA
         Style           =   2
         Caption         =   "Die65t"
         IconAlign       =   3
         iNonThemeStyle  =   2
         BackColor       =   15398133
         MaskColor       =   16711935
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         RoundedBordersByTheme=   0   'False
         ShowFocusRectangle=   0   'False
      End
      Begin UniControls.UniButton csCachLy 
         Height          =   375
         Left            =   3120
         TabIndex        =   30
         Top             =   6000
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         Icon            =   "frmMain.frx":9344
         Style           =   2
         Caption         =   "Ca1ch Ly"
         IconAlign       =   3
         iNonThemeStyle  =   2
         BackColor       =   15398133
         MaskColor       =   16711935
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         RoundedBordersByTheme=   0   'False
         ShowFocusRectangle=   0   'False
      End
      Begin UniControls.UniButton csStop 
         Height          =   375
         Left            =   1680
         TabIndex        =   29
         Top             =   6000
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         Icon            =   "frmMain.frx":9D56
         Style           =   2
         Caption         =   "Du72ng"
         IconAlign       =   3
         iNonThemeStyle  =   2
         Enabled         =   0   'False
         BackColor       =   -2147483643
         MaskColor       =   16711935
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         RoundedBordersByTheme=   0   'False
         ShowFocusRectangle=   0   'False
      End
      Begin UniControls.UniButton csStart 
         Height          =   375
         Left            =   240
         TabIndex        =   28
         Top             =   6000
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         Icon            =   "frmMain.frx":A2F0
         Style           =   2
         Caption         =   "Que1t"
         IconAlign       =   3
         iNonThemeStyle  =   2
         BackColor       =   -2147483643
         MaskColor       =   16711935
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         RoundedBordersByTheme=   0   'False
         ShowFocusRectangle=   0   'False
      End
      Begin UniControls.UniFrame ff 
         Height          =   4815
         Left            =   120
         Top             =   1080
         Width           =   6855
         _ExtentX        =   12091
         _ExtentY        =   8493
         MaskColor       =   16711935
         FrameColor      =   -2147483629
         Style           =   0
         Caption         =   ""
         TextColor       =   13579779
         Alignment       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin UniControls.UniLabel cslblPath 
            Height          =   615
            Left            =   120
            Top             =   1680
            Width           =   6615
            _ExtentX        =   11668
            _ExtentY        =   1085
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
         End
         Begin UniControls.UniCheckBox CScom 
            Height          =   195
            Left            =   3120
            TabIndex        =   21
            Top             =   1080
            Width           =   795
            _ExtentX        =   1402
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
            Caption         =   "*.COM"
            ForeColor       =   0
            ShowFocusRectangle=   0   'False
         End
         Begin UniControls.UniCheckBox CSbat 
            Height          =   195
            Left            =   2280
            TabIndex        =   22
            Top             =   1080
            Width           =   735
            _ExtentX        =   1296
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
            Caption         =   "*.BAT"
            ForeColor       =   0
            ShowFocusRectangle=   0   'False
         End
         Begin UniControls.UniCheckBox CSexe 
            Height          =   195
            Left            =   1440
            TabIndex        =   23
            Top             =   1080
            Width           =   720
            _ExtentX        =   1270
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
            Caption         =   "*.EXE"
            Enabled         =   0   'False
            ForeColor       =   0
            ShowFocusRectangle=   0   'False
         End
         Begin UniControls.UniLabel UniLabel24 
            Height          =   255
            Left            =   240
            Top             =   1080
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   450
            Caption         =   "Kie63u File Que1t:"
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
         Begin UniControls.UniCheckBox CSprocess 
            Height          =   195
            Left            =   240
            TabIndex        =   24
            Top             =   840
            Width           =   3735
            _ExtentX        =   6588
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
            Caption         =   "Que1t ca1c chu7o7ng tri2nh d9ang cha5y trong bo65 nho71."
            ForeColor       =   0
            ShowFocusRectangle=   0   'False
         End
         Begin UniControls.UniCheckBox CSStartUp 
            Height          =   195
            Left            =   240
            TabIndex        =   25
            Top             =   600
            Width           =   3780
            _ExtentX        =   6668
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
            Caption         =   "Que1t ca1c chu7o7ng tri2nh kho73i d9o65ng cu2ng he65 tho61ng."
            ForeColor       =   0
            ShowFocusRectangle=   0   'False
         End
         Begin UniControls.UniLabel UniLabel23 
            Height          =   255
            Left            =   120
            Top             =   240
            Width           =   3255
            _ExtentX        =   5741
            _ExtentY        =   450
            Caption         =   "Tu2y cho5n que1t:"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   16711680
         End
         Begin UniControls.UniFolderView CSFolderView1 
            Height          =   2295
            Left            =   120
            TabIndex        =   26
            Top             =   2400
            Width           =   6615
            _ExtentX        =   11668
            _ExtentY        =   4048
         End
         Begin UniControls.UniLabel UniLabel22 
            Height          =   255
            Left            =   120
            Top             =   1440
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   450
            Caption         =   "Cho5n no7i ca62n que1t:"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   16711680
         End
      End
      Begin UniControls.UniLabel UniLabel21 
         Height          =   255
         Left            =   240
         Top             =   720
         Width           =   6615
         _ExtentX        =   11668
         _ExtentY        =   450
         Caption         =   "- Cho phe1p que1t theo y1 thi1ch cu3a ngu7o72i su73 du5ng."
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   33023
      End
      Begin UniControls.UniLabel UniLabel20 
         Height          =   495
         Left            =   240
         Top             =   360
         Width           =   6615
         _ExtentX        =   11668
         _ExtentY        =   873
         Alignment       =   1
         Caption         =   "Que1t Tu2y Cho5n"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   255
      End
      Begin UniControls.UniListView LVVirus2 
         Height          =   3135
         Left            =   120
         TabIndex        =   27
         Top             =   1320
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
         ForeColor       =   255
         BackColor       =   12648447
         MultiSelect     =   -1  'True
         LabelEdit       =   0   'False
         AutoArrange     =   0   'False
         HeaderButtons   =   0   'False
         HeaderTrackSelect=   0   'False
         HideSelection   =   0   'False
         InfoTips        =   0   'False
      End
      Begin UniControls.UniLabel cslblStatus 
         Height          =   615
         Left            =   120
         Top             =   4800
         Width           =   6855
         _ExtentX        =   12091
         _ExtentY        =   1085
         AutoUnicode     =   0   'False
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
         ForeColor       =   16711680
      End
      Begin UniControls.UniLabel cslblStatus2 
         Height          =   255
         Left            =   240
         Top             =   5640
         Width           =   6615
         _ExtentX        =   11668
         _ExtentY        =   450
         Alignment       =   1
         Caption         =   "Sa84n sa2ng"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   128
      End
   End
   Begin VB.Timer tmrScanTime 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   9000
      Top             =   360
   End
   Begin VB.Timer tmrStartFullScan 
      Enabled         =   0   'False
      Interval        =   789
      Left            =   9360
      Top             =   360
   End
   Begin UniControls.UniFrame fm 
      Height          =   6735
      Index           =   0
      Left            =   3000
      Top             =   1320
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   11880
      MaskColor       =   16711935
      FrameColor      =   -2147483629
      Style           =   0
      Caption         =   "Que1t Toa2n Bo65"
      TextColor       =   13579779
      Alignment       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin UniControls.UniLabel UniLabel13 
         Height          =   255
         Left            =   4920
         Top             =   5400
         Width           =   135
         _ExtentX        =   238
         _ExtentY        =   450
         Caption         =   ":"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   255
      End
      Begin UniControls.UniLabel lblScanTime 
         Height          =   255
         Index           =   0
         Left            =   4680
         Top             =   5400
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   450
         Caption         =   "00"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   255
      End
      Begin UniControls.UniLabel UniLabel12 
         Height          =   255
         Left            =   3240
         Top             =   5400
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         Caption         =   "Tho72i gian que1t:"
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
      Begin UniControls.UniButton cmdSettingFullScan 
         Height          =   255
         Left            =   4920
         TabIndex        =   16
         Top             =   1800
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   450
         Icon            =   "frmMain.frx":AD02
         Style           =   2
         Caption         =   "Ca2i d9a85t ca61u hi2nh que1t"
         IconAlign       =   3
         iNonThemeStyle  =   2
         BackColor       =   -2147483643
         MaskColor       =   16711935
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         RoundedBordersByTheme=   0   'False
         ShowFocusRectangle=   0   'False
      End
      Begin UniControls.UniLabel lblStatus2 
         Height          =   255
         Left            =   240
         Top             =   6360
         Width           =   6615
         _ExtentX        =   11668
         _ExtentY        =   450
         Alignment       =   1
         Caption         =   "Sa84n Sa2ng"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   128
      End
      Begin UniControls.UniLabel lblStatus 
         Height          =   615
         Left            =   120
         Top             =   5640
         Width           =   6855
         _ExtentX        =   12091
         _ExtentY        =   1085
         BackStyle       =   0
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
         ForeColor       =   16711680
         Link            =   ""
      End
      Begin UniControls.UniLabel UniLabel8 
         Height          =   255
         Left            =   120
         Top             =   5400
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   450
         Caption         =   "Ti2nh Tra5ng Que1t:"
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
      Begin UniControls.UniButton cmdFSKillVirus 
         Height          =   375
         Left            =   4440
         TabIndex        =   5
         Top             =   4920
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         Icon            =   "frmMain.frx":AD1E
         Style           =   2
         Caption         =   "Die65t Virus"
         IconAlign       =   3
         iNonThemeStyle  =   2
         BackColor       =   15398133
         MaskColor       =   16711935
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         RoundedBordersByTheme=   0   'False
         ShowFocusRectangle=   0   'False
      End
      Begin UniControls.UniButton cmdFSCachLy 
         Height          =   375
         Left            =   2640
         TabIndex        =   4
         Top             =   4920
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         Icon            =   "frmMain.frx":B2B8
         Style           =   2
         Caption         =   "Ca1ch Ly Virus"
         IconAlign       =   3
         iNonThemeStyle  =   2
         BackColor       =   -2147483643
         MaskColor       =   16711935
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         RoundedBordersByTheme=   0   'False
         ShowFocusRectangle=   0   'False
      End
      Begin UniControls.UniButton cmdFSReport 
         Height          =   375
         Left            =   5880
         TabIndex        =   3
         Top             =   4920
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         Icon            =   "frmMain.frx":BCCA
         Style           =   2
         Caption         =   "Ba1o Ca1o"
         IconAlign       =   3
         iNonThemeStyle  =   2
         Enabled         =   0   'False
         BackColor       =   15398133
         MaskColor       =   16711935
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         RoundedBordersByTheme=   0   'False
         ShowFocusRectangle=   0   'False
      End
      Begin UniControls.UniButton cmdFSStop 
         Height          =   375
         Left            =   1320
         TabIndex        =   2
         Top             =   4920
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         Icon            =   "frmMain.frx":C6DC
         Style           =   2
         Caption         =   "Du72ng"
         IconAlign       =   3
         Enabled         =   0   'False
         BackColor       =   -2147483643
         MaskColor       =   16711935
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         RoundedBordersByTheme=   0   'False
         ShowFocusRectangle=   0   'False
      End
      Begin UniControls.UniButton cmdFSStart 
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   4920
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         Icon            =   "frmMain.frx":CC76
         Style           =   2
         Caption         =   "Que1t"
         IconAlign       =   3
         iNonThemeStyle  =   2
         BackColor       =   -2147483643
         MaskColor       =   16711935
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         RoundedBordersByTheme=   0   'False
         ShowFocusRectangle=   0   'False
      End
      Begin UniControls.UniListView LVVirus1 
         Height          =   2535
         Left            =   120
         TabIndex        =   0
         Top             =   2280
         Width           =   6855
         _ExtentX        =   12091
         _ExtentY        =   4471
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
         ForeColor       =   255
         BackColor       =   12648447
         MultiSelect     =   -1  'True
         LabelEdit       =   0   'False
         AutoArrange     =   0   'False
         HeaderButtons   =   0   'False
         HeaderTrackSelect=   0   'False
         HideSelection   =   0   'False
         InfoTips        =   0   'False
      End
      Begin UniControls.UniLabel UniLabel6 
         Height          =   255
         Left            =   120
         Top             =   1800
         Width           =   6855
         _ExtentX        =   12091
         _ExtentY        =   450
         Caption         =   "Chu1 Y1: Ha4y ca2i d9a85t ca61u hi2nh que1t tru7o71c khi ba81t d9a62u que1t."
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   255
      End
      Begin UniControls.UniLabel UniLabel5 
         Height          =   255
         Left            =   360
         Top             =   1440
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   450
         Caption         =   "- Ca1c Chu7o7ng Tri2nh D9ang Cha5y Trong Bo65 Nho71."
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
      Begin UniControls.UniLabel UniLabel4 
         Height          =   255
         Left            =   360
         Top             =   1200
         Width           =   4575
         _ExtentX        =   8070
         _ExtentY        =   450
         Caption         =   "- Ca1c Chu7o7nh Tri2nh Kho73i D9o65ng Cu2ng He65 Tho61ng."
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
      Begin UniControls.UniLabel UniLabel3 
         Height          =   255
         Left            =   360
         Top             =   960
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   450
         Caption         =   "- Ta61t Ca3 Ca1c O63 D9i4a."
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
         Left            =   120
         Top             =   720
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   450
         Caption         =   "D9i5a Chi3 Que1t:"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   16711680
      End
      Begin UniControls.UniLabel UniLabel1 
         Height          =   495
         Left            =   120
         Top             =   360
         Width           =   6855
         _ExtentX        =   12091
         _ExtentY        =   873
         Alignment       =   1
         BackStyle       =   0
         Caption         =   "Que1t Virus Toa2n Bo65"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   255
      End
      Begin UniControls.UniLabel lblScanTime 
         Height          =   255
         Index           =   1
         Left            =   5040
         Top             =   5400
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   450
         Caption         =   "00"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   255
      End
      Begin UniControls.UniLabel lblScanTime 
         Height          =   255
         Index           =   2
         Left            =   5400
         Top             =   5400
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   450
         Caption         =   "00"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   255
      End
      Begin UniControls.UniLabel UniLabel14 
         Height          =   255
         Left            =   5280
         Top             =   5400
         Width           =   135
         _ExtentX        =   238
         _ExtentY        =   450
         Caption         =   ":"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   255
      End
   End
   Begin UniControls.UniFrame fmMain 
      Height          =   6735
      Left            =   120
      Top             =   1320
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   11880
      MaskColor       =   16711935
      FrameColor      =   -2147483629
      Style           =   0
      Caption         =   "PAV 2009 - Chu71c na8ng"
      TextColor       =   13579779
      Alignment       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin UniControls.UniTreeView TreeChucNang 
         Height          =   6375
         Left            =   120
         TabIndex        =   42
         Top             =   240
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   11245
      End
   End
   Begin VB.Image picLogoPAV 
      Height          =   1335
      Left            =   0
      Picture         =   "frmMain.frx":D688
      Top             =   0
      Width           =   10215
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'<----- API for Search
Private Const vbDot = 46
Private Const MAXDWORD As Long = &HFFFFFFFF
Private Const MAX_PATH As Long = 260
Private Const INVALID_HANDLE_VALUE = -1
Private Const FILE_ATTRIBUTE_DIRECTORY = &H10

Private Type FILETIME
   dwLowDateTime As Long
   dwHighDateTime As Long
End Type

Private Type WIN32_FIND_DATA
   dwFileAttributes As Long
   ftCreationTime As FILETIME
   ftLastAccessTime As FILETIME
   ftLastWriteTime As FILETIME
   nFileSizeHigh As Long
   nFileSizeLow As Long
   dwReserved0 As Long
   dwReserved1 As Long
   cFileName As String * MAX_PATH
   cAlternate As String * 14
End Type

Private Type FILE_PARAMS
   bRecurse As Boolean
   sFileRoot As String
   sFileNameExt As String
   sResult As String
   sMatches As String
   Count As Long
End Type

Private Declare Function FindClose Lib "kernel32" _
  (ByVal hFindFile As Long) As Long
   
Private Declare Function FindFirstFile Lib "kernel32" _
   Alias "FindFirstFileA" _
  (ByVal lpFileName As String, _
   lpFindFileData As WIN32_FIND_DATA) As Long
   
Private Declare Function FindNextFile Lib "kernel32" _
   Alias "FindNextFileA" _
  (ByVal hFindFile As Long, _
   lpFindFileData As WIN32_FIND_DATA) As Long

Private Declare Function GetTickCount Lib "kernel32" () As Long
'-----> API For Search


'-----> Dim for Connect Database
Dim db As Database
Dim rs As Recordset
Dim WS As Workspace
Dim Max As Long
Dim mData As Recordset
'<------ Dim for Connect Database



'---> Dim For FullScan and Custom Scan
Dim xFSStopScan As Boolean
Dim xFSStopScan2 As Boolean
Dim LLVV As UniListView
Dim xTotalFile
Dim HoanThanhFull As Boolean
Dim HoanThanhCus As Boolean
Dim TimeFull As String
Dim TimeCus As String
'<--- Dim For FullScan and Custom Scan

'---> Dim for Timer of Full Scan
Dim xTimeSC
Dim xTime
'<--- Dim for Timer of Full Scan

'---> Check Full scan or Custom Scan
Dim xCustomScan As Boolean
'<--- Check Full scan or Custom Scan

'---> Dim for Process
Public SoLuong
'<--- Dim for Process

'---> Dim for Messenger form
Public xStart
'<--- Dim for Messenger form


Private Function SearchFile(Path, Filename)

   Dim FP As FILE_PARAMS  'holds search parameters
   Dim tstart As Single   'timer var for this routine only
   Dim tend As Single     'timer var for this routine only
   With FP
      .sFileRoot = Path       'start path
      .sFileNameExt = Filename    'file type of interest
      .bRecurse = 1 ' Check1.Value = 1  '1 = recursive search
   End With
   tstart = GetTickCount()
   Call SearchForFiles(FP)
   tend = GetTickCount()

End Function


Private Sub GetFileInformation(FP As FILE_PARAMS)

'On Error Resume Next
DoEvents
   Dim WFD As WIN32_FIND_DATA
   Dim hFile As Long
   Dim sPath As String
   Dim sRoot As String
   Dim sTmp As String
   Dim SKetQua As String
   sRoot = QualifyPath(FP.sFileRoot)
   sPath = sRoot & FP.sFileNameExt
   hFile = FindFirstFile(sPath, WFD)
   If hFile <> INVALID_HANDLE_VALUE Then
      Do
         If Not (WFD.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) = _
                 FILE_ATTRIBUTE_DIRECTORY Then
            FP.Count = FP.Count + 1
            sTmp = TrimNull(WFD.cFileName)
            SKetQua = sRoot & sTmp
If xCustomScan = False And xFSStopScan = True Then Exit Sub
If xCustomScan = True And xFSStopScan2 = True Then Exit Sub

If FileLen(SKetQua) > 0 And FileLen(SKetQua) < (Me.VSLimitSize.Value * 1000000) And modScanVirus.FileExists(SKetQua) = True Then '8350E5A3E24C153DF2275C9F80692773




If xCustomScan = True Then
    cslblStatus.Caption = SKetQua
    If xFSStopScan2 = True Then Exit Sub
Else
    xTotalFile = xTotalFile + 1
    lblStatus.Caption = SKetQua
    
End If

            'SetAttr SKetQua, vbNormal
            
            DoEvents
            Dim AX As String
            AX = modScanVirus.CheckVirus(SKetQua)
            If AX <> "No" Then
            'Virus Found
                'List1.AddItem "!!!!! " & AX & " - " & SKetQua
                
                If xCustomScan = False Then
                    Dim i
                    i = LVVirus1.ListItems.Count + 1
                    LVVirus1.ListItems.Add i, , AX
                    LVVirus1.ListItems(i).SubItems(1).Caption = SKetQua
                    LVVirus1.ListItems(i).SubItems(2).Caption = FileLen(SKetQua) & " Bytes"
                    LVVirus1.ListItems(i).SubItems(3).Caption = CheckProcess(SKetQua)
                    LVVirus1.ListItems(i).SubItems(4).Caption = "---"
                    LVVirus1.ListItems(i).Checked = True
                    
                Else
                    Dim ig
                    ig = LVVirus2.ListItems.Count + 1
                    LVVirus2.ListItems.Add ig, , AX
                    LVVirus2.ListItems(ig).SubItems(1).Caption = SKetQua
                    LVVirus2.ListItems(ig).SubItems(2).Caption = FileLen(SKetQua) & " Bytes"
                    LVVirus2.ListItems(ig).SubItems(3).Caption = CheckProcess(SKetQua)
                    LVVirus2.ListItems(ig).SubItems(4).Caption = "---"
                    
                    LVVirus2.ListItems(ig).Checked = True
                End If
                
                
                
                
            End If ' AX="No"
End If 'FileLen(SKetQua) > 0 And FileLen(SKetQua) < 1000000
    

            'Text1.Text = Text1.Text & SKetQua & vbCrLf
            '*********************************************
            
         End If
      Loop While FindNextFile(hFile, WFD)
      hFile = FindClose(hFile)
   End If
DoEvents
End Sub


Private Sub SearchForFiles(FP As FILE_PARAMS)
  'local working variables
   Dim WFD As WIN32_FIND_DATA
   Dim hFile As Long
   Dim sPath As String
   Dim sRoot As String
   Dim sTmp As String
   sRoot = QualifyPath(FP.sFileRoot)
   sPath = sRoot & "*.*"
   hFile = FindFirstFile(sPath, WFD)
   If hFile <> INVALID_HANDLE_VALUE Then
      Call GetFileInformation(FP)
      Do
         If (WFD.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) Then
            If FP.bRecurse Then
               If Asc(WFD.cFileName) <> vbDot Then
                  FP.sFileRoot = sRoot & TrimNull(WFD.cFileName)
                  Call SearchForFiles(FP)
               End If
            End If
         End If
      Loop While FindNextFile(hFile, WFD)
      hFile = FindClose(hFile)
   End If
End Sub
Private Function TrimNull(startstr As String) As String
   Dim pos As Integer
   pos = InStr(startstr, Chr$(0))
   If pos Then
      TrimNull = Left$(startstr, pos - 1)
      Exit Function
   End If
   TrimNull = startstr
End Function
Private Function QualifyPath(sPath As String) As String
   If Right$(sPath, 1) <> "\" Then
      QualifyPath = sPath & "\"
   Else
      QualifyPath = sPath
   End If
End Function
Public Sub SaveREG()
If FileExists(AppPath & "RegProtect.dat") = True Then modScanVirus.DeleteFile (AppPath & "RegProtect.dat")
With frmMain.atpLVREG
    Dim r
    For r = 1 To .ListItems.Count
        WriteIniFile AppPath & "RegProtect.dat", r, "TenChucNang", Unicode2UTF8(.ListItems(r).Text)
        WriteIniFile AppPath & "RegProtect.dat", r, "KeyGoc", .ListItems(r).SubItems(1).Caption
        WriteIniFile AppPath & "RegProtect.dat", r, "KeyPath", .ListItems(r).SubItems(2).Caption
        WriteIniFile AppPath & "RegProtect.dat", r, "KeyName", .ListItems(r).SubItems(3).Caption
        WriteIniFile AppPath & "RegProtect.dat", r, "KeyData", .ListItems(r).SubItems(4).Caption
    Next r
    WriteIniFile AppPath & "RegProtect.dat", "Other", "Total", .ListItems.Count
    
End With

End Sub

Private Sub atpcmdAutorun_Click()
If Me.atptmrAutorun.Enabled = True Then
    Me.atptmrAutorun.Enabled = False
    Me.atplblStaAutorun.Caption = "D9ang ta81t"
    Me.atpcmdAutorun.Caption = "Mo73 chu71c na8ng na2y"
Else
    Me.atptmrAutorun.Enabled = True
    Me.atplblStaAutorun.Caption = "D9ang mo73"
    Me.atpcmdAutorun.Caption = "Ta81t chu71c na8ng na2y"
End If
End Sub

Private Sub atpcmdREG_Click()
If Me.atptmrREG.Enabled = True Then
    Me.atptmrREG.Enabled = False
    Me.atplbREG.Caption = "D9ang ta81t"
    Me.atpcmdREG.Caption = "Mo73 chu71c na8ng na2y"
Else
    Me.atptmrREG.Enabled = True
    Me.atplbREG.Caption = "D9ang mo73"
    Me.atpcmdREG.Caption = "Ta81t chu71c na8ng na2y"
End If
End Sub

Private Sub atpREGAdd_Click()
pfrmAddREG.Show 1, Me
End Sub

Private Sub atptmrAutorun_Timer()
    DoEvents
    'On Error Resume Next
    Dim Str
    Dim str2
    Dim FSO  As New FileSystemObject
    Dim drv  As Drive
    Dim drvs As Drives
    DoEvents
    Set drvs = FSO.Drives
    For Each drv In drvs
        If UCase(drv.DriveLetter) <> "A" Then
            If FileExists(drv.DriveLetter & ":\autorun.inf") = True Then
                'Code Here
                frmMessenger.zShowMessenger "Pha1t hie65n Autorun!", "Pha1t hie65n ta65p tin tu75 cha5y (Autorun.inf) ta5i o63 d9i4a [" & drv.DriveLetter & ":\] Chu7o7ng tri2nh se4 xo1a no1 ra kho3i he65 tho61ng va2 thie61t la65p ba3o ve65 cho o63 d9i4a na2y ngay ba6y gio72.)", 5000, xvang

                '/////// Tim Nguon Goc Cua Virus ////////
                Dim Xx1        As String
                Dim Xx2        As String
                Dim xFileName1 As String
                Dim xFileName2 As String
                DoEvents
                Xx1 = drv.DriveLetter & ":\" & GetOpenAutorun(drv.DriveLetter & ":\autorun.inf")
                Xx2 = drv.DriveLetter & ":\" & GetShellOpenAutorun(drv.DriveLetter & ":\autorun.inf")
                If Xx1 <> drv.DriveLetter & ":\" And modScanVirus.FileExists(Xx1) = True Then
                    xFileName1 = GetFileName(Xx1)
                    If CheckProcess(xFileName1) <> 0 Then KillProcessById CheckProcess(xFileName1) 'EndTask xFileName1
                    SetAttr Xx1, vbNormal
                    modScanVirus.DeleteFile Xx1
                    frmMessenger.zShowMessenger "Pha1t hie65n Virus!", "D9a4 pha1t hie65n Virus ta5i: [" & drv.DriveLetter & ":\...\" & GetFileName(Xx1) & "]. Ti2nh tra5ng: D9a4 Xo1a", 5000, xvang
                End If

                If Xx2 <> drv.DriveLetter & ":\" And modScanVirus.FileExists(Xx2) = True Then
                    xFileName2 = GetFileName(Xx2)
                    If CheckProcess(xFileName2) <> 0 Then KillProcessById CheckProcess(xFileName2) 'EndTask xFileName2
                    SetAttr Xx2, vbNormal
                    modScanVirus.DeleteFile Xx2
                    frmMessenger.zShowMessenger "Pha1t hie65n Virus!", "D9a4 pha1t hie65n Virus ta5i: [" & drv.DriveLetter & ":\...\" & GetFileName(Xx2) & "]. Ti2nh tra5ng: D9a4 Xo1a", 5000, xvang
                End If

                '////////// End / Tim nguon goc Virus ///////////

                '////// Diet Autorun ////////
                DoEvents
                SetAttr drv.DriveLetter & ":\autorun.inf", vbNormal
                modScanVirus.DeleteFile drv.DriveLetter & ":\autorun.inf"
                MkDir drv.DriveLetter & ":\autorun.inf"
                Str = "cmd /c md \\?\" & drv.DriveLetter & ":\autorun.inf\.PAV.2009."
                Shell Str, vbHide
                SetAttr drv.DriveLetter & ":\autorun.inf", vbHidden + vbReadOnly + vbSystem
                FileCopy AppPath & "PAV2009.ico", drv.DriveLetter & ":\autorun.inf\Icon.ico"
                WriteFileUni drv.DriveLetter & ":\autorun.inf\ThongTin.txt", ToUnicode("Thu7 mu5c na2y la2 thu7 mu5c Autorun gia3, d9u7o75c ta5o ra d9e63 d9a1nh lu72a Virus, nha82m nga8n Virus la6y qua USB." & vbCrLf & "File na2y d9u7o75c ta5o bo73i chu7o7ng tri2nh PAV 2009." & vbCrLf & "Phát ha2nh bo73i: http://qts.come.vn") 'CreateTextFile drv.DriveLetter & ":\autorun.inf\AlwaysProtected.txt", "
                WriteFileUni drv.DriveLetter & ":\autorun.inf\desktop.ini", "[.ShellClassInfo]" & vbCrLf & "IconFile=" & drv.DriveLetter & ":\autorun.inf\Icon.ico" & vbCrLf & "IconIndex = 0"
                SetAttr drv.DriveLetter & ":\autorun.inf\desktop.ini", vbHidden + vbSystem + vbReadOnly
                DoEvents
                '////// End Diet Autorun ////////
            End If

            'Kill drv.DriveLetter & ":\autorun.inf"
        End If
    
    Next

    Set FSO = Nothing
    Set drv = Nothing
    Set drvs = Nothing

    DoEvents
End Sub

Private Sub atptmrREG_Timer()

With frmMain.atpLVREG
    Dim y
    Dim u
    Dim KQ As String
        For y = 1 To .ListItems.Count
            If .ListItems(y).SubItems(1).Caption = "HKEY_CLASSES_ROOT" Then
                    u = &H80000000
                ElseIf .ListItems(y).SubItems(1).Caption = "HKEY_CURRENT_USER" Then
                    u = &H80000001
                ElseIf .ListItems(y).SubItems(1).Caption = "HKEY_LOCAL_MACHINE" Then
                    u = &H80000002
                ElseIf .ListItems(y).SubItems(1).Caption = "HKEY_USERS" Then
                    u = &H80000003
                ElseIf .ListItems(y).SubItems(1).Caption = "HKEY_CURRENT_CONFIG" Then
                    u = &H80000005
            End If
            On Error GoTo ThOaTkHoIvOnGfOr
            If (GetString(u, .ListItems(y).SubItems(2).Caption, .ListItems(y).SubItems(3).Caption) <> .ListItems(y).SubItems(4).Caption) And (GetString(u, .ListItems(y).SubItems(2).Caption, .ListItems(y).SubItems(3).Caption)) <> "" Then
                If UCase(.ListItems(y).SubItems(3).Caption) = "DISABLETASKMGR" Then
                    DeleteValue u, .ListItems(y).SubItems(2).Caption, .ListItems(y).SubItems(3).Caption
                Else
                    SaveString u, .ListItems(y).SubItems(2).Caption, .ListItems(y).SubItems(3).Caption, .ListItems(y).SubItems(4).Caption
                End If
                KQ = .ListItems(y).Text
                GoTo EndATP
            ElseIf (GetDWORD(u, .ListItems(y).SubItems(2).Caption, .ListItems(y).SubItems(3).Caption) <> .ListItems(y).SubItems(4).Caption) Then
                SaveDWORD u, .ListItems(y).SubItems(2).Caption, .ListItems(y).SubItems(3).Caption, .ListItems(y).SubItems(4).Caption
                KQ = .ListItems(y).Text
                GoTo EndATP
            End If
ThOaTkHoIvOnGfOr:
        Next y
End With


If GetString(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon", "Userinit") <> "C:\WINDOWS\system32\userinit.exe," Then
    SaveString HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon", "Userinit", "C:\WINDOWS\system32\userinit.exe,"
    frmMessenger.zShowMessenger "Ba3o ve65 Registry", "Kho1a he65 tho61ng [Winlogon] d9a4 bi5 thay d9o63i so vo71i gia1 tri5 ma85c d9i5nh cu3a no1, d9ie62u na2y co1 the63 ga6y ha5i cho ma1y ti1nh cu3a ba5n. Chu7o7ng tri2nh se4 phu5c ho62i kho1a na2y ve65 gia1 tri5 ban d9a62u ngay ba6y gio72. Tra5ng tha1i: D9a4 phu5c ho62i.", 5000, xTrang
End If

If GetString(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon", "Shell") <> "Explorer.exe" Then
    SaveString HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon", "Shell", "Explorer.exe"
    frmMessenger.zShowMessenger "Ba3o ve65 Registry", "Kho1a he65 tho61ng [Shell] d9a4 bi5 thay d9o63i so vo71i gia1 tri5 ma85c d9i5nh cu3a no1, d9ie62u na2y co1 the63 ga6y ha5i cho ma1y ti1nh cu3a ba5n. Chu7o7ng tri2nh se4 phu5c ho62i kho1a na2y ve65 gia1 tri5 ban d9a62u ngay ba6y gio72. Tra5ng tha1i: D9a4 phu5c ho62i.", 5000, xTrang
End If

Exit Sub
EndATP:
    frmMessenger.zShowMessenger "Ba3o ve65 Registry", "Chu71c na8ng [" & KQ & "] cu3a ma1y ti1nh ba5n d9ang bi5 kho1a, chu7o7ng tri2nh se4 mo73 kho1a cho no1 ngay ba6y gio72. Tra5ng tha1i: D9a4 mo73 kho1a.", 4000, xTrang

End Sub

Private Sub atsAlwaysScanFolder_Click()
With atsAlwaysScanFolder
    If .Value = True Then
        .Caption = "Tu75 d9o65ng que1t va2 ca3nh ba1o Virus trong Thu7 mu5c d9ang d9u7o75c mo73. [D9ang Ba65t]"
        WriteIniFile AppPath & "Setting.ini", "AutoScan", "ScanFolder", .Value
        Load zfrmAutoScanFolder
    Else
        .Caption = "Tu75 d9o65ng que1t va2 ca3nh ba1o Virus trong Thu7 mu5c d9ang d9u7o75c mo73. [D9ang Ta81t]"
        WriteIniFile AppPath & "Setting.ini", "AutoScan", "ScanFolder", .Value
        Unload zfrmAutoScanFolder
    
    End If
End With
End Sub

Private Sub atsAutoScanUSB_Click()
With atsAutoScanUSB
    If .Value = True Then
        .Caption = "Tu75 d9o65ng que1t Virus cho USB khi pha1t hie65n USB ke61t no61i va2o ma1y ti1nh. [D9ang Ba65t]"
        WriteIniFile AppPath & "Setting.ini", "AutoScan", "ScanUSB", .Value
        '////
        Load zfrmScanUSB
        '///
    Else
        .Caption = "Tu75 d9o65ng que1t Virus cho USB khi pha1t hie65n USB ke61t no61i va2o ma1y ti1nh. [D9ang Ta81t]"
        WriteIniFile AppPath & "Setting.ini", "AutoScan", "ScanUSB", .Value
        '///
        Unload zfrmScanUSB
        '///
    
    End If
End With
End Sub

Private Sub atsScanEXE_Click()
With atsScanEXE
    If .Value = True Then
        .Caption = "Tu75 d9o65ng que1t va2 ca3nh ba1o Virus cho nhu74ng File sa81p d9u7o75c mo73. [D9ang Ba65t]"
        WriteIniFile AppPath & "Setting.ini", "AutoScan", "ScanEXE", .Value
        SaveString HKEY_CLASSES_ROOT, "exefile\shell\open\command", "", ChrW(34) & AppPath & "Check_Virus.exe" & ChrW(34) & " " & ChrW(34) & "%1" & ChrW(34)
        SaveString HKEY_CLASSES_ROOT, "batfile\shell\open\command", "", ChrW(34) & AppPath & "Check_Virus.exe" & ChrW(34) & " " & ChrW(34) & "%1" & ChrW(34)
        SaveString HKEY_CLASSES_ROOT, "cmdfile\shell\open\command", "", ChrW(34) & AppPath & "Check_Virus.exe" & ChrW(34) & " " & ChrW(34) & "%1" & ChrW(34)
        SaveString HKEY_CLASSES_ROOT, "comfile\shell\open\command", "", ChrW(34) & AppPath & "Check_Virus.exe" & ChrW(34) & " " & ChrW(34) & "%1" & ChrW(34)
    
    Else
    
        .Caption = "Tu75 d9o65ng que1t va2 ca3nh ba1o Virus cho nhu74ng File sa81p d9u7o75c mo73. [D9ang Ta81t]"
        WriteIniFile AppPath & "Setting.ini", "AutoScan", "ScanEXE", .Value
        SaveString HKEY_CLASSES_ROOT, "exefile\shell\open\command", "", ChrW(34) & "%1" & ChrW(34) & " %*"
        SaveString HKEY_CLASSES_ROOT, "batfile\shell\open\command", "", ChrW(34) & "%1" & ChrW(34) & " %*"
        SaveString HKEY_CLASSES_ROOT, "cmdfile\shell\open\command", "", ChrW(34) & "%1" & ChrW(34) & " %*"
        SaveString HKEY_CLASSES_ROOT, "comfile\shell\open\command", "", ChrW(34) & "%1" & ChrW(34) & " %*"
    
    End If
End With
End Sub

Private Sub atsScanKeylogger_Click()
With atsScanKeylogger
    If .Value = True Then
        .Caption = "Tu75 d9o65ng pha1t hie65n va2 ca3nh ba1o Keylogger. [D9ang Ba65t]"
        WriteIniFile AppPath & "Setting.ini", "AutoScan", "ScanKeylogger", .Value
        '////
        Load zfrmAntiKey
        '///
    Else
        .Caption = "Tu75 d9o65ng pha1t hie65n va2 ca3nh ba1o Keylogger. [D9ang Ta81t]"
        WriteIniFile AppPath & "Setting.ini", "AutoScan", "ScanKeylogger", .Value
        '///
        Unload zfrmAntiKey
        '///
    
    End If
End With
End Sub



Private Sub cmdDelREG_Click()
DelAllChecked Me.atpLVREG
SaveREG
End Sub

Private Sub cmdDelSelected_Click()
If UniMsgBox("Ba5n cha81c cha81n?", vbYesNo) = vbYes Then
DelAllChecked LVVirusEvents
If FileExists(AppPath & "VirusScanLog.log") = True Then modScanVirus.DeleteFile (AppPath & "VirusScanLog.log")
With frmMain.LVVirusEvents
    Dim l
    For l = 1 To .ListItems.Count
        WriteIniFile AppPath & "VirusScanLog.log", l, "TimeQuet", .ListItems(l).Text
        WriteIniFile AppPath & "VirusScanLog.log", l, "KieuQuet", .ListItems(l).SubItems(1).Caption
        WriteIniFile AppPath & "VirusScanLog.log", l, "SoFile", .ListItems(l).SubItems(2).Caption
        WriteIniFile AppPath & "VirusScanLog.log", l, "SoVirus", .ListItems(l).SubItems(3).Caption
        WriteIniFile AppPath & "VirusScanLog.log", l, "KetQua", .ListItems(l).SubItems(4).Caption
        WriteIniFile AppPath & "VirusScanLog.log", "Other", "Total", .ListItems.Count
    Next l
End With

End If
End Sub

Private Sub cmdEventsVirusKillAll_Click()
If UniMsgBox("Ba5n cha81c cha81n muo61n xo1a ta61t ca3?", vbYesNo) = vbYes Then
DelAllLV LVVirusEvents
modScanVirus.DeleteFile AppPath & "RegProtect.dat"
End If
End Sub



Private Sub cmdFSCachLy_Click()
If UniMsgBox("Ba5n co1 muo61n ca1ch ly ca1c Virus d9a4 cho5n kho6ng?", vbYesNo, "Tho6ng Ba1o") = vbYes Then
Dim y, j
Dim x As String
For y = 1 To LVVirus1.ListItems.Count
    If LVVirus1.ListItems(y).Checked = True And FileExists(LVVirus1.ListItems(y).SubItems(1).Caption) = True Then
    
        If LVVirus1.ListItems(y).SubItems(3).Caption <> "0" Then
        'Kill process
            KillProcessById (LVVirus1.ListItems(y).SubItems(3).Caption)
            x = x & " D9a4 ta81t tie61n tri2nh: " & LVVirus1.ListItems(y).SubItems(3).Caption & vbCrLf
        End If
    
        SetAttr LVVirus1.ListItems(y).SubItems(1).Caption, vbNormal
        Name LVVirus1.ListItems(y).SubItems(1).Caption As LVVirus1.ListItems(y).SubItems(1).Caption & ".DaCachLy"
        FileCopy LVVirus1.ListItems(y).SubItems(1).Caption & ".DaCachLy", AppPath & "VungCachLy\" & GetFileName(LVVirus1.ListItems(y).SubItems(1).Caption & ".DaCachLy")
        modScanVirus.DeleteFile LVVirus1.ListItems(y).SubItems(1).Caption & ".DaCachLy"
        x = x & " D9a4 Ca1ch Ly: " & LVVirus1.ListItems(y).SubItems(1).Caption & vbCrLf
    End If
Next y
If Not x = "" Then
DelAllChecked LVVirus1
UniMsgBox x, vbOKOnly, "D9a4 Ca1ch Ly Virus", Me.hWnd
Else
UniMsgBox "Kho6ng co1 Virus na2o d9e63 ca1ch ly.", vbOKOnly, "Tho6ng Ba1o", Me.hWnd
End If

End If 'unimsgbox
End Sub

Private Sub cmdFSKillVirus_Click()
If UniMsgBox("Ba5n co1 cha81c cha81n muo61n xo1a ca1c Virus d9a4 d9a1nh da61u kho6ng?", vbYesNo, "Die65t Virus") = vbYes Then
DoEvents
lblStatus.Caption = "D9ang xo1a..."
Dim y
Dim x As String
For y = 1 To LVVirus1.ListItems.Count
    If LVVirus1.ListItems(y).Checked = True And FileExists(LVVirus1.ListItems(y).SubItems(1).Caption) = True Then
    DoEvents
        If LVVirus1.ListItems(y).SubItems(3).Caption <> "0" Then
        'Kill process
            KillProcessById (LVVirus1.ListItems(y).SubItems(3).Caption)
            x = x & " D9a4 ta81t tie61n tri2nh: " & LVVirus1.ListItems(y).SubItems(3).Caption & vbCrLf
        End If
    DoEvents
        If LVVirus1.ListItems(y).SubItems(4).Caption <> "---" Then
        'Delete Registry Key
            Dim a
            Dim u
            Dim b
            Dim c
            DoEvents
            a = Left(LVVirus1.ListItems(y).SubItems(4).Caption, Len(LVVirus1.ListItems(y).SubItems(4).Caption) - InStrRev(StrReverse(LVVirus1.ListItems(y).SubItems(4).Caption), "-"))
            b = Right(LVVirus1.ListItems(y).SubItems(4).Caption, Len(LVVirus1.ListItems(y).SubItems(4).Caption) - InStrRev(LVVirus1.ListItems(y).SubItems(4).Caption, ":"))
            c = Mid(LVVirus1.ListItems(y).SubItems(4).Caption, Len(LVVirus1.ListItems(y).SubItems(4).Caption) - InStrRev(StrReverse(LVVirus1.ListItems(y).SubItems(4).Caption), "-") + 2, (InStrRev(LVVirus1.ListItems(y).SubItems(4).Caption, ":")) - (Len(LVVirus1.ListItems(y).SubItems(4).Caption) - InStrRev(StrReverse(LVVirus1.ListItems(y).SubItems(4).Caption), "-")) - 2)
            If UCase(a) = "HKEY_CURRENT_USER" Then u = &H80000001
            If UCase(a) = "HKEY_LOCAL_MACHINE" Then u = &H80000002
            DeleteValue u, c, b
        x = x & " D9a4 xo1a Key: " & a & "\" & c & ":" & b & vbCrLf
        End If
        
        'HKEY_CURRENT_USER = &H80000001
        'HKEY_LOCAL_MACHINE = &H80000002
        DoEvents
        SetAttr LVVirus1.ListItems(y).SubItems(1).Caption, vbNormal
        modScanVirus.DeleteFile LVVirus1.ListItems(y).SubItems(1).Caption
        x = x & " D9a4 Xo1a Bo3: " & LVVirus1.ListItems(y).SubItems(1).Caption & vbCrLf
        
        End If
Next y

If Not x = "" Then
DelAllChecked LVVirus1
UniMsgBox x, vbOKOnly, "Nhu74ng vie65c d9a4 la2m.", Me.hWnd

Else
UniMsgBox "Kho6ng co1 Virus na2o d9e63 die65t!", vbOKOnly, "Tho6ng Ba1o", Me.hWnd
End If
lblStatus.Caption = "Sa84n sa2ng"
End If ' Unimsgbox "ban co chac chan ko?"
End Sub

Private Sub cmdFSReport_Click()
frmVirusReport.ShowReport xTotalFile, modLietKeValue.xTotalStartUp, modScanVirus.xTotalProcess, LVVirus1.ListItems.Count, IIf(HoanThanhFull, "D9a4 hoa2n tha2nh", "Chu7a hoa2n tha2nh"), Me.lblScanTime(0).Caption & ":" & Me.lblScanTime(1).Caption & ":" & Me.lblScanTime(2).Caption
End Sub

Private Sub cmdFSStart_Click()

DelAllLV LVVirus1

xTimeSC = 0

HoanThanhFull = True
xFSStopScan = False
cmdFSCachLy.Enabled = False
cmdFSKillVirus.Enabled = False
cmdFSReport.Enabled = False
cmdFSStop.Enabled = True
cmdFSStart.Enabled = False
cmdFSCachLy.Enabled = False
fmMain.Enabled = False
cmdSettingFullScan.Enabled = False
TimeFull = Time & " - " & Day(Date) & "/" & Month(Date) & "/" & Year(Date)

Tray1.ToolTipText = "D9ang que1t toa2n bo65 he65 tho61ng..."


xTotalFile = 0

lblStatus2.Caption = "D9ang Que1t..."



tmrStartFullScan.Enabled = True
lblStatus.Caption = "Chua63n bi5 que1t..."
lblStatus2.Caption = "Vui lo2ng kho6ng cha5y the6m u71ng du5ng na2o trong tho72i gian que1t..."

xTime = 0
tmrScanTime.Enabled = True
End Sub

Private Sub cmdFSStop_Click()
If UniMsgBox("Chu7o7ng tri2nh d9ang que1t Virus, ba5n co1 cha81c cha81n muo61n du72ng la5i kho6ng?", vbYesNo, "Tho6ng Ba1o") = vbYes Then
cmdFSStop.Enabled = False
xFSStopScan = True
HoanThanhFull = False
End If
End Sub

Private Sub cmdFullScan_Click()
HideAllFM
fm(0).Visible = True
End Sub

Private Sub cmdQuetVirus_Click()
    If VirusDuoiRa = False Then
    If ProtectDuoiRa = True Then
        ProtectDuoiRa = False
        tmrProtect.Enabled = True
    End If
    If TienIchDuoiRa = True Then
        TienIchDuoiRa = False
        tmrTienIch.Enabled = True
    End If
    End If

VirusDuoiRa = Not VirusDuoiRa
tmrQuetVirus.Enabled = True


End Sub



Private Sub cmdSettingFullScan_Click()
HideAllFM
fm(3).Visible = True
End Sub



Private Sub HideAllFM()
    Dim i
    For i = 0 To fm.Count - 1
        fm(i).Visible = False
        fm(i).Left = 3000
        fm(i).Top = 1320
    Next i
End Sub


Private Sub csBack_Click()
ff.Visible = True
csBack.Enabled = False
csCachLy.Enabled = False
csKill.Enabled = False
End Sub

Private Sub CSbat_Click()
VSbat.Value = CSbat.Value
End Sub

Private Sub csCachLy_Click()
If UniMsgBox("Ba5n co1 muo61n ca1ch ly ca1c Virus d9a4 cho5n kho6ng?", vbYesNo, "Tho6ng Ba1o") = vbYes Then
Dim y, j
Dim x As String
For y = 1 To LVVirus2.ListItems.Count
    If LVVirus2.ListItems(y).Checked = True And FileExists(LVVirus2.ListItems(y).SubItems(1).Caption) = True Then
    
        If LVVirus2.ListItems(y).SubItems(3).Caption <> "0" Then
        'Kill process
            KillProcessById (LVVirus2.ListItems(y).SubItems(3).Caption)
            x = x & " D9a4 ta81t tie61n tri2nh: " & LVVirus2.ListItems(y).SubItems(3).Caption & vbCrLf
        End If

        Set fss = Nothing
        SetAttr LVVirus2.ListItems(y).SubItems(1).Caption, vbNormal
        Name LVVirus2.ListItems(y).SubItems(1).Caption As LVVirus2.ListItems(y).SubItems(1).Caption & ".DaCachLy"
        FileCopy LVVirus2.ListItems(y).SubItems(1).Caption & ".DaCachLy", AppPath & "VungCachLy\" & GetFileName(LVVirus2.ListItems(y).SubItems(1).Caption & ".DaCachLy")
        modScanVirus.DeleteFile LVVirus2.ListItems(y).SubItems(1).Caption & ".DaCachLy"
        x = x & " D9a4 Ca1ch Ly: " & LVVirus2.ListItems(y).SubItems(1).Caption & vbCrLf
    End If
Next y
If Not x = "" Then
DelAllChecked LVVirus2
UniMsgBox x, vbOKOnly, "D9a4 Ca1ch Ly Virus", Me.hWnd
Else
UniMsgBox "Kho6ng co1 Virus na2o d9e63 ca1ch ly.", vbOKOnly, "Tho6ng Ba1o", Me.hWnd
End If
End If
End Sub

Private Sub CScom_Click()
VScom.Value = CScom.Value
End Sub


Private Sub CSFolderView1_ChangeAfter(ByVal OldPath As String)
cslblPath.Caption = CSFolderView1.Path
End Sub

Private Sub csKill_Click()
If UniMsgBox("Ba5n co1 cha81c cha81n muo61n xo1a ca1c Virus d9a4 d9a1nh da61u kho6ng?", vbYesNo, "Die65t Virus") = vbYes Then
DoEvents
Dim y
Dim x As String
For y = 1 To LVVirus2.ListItems.Count
    If LVVirus2.ListItems(y).Checked = True And FileExists(LVVirus2.ListItems(y).SubItems(1).Caption) = True Then
    DoEvents
        If LVVirus2.ListItems(y).SubItems(3).Caption <> "0" Then
        'Kill process
            KillProcessById (LVVirus2.ListItems(y).SubItems(3).Caption)
            x = x & "D9a4 ta81t tie61n tri2nh: " & LVVirus2.ListItems(y).SubItems(3).Caption & vbCrLf
        End If
        DoEvents
        If LVVirus2.ListItems(y).SubItems(4).Caption <> "---" Then
        'Delete Registry Key
            Dim a
            Dim u
            Dim b
            Dim c
            DoEvents
            a = Left(LVVirus2.ListItems(y).SubItems(4).Caption, Len(LVVirus2.ListItems(y).SubItems(4).Caption) - InStrRev(StrReverse(LVVirus2.ListItems(y).SubItems(4).Caption), "-"))
            b = Right(LVVirus2.ListItems(y).SubItems(4).Caption, Len(LVVirus2.ListItems(y).SubItems(4).Caption) - InStrRev(LVVirus2.ListItems(y).SubItems(4).Caption, ":"))
            c = Mid(LVVirus2.ListItems(y).SubItems(4).Caption, Len(LVVirus2.ListItems(y).SubItems(4).Caption) - InStrRev(StrReverse(LVVirus2.ListItems(y).SubItems(4).Caption), "-") + 2, (InStrRev(LVVirus2.ListItems(y).SubItems(4).Caption, ":")) - (Len(LVVirus2.ListItems(y).SubItems(4).Caption) - InStrRev(StrReverse(LVVirus2.ListItems(y).SubItems(4).Caption), "-")) - 2)
            If UCase(a) = "HKEY_CURRENT_USER" Then u = &H80000001
            If UCase(a) = "HKEY_LOCAL_MACHINE" Then u = &H80000002
            DoEvents
            DeleteValue u, c, b
        x = x & "D9a4 xo1a Key: " & a & "\" & c & ":" & b & vbCrLf
        End If
        
        'HKEY_CURRENT_USER = &H80000001
        'HKEY_LOCAL_MACHINE = &H80000002
            DoEvents
        SetAttr LVVirus2.ListItems(y).SubItems(1).Caption, vbNormal
        modScanVirus.DeleteFile LVVirus2.ListItems(y).SubItems(1).Caption
        x = x & "D9a4 Xo1a Bo3: " & LVVirus2.ListItems(y).SubItems(1).Caption & vbCrLf
        
        End If
Next y

If Not x = "" Then
DelAllChecked LVVirus2
UniMsgBox x, vbOKOnly, "Nhu74ng vie65c d9a4 la2m.", Me.hWnd

Else
UniMsgBox "Kho6ng co1 Virus na2o d9e63 die65t.", vbOKOnly, "Tho6ng Ba1o", Me.hWnd
End If

End If ' Unim
End Sub

Private Sub CSprocess_Click()
Me.VSScanProcess.Value = CSprocess.Value
End Sub

Private Sub csStart_Click()
If cslblPath.Caption = "" Then
    UniMsgBox "Ba5n chu7a cho5n no7i d9e63 que1t!", vbOKOnly, "Tho6ng ba1o"
    Exit Sub
End If

TimeCus = Time & " - " & Day(Date) & "/" & Month(Date) & "/" & Year(Date)
ff.Visible = False
xFSStopScan2 = False
Me.csStart.Enabled = False
Me.csStop.Enabled = True
Me.csCachLy.Enabled = False
Me.csKill.Enabled = False
csBack.Enabled = False
fmMain.Enabled = False
cslblStatus.AutoUnicode = False
HoanThanhCus = True
DelAllLV LVVirus2




If Me.VSScanProcess.Value = True Then
    cslblStatus2.Caption = "D9ang que1t ca1c chu7o7ng tri2nh cha5y trong bo65 nho71..."
    modScanVirus.xScanProcess2
End If

If Me.VSScanStartUp.Value = True Then
    cslblStatus2.Caption = "D9ang que1t ca1c chu7o7ng tri2nh kho73i d9o65ng..."

    cslblStatus.Caption = "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Run"
    GetKeyValue2 "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Run"

    cslblStatus.Caption = "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\RunOnce"
    GetKeyValue2 "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\RunOnce"

    cslblStatus.Caption = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Run"
    GetKeyValue2 "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Run"

    cslblStatus.Caption = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\RunOnce"
    GetKeyValue2 "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\RunOnce"

    cslblStatus.Caption = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\RunOnceEx"
    GetKeyValue2 "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\RunOnceEx"
End If


xCustomScan = True
cslblStatus2.Caption = "D9ang que1t ca1c File..."

SearchFile cslblPath.Caption, "*.exe"
If VSbat.Value = True Then SearchFile cslblPath.Caption, "*.bat"
If VScom.Value = True Then SearchFile cslblPath.Caption, "*.com"
xCustomScan = False


cslblStatus.AutoUnicode = True

cslblStatus2.Caption = "Sa84n Sa2ng"

cslblStatus.Caption = "D9a4 Que1t Xong! Ti2m Tha61y: " & LVVirus2.ListItems.Count & " Ta65p Tin Virus"

Me.csStart.Enabled = True
Me.csStop.Enabled = False
Me.csCachLy.Enabled = True
Me.csKill.Enabled = True
csBack.Enabled = True
fmMain.Enabled = True
PLaySound AppPath & "Sound\ScanDone.wav"
With LVVirusEvents
Dim i
i = .ListItems.Count + 1
    .ListItems.Add i, , TimeCus
    .ListItems(i).SubItems(1).Caption = ToUnicode("Que1t tu2y cho5n (" & Me.cslblPath.Caption & ")")
    .ListItems(i).SubItems(2).Caption = "-"
    .ListItems(i).SubItems(3).Caption = LVVirus2.ListItems.Count
    .ListItems(i).SubItems(4).Caption = IIf(HoanThanhCus, ToUnicode("Hoa2n tha2nh"), ToUnicode("Chu7a hoa2n tha2nh"))
    .AutoUnicode = True
End With



End Sub

Private Sub CSStartUp_Click()
Me.VSScanStartUp.Value = CSStartUp.Value
End Sub

Private Sub csStop_Click()
If UniMsgBox("Chu7o7ng tri2nh d9ang que1t Virus, ba5n co1 cha81c cha81n muo61n du72ng la5i kho6ng?", vbYesNo, "Tho6ng Ba1o") = vbYes Then
csStop.Enabled = False
xFSStopScan2 = True
HoanThanhCus = False


End If
End Sub

Private Sub Form_Load()
'---> Check File Is Running [First of first]
If App.PrevInstance = True Then End
'<--- Check File Is Running [First of first]


'---> Check Exists Database [First]
If FileExists(AppPath & "Data.pav") = False Or FileExists(AppPath & "Check_Virus.exe") = False Or FileExists(AppPath & "Data.str") = False Or FileExists(AppPath & "RegProtect.dat") = False Then
    UniMsgBox "Kho6ng ti2m tha61y nhu74ng File ca62n thie61t cu3a chu7o7ng tri2nh, chu7o7ng tri2nh kho6ng the63 cha5y ba6y gio72.", vbOKOnly, "Kho6ng ti2m tha61y CSDL"
    End
End If

On Error GoTo HeHeChOqUa
MkDir AppPath & "VungCachLy\"
HeHeChOqUa:
'<--- Check Exists Database [First]




'---> Load Setting [Second]
Me.VSbat.Value = ReadIniFile(AppPath & "Setting.ini", "ScanVirus", Me.VSbat.Name, True)
Me.VScmd.Value = ReadIniFile(AppPath & "Setting.ini", "ScanVirus", Me.VScmd.Name, True)
Me.VScom.Value = ReadIniFile(AppPath & "Setting.ini", "ScanVirus", Me.VScom.Name, True)
Me.VSdll.Value = ReadIniFile(AppPath & "Setting.ini", "ScanVirus", Me.VSdll.Name, True)
Me.VSscr.Value = ReadIniFile(AppPath & "Setting.ini", "ScanVirus", Me.VSscr.Name, True)

Me.VSScanProcess.Value = ReadIniFile(AppPath & "Setting.ini", "ScanVirus", Me.VSScanProcess.Name, True)
Me.VSScanStartUp.Value = ReadIniFile(AppPath & "Setting.ini", "ScanVirus", Me.VSScanStartUp.Name, True)

Me.CSbat.Value = ReadIniFile(AppPath & "Setting.ini", "ScanVirus", Me.VSbat.Name, True)
Me.CScom.Value = ReadIniFile(AppPath & "Setting.ini", "ScanVirus", Me.VScom.Name, True)



Me.CSprocess.Value = ReadIniFile(AppPath & "Setting.ini", "ScanVirus", Me.VSScanProcess.Name, True)
Me.CSStartUp.Value = ReadIniFile(AppPath & "Setting.ini", "ScanVirus", Me.VSScanStartUp.Name, True)



Me.VSDontScanSize.Value = ReadIniFile(AppPath & "Setting.ini", "ScanVirus", Me.VSDontScanSize.Name, True)
Me.VSLimitSize.Value = ReadIniFile(AppPath & "Setting.ini", "ScanVirus", Me.VSLimitSize.Name, 1)
Me.atsAlwaysScanFolder.Value = ReadIniFile(AppPath & "Setting.ini", "AutoScan", "ScanFolder", True)
Me.atsScanEXE.Value = ReadIniFile(AppPath & "Setting.ini", "AutoScan", "ScanEXE", False)
Me.atsAutoScanUSB.Value = ReadIniFile(AppPath & "Setting.ini", "AutoScan", "ScanUSB", True)
Me.atsScanKeylogger.Value = ReadIniFile(AppPath & "Setting.ini", "AutoScan", "ScanKeylogger", True)


If ReadIniFile(AppPath & "Setting.ini", "AutoProtect", "Registry", True) = False Then
    Me.atptmrREG.Enabled = False
    Me.atplbREG.Caption = "D9ang ta81t"
    Me.atpcmdREG.Caption = "Mo73 chu71c na8ng na2y"
Else
    Me.atptmrREG.Enabled = True
    Me.atplbREG.Caption = "D9ang mo73"
    Me.atpcmdREG.Caption = "Ta81t chu71c na8ng na2y"
End If

If ReadIniFile(AppPath & "Setting.ini", "AutoProtect", "Autorun", True) = False Then
    Me.atptmrAutorun.Enabled = False
    Me.atplblStaAutorun.Caption = "D9ang ta81t"
    Me.atpcmdAutorun.Caption = "Mo73 chu71c na8ng na2y"
Else
    Me.atptmrAutorun.Enabled = True
    Me.atplblStaAutorun.Caption = "D9ang mo73"
    Me.atpcmdAutorun.Caption = "Ta81t chu71c na8ng na2y"
End If
'<--- Load Setting [Second]

'---> Set Tray TooltipText
Tray1.ToolTipText = "Perfect Antivirus 2009 - Ma1y ti1nh cu3a ba5n d9ang o73 ti2nh tra5ng to61t nha61t!"
'<--- Set Tray TooltipText



'---> Set for "Chuc Nang"

    With TreeChucNang
        .Initialize
        '.InitializeImageList 20, 20
        .HasButtons = True
        .SingleExpand = False
        
        .AddNode , , "a", "Que1t & Die65t Virus", 0, 0
            .AddNode "a", , "0", "Que1t toa2n bo65", 1, 1
            .AddNode "a", , "1", "Que1t tu2y cho5n", 2, 3
            .AddNode "a", , "3", "Ca61u hi2nh que1t", 4, 4
            .AddNode "a", , "4", "Nha65t ky1 que1t Virus"
            

        .AddNode , , "b", "Tu75 d9o65ng ba3o ve65"
            .AddNode "b", , "5", "Ba3o ve65 Registry"
            .AddNode "b", , "2", "Tu75 d9o65ng que1t"
            .AddNode "b", , "6", "Ba3o ve65 Autorun"

        .AddNode , , "c", "Tie65n i1ch he65 tho61ng"
            .AddNode "c", , "7", "Qua3n ly1 tie61n tri2nh"


    End With


'---> Set for "Chuc Nang"

'---> Properties Setting
xFSStopScan = False
'<--- Properties Setting



'---> Connect Database
modScanVirus.ConnectDB
'<--- Connect Database


'---> FullScan
xCustomScan = False
HoanThanhFull = False

LVVirus1.View = eViewDetails
LVVirus1.GridLines = True
LVVirus1.HeaderButtons = False
LVVirus1.CheckBoxes = True

LVVirus1.Columns.Add , , "Virus Name", , 2000
LVVirus1.Columns.Add , , "Path", , 3500
LVVirus1.Columns.Add , , "Size", , 1000
LVVirus1.Columns.Add , , "Process ID", , 1000
LVVirus1.Columns.Add , , "Start Up Key", , 5000
LVVirus1.Refresh
'<--- FullScan

'---> Custom Scan
HoanThanhCus = False
LVVirus2.View = eViewDetails
LVVirus2.GridLines = True
LVVirus2.HeaderButtons = False
LVVirus2.CheckBoxes = True

LVVirus2.Columns.Add , , "Virus Name", , 2000
LVVirus2.Columns.Add , , "Path", , 3500
LVVirus2.Columns.Add , , "Size", , 1000
LVVirus2.Columns.Add , , "Process ID", , 1000
LVVirus2.Columns.Add , , "Start Up Key", , 5000
LVVirus2.Refresh
'<--- Custom Scan


'---> Nhat ky' virus
With LVVirusEvents
    .View = eViewDetails
    .GridLines = True
    .HeaderButtons = False
    .CheckBoxes = True
    .AutoUnicode = True
    .Columns.Add , , "Tho72i gian", , 2200
    .Columns.Add , , "Kie63u que1t", , 2000
    .Columns.Add , , "So61 File", , 1000
    .Columns.Add , , "So61 Virus", , 1000
    .Columns.Add , , "Ke61t qua3", , 2000

Dim k
k = ReadIniFile(AppPath & "VirusScanLog.log", "Other", "Total", 0)
    Dim la
    
    For la = 1 To k
        .ListItems.Add la, , ReadIniFile(AppPath & "VirusScanLog.log", la, "TimeQuet", "")
        .ListItems(la).SubItems(1).Caption = UTF82Unicode(ReadIniFile(AppPath & "VirusScanLog.log", la, "KieuQuet", ""))
        .ListItems(la).SubItems(2).Caption = ReadIniFile(AppPath & "VirusScanLog.log", la, "SoFile", "")
        .ListItems(la).SubItems(3).Caption = ReadIniFile(AppPath & "VirusScanLog.log", la, "SoVirus", "")
        .ListItems(la).SubItems(4).Caption = UTF82Unicode(ReadIniFile(AppPath & "VirusScanLog.log", la, "KetQua", ""))
        
    Next la
End With
'<--- Nhat ky' virus



'---> LV REG
With atpLVREG
    .View = eViewDetails
    .GridLines = True
    .HeaderButtons = False
    .CheckBoxes = True
    .AutoUnicode = True
    .Columns.Add , , "Te6n chu71c na8ng", , 2300
    .Columns.Add , , "Kho1a go61c", , 2000
    .Columns.Add , , "D9u7o72ng da64n kho1a", , 5000
    .Columns.Add , , "Te6n kho1a"
    .Columns.Add , , "Gia1 tri5 ma85c d9i5nh"



    
    Dim F
    For F = 1 To ReadIniFile(AppPath & "RegProtect.dat", "Other", "Total", 0)
        .ListItems.Add F, , UTF82Unicode(ReadIniFile(AppPath & "RegProtect.dat", F, "TenChucNang", ""))
        .ListItems(F).SubItems(1).Caption = ReadIniFile(AppPath & "RegProtect.dat", F, "KeyGoc", "")
        .ListItems(F).SubItems(2).Caption = ReadIniFile(AppPath & "RegProtect.dat", F, "KeyPath", "")
        .ListItems(F).SubItems(3).Caption = ReadIniFile(AppPath & "RegProtect.dat", F, "KeyName", "")
        .ListItems(F).SubItems(4).Caption = ReadIniFile(AppPath & "RegProtect.dat", F, "KeyData", "")
    Next F

End With
'<--- LV REG




'--- LV Pro
LVPro.AutoUnicode = False
LVPro.View = eViewDetails
LVPro.GridLines = True
LVPro.HeaderButtons = False
LVPro.MultiSelect = False
LVPro.Columns.Add , , "Image Name"
LVPro.Columns.Add , , "Path", , 5000
LVPro.Columns.Add , , "Process ID", , 1000
LVPro.Columns.Add , , "Size", , 1000
LVPro.Columns.Add , , "Attributes", , 500
LVPro.Columns.Add , , "Memory Usage", , 1000
LVPro.Columns.Add , , "Priority", , 1000

'Attributes
'<--- LV pro




'---> Hide all fm
HideAllFM
'<--- Hide all fm


'---> Load auto scan


If atsAlwaysScanFolder.Value = True Then
    Load zfrmAutoScanFolder
    atsAlwaysScanFolder.Caption = "Tu75 d9o65ng que1t va2 ca3nh ba1o Virus trong Thu7 mu5c d9ang d9u7o75c mo73. [D9ang Ba65t]"
Else
    atsAlwaysScanFolder.Caption = "Tu75 d9o65ng que1t va2 ca3nh ba1o Virus trong Thu7 mu5c d9ang d9u7o75c mo73. [D9ang Ta81t]"
End If


If atsScanEXE.Value = True Then
        atsScanEXE.Caption = "Tu75 d9o65ng que1t va2 ca3nh ba1o Virus cho nhu74ng File sa81p d9u7o75c mo73. [D9ang Ba65t]"
        SaveString HKEY_CLASSES_ROOT, "exefile\shell\open\command", "", ChrW(34) & AppPath & "Check_Virus.exe" & ChrW(34) & " " & ChrW(34) & "%1" & ChrW(34)
        SaveString HKEY_CLASSES_ROOT, "batfile\shell\open\command", "", ChrW(34) & AppPath & "Check_Virus.exe" & ChrW(34) & " " & ChrW(34) & "%1" & ChrW(34)
        SaveString HKEY_CLASSES_ROOT, "cmdfile\shell\open\command", "", ChrW(34) & AppPath & "Check_Virus.exe" & ChrW(34) & " " & ChrW(34) & "%1" & ChrW(34)
        SaveString HKEY_CLASSES_ROOT, "comfile\shell\open\command", "", ChrW(34) & AppPath & "Check_Virus.exe" & ChrW(34) & " " & ChrW(34) & "%1" & ChrW(34)
Else
        atsScanEXE.Caption = "Tu75 d9o65ng que1t va2 ca3nh ba1o Virus cho nhu74ng File sa81p d9u7o75c mo73. [D9ang Ta81t]"
        SaveString HKEY_CLASSES_ROOT, "exefile\shell\open\command", "", ChrW(34) & "%1" & ChrW(34) & " %*"
        SaveString HKEY_CLASSES_ROOT, "batfile\shell\open\command", "", ChrW(34) & "%1" & ChrW(34) & " %*"
        SaveString HKEY_CLASSES_ROOT, "cmdfile\shell\open\command", "", ChrW(34) & "%1" & ChrW(34) & " %*"
        SaveString HKEY_CLASSES_ROOT, "comfile\shell\open\command", "", ChrW(34) & "%1" & ChrW(34) & " %*"
End If

If atsAutoScanUSB.Value = True Then
    atsAutoScanUSB.Caption = "Tu75 d9o65ng que1t Virus cho USB khi pha1t hie65n USB ke61t no61i va2o ma1y ti1nh. [D9ang Ba65t]"
    Load zfrmScanUSB
Else
    atsAutoScanUSB.Caption = "Tu75 d9o65ng que1t Virus cho USB khi pha1t hie65n USB ke61t no61i va2o ma1y ti1nh. [D9ang Ta81t]"
End If


If atsScanKeylogger.Value = True Then
    Load zfrmAntiKey
    atsScanKeylogger.Caption = "Tu75 d9o65ng pha1t hie65n va2 ca3nh ba1o Keylogger. [D9ang Ba65t]"
Else
    atsScanKeylogger.Caption = "Tu75 d9o65ng pha1t hie65n va2 ca3nh ba1o Keylogger. [D9ang Ta81t]"
End If
'<--- Load auto scan








End Sub



Private Sub Form_Unload(Cancel As Integer)
Cancel = 1

frmMain.Visible = False
App.TaskVisible = False

End Sub






Private Sub xStartFullScanNow()


If Me.VSScanProcess.Value = True Then
    lblStatus2.Caption = "D9ang que1t ca1c chu7o7ng tri2nh cha5y trong bo65 nho71..."

    modScanVirus.xScanProcess

End If

If Me.VSScanStartUp.Value = True Then

    lblStatus2.Caption = "D9ang que1t ca1c chu7o7ng tri2nh kho73i d9o65ng..."
    
    lblStatus.Caption = "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Run"
    GetKeyValue "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Run"
    
    lblStatus.Caption = "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\RunOnce"
    GetKeyValue "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\RunOnce"
    
    lblStatus.Caption = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Run"
    GetKeyValue "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Run"
    
    lblStatus.Caption = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\RunOnce"
    GetKeyValue "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\RunOnce"
    
    lblStatus.Caption = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\RunOnceEx"
    GetKeyValue "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\RunOnceEx"
End If



Dim Str
Dim str2
Dim FSO As New FileSystemObject
    Dim drv As Drive
    Dim drvs As Drives
    'On Error Resume Next
    Set drvs = FSO.Drives
        For Each drv In drvs
            If xFSStopScan = True Then GoTo SsKkIiPp
            If UCase(drv.DriveLetter) <> "A" Then
            
                

                DoEvents
                lblStatus2.Caption = "D9ang que1t File EXE..."
                SearchFile drv.DriveLetter & ":\", "*.exe"
                lblStatus2.Caption = "D9ang que1t File BAT..."
                If Me.VSbat.Value = True Then SearchFile drv.DriveLetter & ":\", "*.bat"
                lblStatus2.Caption = "D9ang que1t File CMD..."
                If Me.VScmd.Value = True Then SearchFile drv.DriveLetter & ":\", "*.cmd"
                lblStatus2.Caption = "D9ang que1t File COM..."
                If Me.VScom.Value = True Then SearchFile drv.DriveLetter & ":\", "*.com"
                lblStatus2.Caption = "D9ang que1t File SCR..."
                If Me.VSscr.Value = True Then SearchFile drv.DriveLetter & ":\", "*.scr"
                lblStatus2.Caption = "D9ang que1t File DLL..."
                If Me.VSdll.Value = True Then SearchFile drv.DriveLetter & ":\", "*.dll"
                
                'Kill drv.DriveLetter & ":\autorun.inf"
                'SearchFile "C:\", "*.exe"

            End If

        Next
SsKkIiPp:
    Set FSO = Nothing
    Set drv = Nothing
    Set drvs = Nothing
    DoEvents


lblStatus.AutoUnicode = True
lblStatus.Caption = "D9a4 Que1t Xong! Ti2m Tha61y: " & LVVirus1.ListItems.Count & " Ta65p Tin Virus Trong To63ng So61 " & xTotalFile & " Ta65p Tin D9a4 Ti2m"
lblStatus2.Caption = "D9a4 que1t xong!"
cmdFSCachLy.Enabled = True
cmdFSKillVirus.Enabled = True
cmdFSReport.Enabled = True
cmdFSStop.Enabled = False
cmdFSStart.Enabled = True
tmrScanTime.Enabled = False
fmMain.Enabled = True
cmdSettingFullScan.Enabled = True
PLaySound AppPath & "Sound\ScanDone.wav"
Tray1.ToolTipText = "Perfect Antivirus 2009 - Ma1y ti1nh cu3a ba5n d9ang o73 ti2nh tra5ng to61t nha61t!"

If frmMain.Visible = False Then
    frmMessenger.zShowMessenger "D9a4 que1t xong", "Chu7o7ng tri2nh d9a4 que1t Virus xong! Ke61t qua3: Ti2m Tha61y: " & LVVirus1.ListItems.Count & " Ta65p Tin Virus Trong To63ng So61 " & xTotalFile & " Ta65p Tin D9a4 Ti2m. Ha4y ba65t chu7o7ng tri2nh le6n va2 xu73 ly1 chu1ng.", 5000, xTrang
End If


With LVVirusEvents
Dim i
i = .ListItems.Count + 1
    .ListItems.Add i, , TimeFull
    .ListItems(i).SubItems(1).Caption = ToUnicode("Que1t toa2n bo65")
    .ListItems(i).SubItems(2).Caption = xTotalFile
    .ListItems(i).SubItems(3).Caption = LVVirus1.ListItems.Count
    .ListItems(i).SubItems(4).Caption = IIf(HoanThanhFull, ToUnicode("D9a4 hoa2n tha2nh"), ToUnicode("Chu7a hoa2n tha2nh"))
    .AutoUnicode = True
End With



End Sub

Private Sub tmrPro_Timer()

Dim theloop As Long
  Dim proc As PROCESSENTRY32
  Dim snap As Long
  Dim exename As String
  Dim ID As Long
   snap = CreateToolhelpSnapshot(TH32CS_SNAPall, 0)
   proc.dwSize = Len(proc)
   theloop = ProcessFirst(snap, proc)

Dim i
i = 0
   While theloop <> 0
        
      ID = proc.th32ProcessID
      theloop = ProcessNext(snap, proc)
      If ProcessPathByPID(proc.th32ProcessID) <> "SYSTEM" Then
      'MsgBox ProcessPathByPID(proc.th32ProcessID)

                  'Set lsv = lstView.ListItems.Add()
                  'lsv.Text = proc.szExeFile
                  'lsv.SubItems(1) = ProcessPathByPID(proc.th32ProcessID)
                  'lsv.SubItems(2) = proc.th32ProcessID
                  i = i + 1
                  If i > LVPro.ListItems.Count Then GoTo ReFeR
            If LVPro.ListItems(i).SubItems(1).Caption <> ProcessPathByPID(proc.th32ProcessID) Then GoTo ReFeR
      End If
   Wend
   CloseHandle snap
    EnumWindows AddressOf EnumWindowsProc, ByVal 0&
    'MsgBox i & "-" & LV1.ListItems.Count - SoLuong
   If i < LVPro.ListItems.Count - SoLuong Then GoTo ReFeR

Exit Sub
ReFeR:
GetProcess LVPro
End Sub

Private Sub tmrScanTime_Timer()
xTime = xTime + 1
Me.lblScanTime(2).Caption = xTime
If Me.lblScanTime(2).Caption > 59 Then
    Me.lblScanTime(1).Caption = lblScanTime(1).Caption + 1
    xTime = 0
End If
If Me.lblScanTime(1).Caption > 59 Then
    Me.lblScanTime(0).Caption = lblScanTime(0).Caption + 1
    Me.lblScanTime(1).Caption = 0
End If
End Sub

Private Sub tmrStartFullScan_Timer()
xTimeSC = xTimeSC + 1
lblStatus.AutoUnicode = True

If xTimeSC = 1 Then
        lblStatus.Caption = "D9ang que1t ca1c kho1a kho73i d9o65ng trong Registry..."
ElseIf xTimeSC = 2 Then
        lblStatus.Caption = "D9ang que1t ca1c tie61n tri2nh d9ang hoa5t d9o65ng..."
ElseIf xTimeSC = 3 Then
        lblStatus.Caption = "Chua63n bi5 que1t file.."
ElseIf xTimeSC = 4 Then
        lblStatus.AutoUnicode = False
        xStartFullScanNow
        tmrStartFullScan.Enabled = False
End If


End Sub



Private Sub UniButton1_Click()
ff.Visible = False
End Sub



Private Sub Tray1_TrayClick(Button As UniControls.stMouseEvent)
If Button = stRightButtonDown Then
    PopupMenu frmMenu.m
ElseIf Button = stLeftButtonDoubleClick Then
    frmMain.Visible = True
    frmMain.Show
    frmMain.WindowState = 0
    App.TaskVisible = True
End If
End Sub

Private Sub TreeChucNang_NodeClick(ByVal hNode As Long)
HideAllFM
If IsNumeric(TreeChucNang.GetNodeKey(hNode)) = False Then Exit Sub
fm(TreeChucNang.GetNodeKey(hNode)).Visible = True

If TreeChucNang.GetNodeKey(hNode) = 7 Then
    tmrPro.Enabled = True
Else
    tmrPro.Enabled = False
End If
End Sub

Private Sub VSbat_Click()
CSbat.Value = VSbat.Value
End Sub

Private Sub VScom_Click()
CScom.Value = VScom.Value
End Sub

Private Sub VSScanProcess_Click()
CSprocess.Value = VSScanProcess.Value
End Sub

Private Sub VSScanStartUp_Click()
CSStartUp.Value = VSScanStartUp.Value
End Sub



