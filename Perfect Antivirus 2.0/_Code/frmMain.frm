VERSION 5.00
Object = "{E8FDD05C-3067-4198-8AEC-1A013A46ABDD}#1.0#0"; "FVUnicodeControl.ocx"
Object = "{6DC1AB90-DCD1-47A1-AB36-924FFD67ADBF}#1.0#0"; "UniControls_v2.0.ocx"
Begin VB.Form frmMain 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Perfect Antivirus v2.0"
   ClientHeight    =   6330
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11040
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
   Picture         =   "frmMain.frx":058A
   ScaleHeight     =   6330
   ScaleWidth      =   11040
   StartUpPosition =   2  'CenterScreen
   Begin UniControls.UniLabel UniLabel76 
      Height          =   255
      Left            =   6840
      Top             =   0
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   450
      Alignment       =   2
      BackStyle       =   0
      Caption         =   "Nha61n F1 d9e63 d9u7o75c tro75 giu1p"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16711680
   End
   Begin UniControls.UniLabel lblSupport 
      Height          =   255
      Left            =   0
      Top             =   6120
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   450
      Alignment       =   1
      Caption         =   "Tha81c ma81c - Go1p y1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16711680
   End
   Begin FVUnicodeControl.FVistaUniFrame FM 
      Height          =   375
      Index           =   19
      Left            =   8880
      TabIndex        =   37
      Top             =   6000
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      Alignment       =   0
      BackColor       =   16777215
      ForeColor       =   255
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Tho6ng tin ta1c gia3"
      AutoUnicode     =   -1  'True
      Begin UniControls.UniLabel UniLabel68 
         Height          =   255
         Left            =   1680
         Top             =   3960
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   450
         Alignment       =   1
         Caption         =   "Xin cha6n tha2nh ca3m o7n!"
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
      Begin UniControls.UniLabel UniLabel67 
         Height          =   255
         Left            =   360
         Top             =   3600
         Width           =   6135
         _ExtentX        =   10821
         _ExtentY        =   450
         Caption         =   "- Ta61t ca3 ca1c ba5n d9a4 d9o1ng go1p y1 kie61n nhie65t ti2nh trong phie6n ba3n PAV d9a62u tie6n!"
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
      Begin UniControls.UniLabel UniLabel66 
         Height          =   255
         Left            =   360
         Top             =   3360
         Width           =   6255
         _ExtentX        =   11033
         _ExtentY        =   450
         Caption         =   "- vnAntivirus (dungcoivb@gmail.com): Vo71i bo65 CSDL ga62n 700.000 ma4 nha65n da5ng Virus"
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
      Begin UniControls.UniLabel UniLabel65 
         Height          =   255
         Left            =   360
         Top             =   3120
         Width           =   6015
         _ExtentX        =   10610
         _ExtentY        =   450
         Caption         =   "- iVB Team (http://caulacbovb.com): Vo71i bo65 d9ie62u khie63n UnicodeControl_2.0.OCX"
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
      Begin UniControls.UniLabel UniLabel64 
         Height          =   255
         Left            =   360
         Top             =   2880
         Width           =   6255
         _ExtentX        =   11033
         _ExtentY        =   450
         Caption         =   "- VirusVN (http://virusvn.com): Vo71i ca1c ma64u Virus mo71i nha61t hie65n nay"
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
      Begin UniControls.UniLabel UniLabel63 
         Height          =   255
         Left            =   360
         Top             =   2640
         Width           =   6255
         _ExtentX        =   11033
         _ExtentY        =   450
         Caption         =   "- PSCode (http://pscode.com): Vo71i Module xu73 ly1 Registry, Process."
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
      Begin UniControls.UniLabel UniLabel62 
         Height          =   255
         Left            =   240
         Top             =   2400
         Width           =   6375
         _ExtentX        =   11245
         _ExtentY        =   450
         BackStyle       =   0
         Caption         =   "D9e63 co1 d9u7o75c chu7o7ng tri2nh na2y, to6i xin cha6n tha2nh gu73i lo72i ca3m o7n d9e61n:"
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
      Begin UniControls.UniLabel lblGopY 
         Height          =   255
         Left            =   1920
         Top             =   1920
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   450
         Caption         =   "Nha61n va2o d9a6y d9e63 gu73i go1p y1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   255
         Link            =   ""
      End
      Begin UniControls.UniLabel UniLabel60 
         Height          =   255
         Left            =   120
         Top             =   1920
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   450
         Alignment       =   2
         Caption         =   "Hoa85c go1p y1 tru75c tie61p:"
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
      Begin UniControls.UniLabel UniLabel59 
         Height          =   210
         Left            =   1920
         Top             =   1440
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   370
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Link            =   "http://www2.shoutmix.com/?qtsoft"
         Style           =   1
      End
      Begin UniControls.UniLabel UniLabel56 
         Height          =   255
         Left            =   480
         Top             =   1440
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         Alignment       =   2
         Caption         =   "Ho2m thu7 go1p y1:"
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
      Begin UniControls.UniLabel UniLabel58 
         Height          =   210
         Left            =   1920
         Top             =   1680
         Width           =   2385
         _ExtentX        =   4207
         _ExtentY        =   370
         Caption         =   "dinhquangtrung90@yahoo.com"
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
      Begin UniControls.UniLabel UniLabel57 
         Height          =   255
         Left            =   840
         Top             =   1680
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   450
         Alignment       =   2
         Caption         =   "Hoa85c Email:"
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
      Begin UniControls.UniLabel UniLabel55 
         Height          =   255
         Left            =   240
         Top             =   1200
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   450
         Caption         =   "Mo5i tha81c ma81c, go1p y1, d9a1nh gia1 ve62 chu7o7ng tri2nh xin gu73i d9e61n:"
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
      Begin UniControls.UniLabel UniLabel54 
         Height          =   210
         Left            =   2160
         Top             =   840
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   370
         Caption         =   "http://phanmemvn.net"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Link            =   "http://phanmemvn.net"
         Style           =   1
      End
      Begin UniControls.UniLabel UniLabel53 
         Height          =   255
         Left            =   1440
         Top             =   840
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   450
         Caption         =   "Website:"
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
      Begin UniControls.UniLabel UniLabel52 
         Height          =   255
         Left            =   1440
         Top             =   600
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   450
         Caption         =   "Email: dinhquangtrung90@yahoo.com"
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
      Begin UniControls.UniLabel UniLabel51 
         Height          =   255
         Left            =   1440
         Top             =   360
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   450
         Caption         =   "Ta1c gia3: D9inh Quang Trung (12/12/1993)"
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
      Begin VB.Line Line6 
         X1              =   240
         X2              =   4200
         Y1              =   2280
         Y2              =   2280
      End
      Begin VB.Image Image1 
         Height          =   930
         Left            =   120
         Picture         =   "frmMain.frx":11B11
         Stretch         =   -1  'True
         Top             =   240
         Width           =   1260
      End
   End
   Begin FVUnicodeControl.FVistaUniFrame FM 
      Height          =   135
      Index           =   18
      Left            =   8880
      TabIndex        =   36
      Top             =   5880
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   238
      Alignment       =   0
      BackColor       =   16777215
      ForeColor       =   255
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Tho6ng tin chu7o7ng tri2nh"
      AutoUnicode     =   -1  'True
      Begin FVUnicodeControl.FVistaUniButton cmdCheckForUpdate 
         Height          =   375
         Left            =   5160
         TabIndex        =   91
         Top             =   3840
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         BackColor       =   -2147483633
         ButtonStyle     =   3
         Caption         =   "Kie63m tra ca65p nha65t"
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
      Begin FVUnicodeControl.FVistaUniLabel FVistaUniLabel1 
         Height          =   255
         Left            =   5040
         TabIndex        =   90
         Top             =   360
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   450
         Caption         =   "2.0.9 (07/06/2011)"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483636
      End
      Begin UniControls.UniLabel UniLabel50 
         Height          =   255
         Left            =   3480
         Top             =   2160
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   450
         Caption         =   "- Ba3o ve65 Registry"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   -1  'True
         EndProperty
      End
      Begin UniControls.UniLabel UniLabel49 
         Height          =   255
         Left            =   360
         Top             =   2400
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   450
         Caption         =   "- Que1t tu2y cho5n"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   -1  'True
         EndProperty
      End
      Begin UniControls.UniLabel UniLabel48 
         Height          =   255
         Left            =   5280
         Top             =   1680
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   450
         BackStyle       =   0
         Caption         =   "Na6ng ca61p"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   49152
      End
      Begin UniControls.UniLabel UniLabel47 
         Height          =   255
         Left            =   4080
         Top             =   1680
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   450
         BackStyle       =   0
         Caption         =   "Xo1a bo3"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   -1  'True
         EndProperty
      End
      Begin UniControls.UniLabel UniLabel24 
         Height          =   255
         Left            =   2280
         Top             =   1680
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         BackStyle       =   0
         Caption         =   "Chu71c na8ng mo71i"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   255
      End
      Begin UniControls.UniLabel UniLabel46 
         Height          =   255
         Left            =   600
         Top             =   1680
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         BackStyle       =   0
         Caption         =   "Chu71c na8ng cu4"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   16711680
      End
      Begin UniControls.UniLabel UniLabel45 
         Height          =   255
         Left            =   3480
         Top             =   4080
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   450
         Caption         =   "- Que1t ma64u co1 sa84n"
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
      Begin UniControls.UniLabel UniLabel44 
         Height          =   255
         Left            =   3480
         Top             =   3840
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   450
         Caption         =   "- Kie63m tra he65 tho61ng"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   49152
      End
      Begin UniControls.UniLabel UniLabel43 
         Height          =   255
         Left            =   3480
         Top             =   3600
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   450
         Caption         =   "- Qua3n ly1 kho73i d9o65ng"
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
      Begin UniControls.UniLabel UniLabel42 
         Height          =   255
         Left            =   3480
         Top             =   3360
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   450
         Caption         =   "- Qua3n ly1 tie61n tri2nh"
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
      Begin UniControls.UniLabel UniLabel41 
         Height          =   255
         Left            =   3480
         Top             =   3120
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   450
         Caption         =   "- Ta8ng to61c ma1y ti1nh (13 chu71c na8ng)"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   49152
      End
      Begin UniControls.UniLabel UniLabel40 
         Height          =   255
         Left            =   3480
         Top             =   2880
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   450
         Caption         =   "- Tinh chi3nh Registry (27 chu71c na8ng)"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   49152
      End
      Begin UniControls.UniLabel UniLabel39 
         Height          =   255
         Left            =   3480
         Top             =   2640
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   450
         Caption         =   "- Phu5c ho62i du74 lie65u"
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
      Begin UniControls.UniLabel UniLabel38 
         Height          =   255
         Left            =   3480
         Top             =   2400
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   450
         Caption         =   "- Ba3o ve65 Autorun"
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
      Begin UniControls.UniLabel UniLabel37 
         Height          =   255
         Left            =   360
         Top             =   4080
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   450
         Caption         =   "- Ba3o ve65 tho72i gian thu75c"
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
      Begin UniControls.UniLabel UniLabel36 
         Height          =   255
         Left            =   360
         Top             =   3840
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   450
         Caption         =   "- Danh sa1ch ca1c File tin tu7o73ng"
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
      Begin UniControls.UniLabel UniLabel35 
         Height          =   255
         Left            =   360
         Top             =   3600
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   450
         Caption         =   "- Cho phe1p the6m ma64u Virus"
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
      Begin UniControls.UniLabel UniLabel34 
         Height          =   255
         Left            =   360
         Top             =   3360
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   450
         Caption         =   "- Que1t tie61n tri2nh he65 tho61ng"
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
      Begin UniControls.UniLabel UniLabel33 
         Height          =   255
         Left            =   360
         Top             =   3120
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   450
         Caption         =   "- Que1t ca1c file kho73i d9o65ng"
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
      Begin UniControls.UniLabel UniLabel32 
         Height          =   255
         Left            =   360
         Top             =   2880
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   450
         Caption         =   "- Qua3n ly1 nha65t ky1 que1t"
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
      Begin UniControls.UniLabel UniLabel31 
         Height          =   255
         Left            =   360
         Top             =   2640
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   450
         Caption         =   "- Ma4 ho1a va2 ca1ch ly Virus"
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
      Begin UniControls.UniLabel UniLabel30 
         Height          =   255
         Left            =   360
         Top             =   2160
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   450
         Caption         =   "- Que1t va2 die65t Virus"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   49152
      End
      Begin UniControls.UniLabel UniLabel29 
         Height          =   255
         Left            =   1800
         Top             =   1440
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   450
         Caption         =   "Ca1c chu71c trong phie6n ba3n 2.0 na2y"
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
      Begin UniControls.UniLabel UniLabel28 
         Height          =   255
         Left            =   2400
         Top             =   1080
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   450
         BackStyle       =   0
         Caption         =   "Tu7o7ng thi1ch: Windows XP"
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
      Begin UniControls.UniLabel UniLabel27 
         Height          =   255
         Left            =   2400
         Top             =   840
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   450
         Caption         =   "Pha1t ha2nh: Perfect Software"
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
      Begin UniControls.UniLabel UniLabel26 
         Height          =   255
         Left            =   480
         Top             =   1080
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   450
         Caption         =   "Phie6n ba3n: 2.0"
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
      Begin UniControls.UniLabel UniLabel25 
         Height          =   255
         Left            =   480
         Top             =   840
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   450
         Caption         =   "Ki1ch thu7o71c: 2MB"
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
      Begin UniControls.UniLabel UniLabel23 
         Height          =   255
         Left            =   4200
         Top             =   600
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   450
         Caption         =   "Phie6n ba3n II"
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
      Begin UniControls.UniLabel UniLabel22 
         Height          =   375
         Left            =   1440
         Top             =   240
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   661
         Alignment       =   1
         Caption         =   "Perfect Antivirus 2009"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   16711680
      End
      Begin VB.Line Line5 
         BorderColor     =   &H00FF0000&
         X1              =   6240
         X2              =   6600
         Y1              =   1920
         Y2              =   1560
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00FF0000&
         X1              =   120
         X2              =   480
         Y1              =   1560
         Y2              =   1920
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00FF0000&
         X1              =   3120
         X2              =   3120
         Y1              =   2160
         Y2              =   4320
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FF0000&
         X1              =   480
         X2              =   6240
         Y1              =   1920
         Y2              =   1920
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         X1              =   120
         X2              =   6600
         Y1              =   1560
         Y2              =   1560
      End
   End
   Begin FVUnicodeControl.FVistaUniFrame FM 
      Height          =   615
      Index           =   17
      Left            =   8880
      TabIndex        =   35
      Top             =   5760
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   1085
      Alignment       =   0
      BackColor       =   16777215
      ForeColor       =   255
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Ca2i d9a85t chung"
      AutoUnicode     =   -1  'True
      Begin FVUnicodeControl.FVistaUniFrame FMCapNhat 
         Height          =   2055
         Left            =   360
         TabIndex        =   79
         Top             =   2160
         Width           =   6015
         _ExtentX        =   10610
         _ExtentY        =   3625
         Alignment       =   0
         BackColor       =   16777215
         ForeColor       =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Ca65p nha65t"
         AutoUnicode     =   -1  'True
         Begin UniControls.UniLabel UniLabel75 
            Height          =   255
            Left            =   1080
            Top             =   120
            Width           =   4935
            _ExtentX        =   8705
            _ExtentY        =   450
            BackStyle       =   0
            Caption         =   "(Bo3 d9a1nh da61u ca1c mu5c be6n du7o71i ne61u ma1y ti1nh cu3a ba5n kho6ng ke61t no61i Internet)"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   8421504
         End
         Begin FVUnicodeControl.FVistaUniCheckbox CHK1 
            Height          =   195
            Index           =   3
            Left            =   240
            TabIndex        =   87
            Top             =   1560
            Width           =   2925
            _ExtentX        =   5159
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
            Caption         =   "Tu75 d9o65ng the6m va2o ca1c ma64u Virus mo71i"
            ForeColor       =   0
            ShowFocusRectangle=   0   'False
         End
         Begin FVUnicodeControl.FVistaUniCheckbox CHK1 
            Height          =   195
            Index           =   2
            Left            =   240
            TabIndex        =   86
            Top             =   1200
            Width           =   3930
            _ExtentX        =   6932
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
            Caption         =   "Ho3i y1 kie61n ngu7o72i du2ng khi ca65p nha65t ca1c ba3n Update"
            ForeColor       =   0
            ShowFocusRectangle=   0   'False
         End
         Begin FVUnicodeControl.FVistaUniCheckbox CHK1 
            Height          =   195
            Index           =   1
            Left            =   240
            TabIndex        =   85
            Top             =   840
            Width           =   2340
            _ExtentX        =   4128
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
            Caption         =   "Ca65t nha65t khi ke61t no61i Internet"
            ForeColor       =   0
            ShowFocusRectangle=   0   'False
         End
         Begin FVUnicodeControl.FVistaUniCheckbox CHK1 
            Height          =   195
            Index           =   0
            Left            =   240
            TabIndex        =   84
            Top             =   480
            Width           =   3795
            _ExtentX        =   6694
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
            Caption         =   "Tu75 d9o65ng ca65p nha65t mo64i khi kho73i d9o65ng chu7o7ng tri2nh"
            ForeColor       =   0
            ShowFocusRectangle=   0   'False
         End
      End
      Begin FVUnicodeControl.FVistaUniFrame fmCauHinh 
         Height          =   1695
         Left            =   360
         TabIndex        =   78
         Top             =   360
         Width           =   6015
         _ExtentX        =   10610
         _ExtentY        =   2990
         Alignment       =   0
         BackColor       =   16777215
         ForeColor       =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Ca61u hi2nh"
         AutoUnicode     =   -1  'True
         Begin FVUnicodeControl.FVistaUniCheckbox chkCheDoBaoVe 
            Height          =   195
            Left            =   240
            TabIndex        =   89
            Top             =   840
            Width           =   2940
            _ExtentX        =   5186
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
            Caption         =   "Ki1ch hoa5t che61 d9o65 ba3o ve65 thu7o72ng tru75c"
            ForeColor       =   0
            ShowFocusRectangle=   0   'False
         End
         Begin FVUnicodeControl.FVistaUniCheckbox CHK0 
            Height          =   195
            Index           =   3
            Left            =   240
            TabIndex        =   83
            Top             =   1320
            Width           =   5175
            _ExtentX        =   9128
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
            Caption         =   "Su73 du5ng chu71c na8ng que1t nhanh (Right Click => Scan with PAV2009)"
            ForeColor       =   0
            ShowFocusRectangle=   0   'False
         End
         Begin FVUnicodeControl.FVistaUniCheckbox CHK0 
            Height          =   195
            Index           =   2
            Left            =   240
            TabIndex        =   82
            Top             =   1080
            Width           =   3615
            _ExtentX        =   6376
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
            Caption         =   "Cho phe1p xua61t hie65n tho6ng ba1o o73 go1c ma2n hi2nh"
            ForeColor       =   0
            ShowFocusRectangle=   0   'False
         End
         Begin FVUnicodeControl.FVistaUniCheckbox CHK0 
            Height          =   195
            Index           =   1
            Left            =   240
            TabIndex        =   81
            Top             =   600
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
            Caption         =   "Tu75 d9o65ng phu5c ho62i he65 tho61ng mo64i la62n kho73i d9o65ng"
            ForeColor       =   0
            ShowFocusRectangle=   0   'False
         End
         Begin FVUnicodeControl.FVistaUniCheckbox CHK0 
            Height          =   195
            Index           =   0
            Left            =   240
            TabIndex        =   80
            Top             =   360
            Width           =   2085
            _ExtentX        =   3678
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
            Caption         =   "Kho73i d9o65ng cu2ng Windows"
            ForeColor       =   0
            ShowFocusRectangle=   0   'False
         End
         Begin VB.Image Image2 
            Height          =   1080
            Left            =   4800
            Picture         =   "frmMain.frx":12072
            Top             =   0
            Width           =   1080
         End
      End
   End
   Begin FVUnicodeControl.FVistaUniFrame FM 
      Height          =   135
      Index           =   13
      Left            =   8880
      TabIndex        =   26
      Top             =   5640
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   238
      Alignment       =   0
      BackColor       =   16777215
      ForeColor       =   255
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Que1t ma64u"
      AutoUnicode     =   -1  'True
      Begin FVUnicodeControl.FVistaUniButton cmdVirusRemoveAll 
         Height          =   615
         Left            =   2160
         TabIndex        =   75
         Top             =   1800
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   1085
         BackColor       =   14737632
         ButtonStyle     =   3
         Caption         =   "Que1t ma64u Virus"
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
      Begin UniControls.UniLabel UniLabel18 
         Height          =   375
         Left            =   120
         Top             =   240
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   661
         Caption         =   "Que1t Virus vo71i ma64u co1 sa84n, giu1p loa5i bo3 ta65n go61c Virus"
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
      Begin VB.Image Image3 
         Height          =   720
         Left            =   1320
         Picture         =   "frmMain.frx":130DF
         Top             =   1800
         Width           =   720
      End
   End
   Begin FVUnicodeControl.FVistaUniFrame FM 
      Height          =   135
      Index           =   12
      Left            =   8880
      TabIndex        =   25
      Top             =   5520
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   238
      Alignment       =   0
      BackColor       =   16777215
      ForeColor       =   255
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Kie63m tra he65 tho61ng"
      AutoUnicode     =   -1  'True
      Begin FVUnicodeControl.FVistaUniButton cmdPerfectSystemReporter 
         Height          =   615
         Left            =   2160
         TabIndex        =   74
         Top             =   1800
         Width           =   2505
         _ExtentX        =   4419
         _ExtentY        =   1085
         BackColor       =   14737632
         ButtonStyle     =   3
         Caption         =   "Kie63m tra he65 tho61ng"
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
      Begin UniControls.UniLabel UniLabel17 
         Height          =   375
         Left            =   120
         Top             =   240
         Width           =   6495
         _ExtentX        =   11456
         _ExtentY        =   661
         Caption         =   "Ta5o file kie63m tra ti2nh tra5ng hie65n ta5i cu3a ma1y ti1nh"
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
      Begin VB.Image Image4 
         Height          =   720
         Left            =   1320
         Picture         =   "frmMain.frx":29A91
         Top             =   1800
         Width           =   720
      End
   End
   Begin FVUnicodeControl.FVistaUniFrame FM 
      Height          =   135
      Index           =   11
      Left            =   8880
      TabIndex        =   24
      Top             =   5400
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   238
      Alignment       =   0
      BackColor       =   16777215
      ForeColor       =   255
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Qua3n ly1 kho73i d9o65ng"
      AutoUnicode     =   -1  'True
      Begin UniControls.UniLabel UniLabel16 
         Height          =   375
         Left            =   120
         Top             =   240
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   661
         Caption         =   "Qua3n ly1 ca1c chu7o7ng tri2nh kho73i d9o65ng cu2ng he65 tho61ng"
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
      Begin FVUnicodeControl.FVistaUniButton cmdPerfectStartUpManager 
         Height          =   615
         Left            =   2160
         TabIndex        =   73
         Top             =   1800
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   1085
         BackColor       =   14737632
         ButtonStyle     =   3
         Caption         =   "Qua3n ly1 kho73i d9o65ng"
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
      Begin VB.Image Image5 
         Height          =   720
         Left            =   1320
         Picture         =   "frmMain.frx":2A468
         Top             =   1680
         Width           =   720
      End
   End
   Begin FVUnicodeControl.FVistaUniFrame FM 
      Height          =   135
      Index           =   10
      Left            =   8880
      TabIndex        =   21
      Top             =   5280
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   238
      Alignment       =   0
      BackColor       =   16777215
      ForeColor       =   255
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Qua3n ly1 tie61n tri2nh"
      AutoUnicode     =   -1  'True
      Begin FVUnicodeControl.FVistaUniButton cmdPerfectProcessManager 
         Height          =   615
         Left            =   2160
         TabIndex        =   72
         Top             =   1800
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   1085
         BackColor       =   14737632
         ButtonStyle     =   3
         Caption         =   "Qua3n ly1 tie61n tri2nh"
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
      Begin UniControls.UniLabel UniLabel15 
         Height          =   375
         Left            =   120
         Top             =   240
         Width           =   6495
         _ExtentX        =   11456
         _ExtentY        =   661
         Caption         =   "Qua3n ly1 ca1c chu7o7ng tri2nh d9ang cha5y trong he65 tho61ng"
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
      Begin VB.Image Image6 
         Height          =   720
         Left            =   1320
         Picture         =   "frmMain.frx":2AB68
         Top             =   1800
         Width           =   720
      End
   End
   Begin FVUnicodeControl.FVistaUniFrame FM 
      Height          =   615
      Index           =   9
      Left            =   8880
      TabIndex        =   20
      Top             =   5160
      Visible         =   0   'False
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   1085
      Alignment       =   0
      BackColor       =   16777215
      ForeColor       =   255
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Ta8ng to61c ma1y ti1nh"
      AutoUnicode     =   -1  'True
      Begin UniControls.UniLabel UniLabel21 
         Height          =   855
         Left            =   840
         Top             =   840
         Width           =   5055
         _ExtentX        =   8916
         _ExtentY        =   1508
         Alignment       =   1
         Caption         =   "Ha4y su73 du5ng chu71c na8ng na2y 1 tua62n 1 la62n d9e63 d9a3m ba3o ma1y ti1nh hoa5t d9o65ng nhanh va2 o63n d9i5nh"
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
      Begin UniControls.UniLabel UniLabel20 
         Height          =   495
         Left            =   120
         Top             =   240
         Width           =   6375
         _ExtentX        =   11245
         _ExtentY        =   873
         Caption         =   "Ta8ng to61c ma1y ti1nh ba82ng ca1c do5n de5p ra1c, chi3nh su73a Registry..."
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
      Begin FVUnicodeControl.FVistaUniButton cmdTangTocMayTinh 
         Height          =   735
         Left            =   2400
         TabIndex        =   77
         Top             =   1920
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   1296
         BackColor       =   14737632
         ButtonStyle     =   3
         Caption         =   "Ta8ng to61c ma1y ti1nh"
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
      Begin VB.Image Image7 
         Height          =   720
         Left            =   1560
         Picture         =   "frmMain.frx":2B4E9
         Top             =   1920
         Width           =   720
      End
   End
   Begin FVUnicodeControl.FVistaUniFrame FM 
      Height          =   495
      Index           =   8
      Left            =   8880
      TabIndex        =   19
      Top             =   5040
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   873
      Alignment       =   0
      BackColor       =   16777215
      ForeColor       =   255
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Tinh chi3nh Registry"
      AutoUnicode     =   -1  'True
      Begin FVUnicodeControl.FVistaUniButton cmdTinhChinhRegistry 
         Height          =   735
         Left            =   2400
         TabIndex        =   76
         Top             =   1920
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   1296
         BackColor       =   14737632
         ButtonStyle     =   3
         Caption         =   "Tinh chi3nh Registry"
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
      Begin UniControls.UniLabel UniLabel19 
         Height          =   615
         Left            =   240
         Top             =   360
         Width           =   6135
         _ExtentX        =   10821
         _ExtentY        =   1085
         Caption         =   "Thu75c hie65n ca1c tinh chi3nh, mo73 kho1a ca1c chu71c na8ng cu3a he65 d9ie62u ha2nh tho6ng qua Registry"
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
      Begin VB.Image Image8 
         Height          =   690
         Left            =   1440
         Picture         =   "frmMain.frx":2BEBB
         Top             =   1920
         Width           =   720
      End
   End
   Begin FVUnicodeControl.FVistaUniFrame FM 
      Height          =   615
      Index           =   16
      Left            =   8880
      TabIndex        =   29
      Top             =   4920
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   1085
      Alignment       =   0
      BackColor       =   16777215
      ForeColor       =   255
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Phu5c ho62i du74 lie65u"
      AutoUnicode     =   -1  'True
      Begin FVUnicodeControl.FVistaUniButton cmdPhucHoiDuLieu 
         Height          =   615
         Left            =   1200
         TabIndex        =   71
         Top             =   2040
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   1085
         BackColor       =   14737632
         ButtonStyle     =   3
         Caption         =   "Nha61n va2o d9a6y d9e63 su73 du5ng chu71c na8ng phu5c ho62i du74 lie65u"
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
      Begin UniControls.UniLabel UniLabel14 
         Height          =   495
         Left            =   120
         Top             =   360
         Width           =   6495
         _ExtentX        =   11456
         _ExtentY        =   873
         Caption         =   "Chu71c na8ng giu1p phu5c ho62i du74 lie65u bi5 ma61t sau khi die65t Virus."
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
      Begin VB.Image Image9 
         Height          =   960
         Left            =   240
         Picture         =   "frmMain.frx":2C7FB
         Top             =   1800
         Width           =   960
      End
   End
   Begin FVUnicodeControl.FVistaUniFrame FM 
      Height          =   495
      Index           =   15
      Left            =   8880
      TabIndex        =   28
      Top             =   4800
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   873
      Alignment       =   0
      BackColor       =   16777215
      ForeColor       =   255
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Ba3o ve65 Autorun"
      AutoUnicode     =   -1  'True
      Begin UniControls.UniLabel lblAT 
         Height          =   375
         Left            =   1920
         Top             =   1920
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   661
         Alignment       =   1
         BackStyle       =   0
         Caption         =   "(D9ang ba65t)"
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
      Begin UniControls.UniLabel UniLabel13 
         Height          =   375
         Left            =   120
         Top             =   240
         Width           =   6495
         _ExtentX        =   11456
         _ExtentY        =   661
         Caption         =   "Che61 d9o65 tu75 d9o65ng ba3o ve65 va2 xo1a Autorun."
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
      Begin FVUnicodeControl.FVistaUniButton cmdOnOffAT 
         Height          =   375
         Index           =   0
         Left            =   1920
         TabIndex        =   69
         Top             =   2280
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   661
         BackColor       =   14737632
         ButtonStyle     =   3
         Caption         =   "Ba65t che61 d9o65 ba3o ve65 Autorun"
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
      Begin FVUnicodeControl.FVistaUniButton cmdOnOffAT 
         Height          =   375
         Index           =   1
         Left            =   1920
         TabIndex        =   70
         Top             =   2640
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   661
         BackColor       =   14737632
         ButtonStyle     =   3
         Caption         =   "Ta81t che61 d9o65 ba3o ve65 Autorun"
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
      Begin VB.Image Image10 
         Height          =   720
         Left            =   1080
         Picture         =   "frmMain.frx":2D399
         Top             =   2280
         Width           =   720
      End
   End
   Begin FVUnicodeControl.FVistaUniFrame FM 
      Height          =   975
      Index           =   14
      Left            =   8880
      TabIndex        =   27
      Top             =   4680
      Visible         =   0   'False
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   1720
      Alignment       =   0
      BackColor       =   16777215
      ForeColor       =   255
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Ba3o ve65 tho72i gian thu75c"
      AutoUnicode     =   -1  'True
      Begin UniControls.UniLabel lblRTP 
         Height          =   375
         Left            =   1920
         Top             =   1920
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   661
         Alignment       =   1
         BackStyle       =   0
         Caption         =   "(D9ang ba65t)"
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
      Begin FVUnicodeControl.FVistaUniButton cmdOnOffRTP 
         Height          =   375
         Index           =   1
         Left            =   1920
         TabIndex        =   68
         Top             =   2640
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   661
         BackColor       =   14737632
         ButtonStyle     =   3
         Caption         =   "Ta81t che61 d9o65 ba3o ve65 tho72i gian thu75c"
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
      Begin UniControls.UniLabel UniLabel12 
         Height          =   615
         Left            =   1200
         Top             =   1320
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   1085
         Alignment       =   1
         Caption         =   "Khi pha1t hie65n Virus, chu7o7ng tri2nh se4 tho6ng ba1o ra giu74a ma2n hi2nh cho ba5n bie61t ki5p tho72i"
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
      End
      Begin FVUnicodeControl.FVistaUniButton cmdOnOffRTP 
         Height          =   375
         Index           =   0
         Left            =   1920
         TabIndex        =   67
         Top             =   2280
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   661
         BackColor       =   14737632
         ButtonStyle     =   3
         Caption         =   "Ba65t che61 d9o65 ba3o ve65 tho72i gian thu75c"
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
      Begin UniControls.UniLabel UniLabel11 
         Height          =   735
         Left            =   120
         Top             =   240
         Width           =   6495
         _ExtentX        =   11456
         _ExtentY        =   1296
         Caption         =   $"frmMain.frx":2DC18
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
      Begin VB.Image Image11 
         Height          =   720
         Left            =   1080
         Picture         =   "frmMain.frx":2DCDF
         Top             =   2280
         Width           =   720
      End
   End
   Begin FVUnicodeControl.FVistaUniFrame FM 
      Height          =   495
      Index           =   7
      Left            =   8880
      TabIndex        =   18
      Top             =   4560
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   873
      Alignment       =   0
      BackColor       =   16777215
      ForeColor       =   255
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Danh sách tin tu7o73ng"
      AutoUnicode     =   -1  'True
      Begin FVUnicodeControl.FVistaUniButton cmdGetTinTuong 
         Height          =   255
         Left            =   5400
         TabIndex        =   66
         Top             =   240
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   450
         BackColor       =   14737632
         ButtonStyle     =   3
         Caption         =   "La61y tho6ng tin"
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
      Begin FVUnicodeControl.FVistaUniButton cmdDeleteAllTinTuong 
         Height          =   375
         Left            =   3600
         TabIndex        =   65
         Top             =   3960
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         BackColor       =   14737632
         ButtonStyle     =   3
         Caption         =   "Xo1a he61t"
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
      Begin FVUnicodeControl.FVistaUniButton cmdAddFolderTinTuong 
         Height          =   375
         Left            =   1920
         TabIndex        =   64
         Top             =   3960
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         BackColor       =   14737632
         ButtonStyle     =   3
         Caption         =   "The6m va2o thu7 mu5c"
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
      Begin FVUnicodeControl.FVistaUniButton cmdAddFileTinTuong 
         Height          =   375
         Left            =   120
         TabIndex        =   63
         Top             =   3960
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         BackColor       =   14737632
         ButtonStyle     =   3
         Caption         =   "The6m va2o ta65p tin"
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
      Begin UniControls.UniListBox lstTinTuong 
         Height          =   3255
         Left            =   120
         TabIndex        =   62
         Top             =   600
         Width           =   6495
         _ExtentX        =   11456
         _ExtentY        =   5741
         IconMaskColor   =   16711935
         AutoUnicode     =   0   'False
         Picture         =   "frmMain.frx":2E55E
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   3
         RowHeight       =   19
         FlatScrollBar   =   -1  'True
         AutoHideScrollBars=   -1  'True
      End
      Begin UniControls.UniLabel UniLabel10 
         Height          =   375
         Left            =   120
         Top             =   240
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   661
         BackStyle       =   0
         Caption         =   "Danh sa1ch nhu74ng File tin tu7o73ng"
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
   Begin FVUnicodeControl.FVistaUniFrame FM 
      Height          =   615
      Index           =   6
      Left            =   8880
      TabIndex        =   17
      Top             =   4440
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   1085
      Alignment       =   0
      BackColor       =   16777215
      ForeColor       =   255
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Du74 lie65u Virus"
      AutoUnicode     =   -1  'True
      Begin FVUnicodeControl.FVistaUniButton cmdGetListUserData 
         Height          =   255
         Left            =   5400
         TabIndex        =   61
         Top             =   240
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   450
         BackColor       =   14737632
         ButtonStyle     =   3
         Caption         =   "La61y tho6ng tin"
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
      Begin FVUnicodeControl.FVistaUniButton cmdDelData 
         Height          =   375
         Left            =   1680
         TabIndex        =   60
         Top             =   3960
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         BackColor       =   14737632
         ButtonStyle     =   3
         Caption         =   "Xo1a ma64u Virus"
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
      Begin FVUnicodeControl.FVistaUniButton cmdAddData 
         Height          =   375
         Left            =   120
         TabIndex        =   59
         Top             =   3960
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         BackColor       =   14737632
         ButtonStyle     =   3
         Caption         =   "The6m ma64u Virus"
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
      Begin UniControls.UniListBox lstUserData 
         Height          =   3255
         Left            =   120
         TabIndex        =   58
         Top             =   600
         Width           =   6495
         _ExtentX        =   11456
         _ExtentY        =   5741
         IconMaskColor   =   16711935
         Picture         =   "frmMain.frx":2E57A
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   3
         RowHeight       =   19
         FlatScrollBar   =   -1  'True
         AutoHideScrollBars=   -1  'True
      End
      Begin UniControls.UniLabel UniLabel9 
         Height          =   255
         Left            =   240
         Top             =   240
         Width           =   5175
         _ExtentX        =   9128
         _ExtentY        =   450
         Caption         =   "Danh sa1ch ca1c ma64u Virus do ngu7o72i du2ng the6m va2o:"
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
   Begin FVUnicodeControl.FVistaUniFrame FM 
      Height          =   375
      Index           =   5
      Left            =   8880
      TabIndex        =   16
      Top             =   4320
      Visible         =   0   'False
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   661
      Alignment       =   0
      BackColor       =   16777215
      ForeColor       =   255
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Ca61u hi2nh que1t"
      AutoUnicode     =   -1  'True
      Begin UniControls.UniLabel UniLabel8 
         Height          =   255
         Left            =   120
         Top             =   3840
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   450
         Caption         =   "Nha61n va2o nu1t Lu7u la5i d9e63 thay d9o63i co1 hie65u lu75c"
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
      Begin FVUnicodeControl.FVistaUniButton cmdSaveCauHinh 
         Height          =   375
         Left            =   4560
         TabIndex        =   57
         Top             =   3840
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         BackColor       =   -2147483633
         ButtonStyle     =   3
         Caption         =   "Lu7u la5i"
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
      Begin FVUnicodeControl.FVistaUniFrame FMThietLapKhiQuet 
         Height          =   1815
         Left            =   120
         TabIndex        =   51
         Top             =   1920
         Width           =   5175
         _ExtentX        =   9128
         _ExtentY        =   3201
         Alignment       =   0
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Ca1c thie61t la65p khi que1t"
         AutoUnicode     =   -1  'True
         Begin FVUnicodeControl.FVistaUniCheckbox Chk2 
            Height          =   195
            Index           =   4
            Left            =   240
            TabIndex        =   56
            Top             =   1200
            Width           =   3885
            _ExtentX        =   6853
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
            Caption         =   "Phu5c ho62i ca1c kho1a Registry bi5 hu7 ho3ng sau khi die65t"
            ForeColor       =   0
            ShowFocusRectangle=   0   'False
         End
         Begin FVUnicodeControl.FVistaUniCheckbox Chk2 
            Height          =   195
            Index           =   3
            Left            =   240
            TabIndex        =   55
            Top             =   960
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
            Caption         =   "Ta5o file ba1o ca1o sau khi die65t Virus"
            ForeColor       =   0
            ShowFocusRectangle=   0   'False
         End
         Begin FVUnicodeControl.FVistaUniCheckbox Chk2 
            Height          =   195
            Index           =   2
            Left            =   240
            TabIndex        =   54
            Top             =   720
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
            Caption         =   "Bo3 qua nhu74ng file co1 dung lu7o75ng lo71n ho7n 10MB"
            ForeColor       =   0
            ShowFocusRectangle=   0   'False
         End
         Begin FVUnicodeControl.FVistaUniCheckbox Chk2 
            Height          =   195
            Index           =   1
            Left            =   240
            TabIndex        =   53
            Top             =   480
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
            Caption         =   "Que1t ca1c tie61n tri2nh d9ang cha5y trong bo65 nho71"
            ForeColor       =   0
            ShowFocusRectangle=   0   'False
         End
         Begin FVUnicodeControl.FVistaUniCheckbox Chk2 
            Height          =   195
            Index           =   0
            Left            =   240
            TabIndex        =   52
            Top             =   240
            Width           =   3015
            _ExtentX        =   5318
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
            Caption         =   "Que1t ca1c file kho73i d9o65ng cu2ng he65 tho61ng"
            ForeColor       =   0
            ShowFocusRectangle=   0   'False
         End
      End
      Begin FVUnicodeControl.FVistaUniFrame fmDuoiFile 
         Height          =   975
         Left            =   120
         TabIndex        =   47
         Top             =   840
         Width           =   5175
         _ExtentX        =   9128
         _ExtentY        =   1720
         Alignment       =   0
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "D9uo6i file se4 que1t"
         AutoUnicode     =   -1  'True
         Begin VB.ComboBox cboEXT2 
            Height          =   315
            Left            =   2160
            Style           =   2  'Dropdown List
            TabIndex        =   50
            Top             =   480
            Width           =   2895
         End
         Begin FVUnicodeControl.FVistaUniOption optEXT2 
            Height          =   195
            Index           =   1
            Left            =   240
            TabIndex        =   49
            Top             =   480
            Width           =   1770
            _ExtentX        =   3122
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
            ShowFocusRectangle=   0   'False
            Caption         =   "Chi3 que1t file co1 d9uo6i:"
            ForeColor       =   0
         End
         Begin FVUnicodeControl.FVistaUniOption optEXT2 
            Height          =   195
            Index           =   0
            Left            =   240
            TabIndex        =   48
            Top             =   240
            Width           =   2325
            _ExtentX        =   4101
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
            ShowFocusRectangle=   0   'False
            Caption         =   "Ta61t ca3 ca1c loa5i d9uo6i file (*.*)"
            ForeColor       =   0
         End
      End
      Begin UniControls.UniLabel UniLabel7 
         Height          =   495
         Left            =   240
         Top             =   240
         Width           =   6255
         _ExtentX        =   11033
         _ExtentY        =   873
         Caption         =   "Ca1c ca61u hi2nh ba5n lu7u o73 d9a6y se4 d9u7o75c lu7u la5i o73 bu7o71c ""Ca2i d9a85t ca61u hi2nh"" khi que1t Virus."
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
   Begin FVUnicodeControl.FVistaUniFrame FM 
      Height          =   495
      Index           =   2
      Left            =   8880
      TabIndex        =   13
      Top             =   4200
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   873
      Alignment       =   0
      BackColor       =   16777215
      ForeColor       =   255
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Nha65t ky1 que1t"
      AutoUnicode     =   -1  'True
      Begin FVUnicodeControl.FVistaUniButton cmdXemChiTiet 
         Height          =   375
         Left            =   4920
         TabIndex        =   46
         Top             =   3960
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         BackColor       =   14737632
         ButtonStyle     =   3
         Caption         =   "Xem Chi Tie61t"
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
      Begin FVUnicodeControl.FVistaUniButton cmdLayThongTinNhatKy 
         Height          =   255
         Left            =   5400
         TabIndex        =   45
         Top             =   240
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   450
         BackColor       =   14737632
         ButtonStyle     =   3
         Caption         =   "La61y tho6ng tin"
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
      Begin UniControls.UniListBox lstNhatKy 
         Height          =   3255
         Left            =   120
         TabIndex        =   44
         Top             =   600
         Width           =   6495
         _ExtentX        =   11456
         _ExtentY        =   5741
         IconMaskColor   =   16711935
         Picture         =   "frmMain.frx":2E596
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   3
         SortOrder       =   -1
         RowHeight       =   19
         GridLines       =   -1  'True
         GridColor       =   16761024
         FlatScrollBar   =   -1  'True
         AutoHideScrollBars=   -1  'True
      End
      Begin UniControls.UniLabel UniLabel2 
         Height          =   375
         Left            =   120
         Top             =   240
         Width           =   6495
         _ExtentX        =   11456
         _ExtentY        =   661
         Caption         =   "Danh sa1ch ca1c la62n que1t Virus tru7o71c d9a6y"
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
   Begin FVUnicodeControl.FVistaUniFrame FM 
      Height          =   495
      Index           =   1
      Left            =   8880
      TabIndex        =   12
      Top             =   4080
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   873
      Alignment       =   0
      BackColor       =   16777215
      ForeColor       =   255
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Vu2ng ca1ch ly"
      AutoUnicode     =   -1  'True
      Begin FVUnicodeControl.FVistaUniButton cmdLayThongTinCachLy 
         Height          =   255
         Left            =   5400
         TabIndex        =   43
         Top             =   240
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   450
         BackColor       =   14737632
         ButtonStyle     =   3
         Caption         =   "La61y tho6ng tin"
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
      Begin FVUnicodeControl.FVistaUniButton cmdXoaVirusDaChon 
         Height          =   375
         Left            =   1920
         TabIndex        =   42
         Top             =   3960
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         BackColor       =   14737632
         ButtonStyle     =   3
         Caption         =   "Xo1a Virus d9a4 cho5n"
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
      Begin FVUnicodeControl.FVistaUniButton cmdXoaTatCa 
         Height          =   375
         Left            =   4800
         TabIndex        =   41
         Top             =   3960
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         BackColor       =   14737632
         ButtonStyle     =   3
         Caption         =   "Xo1a ta61t ca3"
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
      Begin FVUnicodeControl.FVistaUniButton cmdPhucHoiTatCa 
         Height          =   375
         Left            =   3480
         TabIndex        =   40
         Top             =   3960
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         BackColor       =   14737632
         ButtonStyle     =   3
         Caption         =   "Phu5c ho62i ta61t ca3"
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
      Begin FVUnicodeControl.FVistaUniButton cmdPhucHoiVirusDaChon 
         Height          =   375
         Left            =   120
         TabIndex        =   39
         Top             =   3960
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         BackColor       =   14737632
         ButtonStyle     =   3
         Caption         =   "Phu5c ho62i virus d9a4 cho5n"
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
      Begin UniControls.UniLabel UniLabel1 
         Height          =   375
         Left            =   120
         Top             =   240
         Width           =   6255
         _ExtentX        =   11033
         _ExtentY        =   661
         BackStyle       =   0
         Caption         =   "Danh sa1ch nhu74ng Virus d9a4 d9u7o75c ma4 ho1a va2 ca1ch ly"
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
      Begin UniControls.UniListBox lstCachLy 
         Height          =   3255
         Left            =   120
         TabIndex        =   30
         Top             =   600
         Width           =   6495
         _ExtentX        =   11456
         _ExtentY        =   5741
         IconMaskColor   =   16711935
         AutoUnicode     =   0   'False
         Picture         =   "frmMain.frx":2E5B2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   3
         RowHeight       =   23
         GridLines       =   -1  'True
         GridColor       =   16761024
         FlatScrollBar   =   -1  'True
         AutoHideScrollBars=   -1  'True
      End
   End
   Begin FVUnicodeControl.FVistaUniFrame FM 
      Height          =   135
      Index           =   0
      Left            =   8880
      TabIndex        =   11
      Top             =   3960
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   238
      Alignment       =   0
      BackColor       =   16777215
      ForeColor       =   255
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Que1t Virus"
      AutoUnicode     =   -1  'True
      Begin UniControls.UniLabel UniLabel6 
         Height          =   255
         Left            =   240
         Top             =   240
         Width           =   6255
         _ExtentX        =   11033
         _ExtentY        =   450
         BackStyle       =   0
         Caption         =   "Nha61n va2o nu1t be6n du7o71i d9e63 d9e61n chu7o7ng tri2nh que1t"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Link            =   ""
      End
      Begin FVUnicodeControl.FVistaUniButton Scan 
         Height          =   495
         Left            =   1680
         TabIndex        =   38
         Top             =   2520
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   873
         BackColor       =   12632319
         ButtonStyle     =   3
         Caption         =   ">> Nha61n va2o d9a6y d9e63 que1t <<"
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
      Begin VB.Image Image12 
         Height          =   1920
         Left            =   2400
         Picture         =   "frmMain.frx":2E5CE
         Top             =   600
         Width           =   1920
      End
   End
   Begin VB.FileListBox File1 
      Height          =   285
      Hidden          =   -1  'True
      Left            =   600
      System          =   -1  'True
      TabIndex        =   10
      Top             =   0
      Visible         =   0   'False
      Width           =   615
   End
   Begin FVUnicodeControl.FVistaUniButton Cmd 
      Height          =   615
      Index           =   2
      Left            =   0
      TabIndex        =   5
      Top             =   2400
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   1085
      BackColor       =   14737632
      ButtonShape     =   3
      ButtonStyle     =   3
      Caption         =   "Ca61u hi2nh que1t"
      Effects         =   3
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
   Begin FVUnicodeControl.FVistaUniButton Cmd 
      Height          =   615
      Index           =   1
      Left            =   0
      TabIndex        =   4
      Top             =   1800
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   1085
      BackColor       =   14737632
      ButtonShape     =   3
      ButtonStyle     =   3
      Caption         =   "Que1t Virus"
      Effects         =   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   192
   End
   Begin FVUnicodeControl.FVistaUniButton Cmd 
      Height          =   615
      Index           =   0
      Left            =   0
      TabIndex        =   3
      Top             =   1200
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   1085
      BackColor       =   16761024
      ButtonShape     =   3
      ButtonStyle     =   3
      Caption         =   "To63ng qua1t"
      Effects         =   3
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
   Begin FVUnicodeControl.FVistaUniButton Cmd 
      Height          =   615
      Index           =   4
      Left            =   0
      TabIndex        =   6
      Top             =   3600
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   1085
      BackColor       =   14737632
      ButtonShape     =   3
      ButtonStyle     =   3
      Caption         =   "To61i u7u he65 tho61ng"
      Effects         =   3
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
   Begin FVUnicodeControl.FVistaUniButton Cmd 
      Height          =   615
      Index           =   5
      Left            =   0
      TabIndex        =   7
      Top             =   4200
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   1085
      BackColor       =   14737632
      ButtonShape     =   3
      ButtonStyle     =   3
      Caption         =   "Co6ng cu5 ho64 tro75"
      Effects         =   3
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
   Begin FVUnicodeControl.FVistaUniFrame FM 
      Height          =   615
      Index           =   4
      Left            =   8880
      TabIndex        =   15
      Top             =   3840
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   1085
      Alignment       =   0
      BackColor       =   16777215
      ForeColor       =   255
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Tho6ng tin ma1y ti1nh"
      AutoUnicode     =   -1  'True
   End
   Begin FVUnicodeControl.FVistaUniButton Cmd 
      Height          =   615
      Index           =   3
      Left            =   0
      TabIndex        =   22
      Top             =   3000
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   1085
      BackColor       =   14737632
      ButtonShape     =   3
      ButtonStyle     =   3
      Caption         =   "Ba3o ve65 ma1y ti1nh"
      Effects         =   3
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
   Begin FVUnicodeControl.FVistaUniFrame FM 
      Height          =   375
      Index           =   3
      Left            =   8880
      TabIndex        =   14
      Top             =   3720
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      Alignment       =   0
      BackColor       =   16777215
      ForeColor       =   255
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Tho6ng tin chung"
      AutoUnicode     =   -1  'True
      Begin UniControls.UniLabel UniLabel70 
         Height          =   255
         Left            =   720
         Top             =   2160
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   450
         BackStyle       =   0
         Caption         =   "Ba3o ve65 tho72i gian thu75c:"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   12582912
      End
      Begin UniControls.UniLabel UniLabel74 
         Height          =   255
         Left            =   720
         Top             =   4080
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   450
         BackStyle       =   0
         Caption         =   "Tu75 d9o65ng ca65p nha65t:"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   12582912
      End
      Begin UniControls.UniLabel UniLabel73 
         Height          =   255
         Left            =   720
         Top             =   3600
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   450
         BackStyle       =   0
         Caption         =   "Kho73i d9o65ng cu2ng Windows:"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   12582912
      End
      Begin UniControls.UniLabel UniLabel72 
         Height          =   255
         Left            =   720
         Top             =   3120
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   450
         BackStyle       =   0
         Caption         =   "Tu75 d9o65ng kie63m tra ca1c ke61t no61i va2o ma1y ti1nh:"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   12582912
      End
      Begin UniControls.UniLabel UniLabel71 
         Height          =   255
         Left            =   720
         Top             =   2640
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   450
         BackStyle       =   0
         Caption         =   "Ba3o ve65 USB va2 die65t Autorun:"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   12582912
      End
      Begin UniControls.UniLabel lbl1TinhTrangMayTinh 
         Height          =   255
         Left            =   600
         Top             =   1560
         Width           =   5295
         _ExtentX        =   9340
         _ExtentY        =   450
         Alignment       =   1
         BackStyle       =   0
         Caption         =   "Ma1y hoa5t d9o65ng to61t!"
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
      Begin UniControls.UniLabel lbl1Memory 
         Height          =   255
         Left            =   3240
         Top             =   1320
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   450
         Caption         =   "Total Memory"
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
      Begin UniControls.UniLabel lbl1Process 
         Height          =   255
         Left            =   3240
         Top             =   1080
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   450
         Caption         =   "Total Process"
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
      Begin UniControls.UniLabel lbl1UserName 
         Height          =   255
         Left            =   3240
         Top             =   840
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   450
         AutoUnicode     =   0   'False
         Caption         =   "User Name"
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
      Begin UniControls.UniLabel lbl1ComputerName 
         Height          =   255
         Left            =   3240
         Top             =   600
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   450
         AutoUnicode     =   0   'False
         Caption         =   "Computer Name"
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
      Begin UniControls.UniLabel UniLabel69 
         Height          =   255
         Left            =   600
         Top             =   1320
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   450
         Alignment       =   2
         Caption         =   "To63ng dung lu7o7ng bo65 nho71 RAM:"
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
      Begin UniControls.UniLabel UniLabel61 
         Height          =   255
         Left            =   600
         Top             =   1080
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   450
         Alignment       =   2
         Caption         =   "To63ng so61 chu7o7ng tri2nh d9ang cha5y:"
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
      Begin UniControls.UniLabel UniLabel5 
         Height          =   255
         Left            =   1080
         Top             =   840
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   450
         Alignment       =   2
         Caption         =   "Te6n ngu7o72i du2ng:"
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
         Left            =   1440
         Top             =   600
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   450
         Alignment       =   2
         Caption         =   "Te6n ma1y ti1nh:"
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
      Begin FVUnicodeControl.FVistaUniProgressbar ProTinhTrang 
         Height          =   225
         Left            =   600
         Top             =   1800
         Width           =   5295
         _ExtentX        =   9340
         _ExtentY        =   397
         Max             =   100
         Value           =   0
         TStyle          =   2
         Min             =   0
         Style           =   1
         Text            =   "To61c d9o65 xu73 ly1 hie65n ta5i:"
         Align           =   1
      End
      Begin UniControls.UniLabel UniLabel3 
         Height          =   375
         Left            =   1320
         Top             =   120
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   661
         Alignment       =   1
         BackStyle       =   0
         Caption         =   "Tho6ng tin chung"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   16576
      End
      Begin VB.Image img0 
         Height          =   495
         Index           =   4
         Left            =   5160
         Top             =   3960
         Width           =   975
      End
      Begin VB.Image img0 
         Height          =   495
         Index           =   3
         Left            =   5160
         Top             =   3480
         Width           =   975
      End
      Begin VB.Image img0 
         Height          =   495
         Index           =   2
         Left            =   5160
         Top             =   3000
         Width           =   975
      End
      Begin VB.Image img0 
         Height          =   495
         Index           =   1
         Left            =   5160
         Top             =   2520
         Width           =   975
      End
      Begin VB.Image img0 
         Height          =   495
         Index           =   0
         Left            =   5160
         Top             =   2040
         Width           =   975
      End
      Begin VB.Line Line11 
         X1              =   2400
         X2              =   5040
         Y1              =   4200
         Y2              =   4200
      End
      Begin VB.Line Line10 
         X1              =   3000
         X2              =   5040
         Y1              =   3720
         Y2              =   3720
      End
      Begin VB.Line Line9 
         X1              =   4440
         X2              =   5040
         Y1              =   3240
         Y2              =   3240
      End
      Begin VB.Line Line8 
         X1              =   3360
         X2              =   5040
         Y1              =   2760
         Y2              =   2760
      End
      Begin VB.Line Line7 
         X1              =   2760
         X2              =   5040
         Y1              =   2280
         Y2              =   2280
      End
   End
   Begin FVUnicodeControl.FVistaUniButton Cmd 
      Height          =   615
      Index           =   6
      Left            =   0
      TabIndex        =   31
      Top             =   4800
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   1085
      BackColor       =   14737632
      ButtonShape     =   3
      ButtonStyle     =   3
      Caption         =   "Ca2i d9a85t chung"
      Effects         =   3
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
   Begin FVUnicodeControl.FVistaUniButton Cmd 
      Height          =   615
      Index           =   7
      Left            =   0
      TabIndex        =   32
      Top             =   5400
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   1085
      BackColor       =   14737632
      ButtonShape     =   3
      ButtonStyle     =   3
      Caption         =   "Tho6ng tin"
      Effects         =   3
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
   Begin PerfectAntivirus.WM11ToolBar MenuBar 
      Height          =   555
      Index           =   7
      Left            =   1920
      TabIndex        =   34
      Top             =   5400
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   979
      DDBCaption      =   ""
   End
   Begin PerfectAntivirus.WM11ToolBar MenuBar 
      Height          =   555
      Index           =   6
      Left            =   1920
      TabIndex        =   33
      Top             =   4800
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   979
      DDBCaption      =   ""
   End
   Begin PerfectAntivirus.WM11ToolBar MenuBar 
      Height          =   555
      Index           =   3
      Left            =   1920
      TabIndex        =   23
      Top             =   3000
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   979
      DDBCaption      =   ""
   End
   Begin PerfectAntivirus.WM11ToolBar MenuBar 
      Height          =   555
      Index           =   1
      Left            =   1920
      TabIndex        =   0
      Top             =   1800
      Width           =   285
      _ExtentX        =   503
      _ExtentY        =   979
      DDBCaption      =   ""
   End
   Begin PerfectAntivirus.WM11ToolBar MenuBar 
      Height          =   555
      Index           =   2
      Left            =   1920
      TabIndex        =   1
      Top             =   2400
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   979
      DDBCaption      =   ""
   End
   Begin PerfectAntivirus.WM11ToolBar MenuBar 
      Height          =   555
      Index           =   0
      Left            =   1920
      TabIndex        =   2
      Top             =   1200
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   979
      DDBCaption      =   ""
   End
   Begin PerfectAntivirus.WM11ToolBar MenuBar 
      Height          =   555
      Index           =   4
      Left            =   1920
      TabIndex        =   8
      Top             =   3600
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   979
      DDBCaption      =   ""
   End
   Begin PerfectAntivirus.WM11ToolBar MenuBar 
      Height          =   555
      Index           =   5
      Left            =   1920
      TabIndex        =   9
      Top             =   4200
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   979
      DDBCaption      =   ""
   End
   Begin FVUnicodeControl.FVistaUniLabel FVistaUniLabel2 
      Height          =   255
      Left            =   6960
      TabIndex        =   92
      Top             =   990
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   450
      BackStyle       =   0
      Caption         =   "2.0.9 (07/06/2011)"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   255
   End
   Begin VB.Image IMGON 
      Height          =   435
      Left            =   0
      Picture         =   "frmMain.frx":2FEEF
      Stretch         =   -1  'True
      Top             =   0
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Image IMGOFF 
      Height          =   435
      Left            =   0
      Picture         =   "frmMain.frx":305C0
      Top             =   480
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "?????????"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4320
      TabIndex        =   88
      Top             =   3480
      Width           =   2655
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public xScanning As Boolean
Public xThoat As Integer
Public Sub Scan4Virus()
Unload frmScan
frmScan.Show
End Sub


Private Sub CHK0_Click(Index As Integer)
SaveSetting "PAV2009", "Setting", "StartUP", CHK0(0).Value
If CHK0(0).Value = True Then
    SaveString HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Run", "PAV2009", AppPath & App.EXEName & ".exe /task"
Else
    DeleteValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Run", "PAV2009"
End If
If CHK0(3).Value = True Then
    SaveString HKEY_CLASSES_ROOT, "Folder\shell\Scan Virus With PAV\command", "", ChrW(34) & AppPath & App.EXEName & ".exe" & ChrW(34) & " %1"
    SaveString HKEY_CLASSES_ROOT, "*\shell\Scan Virus With PAV\command", "", ChrW(34) & AppPath & App.EXEName & ".exe" & ChrW(34) & " %1"
    
Else
    DeleteKey HKEY_CLASSES_ROOT, "Folder\shell\Scan Virus With PAV\command"
    DeleteKey HKEY_CLASSES_ROOT, "Folder\shell\Scan Virus With PAV"
    '=
    DeleteKey HKEY_CLASSES_ROOT, "*\shell\Scan Virus With PAV\command"
    DeleteKey HKEY_CLASSES_ROOT, "*\shell\Scan Virus With PAV"
    
End If
SaveSetting "PAV2009", "Setting", "PhucHoi", CHK0(1).Value
SaveSetting "PAV2009", "Setting", "ScanMenu", CHK0(3).Value

End Sub

Private Sub CHK1_Click(Index As Integer)
SaveSetting "PAV2009", "Update", Index, CHK1(Index).Value
End Sub

Private Sub Chk2_Click(Index As Integer)
cmdSaveCauHinh.Enabled = True
End Sub

Private Sub chkCheDoBaoVe_Click()
chkCheDoBaoVe.Value = True
End Sub

Private Sub Cmd_Click(Index As Integer)
ShowMenu Index
'&H00FFC0C0&
If Index = 0 Then GetInfo


End Sub

Private Sub cmdAddData_Click()
On Error Resume Next
CreateFolderUserData
Dim xPathVirus As String
Dim xVirusName As String
xPathVirus = MoFile(Me, ToUnicode("Cho5n ma64u Virus"), "All File|*.*", "C:\Documents and Settings\" & Environ("USERNAME") & "\Desktop")
If modMain.FileExists(xPathVirus) = True Then
    xVirusName = UniInputBox("Ha4y nha65p te6n Virus:" & vbCrLf & xPathVirus, "Enter Name", "Virus.User." & modMain.GetFileName(xPathVirus))
    If xVirusName <> "" Then
        If UniMsgBox("Ba5n co1 muo61n the6m file na2y va2o CSDL Virus hay kho6ng?" & vbCrLf & xVirusName & ": " & xPathVirus, vbYesNo, "?", Me.hwnd) = vbYes Then
            'add here
            If AddVirus(xVirusName, GetMD5(xPathVirus)) = True Then
                modReadWrite.WriteFileUni AppPath & "UserData\" & File2Str(xVirusName), GetMD5(xPathVirus)
                GetListData
                UniMsgBox "D9a4 the6m va2o CSDL Virus!", vbOKOnly, "OK!"
            Else
                UniMsgBox "Xa3y ra lo64i, ca1c nguye6n nha6n co1 the63 xa4y ra lo64i:" & vbCrLf & " - Virus d9a4 to62n ta5i!" & vbCrLf & " - Kho6ng ti2m tha61y CSDL" & vbCrLf & " - Ke61t no61i kho6ng tha2nh co6ng!" & vbCrLf & " - File ba5n cho5n kho6ng co1 du74 lie65u be6n trong." & vbCrLf & vbCrLf & "Xin thu73 la5i la62n nu74a!", vbOKOnly, "Error"
            End If
        End If
    End If
End If
End Sub

Private Sub cmdAddFileTinTuong_Click()
On Error Resume Next
Dim xPath As String
xPath = MoFile(frmMain, "Mo73 file", "All File|*.*", AppPath)
If xPath <> "" Then
    CreateFolDd
    'MsgBox StripNulls(xPath)
    modReadWrite.WriteFileUni AppPath & "TinTuong\" & GetMD5(xPath) & ".USER", StripNulls(xPath)
    GetListTinTuong
    UniMsgBox "D9a4 the6m va2o danh sa1ch tin tu7o73ng!", vbOKOnly, "OK", Me.hwnd
End If
GetListTinTuong
End Sub

Private Sub cmdAddFolderTinTuong_Click()
On Error Resume Next
Dim YPATH As String
YPATH = ChonThuMuc(frmMain)
If YPATH <> "" Then
    YPATH = FixPath(YPATH)
    File1.Path = YPATH
    File1.Refresh
    CreateFolDd
    Dim Uu As Integer
    For Uu = 0 To File1.ListCount - 1
        'lstTinTuong.AddItem YPATH & File1.List(Uu)
        modReadWrite.WriteFileUni AppPath & "TinTuong\" & GetMD5(YPATH & File1.List(Uu)) & ".USER", StripNulls(YPATH & File1.List(Uu))
    Next Uu
    GetListTinTuong
    UniMsgBox "D9a4 the6m va2o danh sa1ch tin tu7o73ng!", vbOKOnly, "OK", Me.hwnd
    
End If
End Sub

Private Sub cmdCheckForUpdate_Click()
    Shell AppPath & "PAVUpdate.exe", vbNormalFocus
    UniMsgBox "Chu7o7ng tri2nh se4 cha5y file PAVUpdate.exe d9e63 kie63m tra ca65p nha65t!" & vbCrLf & " Ne61u kho6ng co1 tho6ng ba1o gi2 xua61t hie65n sau 5 gia6y nu74a thi2 co1 nghi4a la2 ba5n d9ang su73 du5ng phie6n ba4n PAV mo71i nha61t!", vbOKOnly, "Tho6ng ba1o"
    
End Sub

Private Sub cmdDelData_Click()
On Error Resume Next
If lstUserData.ListIndex <> -1 Then
    'Delete
    If DeleteVirus(ReadFileUni(AppPath & "UserData\" & GetFileNhatKy(lstUserData.List(lstUserData.ListIndex)))) = True Then
        tXoaFile AppPath & "UserData\" & GetFileNhatKy(lstUserData.List(lstUserData.ListIndex))
        GetListData
        UniMsgBox "D9a4 xo1a xong!", vbOKOnly, "OK", Me.hwnd
    Else
        UniMsgBox "Lo64i!", vbOKOnly, "Error!", Me.hwnd
    End If
End If
End Sub

Private Sub cmdDeleteAllTinTuong_Click()
On Error Resume Next
Kill AppPath & "TinTuong\*.*"
GetListTinTuong
UniMsgBox "D9a4 xo1a he61t!", vbOKOnly, "OK!"
End Sub

Private Sub cmdGetListUserData_Click()
GetListData
End Sub

Private Sub cmdGetTinTuong_Click()
GetListTinTuong
End Sub

Private Sub cmdLayThongTinCachLy_Click()
GetListCachLy
End Sub

Private Sub cmdLayThongTinNhatKy_Click()
GetListNhatKy
End Sub

Private Sub cmdOnOffAT_Click(Index As Integer)
On Error Resume Next
If Index = 0 Then
    Load frmAutorun
    cmdOnOffAT(0).Enabled = False
    cmdOnOffAT(1).Enabled = True
    lblAT.Caption = "(D9ang ba65t)"
Else
    Unload frmAutorun
    cmdOnOffAT(1).Enabled = False
    cmdOnOffAT(0).Enabled = True
    lblAT.Caption = "(D9ang ta81t)"
End If
SaveSetting "PAV2009", "AutorunProtect", "OnOff", cmdOnOffAT(1).Enabled
End Sub

Private Sub cmdOnOffRTP_Click(Index As Integer)
On Error Resume Next
If Index = 0 Then
    Load frmRTP
    lblRTP.Caption = "(D9ang ba65t)"
    cmdOnOffRTP(0).Enabled = False
    cmdOnOffRTP(1).Enabled = True
Else
    Unload frmRTP
    lblRTP.Caption = "(D9ang ta81t)"
    cmdOnOffRTP(1).Enabled = False
    cmdOnOffRTP(0).Enabled = True
End If
SaveSetting "PAV2009", "RealTimeProtection", "OnOff", cmdOnOffRTP(1).Enabled
End Sub

Private Sub cmdPerfectProcessManager_Click()
If modMain.FileExists(AppPath & "PPM.exe") = False Then
    UniMsgBox "Kho6ng ti2m tha61y file PPM.exe!", vbOKOnly + vbCritical, "Error!!!", Me.hwnd
Else
    Shell AppPath & "PPM.exe", vbNormalFocus
End If
End Sub

Private Sub cmdPerfectStartUpManager_Click()
If modMain.FileExists(AppPath & "PSM.exe") = False Then
    UniMsgBox "Kho6ng ti2m tha61y file PSM.exe!", vbOKOnly + vbCritical, "Error!!!", Me.hwnd
Else
    Shell AppPath & "PSM.exe", vbNormalFocus
End If
End Sub

Private Sub cmdPerfectSystemReporter_Click()
If modMain.FileExists(AppPath & "PSR.exe") = False Then
    UniMsgBox "Kho6ng ti2m tha61y file PSR.exe!", vbOKOnly + vbCritical, "Error!!!", Me.hwnd
Else
    Shell AppPath & "PSR.exe", vbNormalFocus
End If
End Sub

Private Sub cmdPhucHoiDuLieu_Click()
frmPhucHoiDuLieu.Show , Me
End Sub

Private Sub cmdPhucHoiTatCa_Click()
If UniMsgBox("Ba5n co1 muo61n phu5c ho62i ta61t ca3 ca1c Virus co1 trong danh sa1ch?", vbYesNo, "?") = vbNo Then Exit Sub
Dim Iu As Integer
For Iu = 0 To lstCachLy.ListCount - 1
    PhucHoiCachLy lstCachLy.List(Iu)
Next Iu
UniMsgBox "D9a4 phu5c ho62i ta61t ca3!"
GetListCachLy
End Sub

Private Sub cmdPhucHoiVirusDaChon_Click()
If lstCachLy.ListIndex <> -1 Then
    PhucHoiCachLy lstCachLy.List(lstCachLy.ListIndex)
    UniMsgBox "D9a4 phu5c ho62i xong!", vbOKOnly, "OK!", Me.hwnd
    GetListCachLy
End If
End Sub

Private Sub cmdSaveCauHinh_Click()
On Error Resume Next
SaveSetting "PAV2009", "CauHinhVirus", "OPTEXT0", optEXT2(0).Value
SaveSetting "PAV2009", "CauHinhVirus", "OPTEXT1", optEXT2(1).Value
SaveSetting "PAV2009", "CauHinhVirus", "IndexEXT", cboEXT2.ListIndex
SaveSetting "PAV2009", "CauHinhVirus", "CboEn", cboEXT2.Enabled
SaveSetting "PAV2009", "CauHinhVirus", "Check0", Chk2(0).Value
SaveSetting "PAV2009", "CauHinhVirus", "Check1", Chk2(1).Value
SaveSetting "PAV2009", "CauHinhVirus", "Check2", Chk2(2).Value
SaveSetting "PAV2009", "CauHinhVirus", "Check3", Chk2(3).Value
SaveSetting "PAV2009", "CauHinhVirus", "Check4", Chk2(4).Value

cmdSaveCauHinh.Enabled = False
UniMsgBox "D9a4 lu7u la5i ca61u hi2nh!", vbOKOnly, "OK!", Me.hwnd
End Sub

Private Sub cmdTangTocMayTinh_Click()
frmTangToc.Show , Me
End Sub

Private Sub cmdTinhChinhRegistry_Click()
frmRegistry.Show , Me
End Sub

Private Sub cmdVirusRemoveAll_Click()
If modMain.FileExists(AppPath & "VRA.exe") = False Then
    UniMsgBox "Kho6ng ti2m tha61y file VRA.exe!", vbOKOnly + vbCritical, "Error!!!", Me.hwnd
Else
    Shell AppPath & "VRA.exe", vbNormalFocus
End If
End Sub

Private Sub cmdXemChiTiet_Click()
'Text1.Text = AppPath & "NhatKy\" & GetFileNhatKy(lstNhatKy.List(lstNhatKy.ListIndex))
Shell "notepad " & AppPath & "NhatKy\" & GetFileNhatKy(lstNhatKy.List(lstNhatKy.ListIndex)), vbNormalFocus
End Sub

Private Sub cmdXoaTatCa_Click()
If UniMsgBox("Ba5n co1 muo61n xo1a ta61t ca3 ca1c Virus co1 trong danh sa1ch?", vbYesNo, "?") = vbNo Then Exit Sub
Dim IuX As Integer
For IuX = 0 To lstCachLy.ListCount - 1
    tXoaFile Path2File(lstCachLy.List(IuX))
Next IuX
UniMsgBox "D9a4 xo1a ta61t ca3!"
GetListCachLy
End Sub

Private Sub cmdXoaVirusDaChon_Click()
If lstCachLy.ListIndex = -1 Then Exit Sub
tXoaFile Path2File(lstCachLy.List(lstCachLy.ListIndex))
UniMsgBox "D9a4 xo1a xong!", vbOKOnly, "OK!", Me.hwnd
GetListCachLy
End Sub


Private Sub Form_Load()
CaiDatBatDau
End Sub

Private Sub CaiDatBatDau()
Load frmMenu
MenuBar(1).AddButton ToUnicode("Que1t Virus")
MenuBar(1).AddSeparator
MenuBar(1).AddButton ToUnicode("Vu2ng ca1ch ly")
MenuBar(1).AddSeparator
MenuBar(1).AddButton ToUnicode("Nha65t ky1 que1t")
MenuBar(1).AddSeparator

MenuBar(2).AddButton ToUnicode("Ca61u hi2nh que1t")
MenuBar(2).AddSeparator
MenuBar(2).AddButton ToUnicode("Du74 lie65u Virus")
MenuBar(2).AddSeparator
MenuBar(2).AddButton ToUnicode("Danh sách tin tu7o73ng")
MenuBar(2).AddSeparator

MenuBar(3).AddButton ToUnicode("Ba3o ve65 tho72i gian thu75c")
MenuBar(3).AddSeparator
MenuBar(3).AddButton ToUnicode("Ba3o ve65 Autorun"), , , 1500
MenuBar(3).AddSeparator
MenuBar(3).AddButton ToUnicode("Phu5c ho62i du74 lie65u"), , , 1500
MenuBar(3).AddSeparator


MenuBar(0).AddButton ToUnicode("Tho6ng tin chung"), , , 1500
MenuBar(0).AddSeparator


MenuBar(4).AddButton ToUnicode("Tinh chi3nh Registry")
MenuBar(4).AddSeparator
MenuBar(4).AddButton ToUnicode("Ta8ng to61c ma1y ti1nh")
MenuBar(4).AddSeparator

MenuBar(5).AddButton ToUnicode("Qua3n ly1 tie61n tri2nh"), , , 1800
MenuBar(5).AddSeparator
MenuBar(5).AddButton ToUnicode("Qua3n ly1 kho73i d9o65ng")
MenuBar(5).AddSeparator
MenuBar(5).AddButton ToUnicode("Kie63m tra he65 tho61ng")
MenuBar(5).AddSeparator
MenuBar(5).AddButton ToUnicode("Que1t ma64u")
MenuBar(5).AddSeparator

MenuBar(6).AddButton ToUnicode("Ca2i d9a85t chung")
MenuBar(6).AddSeparator

MenuBar(7).AddButton ToUnicode("Tho6ng tin chu7o7ng tri2nh")
MenuBar(7).AddSeparator
MenuBar(7).AddButton ToUnicode("Tho6ng tin ta1c gia3")
MenuBar(7).AddSeparator


Dim KiK As Integer
For KiK = 0 To FM.Count - 1
    FM(KiK).Left = Cmd(0).Width + 120
    FM(KiK).Top = 1800
    FM(KiK).Height = 4455
    FM(KiK).Width = 6735
Next KiK

Me.Width = 8970
Me.Height = 6810
xThoat = 1

Dim xAx As Integer
For xAx = 0 To MenuBar.Count - 1
MenuBar(xAx).Top = 1200
MenuBar(xAx).Left = Cmd(0).Width
MenuBar(xAx).Width = Me.Width - Cmd(0).Width
MenuBar(xAx).ActiveButton = 0
Next xAx

'== Load setting

optEXT2(0).Value = GetSetting("PAV2009", "CauHinhVirus", "OPTEXT0", False)
optEXT2(1).Value = GetSetting("PAV2009", "CauHinhVirus", "OPTEXT1", True)
cboEXT2.Enabled = GetSetting("PAV2009", "CauHinhVirus", "CboEn", True)
Chk2(0).Value = GetSetting("PAV2009", "CauHinhVirus", "Check0", True)
Chk2(1).Value = GetSetting("PAV2009", "CauHinhVirus", "Check1", True)
Chk2(2).Value = GetSetting("PAV2009", "CauHinhVirus", "Check2", True)
Chk2(3).Value = GetSetting("PAV2009", "CauHinhVirus", "Check3", True)
Chk2(4).Value = GetSetting("PAV2009", "CauHinhVirus", "Check4", True)

cboEXT2.AddItem "*.EXE - Application"
cboEXT2.AddItem "*.BAT - MS-DOS Batch File"
cboEXT2.AddItem "*.CMD - Windows NT Command Script"
cboEXT2.AddItem "*.COM - MS-DOS Application"
cboEXT2.AddItem "*.DLL - Application Extension"
cboEXT2.AddItem "*.OCX - ActiveX Control"
cboEXT2.AddItem "*.PIF - Shortcut to MS-DOS Program"
cboEXT2.AddItem "*.SCR - Screen Saver"
cboEXT2.ListIndex = GetSetting("PAV2009", "CauHinhVirus", "IndexEXT", "0")

If GetSetting("PAV2009", "RealTimeProtection", "OnOff", True) = True Then
    lblRTP.Caption = "(D9ang ba65t)"
    cmdOnOffRTP(0).Enabled = False
    cmdOnOffRTP(1).Enabled = True
    img0(0).Picture = IMGON.Picture
    Load frmRTP
Else
    lblRTP.Caption = "(D9ang ta81t)"
    cmdOnOffRTP(1).Enabled = False
    cmdOnOffRTP(0).Enabled = True
    img0(0).Picture = IMGOFF.Picture
End If

If GetSetting("PAV2009", "AutorunProtect", "OnOff", True) = True Then
    lblAT.Caption = "(D9ang ba65t)"
    cmdOnOffAT(0).Enabled = False
    cmdOnOffAT(1).Enabled = True
    Load frmAutorun
    img0(1).Picture = IMGON.Picture
Else
    lblAT.Caption = "(D9ang ta81t)"
    cmdOnOffAT(1).Enabled = False
    cmdOnOffAT(0).Enabled = True
    img0(1).Picture = IMGOFF.Picture
End If

'SaveSetting "PAV2009", "Setting", "StartUP", CHK0(0).Value
If GetSetting("PAV2009", "Setting", "StartUP", True) = True Then
    img0(3).Picture = IMGON.Picture
    CHK0(0).Value = True
    SaveString HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Run", "PAV2009", AppPath & App.EXEName & ".exe /task"
Else
    img0(3).Picture = IMGOFF.Picture
    CHK0(0).Value = False
    DeleteValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Run", "PAV2009"
End If

'SaveSetting "PAV2009", "Setting", "ScanMenu", CHK0(3).Value
If GetSetting("PAV2009", "Setting", "ScanMenu", True) = True Then
    CHK0(3).Value = True
    SaveString HKEY_CLASSES_ROOT, "Folder\shell\Scan Virus With PAV\command", "", ChrW(34) & AppPath & App.EXEName & ".exe" & ChrW(34) & " %1"
    '=
    SaveString HKEY_CLASSES_ROOT, "*\shell\Scan Virus With PAV\command", "", ChrW(34) & AppPath & App.EXEName & ".exe" & ChrW(34) & " %1"
    
Else
    DeleteKey HKEY_CLASSES_ROOT, "Folder\shell\Scan Virus With PAV\command"
    DeleteKey HKEY_CLASSES_ROOT, "Folder\shell\Scan Virus With PAV"
    '=
    DeleteKey HKEY_CLASSES_ROOT, "*\shell\Scan Virus With PAV\command"
    DeleteKey HKEY_CLASSES_ROOT, "*\shell\Scan Virus With PAV"
    
    CHK0(3).Value = False
End If

'SaveSetting "PAV2009", "Update", Index, CHK1(Index).Value
Dim J As Integer
For J = 0 To 3
    CHK1(J).Value = GetSetting("PAV2009", "Update", J, True)
Next J
If CHK1(0).Value = True Then
    If modMain.FileExists(AppPath & "PAVUPDATE.EXE") = True Then
        Shell AppPath & "PAVUPDATE.EXE", vbNormalFocus
    Else
        UniMsgBox "Kho6ng ti2m tha61y file PAVUPDATE.EXE!", vbOKOnly, "Error!", Me.hwnd
    End If
End If

Load frmProtect

img0(2).Picture = IMGON.Picture
img0(4).Picture = IMGON.Picture

GetListNhatKy
GetListCachLy
GetListData
GetListTinTuong
GetComputerInfo
'== End load setting
ShowMenu 0
End Sub

Public Sub ShowMenu(xIndex)
Dim Ii As Integer
For Ii = 0 To MenuBar.Count - 1
    MenuBar(Ii).Visible = False
Next Ii
MenuBar(xIndex).Visible = True


Dim Uy As Integer
For Uy = 0 To FM.Count - 1
    If ToUnicode(FM(Uy).Caption) = ToUnicode(MenuBar(xIndex).ButtonCaption(MenuBar(xIndex).ActiveButton)) Then
        FM(Uy).Visible = True
    Else
        FM(Uy).Visible = False
    End If
Next Uy

Dim Yu As Integer
For Yu = 0 To Cmd.Count - 1
    Cmd(Yu).BackColor = &HE0E0E0
Next Yu
Cmd(xIndex).BackColor = &HFFC0C0
On Error Resume Next
Cmd(xIndex).SetFocus
End Sub


Private Sub Form_Unload(Cancel As Integer)
Cancel = xThoat
Me.Hide
Load frmMenu
App.TaskVisible = False
End Sub


Private Sub lblGopY_Click()
If modMain.FileExists(AppPath & "GOPY.EXE") = False Then
    UniMsgBox "Kho6ng ti2m tha61y file GOPY.EXE!", vbOKOnly, "Error!", Me.hwnd
Else
    Shell AppPath & "GOPY.EXE", vbNormalFocus
End If
End Sub


Private Sub lblSupport_Click()
If modMain.FileExists(AppPath & "GOPY.exe") = False Then
    UniMsgBox "Kho6ng ti2m tha61y file GOPY.EXE", vbOKOnly, "Error!", Me.hwnd
Else
    Shell AppPath & "GOPY.EXE", vbNormalFocus
End If
End Sub



Private Sub MenuBar_ButtonClick(Index As Integer, ButtonIndex As Integer)
Dim Uy As Integer
For Uy = 0 To FM.Count - 1
    If ToUnicode(FM(Uy).Caption) = ToUnicode(MenuBar(Index).ButtonCaption(ButtonIndex)) Then
        FM(Uy).Visible = True
    Else
        FM(Uy).Visible = False
    End If
Next Uy
If Index = 0 Then
GetInfo
End If
End Sub



Private Sub optEXT2_Click(Index As Integer)
If optEXT2(1).Value = True Then
    cboEXT2.Enabled = True
Else
    cboEXT2.Enabled = False
End If
cmdSaveCauHinh.Enabled = True
End Sub

Private Sub Scan_Click()
frmScan.Show
xScanning = True
Me.Hide
End Sub

Public Sub GetInfo()
If GetSetting("PAV2009", "RealTimeProtection", "OnOff", True) = True Then
    img0(0).Picture = IMGON.Picture
Else
    img0(0).Picture = IMGOFF.Picture
End If

If GetSetting("PAV2009", "AutorunProtect", "OnOff", True) = True Then
    img0(1).Picture = IMGON.Picture
Else
    img0(1).Picture = IMGOFF.Picture
End If

If GetSetting("PAV2009", "Setting", "StartUP", True) = True Then
    img0(3).Picture = IMGON.Picture
Else
    img0(3).Picture = IMGOFF.Picture
End If
Me.lbl1TinhTrangMayTinh.Caption = CheckComputerHeal
Me.lbl1Process.Caption = GetTotalProcess
End Sub

Public Sub GetComputerInfo()
Me.lbl1UserName.Caption = Environ("USERNAME")
Me.lbl1ComputerName.Caption = GetComputer
Me.lbl1Memory.Caption = GetRAMTotal
End Sub

