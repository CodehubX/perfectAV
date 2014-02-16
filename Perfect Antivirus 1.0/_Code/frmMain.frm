VERSION 5.00
Object = "{D1ECD258-D329-4388-AB83-DEC261A66B86}#1.0#0"; "UniControls_v2.0.ocx"
Begin VB.Form frmMain 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Perfect Antivirus 2009"
   ClientHeight    =   8160
   ClientLeft      =   2625
   ClientTop       =   1080
   ClientWidth     =   10215
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
   ScaleHeight     =   8160
   ScaleWidth      =   10215
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox icoFull 
      Height          =   495
      Index           =   17
      Left            =   5880
      Picture         =   "frmMain.frx":058A
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   76
      Top             =   120
      Visible         =   0   'False
      Width           =   495
   End
   Begin UniControls.UniFrame fm 
      Height          =   6735
      Index           =   12
      Left            =   -2640
      Top             =   1440
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   11880
      MaskColor       =   16711935
      FrameColor      =   -2147483629
      Style           =   0
      Caption         =   "Tho6ng Tin Pha62n Me62m"
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
      Begin UniControls.UniButton cmdViewErr 
         Height          =   255
         Left            =   240
         TabIndex        =   86
         Top             =   6000
         Width           =   6735
         _ExtentX        =   11880
         _ExtentY        =   450
         Icon            =   "frmMain.frx":0C74
         Style           =   2
         Caption         =   "Nha61n va2o d9a6y d9e63 xem ca1c lo64i xa3y ra trong qua1 tri2nh su73 du5ng cu3a ba5n"
         IconAlign       =   3
         iNonThemeStyle  =   2
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
      Begin UniControls.UniLabel UniLabel86 
         Height          =   495
         Left            =   120
         Top             =   5520
         Width           =   6615
         _ExtentX        =   11668
         _ExtentY        =   873
         Caption         =   $"frmMain.frx":0C90
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
      Begin UniControls.UniLabel UniLabel63 
         Height          =   255
         Left            =   600
         Top             =   6360
         Width           =   6015
         _ExtentX        =   10610
         _ExtentY        =   450
         Alignment       =   1
         Caption         =   "Ca3m o7n ba5n d9a4 su73 du5ng chu7o7ng tri2nh na2y."
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
      Begin UniControls.UniLabel UniLabel62 
         Height          =   855
         Left            =   120
         Top             =   4680
         Width           =   6855
         _ExtentX        =   12091
         _ExtentY        =   1508
         Caption         =   $"frmMain.frx":0D43
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
      Begin UniControls.UniLabel UniLabel60 
         Height          =   255
         Left            =   480
         Top             =   3480
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   450
         Caption         =   "Vo71i bo65 UnicodeControls 2.0."
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   16512
      End
      Begin UniControls.UniLabel UniLabel59 
         Height          =   255
         Left            =   120
         Top             =   3240
         Width           =   4815
         _ExtentX        =   8493
         _ExtentY        =   450
         Caption         =   "Nho1m iVB (http://caulacbovb.com)"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483635
      End
      Begin UniControls.UniLabel UniLabel55 
         Height          =   255
         Left            =   120
         Top             =   2760
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   450
         Caption         =   "Forum: http://virusvn.com"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483635
      End
      Begin UniControls.UniLabel UniLabel54 
         Height          =   255
         Left            =   480
         Top             =   2520
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   450
         Caption         =   "Vo71i ca1c module xu73 ly1 Registry, Process..."
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   16512
      End
      Begin UniControls.UniLabel UniLabel53 
         Height          =   255
         Left            =   120
         Top             =   2280
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   450
         Caption         =   "Website: http://pscode.com"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483635
      End
      Begin UniControls.UniLabel UniLabel52 
         Height          =   375
         Left            =   120
         Top             =   1800
         Width           =   6735
         _ExtentX        =   11880
         _ExtentY        =   661
         Caption         =   "D9e63 co1 d9u7o75c chu7o7ng tri2nh na2y, to6i xin cha6n tha2nh gu73i lo72i ca3m o7n d9e61n:"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   16576
      End
      Begin UniControls.UniLabel UniLabel51 
         Height          =   255
         Left            =   960
         Top             =   1320
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   450
         Alignment       =   2
         Caption         =   "Website:"
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
      Begin UniControls.UniLabel UniLabel49 
         Height          =   255
         Left            =   1080
         Top             =   840
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   450
         Alignment       =   2
         Caption         =   "Ta1c Gia3:"
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
      Begin UniControls.UniLabel UniLabel45 
         Height          =   255
         Left            =   2040
         Top             =   840
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   450
         Caption         =   "D9inh Quang Trung"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   12632064
      End
      Begin UniControls.UniLabel UniLabel44 
         Height          =   375
         Left            =   1320
         Top             =   360
         Width           =   4695
         _ExtentX        =   8281
         _ExtentY        =   661
         Caption         =   "Tho6ng tin ve62 pha62n me62m Perfect Antivirus 2009"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   49152
      End
      Begin UniControls.UniLabel UniLabel46 
         Height          =   255
         Left            =   3960
         Top             =   840
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   450
         Caption         =   "Lo71p 10T2 Tru7o72ng THPT D9o62ng Phu1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   12632064
      End
      Begin UniControls.UniLabel UniLabel47 
         Height          =   255
         Left            =   2040
         Top             =   1080
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   450
         Caption         =   "DinhQuangTrung90@Yahoo.Com"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   12632064
      End
      Begin UniControls.UniLabel UniLabel48 
         Height          =   255
         Left            =   2040
         Top             =   1320
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   450
         Caption         =   "http://qts.come.vn"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   12632064
      End
      Begin UniControls.UniLabel UniLabel50 
         Height          =   255
         Left            =   1080
         Top             =   1080
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   450
         Alignment       =   2
         Caption         =   "Email:"
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
      Begin UniControls.UniLabel UniLabel56 
         Height          =   255
         Left            =   480
         Top             =   3000
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   450
         Caption         =   "Vo71i mo65t so61 ma64u Virus mo71i hie65n nay."
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   16512
      End
      Begin UniControls.UniLabel UniLabel57 
         Height          =   255
         Left            =   120
         Top             =   3720
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   450
         Caption         =   "Le6 Nguye6n Du4ng (dungcoivb@gmail.com)"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483635
      End
      Begin UniControls.UniLabel UniLabel58 
         Height          =   255
         Left            =   480
         Top             =   3960
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   450
         Caption         =   "Vo71i bo65 co7 so73 du74 lie65u tu72 vnAntivirus."
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   16512
      End
      Begin UniControls.UniLabel UniLabel73 
         Height          =   255
         Left            =   120
         Top             =   4200
         Width           =   6735
         _ExtentX        =   11880
         _ExtentY        =   450
         Caption         =   "D9o64 Ngo5c Hoa2ng (amduongvotinh_danhlangquen_001@yahoo.com.vn)"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483635
      End
      Begin UniControls.UniLabel UniLabel74 
         Height          =   255
         Left            =   480
         Top             =   4440
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   450
         Caption         =   "Vo71i mo65t so61 ma64u Keylogger mo71i hie65n nay."
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   16512
      End
      Begin VB.Image Image3 
         Height          =   720
         Left            =   240
         Picture         =   "frmMain.frx":0EB4
         Top             =   840
         Width           =   720
      End
   End
   Begin VB.PictureBox picOn 
      Height          =   495
      Left            =   6360
      Picture         =   "frmMain.frx":6696
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   69
      Top             =   600
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox PicOff 
      Height          =   495
      Left            =   6720
      Picture         =   "frmMain.frx":6D16
      ScaleHeight     =   435
      ScaleWidth      =   555
      TabIndex        =   68
      Top             =   120
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox icoFull 
      Height          =   495
      Index           =   16
      Left            =   4080
      Picture         =   "frmMain.frx":7351
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   62
      Top             =   600
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox icoFull 
      Height          =   495
      Index           =   15
      Left            =   3600
      Picture         =   "frmMain.frx":7A3B
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   61
      Top             =   600
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox icoFull 
      Height          =   495
      Index           =   14
      Left            =   3120
      Picture         =   "frmMain.frx":8125
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   60
      Top             =   600
      Visible         =   0   'False
      Width           =   495
   End
   Begin UniControls.UniFrame fm 
      Height          =   6735
      Index           =   11
      Left            =   9000
      Top             =   7920
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   11880
      MaskColor       =   16711935
      FrameColor      =   -2147483629
      Style           =   0
      Caption         =   "Ca2i D9a85t Ca61u Hi2nh"
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
      Begin UniControls.UniFrame fmGopY 
         Height          =   1695
         Left            =   120
         Top             =   4920
         Width           =   6855
         _ExtentX        =   12091
         _ExtentY        =   2990
         MaskColor       =   16711935
         FrameColor      =   -2147483629
         Style           =   0
         Caption         =   "Gu73i Tha81c Ma81c/Go1p Y1"
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
         Begin UniControls.UniLabel UniLabel43 
            Height          =   255
            Left            =   120
            Top             =   840
            Width           =   6615
            _ExtentX        =   11668
            _ExtentY        =   450
            Alignment       =   1
            Caption         =   "Mo5i su75 go1p y1 cu3a ba5n d9e62u la2 ne62n ta3ng cho su75 pha1t trie63n cu3a chu7o7ng tri2nh."
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
         Begin UniControls.UniLabel UniLabel42 
            Height          =   495
            Left            =   120
            Top             =   360
            Width           =   6615
            _ExtentX        =   11668
            _ExtentY        =   873
            Alignment       =   1
            Caption         =   $"frmMain.frx":880F
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
         Begin UniControls.UniButton cmdGotoGopY 
            Height          =   375
            Left            =   2520
            TabIndex        =   67
            Top             =   1200
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   661
            Icon            =   "frmMain.frx":88CF
            Style           =   2
            Caption         =   "Gu73i Tha81c Ma81c/Go1p Y1"
            IconAlign       =   3
            iNonThemeStyle  =   2
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
      End
      Begin UniControls.UniFrame fmHelp 
         Height          =   1335
         Left            =   120
         Top             =   3480
         Width           =   6855
         _ExtentX        =   12091
         _ExtentY        =   2355
         MaskColor       =   16711935
         FrameColor      =   -2147483629
         Style           =   0
         Caption         =   "Giu1p D9o74"
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
         Begin UniControls.UniLabel UniLabel41 
            Height          =   375
            Left            =   240
            Top             =   240
            Width           =   6495
            _ExtentX        =   11456
            _ExtentY        =   661
            Alignment       =   1
            Caption         =   "Ha4y d9o5c ky4 hu7o71ng da64n su73 du5ng d9e63 co1 the63 su73 du5ng hie65u qua3 chu7o7ng tri2nh na2y."
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
         Begin UniControls.UniButton cmdGoToHelp 
            Height          =   375
            Left            =   2520
            TabIndex        =   66
            Top             =   840
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   661
            Icon            =   "frmMain.frx":88EB
            Style           =   2
            Caption         =   "Hu7o71ng Da64n Su73 Du5ng"
            IconAlign       =   3
            iNonThemeStyle  =   2
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
      End
      Begin UniControls.UniFrame fmAutoUpdate 
         Height          =   1215
         Left            =   120
         Top             =   2160
         Width           =   6855
         _ExtentX        =   12091
         _ExtentY        =   2143
         MaskColor       =   16711935
         FrameColor      =   -2147483629
         Style           =   0
         Caption         =   "Tu75 D9o65ng Ca65p Nha65t"
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
         Begin UniControls.UniButton cmdUpdateOff 
            Height          =   375
            Left            =   4080
            TabIndex        =   80
            Top             =   720
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   661
            Icon            =   "frmMain.frx":8907
            Style           =   2
            Caption         =   "Ca65p Nha65t Offline"
            IconAlign       =   3
            iNonThemeStyle  =   2
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
         Begin UniControls.UniButton cmdUpdateOnline 
            Height          =   375
            Left            =   1320
            TabIndex        =   79
            Top             =   720
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   661
            Icon            =   "frmMain.frx":8923
            Style           =   2
            Caption         =   "Ca65p nha65t Online"
            IconAlign       =   3
            iNonThemeStyle  =   2
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
         Begin UniControls.UniCheckBox chkAutoUpdate 
            Height          =   195
            Left            =   1320
            TabIndex        =   65
            Top             =   360
            Width           =   3690
            _ExtentX        =   6509
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
            Caption         =   "Tu75 D9o65ng Ca65p Nha65t Mo64i Khi Cha5y Chu7o7ng Tri2nh."
            ForeColor       =   0
            ShowFocusRectangle=   0   'False
         End
         Begin VB.Image Image2 
            Height          =   720
            Left            =   240
            Picture         =   "frmMain.frx":893F
            Top             =   360
            Width           =   720
         End
      End
      Begin UniControls.UniFrame fmsetting 
         Height          =   1335
         Left            =   120
         Top             =   720
         Width           =   6855
         _ExtentX        =   12091
         _ExtentY        =   2355
         MaskColor       =   16711935
         FrameColor      =   -2147483629
         Style           =   0
         Caption         =   "Ca2i D9a85t Chung"
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
         Begin UniControls.UniCheckBox chkUseFastScan 
            Height          =   195
            Left            =   1320
            TabIndex        =   78
            Top             =   840
            Width           =   5370
            _ExtentX        =   9472
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
            Caption         =   "Su73 du5ng chu71c na8ng que1t nhanh Folder (Khi click chuo65t pha3i va2o Folder)"
            ForeColor       =   0
            ShowFocusRectangle=   0   'False
         End
         Begin UniControls.UniCheckBox chkShowFlash 
            Height          =   195
            Left            =   1320
            TabIndex        =   64
            Top             =   600
            Width           =   3225
            _ExtentX        =   5689
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
            Caption         =   "Hie65n Flash Screen khi cha5y chu7o7ng tri2nh."
            ForeColor       =   0
            ShowFocusRectangle=   0   'False
         End
         Begin UniControls.UniCheckBox chkAutoStart 
            Height          =   195
            Left            =   1320
            TabIndex        =   63
            Top             =   360
            Width           =   2700
            _ExtentX        =   4763
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
            Caption         =   "Tu75 Cha5y Khi Windows Kho73i D9o65ng."
            ForeColor       =   0
            ShowFocusRectangle=   0   'False
         End
         Begin VB.Image Image1 
            Height          =   720
            Left            =   240
            Picture         =   "frmMain.frx":F191
            Top             =   360
            Width           =   720
         End
      End
      Begin UniControls.UniLabel UniLabel40 
         Height          =   375
         Left            =   120
         Top             =   360
         Width           =   6855
         _ExtentX        =   12091
         _ExtentY        =   661
         Alignment       =   1
         Caption         =   "Ca2i D9a85t Ca61u Hi2nh Cho PAV 2009"
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
      Caption         =   "UniFrame1"
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
         Height          =   6735
         Left            =   0
         TabIndex        =   59
         Top             =   0
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   11880
      End
   End
   Begin UniControls.UniFrame fm 
      Height          =   6735
      Index           =   10
      Left            =   9000
      Top             =   7800
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   11880
      MaskColor       =   16711935
      FrameColor      =   -2147483629
      Style           =   0
      Caption         =   "Kie63m Tra He65 Tho61ng"
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
      Begin UniControls.UniLabel UniLabel39 
         Height          =   1335
         Left            =   480
         Top             =   960
         Width           =   6135
         _ExtentX        =   10821
         _ExtentY        =   2355
         Caption         =   $"frmMain.frx":1530B
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
      Begin UniControls.UniButton cmdKiemTraHeThong 
         Height          =   615
         Left            =   2280
         TabIndex        =   58
         Top             =   2640
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   1085
         Icon            =   "frmMain.frx":15421
         Style           =   2
         Caption         =   "Kie63m Tra He65 Tho61ng"
         IconAlign       =   3
         iNonThemeStyle  =   2
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
      Begin UniControls.UniLabel UniLabel38 
         Height          =   375
         Left            =   600
         Top             =   360
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   661
         Alignment       =   1
         Caption         =   "Nha61n va2o nu1t be6n du7o71i d9e63 va2o mu5c Kie63m Tra He65 Tho61ng"
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
   Begin VB.PictureBox icoFull 
      Height          =   495
      Index           =   13
      Left            =   5400
      Picture         =   "frmMain.frx":1543D
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   57
      Top             =   120
      Visible         =   0   'False
      Width           =   495
   End
   Begin UniControls.UniFrame fm 
      Height          =   6735
      Index           =   9
      Left            =   9000
      Top             =   7680
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   11880
      MaskColor       =   16711935
      FrameColor      =   -2147483629
      Style           =   0
      Caption         =   "Qua3n ly1 kho73i d9o65ng"
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
      Begin UniControls.UniButton cmdStartUp 
         Height          =   615
         Left            =   2280
         TabIndex        =   56
         Top             =   2640
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   1085
         Icon            =   "frmMain.frx":15B27
         Style           =   2
         Caption         =   "Qua3n Ly1 Kho73i D9o65ng"
         IconAlign       =   3
         iNonThemeStyle  =   2
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
      Begin UniControls.UniLabel UniLabel37 
         Height          =   375
         Left            =   600
         Top             =   360
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   661
         Alignment       =   1
         Caption         =   "Nha61n va2o nu1t be6n du7o71i d9e63 va2o mu5c qua3n ly1 kho73i d9o65ng"
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
   Begin VB.PictureBox icoFull 
      Height          =   495
      Index           =   12
      Left            =   4920
      Picture         =   "frmMain.frx":15B43
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   55
      Top             =   120
      Visible         =   0   'False
      Width           =   495
   End
   Begin UniControls.UniFrame fm 
      Height          =   6735
      Index           =   8
      Left            =   9000
      Top             =   7560
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   11880
      MaskColor       =   16711935
      FrameColor      =   -2147483629
      Style           =   0
      Caption         =   "Qua3n ly1 Files"
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
      Begin UniControls.UniButton cmdQuanLyFile 
         Height          =   615
         Left            =   2280
         TabIndex        =   54
         Top             =   2640
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   1085
         Icon            =   "frmMain.frx":1622D
         Style           =   2
         Caption         =   "Qua3n Ly1 Ta65p Tin"
         IconAlign       =   3
         iNonThemeStyle  =   2
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
      Begin UniControls.UniLabel UniLabel33 
         Height          =   375
         Left            =   600
         Top             =   360
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   661
         Alignment       =   1
         Caption         =   "Nha61n va2o nu1t be6n du7o71i d9e63 va2o mu5c Qua3n Ly1 Ta65p Tin"
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
   Begin VB.PictureBox icoFull 
      Height          =   495
      Index           =   11
      Left            =   4440
      Picture         =   "frmMain.frx":16249
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   53
      Top             =   120
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox icoFull 
      Height          =   495
      Index           =   10
      Left            =   3960
      Picture         =   "frmMain.frx":16933
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   52
      Top             =   120
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox icoFull 
      Height          =   495
      Index           =   9
      Left            =   960
      Picture         =   "frmMain.frx":1701D
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   51
      Top             =   0
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox icoFull 
      Height          =   495
      Index           =   8
      Left            =   1920
      Picture         =   "frmMain.frx":17707
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   50
      Top             =   600
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox icoFull 
      Height          =   495
      Index           =   7
      Left            =   1440
      Picture         =   "frmMain.frx":17DF1
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   49
      Top             =   600
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox icoFull 
      Height          =   495
      Index           =   6
      Left            =   960
      Picture         =   "frmMain.frx":184DB
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   48
      Top             =   600
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox icoFull 
      Height          =   495
      Index           =   5
      Left            =   2400
      Picture         =   "frmMain.frx":18BC5
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   47
      Top             =   0
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox icoFull 
      Height          =   495
      Index           =   4
      Left            =   1920
      Picture         =   "frmMain.frx":192AF
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   46
      Top             =   0
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox icoFull 
      Height          =   495
      Index           =   3
      Left            =   1440
      Picture         =   "frmMain.frx":19999
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   45
      Top             =   0
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox icoFull 
      Height          =   495
      Index           =   2
      Left            =   3480
      Picture         =   "frmMain.frx":1A083
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   44
      Top             =   120
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox icoFull 
      Height          =   495
      Index           =   1
      Left            =   480
      Picture         =   "frmMain.frx":1A76D
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   43
      Top             =   600
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox icoFull 
      Height          =   495
      Index           =   0
      Left            =   480
      Picture         =   "frmMain.frx":1AE57
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   42
      Top             =   0
      Visible         =   0   'False
      Width           =   495
   End
   Begin UniControls.UniFrame fm 
      Height          =   6735
      Index           =   7
      Left            =   9000
      Top             =   7440
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
      Begin UniControls.UniLabel UniLabel31 
         Height          =   375
         Left            =   840
         Top             =   360
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   661
         Alignment       =   1
         Caption         =   "Nha61n va2o nu1t be6n du7o71i d9e63 va2o mu5c qua3n ly1 tie61n tri2nh."
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
      Begin UniControls.UniButton cmdQuanly 
         Height          =   615
         Left            =   2280
         TabIndex        =   41
         Top             =   2640
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   1085
         Icon            =   "frmMain.frx":1B541
         Style           =   2
         Caption         =   "Qua3n ly1 tie61n tri2nh"
         IconAlign       =   3
         iNonThemeStyle  =   2
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
   End
   Begin UniControls.UniFrame fm 
      Height          =   6735
      Index           =   6
      Left            =   2520
      Top             =   120
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
      Begin UniControls.UniLabel UniLabel76 
         Height          =   735
         Left            =   960
         Top             =   5280
         Width           =   5775
         _ExtentX        =   10186
         _ExtentY        =   1296
         Caption         =   $"frmMain.frx":1B55D
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
      Begin UniControls.UniLabel UniLabel75 
         Height          =   615
         Left            =   960
         Top             =   4560
         Width           =   6015
         _ExtentX        =   10610
         _ExtentY        =   1085
         Caption         =   $"frmMain.frx":1B667
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
      Begin UniControls.UniLabel UniLabel61 
         Height          =   255
         Left            =   240
         Top             =   4560
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   450
         Caption         =   "Chu1 Y1:"
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
      Begin UniControls.UniCheckBox chkAutoAddAutorun 
         Height          =   570
         Left            =   240
         TabIndex        =   70
         Top             =   3840
         Width           =   6360
         _ExtentX        =   11218
         _ExtentY        =   1005
         AutoSize        =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Tu75 d9o65ng the6m nhu74ng loa5i Virus chu7a bie61t va2o co7 so73 du74 lie65u cu3a chu7o7ng tri2nh."
         ForeColor       =   255
         ShowFocusRectangle=   0   'False
      End
      Begin UniControls.UniLabel UniLabel36 
         Height          =   495
         Left            =   360
         Top             =   2040
         Width           =   6375
         _ExtentX        =   11245
         _ExtentY        =   873
         Caption         =   $"frmMain.frx":1B755
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
         TabIndex        =   40
         Top             =   2760
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   1296
         Icon            =   "frmMain.frx":1B802
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
         Caption         =   $"frmMain.frx":1B81E
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
      Left            =   9000
      Top             =   7200
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
         Interval        =   10000
         Left            =   5160
         Top             =   600
      End
      Begin UniControls.UniButton cmdDelREG 
         Height          =   375
         Left            =   2280
         TabIndex        =   39
         Top             =   6240
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   661
         Icon            =   "frmMain.frx":1B8E7
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
         TabIndex        =   38
         Top             =   6240
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   661
         Icon            =   "frmMain.frx":1BE81
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
         TabIndex        =   37
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
         FullRowSelect   =   -1  'True
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
         Caption         =   $"frmMain.frx":1C41B
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
         TabIndex        =   36
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
      Left            =   9000
      Top             =   7080
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
         TabIndex        =   35
         Top             =   6240
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         Icon            =   "frmMain.frx":1C4F9
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
         TabIndex        =   34
         Top             =   6240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         Icon            =   "frmMain.frx":1CA93
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
         TabIndex        =   33
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
      Left            =   9000
      Top             =   6960
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
         TabIndex        =   16
         Top             =   4080
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   661
         Icon            =   "frmMain.frx":1D02D
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
         Text            =   "4"
         Max             =   4
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
      Left            =   9000
      Top             =   6840
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
         TabIndex        =   32
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
         TabIndex        =   19
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
         Caption         =   $"frmMain.frx":1DA3F
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
         Caption         =   $"frmMain.frx":1DAC6
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
         TabIndex        =   18
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
         TabIndex        =   17
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
         Caption         =   $"frmMain.frx":1DC42
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
      Icon            =   "frmMain.frx":1DD71
   End
   Begin UniControls.UniFrame fm 
      Height          =   6735
      Index           =   1
      Left            =   9000
      Top             =   6720
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
         TabIndex        =   31
         Top             =   6000
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         Icon            =   "frmMain.frx":1E30B
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
         TabIndex        =   30
         Top             =   6000
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         Icon            =   "frmMain.frx":1ED1D
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
         TabIndex        =   29
         Top             =   6000
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         Icon            =   "frmMain.frx":1F2B7
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
         TabIndex        =   28
         Top             =   6000
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         Icon            =   "frmMain.frx":1FCC9
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
         TabIndex        =   27
         Top             =   6000
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         Icon            =   "frmMain.frx":20263
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
            TabIndex        =   20
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
            TabIndex        =   21
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
            TabIndex        =   22
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
            TabIndex        =   23
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
            TabIndex        =   24
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
            TabIndex        =   25
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
         TabIndex        =   26
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
      Left            =   9000
      Top             =   6600
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
      Begin UniControls.UniButton cmdSettingFullScan 
         Height          =   315
         Left            =   4920
         TabIndex        =   77
         Top             =   1680
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   556
         Icon            =   "frmMain.frx":20C75
         Style           =   2
         Caption         =   "Ca2i d9a85t ca61u hi2nh que1t"
         IconAlign       =   3
         iNonThemeStyle  =   2
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
         Icon            =   "frmMain.frx":20C91
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
         Icon            =   "frmMain.frx":2122B
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
         Icon            =   "frmMain.frx":21C3D
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
         Icon            =   "frmMain.frx":2264F
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
         Icon            =   "frmMain.frx":22BE9
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
   Begin UniControls.UniFrame fm 
      Height          =   6615
      Index           =   13
      Left            =   9000
      Top             =   6480
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   11668
      MaskColor       =   16711935
      FrameColor      =   -2147483629
      Style           =   0
      Caption         =   "Tie65n I1ch"
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
      Begin UniControls.UniCommonDialog Dialog1 
         Left            =   6240
         Top             =   1680
         _ExtentX        =   714
         _ExtentY        =   688
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
      Begin UniControls.UniButton cmdDelRe 
         Height          =   375
         Left            =   5640
         TabIndex        =   85
         Top             =   3000
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         Icon            =   "frmMain.frx":235FB
         Style           =   2
         Caption         =   "Xo1a Bo3"
         IconAlign       =   3
         iNonThemeStyle  =   2
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
      Begin UniControls.UniButton cmdAddRebot 
         Height          =   375
         Left            =   5640
         TabIndex        =   84
         Top             =   2520
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         Icon            =   "frmMain.frx":23B95
         Style           =   2
         Caption         =   "The6m Va2o"
         IconAlign       =   3
         iNonThemeStyle  =   2
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
      Begin UniControls.UniListBox List1 
         Height          =   1935
         Left            =   240
         TabIndex        =   83
         Top             =   2520
         Width           =   5295
         _ExtentX        =   9340
         _ExtentY        =   3413
         IconMaskColor   =   16711935
         Picture         =   "frmMain.frx":2412F
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         RowHeight       =   19
      End
      Begin UniControls.UniButton cmdRebotDel 
         Height          =   375
         Left            =   5640
         TabIndex        =   82
         Top             =   4080
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         Icon            =   "frmMain.frx":2414B
         Style           =   2
         Caption         =   "Xa1c nha65n"
         IconAlign       =   3
         iNonThemeStyle  =   2
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
      Begin UniControls.UniLabel UniLabel84 
         Height          =   255
         Left            =   120
         Top             =   2040
         Width           =   6855
         _ExtentX        =   12091
         _ExtentY        =   450
         Alignment       =   1
         Caption         =   "Xoa1 File Kho6ng The63 Xo1a"
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
      Begin UniControls.UniButton cmdTTCheckAll 
         Height          =   375
         Left            =   3720
         TabIndex        =   75
         Top             =   720
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         Icon            =   "frmMain.frx":246E5
         Style           =   2
         Caption         =   "D9a1nh da61u ta61t ca3"
         IconAlign       =   3
         iNonThemeStyle  =   2
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
      Begin UniControls.UniCheckBox chkTangToc 
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   72
         Top             =   840
         Width           =   2775
         _ExtentX        =   4895
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
         Caption         =   "Ta8ng to61c d9o65 truy xua61t Start Menu."
         ForeColor       =   0
         ShowFocusRectangle=   0   'False
      End
      Begin UniControls.UniButton cmdTangToc 
         Height          =   375
         Left            =   3720
         TabIndex        =   71
         Top             =   1080
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         Icon            =   "frmMain.frx":24C7F
         Style           =   2
         Caption         =   "Thu75c Hie65n"
         IconAlign       =   3
         iNonThemeStyle  =   2
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
      Begin UniControls.UniLabel UniLabel77 
         Height          =   375
         Left            =   120
         Top             =   360
         Width           =   6855
         _ExtentX        =   12091
         _ExtentY        =   661
         Alignment       =   1
         Caption         =   "Ca1c thie61t la65p giu1p ta8ng to61c ma1y ti1nh"
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
      Begin UniControls.UniCheckBox chkTangToc 
         Height          =   195
         Index           =   7
         Left            =   240
         TabIndex        =   73
         Top             =   1080
         Width           =   2985
         _ExtentX        =   5265
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
         Caption         =   "Ta8ng to61c ta81t ma1y (tu2y ca61u hi2nh ma1y)."
         ForeColor       =   0
         ShowFocusRectangle=   0   'False
      End
      Begin UniControls.UniCheckBox chkTangToc 
         Height          =   195
         Index           =   8
         Left            =   240
         TabIndex        =   74
         Top             =   1320
         Width           =   3465
         _ExtentX        =   6112
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
         Caption         =   "Ta8ng to61c kho73i d9o65ng ma1y (tu2y ca61u hi2nh ma1y)."
         ForeColor       =   0
         ShowFocusRectangle=   0   'False
      End
   End
   Begin UniControls.UniFrame mf 
      Height          =   6615
      Left            =   3000
      Top             =   1440
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   11668
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
      Begin UniControls.UniLabel xlblComputerName 
         Height          =   255
         Left            =   2760
         Top             =   1200
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   450
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   49152
      End
      Begin UniControls.ProgressBar xProcessRAM 
         Height          =   255
         Left            =   240
         TabIndex        =   81
         Top             =   2760
         Width           =   6615
         _ExtentX        =   11668
         _ExtentY        =   450
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BrushStyle      =   0
         Color           =   16744703
         ShowText        =   -1  'True
         Value           =   100
      End
      Begin UniControls.UniLabel UniLabel83 
         Height          =   255
         Left            =   120
         Top             =   1680
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   450
         Alignment       =   2
         Caption         =   "Ca1c chu7o7ng tri2nh d9ang cha5y:"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   33023
      End
      Begin UniControls.UniLabel UniLabel80 
         Height          =   255
         Left            =   1320
         Top             =   1200
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         Alignment       =   2
         Caption         =   "Te6n ma1y ti1nh:"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   33023
      End
      Begin UniControls.UniLabel UniLabel79 
         Height          =   255
         Left            =   120
         Top             =   720
         Width           =   6855
         _ExtentX        =   12091
         _ExtentY        =   450
         Alignment       =   1
         Caption         =   "Tho6ng tin ve62 ma1y ti1nh"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   16744576
      End
      Begin UniControls.UniLabel UniLabel78 
         Height          =   255
         Left            =   240
         Top             =   3240
         Width           =   6735
         _ExtentX        =   11880
         _ExtentY        =   450
         Alignment       =   1
         Caption         =   "Tra5ng tha1i ca1c chu71c na8ng"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   16744576
      End
      Begin UniControls.UniLabel UniLabel71 
         Height          =   255
         Left            =   240
         Top             =   5880
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   450
         Caption         =   "Kho73i D9o65ng Cu2ng Ma1y:..............................................................."
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   12582912
      End
      Begin UniControls.UniLabel UniLabel70 
         Height          =   255
         Left            =   240
         Top             =   5520
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   450
         Caption         =   "Tho6ng Ba1o Khi Pha1t Hie65n USB Ke61t No61i:....................................."
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   12582912
      End
      Begin UniControls.UniLabel UniLabel69 
         Height          =   255
         Left            =   240
         Top             =   5160
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   450
         Caption         =   "Tu75 D9o65ng Pha1t Hie65n && Ca3nh Ba1o Virus:...................................."
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   12582912
      End
      Begin UniControls.UniLabel UniLabel68 
         Height          =   255
         Left            =   240
         Top             =   4800
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   450
         Caption         =   "Tu75 D9o65ng Que1t Trong Thu7 Mu5c D9ang Mo73:................................."
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   12582912
      End
      Begin UniControls.UniLabel UniLabel67 
         Height          =   255
         Left            =   240
         Top             =   4440
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   450
         Caption         =   "Que1t File Tru7o71c Khi Cha5y:........................................................."
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   12582912
      End
      Begin UniControls.UniLabel UniLabel66 
         Height          =   255
         Left            =   240
         Top             =   4080
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   450
         Caption         =   "Tu75 D9o65ng Ba3o Ve65 Autorun:......................................................."
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   12582912
      End
      Begin UniControls.UniLabel UniLabel65 
         Height          =   255
         Left            =   240
         Top             =   3720
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   450
         Caption         =   "Tu75 D9o65ng Ba3o Ve65 Registry:......................................................."
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   12582912
      End
      Begin UniControls.UniLabel UniLabel64 
         Height          =   495
         Left            =   240
         Top             =   120
         Width           =   6615
         _ExtentX        =   11668
         _ExtentY        =   873
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
      Begin UniControls.UniLabel UniLabel72 
         Height          =   255
         Left            =   240
         Top             =   6240
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   450
         Caption         =   "Tu75 D9o65ng Update:......................................................................"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   12582912
      End
      Begin UniControls.UniLabel UniLabel81 
         Height          =   255
         Left            =   1680
         Top             =   1440
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   450
         Alignment       =   2
         Caption         =   "Te6n User:"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   33023
      End
      Begin UniControls.UniLabel UniLabel82 
         Height          =   255
         Left            =   480
         Top             =   1920
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   450
         Alignment       =   2
         Caption         =   "Dung lu7o75ng bo65 nho71 RAM:"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   33023
      End
      Begin UniControls.UniLabel UniLabel85 
         Height          =   255
         Left            =   840
         Top             =   2160
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   450
         Alignment       =   2
         Caption         =   "Ti2nh tra5ng ma1y ti1nh:"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   16711935
      End
      Begin UniControls.UniLabel xlblUserName 
         Height          =   255
         Left            =   2760
         Top             =   1440
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   450
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   49152
      End
      Begin UniControls.UniLabel xlblProcess 
         Height          =   255
         Left            =   2760
         Top             =   1680
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   450
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   49152
      End
      Begin UniControls.UniLabel xlblRAM 
         Height          =   255
         Left            =   2760
         Top             =   1920
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   450
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   49152
      End
      Begin UniControls.UniLabel xlblTinhTrang 
         Height          =   615
         Left            =   2760
         Top             =   2160
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   1085
         BackStyle       =   0
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   49152
      End
      Begin VB.Image PicOnOff 
         Height          =   375
         Index           =   7
         Left            =   6240
         Stretch         =   -1  'True
         Top             =   6120
         Width           =   375
      End
      Begin VB.Image PicOnOff 
         Height          =   375
         Index           =   6
         Left            =   6240
         Stretch         =   -1  'True
         Top             =   5760
         Width           =   375
      End
      Begin VB.Image PicOnOff 
         Height          =   375
         Index           =   5
         Left            =   6240
         Stretch         =   -1  'True
         Top             =   5400
         Width           =   375
      End
      Begin VB.Image PicOnOff 
         Height          =   375
         Index           =   4
         Left            =   6240
         Stretch         =   -1  'True
         Top             =   5040
         Width           =   375
      End
      Begin VB.Image PicOnOff 
         Height          =   375
         Index           =   3
         Left            =   6240
         Stretch         =   -1  'True
         Top             =   4680
         Width           =   375
      End
      Begin VB.Image PicOnOff 
         Height          =   375
         Index           =   2
         Left            =   6240
         Stretch         =   -1  'True
         Top             =   4320
         Width           =   375
      End
      Begin VB.Image PicOnOff 
         Height          =   375
         Index           =   1
         Left            =   6240
         Stretch         =   -1  'True
         Top             =   3960
         Width           =   375
      End
      Begin VB.Image PicOnOff 
         Height          =   375
         Index           =   0
         Left            =   6240
         Stretch         =   -1  'True
         Top             =   3600
         Width           =   375
      End
   End
   Begin VB.Image picLogoPAV 
      Height          =   1335
      Left            =   0
      Picture         =   "frmMain.frx":25219
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

Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
   
Private Declare Function FindFirstFile _
                Lib "kernel32" _
                Alias "FindFirstFileA" (ByVal lpFileName As String, _
                                        lpFindFileData As WIN32_FIND_DATA) As Long
   
Private Declare Function FindNextFile _
                Lib "kernel32" _
                Alias "FindNextFileA" (ByVal hFindFile As Long, _
                                       lpFindFileData As WIN32_FIND_DATA) As Long

Private Declare Function GetTickCount Lib "kernel32" () As Long
'-----> API For Search

Private Declare Function ShellExecute _
                Lib "shell32.dll" _
                Alias "ShellExecuteA" (ByVal hWnd As Long, _
                                       ByVal lpOperation As String, _
                                       ByVal lpFile As String, _
                                       ByVal lpParameters As String, _
                                       ByVal lpDirectory As String, _
                                       ByVal nShowCmd As Long) As Long

'-----> Dim for Connect Database
Dim db           As Database

Dim rs           As Recordset

Dim WS           As Workspace

Dim Max          As Long

Dim mData        As Recordset

'<------ Dim for Connect Database

'---> Dim For FullScan and Custom Scan
Dim xFSStopScan  As Boolean

Dim xFSStopScan2 As Boolean

Dim LLVV         As UniListView

Dim xTotalFile

Dim HoanThanhFull As Boolean

Dim HoanThanhCus  As Boolean

Dim TimeFull      As String

Dim TimeCus       As String

'<--- Dim For FullScan and Custom Scan

'---> Dim for Timer of Full Scan
Dim xTimeSC

Dim xTime

'<--- Dim for Timer of Full Scan

'---> Check Full scan or Custom Scan
Dim xCustomScan As Boolean

'<--- Check Full scan or Custom Scan

'---> Dim for Messenger form
Public xStart
'<--- Dim for Messenger form

Public xTask As Boolean

Private Function SearchFile(Path, FileName)

        '<EhHeader>
        On Error GoTo SearchFile_Err

        '</EhHeader>

        Dim FP     As FILE_PARAMS  'holds search parameters

        Dim tstart As Single   'timer var for this routine only

        Dim tend   As Single     'timer var for this routine only

100     With FP
102         .sFileRoot = Path       'start path
104         .sFileNameExt = FileName    'file type of interest
106         .bRecurse = 1 ' Check1.Value = 1  '1 = recursive search
        End With

108     tstart = GetTickCount()
110     Call SearchForFiles(FP)
112     tend = GetTickCount()

        '<EhFooter>
        Exit Function

SearchFile_Err:
        WriteErr Time & "-" & Date & " - " & Err.Description & " - " & "in PerfectAntivirus2009.frmMain.SearchFile " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Function

Private Sub GetFileInformation(FP As FILE_PARAMS)

        '<EhHeader>
        On Error GoTo GetFileInformation_Err

        '</EhHeader>

        'On Error Resume Next
100     DoEvents

        Dim WFD     As WIN32_FIND_DATA

        Dim hFile   As Long

        Dim sPath   As String

        Dim sRoot   As String

        Dim sTmp    As String

        Dim SKetQua As String

102     sRoot = QualifyPath(FP.sFileRoot)
104     sPath = sRoot & FP.sFileNameExt
106     hFile = FindFirstFile(sPath, WFD)

108     If hFile <> INVALID_HANDLE_VALUE Then

            Do

110             If Not (WFD.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) = FILE_ATTRIBUTE_DIRECTORY Then
112                 FP.Count = FP.Count + 1
114                 sTmp = TrimNull(WFD.cFileName)
116                 SKetQua = sRoot & sTmp

118                 If xCustomScan = False And xFSStopScan = True Then GoTo DuNgQuEt
120                 If xCustomScan = True And xFSStopScan2 = True Then GoTo DuNgQuEt
122                 If modScanVirus.FileExists(SKetQua) = True Then
124                     If FileLen(SKetQua) > 0 And FileLen(SKetQua) < (Me.VSLimitSize.Value * 1000000) Then '8350E5A3E24C153DF2275C9F80692773

126                         If xCustomScan = True Then
128                             cslblStatus.Caption = SKetQua

130                             If xFSStopScan2 = True Then Exit Sub
                            Else
132                             xTotalFile = xTotalFile + 1
134                             lblStatus.Caption = SKetQua
    
                            End If

                            'SetAttr SKetQua, vbNormal
            
136                         DoEvents

                            Dim AX As String

138                         AX = modScanVirus.CheckVirus(SKetQua)

140                         If AX <> "No" Then
                                'Virus Found
                                'List1.AddItem "!!!!! " & AX & " - " & SKetQua
                
142                             If xCustomScan = False Then

                                    Dim I

144                                 I = LVVirus1.ListItems.Count + 1
146                                 LVVirus1.ListItems.Add I, , AX
148                                 LVVirus1.ListItems(I).SubItems(1).Caption = SKetQua
150                                 LVVirus1.ListItems(I).SubItems(2).Caption = FileLen(SKetQua) & " Bytes"
152                                 LVVirus1.ListItems(I).SubItems(3).Caption = CheckProcess(SKetQua)
154                                 LVVirus1.ListItems(I).SubItems(4).Caption = "---"
156                                 LVVirus1.ListItems(I).Checked = True
                    
                                Else

                                    Dim ig

158                                 ig = LVVirus2.ListItems.Count + 1
160                                 LVVirus2.ListItems.Add ig, , AX
162                                 LVVirus2.ListItems(ig).SubItems(1).Caption = SKetQua
164                                 LVVirus2.ListItems(ig).SubItems(2).Caption = FileLen(SKetQua) & " Bytes"
166                                 LVVirus2.ListItems(ig).SubItems(3).Caption = CheckProcess(SKetQua)
168                                 LVVirus2.ListItems(ig).SubItems(4).Caption = "---"
                    
170                                 LVVirus2.ListItems(ig).Checked = True
                                End If
                
                            End If ' AX="No"
                        End If 'FileLen(SKetQua) > 0 And FileLen(SKetQua) < 1000000
                    End If ' file exists

                    'Text1.Text = Text1.Text & SKetQua & vbCrLf
                    '*********************************************
            
                End If

172         Loop While FindNextFile(hFile, WFD)

174         hFile = FindClose(hFile)
        End If

176     DoEvents
DuNgQuEt:

        '<EhFooter>
        Exit Sub

GetFileInformation_Err:
        WriteErr Time & "-" & Date & " - " & Err.Description & " - " & "in PerfectAntivirus2009.frmMain.GetFileInformation " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub SearchForFiles(FP As FILE_PARAMS)

        'local working variables
        '<EhHeader>
        On Error GoTo SearchForFiles_Err

        '</EhHeader>
        Dim WFD   As WIN32_FIND_DATA

        Dim hFile As Long

        Dim sPath As String

        Dim sRoot As String

        Dim sTmp  As String

100     sRoot = QualifyPath(FP.sFileRoot)
102     sPath = sRoot & "*.*"
104     hFile = FindFirstFile(sPath, WFD)

106     If hFile <> INVALID_HANDLE_VALUE Then
108         Call GetFileInformation(FP)

            Do

110             If (WFD.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) Then
112                 If FP.bRecurse Then
114                     If Asc(WFD.cFileName) <> vbDot Then
116                         FP.sFileRoot = sRoot & TrimNull(WFD.cFileName)
118                         Call SearchForFiles(FP)
                        End If
                    End If
                End If

120         Loop While FindNextFile(hFile, WFD)

122         hFile = FindClose(hFile)
        End If

        '<EhFooter>
        Exit Sub

SearchForFiles_Err:
        WriteErr Time & "-" & Date & " - " & Err.Description & " - " & "in PerfectAntivirus2009.frmMain.SearchForFiles " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Function TrimNull(startstr As String) As String

        '<EhHeader>
        On Error GoTo TrimNull_Err

        '</EhHeader>
        Dim pos As Integer

100     pos = InStr(startstr, Chr$(0))

102     If pos Then
104         TrimNull = Left$(startstr, pos - 1)

            Exit Function

        End If

106     TrimNull = startstr

        '<EhFooter>
        Exit Function

TrimNull_Err:
        WriteErr Time & "-" & Date & " - " & Err.Description & " - " & "in PerfectAntivirus2009.frmMain.TrimNull " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Function

Private Function QualifyPath(sPath As String) As String

        '<EhHeader>
        On Error GoTo QualifyPath_Err

        '</EhHeader>
100     If Right$(sPath, 1) <> "\" Then
102         QualifyPath = sPath & "\"
        Else
104         QualifyPath = sPath
        End If

        '<EhFooter>
        Exit Function

QualifyPath_Err:
        WriteErr Time & "-" & Date & " - " & Err.Description & " - " & "in PerfectAntivirus2009.frmMain.QualifyPath " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Function

Public Sub SaveREG()

        '<EhHeader>
        On Error GoTo SaveREG_Err

        '</EhHeader>
100     If FileExists(AppPath & "RegProtect.dat") = True Then modScanVirus.DeleteFile (AppPath & "RegProtect.dat")

102     With frmMain.atpLVREG

            Dim r

104         For r = 1 To .ListItems.Count
106             WriteIniFile AppPath & "RegProtect.dat", r, "TenChucNang", Unicode2UTF8(.ListItems(r).Text)
108             WriteIniFile AppPath & "RegProtect.dat", r, "KeyGoc", .ListItems(r).SubItems(1).Caption
110             WriteIniFile AppPath & "RegProtect.dat", r, "KeyPath", .ListItems(r).SubItems(2).Caption
112             WriteIniFile AppPath & "RegProtect.dat", r, "KeyName", .ListItems(r).SubItems(3).Caption
114             WriteIniFile AppPath & "RegProtect.dat", r, "KeyData", .ListItems(r).SubItems(4).Caption
116         Next r

118         WriteIniFile AppPath & "RegProtect.dat", "Other", "Total", .ListItems.Count
    
        End With

        '<EhFooter>
        Exit Sub

SaveREG_Err:
        WriteErr Time & "-" & Date & " - " & Err.Description & " - " & "in PerfectAntivirus2009.frmMain.SaveREG " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub atpcmdAutorun_Click()

        '<EhHeader>
        On Error GoTo atpcmdAutorun_Click_Err

        '</EhHeader>
100     If Me.atptmrAutorun.Enabled = True Then
102         Me.atptmrAutorun.Enabled = False
104         Me.atplblStaAutorun.Caption = "D9ang ta81t"
106         Me.atpcmdAutorun.Caption = "Mo73 chu71c na8ng na2y"
        Else
108         Me.atptmrAutorun.Enabled = True
110         Me.atplblStaAutorun.Caption = "D9ang mo73"
112         Me.atpcmdAutorun.Caption = "Ta81t chu71c na8ng na2y"
        End If

        '<EhFooter>
        Exit Sub

atpcmdAutorun_Click_Err:
        WriteErr Time & "-" & Date & " - " & Err.Description & " - " & "in PerfectAntivirus2009.frmMain.atpcmdAutorun_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub atpcmdREG_Click()

        '<EhHeader>
        On Error GoTo atpcmdREG_Click_Err

        '</EhHeader>
100     If Me.atptmrREG.Enabled = True Then
102         Me.atptmrREG.Enabled = False
104         Me.atplbREG.Caption = "D9ang ta81t"
106         Me.atpcmdREG.Caption = "Mo73 chu71c na8ng na2y"
        Else
108         Me.atptmrREG.Enabled = True
110         Me.atplbREG.Caption = "D9ang mo73"
112         Me.atpcmdREG.Caption = "Ta81t chu71c na8ng na2y"
        End If

        '<EhFooter>
        Exit Sub

atpcmdREG_Click_Err:
        WriteErr Time & "-" & Date & " - " & Err.Description & " - " & "in PerfectAntivirus2009.frmMain.atpcmdREG_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub atpREGAdd_Click()

        '<EhHeader>
        On Error GoTo atpREGAdd_Click_Err

        '</EhHeader>

100     pfrmAddREG.Show

        '<EhFooter>
        Exit Sub

atpREGAdd_Click_Err:
        WriteErr Time & "-" & Date & " - " & Err.Description & " - " & "in PerfectAntivirus2009.frmMain.atpREGAdd_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub atptmrAutorun_Timer()

    '<EhHeader>
    On Error Resume Next

    '</EhHeader>
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

        If UCase(drv.DriveLetter) <> "A" And drv.DriveType <> CDRom Then
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

                Dim Xa As New frmAddAutorun

                If Xx1 <> drv.DriveLetter & ":\" And modScanVirus.FileExists(Xx1) = True Then
                
                    If modScanVirus.CheckVirus(Xx1) = "No" Then
                        '////////////////
                        Xa.Show
                        Xa.txtPath.Text = Xx1
                        Xa.lblMain.Caption = "Chu7o7ng tri2nh pha1t hie65n tha61y file [" & GetFileName(Xx1) & "] co1 the63 la2 Virus."
                        GetIconFromFile Xx1, Xa.picIcon
                        Xa.lblMD5.Caption = GetMD5(Xx1)
                        '////////////////////
                    End If
                    
                    xFileName1 = GetFileName(Xx1)

                    If CheckProcess(xFileName1) <> 0 Then KillProcessById CheckProcess(xFileName1) 'EndTask xFileName1
                    SetAttr Xx1, vbNormal
                    modScanVirus.DeleteFile Xx1
                    frmMessenger.zShowMessenger "Pha1t hie65n Virus!", "D9a4 pha1t hie65n Virus ta5i: [" & drv.DriveLetter & ":\...\" & GetFileName(Xx1) & "]. Ti2nh tra5ng: D9a4 Xo1a", 5000, xvang

                End If

                If Xx2 <> drv.DriveLetter & ":\" And modScanVirus.FileExists(Xx2) = True Then
                
                    If modScanVirus.CheckVirus(Xx2) = "No" Then
                        '////////////////
                        Xa.Show
                        Xa.txtPath.Text = Xx2
                        Xa.lblMain.Caption = "Chu7o7ng tri2nh pha1t hie65n tha61y file [" & GetFileName(Xx2) & "] co1 the63 la2 Virus."
                        GetIconFromFile Xx2, Xa.picIcon
                        Xa.lblMD5.Caption = GetMD5(Xx2)
                        '////////////////////
                    End If
                    
                    xFileName2 = GetFileName(Xx2)

                    If CheckProcess(xFileName2) <> 0 Then KillProcessById CheckProcess(xFileName2) 'EndTask xFileName2
                    SetAttr Xx2, vbNormal
                    modScanVirus.DeleteFile Xx2
                    frmMessenger.zShowMessenger "Pha1t hie65n Virus!", "D9a4 pha1t hie65n Virus ta5i: [" & drv.DriveLetter & ":\...\" & GetFileName(Xx2) & "]. Ti2nh tra5ng: D9a4 Xo1a", 5000, xvang
                End If

                '////////// End / Tim nguon goc Virus ///////////

                '////// Diet Autorun ////////
                On Error GoTo KhOnGtHeXoAaUtOrUn

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

    Exit Sub

KhOnGtHeXoAaUtOrUn:
    frmMessenger.zShowMessenger "Ba3o ve65 Autorun", "Kho6ng the63 xo1a d9u7o75c: [" & drv.DriveLetter & ":\autorun.inf]", 5000, xvang

    Resume Next

End Sub

Private Sub atptmrREG_Timer()

    '<EhHeader>
    On Error Resume Next

    '</EhHeader>

    With frmMain.atpLVREG

        Dim Y

        Dim u

        Dim KQ As String

BaTdAuVoNgFoR:

        For Y = 1 To .ListItems.Count

            If .ListItems(Y).SubItems(1).Caption = "HKEY_CLASSES_ROOT" Then
                u = &H80000000
            ElseIf .ListItems(Y).SubItems(1).Caption = "HKEY_CURRENT_USER" Then
                u = &H80000001
            ElseIf .ListItems(Y).SubItems(1).Caption = "HKEY_LOCAL_MACHINE" Then
                u = &H80000002
            ElseIf .ListItems(Y).SubItems(1).Caption = "HKEY_USERS" Then
                u = &H80000003
            ElseIf .ListItems(Y).SubItems(1).Caption = "HKEY_CURRENT_CONFIG" Then
                u = &H80000005
            End If

            On Error GoTo ThOaTkHoIvOnGfOr

            If (GetString(u, .ListItems(Y).SubItems(2).Caption, .ListItems(Y).SubItems(3).Caption) <> .ListItems(Y).SubItems(4).Caption) And (GetString(u, .ListItems(Y).SubItems(2).Caption, .ListItems(Y).SubItems(3).Caption)) <> "" Then
                If UCase(.ListItems(Y).SubItems(3).Caption) = "DISABLETASKMGR" Then
                    DeleteValue u, .ListItems(Y).SubItems(2).Caption, .ListItems(Y).SubItems(3).Caption
                Else
                    SaveString u, .ListItems(Y).SubItems(2).Caption, .ListItems(Y).SubItems(3).Caption, .ListItems(Y).SubItems(4).Caption
                End If

                KQ = .ListItems(Y).Text
                GoTo EndATP
            ElseIf (GetDWORD(u, .ListItems(Y).SubItems(2).Caption, .ListItems(Y).SubItems(3).Caption) <> .ListItems(Y).SubItems(4).Caption) Then
                SaveDWORD u, .ListItems(Y).SubItems(2).Caption, .ListItems(Y).SubItems(3).Caption, .ListItems(Y).SubItems(4).Caption
                KQ = .ListItems(Y).Text
                GoTo EndATP
            End If

ThOaTkHoIvOnGfOr:
        Next Y

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
    KQ = ""
    GoTo BaTdAuVoNgFoR
End Sub

Private Sub atsAlwaysScanFolder_Click()

        '<EhHeader>
        On Error GoTo atsAlwaysScanFolder_Click_Err

        '</EhHeader>
100     With atsAlwaysScanFolder

102         If .Value = True Then
104             .Caption = "Tu75 d9o65ng que1t va2 ca3nh ba1o Virus trong Thu7 mu5c d9ang d9u7o75c mo73. [D9ang Ba65t]"
106             WriteIniFile AppPath & "Setting.ini", "AutoScan", "ScanFolder", .Value
108             Load zfrmAutoScanFolder
            Else
110             .Caption = "Tu75 d9o65ng que1t va2 ca3nh ba1o Virus trong Thu7 mu5c d9ang d9u7o75c mo73. [D9ang Ta81t]"
112             WriteIniFile AppPath & "Setting.ini", "AutoScan", "ScanFolder", .Value
114             Unload zfrmAutoScanFolder
    
            End If

        End With

        '<EhFooter>
        Exit Sub

atsAlwaysScanFolder_Click_Err:
        WriteErr Time & "-" & Date & " - " & Err.Description & " - " & "in PerfectAntivirus2009.frmMain.atsAlwaysScanFolder_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub atsAutoScanUSB_Click()

        '<EhHeader>
        On Error GoTo atsAutoScanUSB_Click_Err

        '</EhHeader>
100     With atsAutoScanUSB

102         If .Value = True Then
104             .Caption = "Tu75 d9o65ng que1t Virus cho USB khi pha1t hie65n USB ke61t no61i va2o ma1y ti1nh. [D9ang Ba65t]"
106             WriteIniFile AppPath & "Setting.ini", "AutoScan", "ScanUSB", .Value
                '////
108             Load zfrmScanUSB
                '///
            Else
110             .Caption = "Tu75 d9o65ng que1t Virus cho USB khi pha1t hie65n USB ke61t no61i va2o ma1y ti1nh. [D9ang Ta81t]"
112             WriteIniFile AppPath & "Setting.ini", "AutoScan", "ScanUSB", .Value
                '///
114             Unload zfrmScanUSB
                '///
    
            End If

        End With

        '<EhFooter>
        Exit Sub

atsAutoScanUSB_Click_Err:
        WriteErr Time & "-" & Date & " - " & Err.Description & " - " & "in PerfectAntivirus2009.frmMain.atsAutoScanUSB_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub atsScanEXE_Click()

        '<EhHeader>
        On Error GoTo atsScanEXE_Click_Err

        '</EhHeader>
100     With atsScanEXE

102         If .Value = True Then
104             .Caption = "Tu75 d9o65ng que1t va2 ca3nh ba1o Virus cho nhu74ng File sa81p d9u7o75c mo73. [D9ang Ba65t]"
106             WriteIniFile AppPath & "Setting.ini", "AutoScan", "ScanEXE", .Value
108             SaveString HKEY_CLASSES_ROOT, "exefile\shell\open\command", "", ChrW(34) & AppPath & "Check_Virus.exe" & ChrW(34) & " " & ChrW(34) & "%1" & ChrW(34)
110             SaveString HKEY_CLASSES_ROOT, "batfile\shell\open\command", "", ChrW(34) & AppPath & "Check_Virus.exe" & ChrW(34) & " " & ChrW(34) & "%1" & ChrW(34)
112             SaveString HKEY_CLASSES_ROOT, "cmdfile\shell\open\command", "", ChrW(34) & AppPath & "Check_Virus.exe" & ChrW(34) & " " & ChrW(34) & "%1" & ChrW(34)
114             SaveString HKEY_CLASSES_ROOT, "comfile\shell\open\command", "", ChrW(34) & AppPath & "Check_Virus.exe" & ChrW(34) & " " & ChrW(34) & "%1" & ChrW(34)
    
            Else
    
116             .Caption = "Tu75 d9o65ng que1t va2 ca3nh ba1o Virus cho nhu74ng File sa81p d9u7o75c mo73. [D9ang Ta81t]"
118             WriteIniFile AppPath & "Setting.ini", "AutoScan", "ScanEXE", .Value
120             SaveString HKEY_CLASSES_ROOT, "exefile\shell\open\command", "", ChrW(34) & "%1" & ChrW(34) & " %*"
122             SaveString HKEY_CLASSES_ROOT, "batfile\shell\open\command", "", ChrW(34) & "%1" & ChrW(34) & " %*"
124             SaveString HKEY_CLASSES_ROOT, "cmdfile\shell\open\command", "", ChrW(34) & "%1" & ChrW(34) & " %*"
126             SaveString HKEY_CLASSES_ROOT, "comfile\shell\open\command", "", ChrW(34) & "%1" & ChrW(34) & " %*"
    
            End If

        End With

        '<EhFooter>
        Exit Sub

atsScanEXE_Click_Err:
        WriteErr Time & "-" & Date & " - " & Err.Description & " - " & "in PerfectAntivirus2009.frmMain.atsScanEXE_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub atsScanKeylogger_Click()

        '<EhHeader>
        On Error GoTo atsScanKeylogger_Click_Err

        '</EhHeader>
100     With atsScanKeylogger

102         If .Value = True Then
104             .Caption = "Tu75 d9o65ng pha1t hie65n va2 ca3nh ba1o Keylogger. [D9ang Ba65t]"
106             WriteIniFile AppPath & "Setting.ini", "AutoScan", "ScanKeylogger", .Value
                '////
108             Load zfrmAntiKey
                '///
            Else
110             .Caption = "Tu75 d9o65ng pha1t hie65n va2 ca3nh ba1o Keylogger. [D9ang Ta81t]"
112             WriteIniFile AppPath & "Setting.ini", "AutoScan", "ScanKeylogger", .Value
                '///
114             Unload zfrmAntiKey
                '///
    
            End If

        End With

        '<EhFooter>
        Exit Sub

atsScanKeylogger_Click_Err:
        WriteErr Time & "-" & Date & " - " & Err.Description & " - " & "in PerfectAntivirus2009.frmMain.atsScanKeylogger_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub chkAutoStart_Click()

        '<EhHeader>
        On Error GoTo chkAutoStart_Click_Err

        '</EhHeader>
100     If chkAutoStart.Value = True Then
102         SaveString HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Run", "PAV2009", AppPath & "PAV2009.exe /task"
        Else
104         DeleteValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Run", "PAV2009"
        End If

        '<EhFooter>
        Exit Sub

chkAutoStart_Click_Err:
        WriteErr Time & "-" & Date & " - " & Err.Description & " - " & "in PerfectAntivirus2009.frmMain.chkAutoStart_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub chkUseFastScan_Click()

        '<EhHeader>
        On Error GoTo chkUseFastScan_Click_Err

        '</EhHeader>
100     If chkUseFastScan.Value = True Then
102         SaveString HKEY_CLASSES_ROOT, "Folder\shell\[PAV 2009] Quét Virus\command", "", ChrW(34) & AppPath & "PAVMiniScan.exe" & ChrW(34) & " " & ChrW(34) & "%1" & ChrW(34)
        Else
104         DeleteKey HKEY_CLASSES_ROOT, "Folder\shell\[PAV 2009] Quét Virus\command"
106         DeleteKey HKEY_CLASSES_ROOT, "Folder\shell\[PAV 2009] Quét Virus"
        End If

        '<EhFooter>
        Exit Sub

chkUseFastScan_Click_Err:
        WriteErr Time & "-" & Date & " - " & Err.Description & " - " & "in PerfectAntivirus2009.frmMain.chkUseFastScan_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub cmdAddRebot_Click()

        '<EhHeader>
        On Error GoTo cmdAddRebot_Click_Err

        '</EhHeader>
        Dim strApp As String

100     Dialog1.FileName = ""
102     Dialog1.Filter = "All File (*.*)|*.*|"
104     Dialog1.ShowOpen
106     strApp = Dialog1.FileName

108     If strApp <> "" Then List1.AddItem strApp

        '<EhFooter>
        Exit Sub

cmdAddRebot_Click_Err:
        WriteErr Time & "-" & Date & " - " & Err.Description & " - " & "in PerfectAntivirus2009.frmMain.cmdAddRebot_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub cmdDelRe_Click()

        '<EhHeader>
        On Error GoTo cmdDelRe_Click_Err

        '</EhHeader>
100     If List1.ListCount <> 0 Then List1.Remove List1.ListIndex

        '<EhFooter>
        Exit Sub

cmdDelRe_Click_Err:
        WriteErr Time & "-" & Date & " - " & Err.Description & " - " & "in PerfectAntivirus2009.frmMain.cmdDelRe_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub cmdDelREG_Click()

        '<EhHeader>
        On Error GoTo cmdDelREG_Click_Err

        '</EhHeader>

100     DelAllChecked Me.atpLVREG
102     SaveREG

        '<EhFooter>
        Exit Sub

cmdDelREG_Click_Err:
        WriteErr Time & "-" & Date & " - " & Err.Description & " - " & "in PerfectAntivirus2009.frmMain.cmdDelREG_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub cmdDelSelected_Click()

        '<EhHeader>
        On Error GoTo cmdDelSelected_Click_Err

        '</EhHeader>
100     If UniMsgBox("Ba5n cha81c cha81n?", vbYesNo + vbQuestion) = vbYes Then
102         DelAllChecked LVVirusEvents

104         If FileExists(AppPath & "VirusScanLog.log") = True Then modScanVirus.DeleteFile (AppPath & "VirusScanLog.log")

106         With frmMain.LVVirusEvents

                Dim l

108             For l = 1 To .ListItems.Count
110                 WriteIniFile AppPath & "VirusScanLog.log", l, "TimeQuet", .ListItems(l).Text
112                 WriteIniFile AppPath & "VirusScanLog.log", l, "KieuQuet", .ListItems(l).SubItems(1).Caption
114                 WriteIniFile AppPath & "VirusScanLog.log", l, "SoFile", .ListItems(l).SubItems(2).Caption
116                 WriteIniFile AppPath & "VirusScanLog.log", l, "SoVirus", .ListItems(l).SubItems(3).Caption
118                 WriteIniFile AppPath & "VirusScanLog.log", l, "KetQua", .ListItems(l).SubItems(4).Caption
120                 WriteIniFile AppPath & "VirusScanLog.log", "Other", "Total", .ListItems.Count
122             Next l

            End With

        End If

        '<EhFooter>
        Exit Sub

cmdDelSelected_Click_Err:
        WriteErr Time & "-" & Date & " - " & Err.Description & " - " & "in PerfectAntivirus2009.frmMain.cmdDelSelected_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub cmdEventsVirusKillAll_Click()

        '<EhHeader>
        On Error GoTo cmdEventsVirusKillAll_Click_Err

        '</EhHeader>
100     If UniMsgBox("Ba5n cha81c cha81n muo61n xo1a ta61t ca3?", vbYesNo + vbQuestion) = vbYes Then
102         DelAllLV LVVirusEvents
104         modScanVirus.DeleteFile AppPath & "VirusScanLog.log"
        End If

        '<EhFooter>
        Exit Sub

cmdEventsVirusKillAll_Click_Err:
        WriteErr Time & "-" & Date & " - " & Err.Description & " - " & "in PerfectAntivirus2009.frmMain.cmdEventsVirusKillAll_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub cmdFSCachLy_Click()

        '<EhHeader>
        On Error GoTo cmdFSCachLy_Click_Err

        '</EhHeader>
100     If UniMsgBox("Ba5n co1 muo61n ca1ch ly ca1c Virus d9a4 cho5n kho6ng?", vbYesNo + vbQuestion, "Tho6ng Ba1o") = vbYes Then

            On Error GoTo HeHeChOqUa2

102         MkDir AppPath & "VungCachLy\"
HeHeChOqUa2:

            Dim Y, j

            Dim X As String

104         For Y = 1 To LVVirus1.ListItems.Count

106             If LVVirus1.ListItems(Y).Checked = True And FileExists(LVVirus1.ListItems(Y).SubItems(1).Caption) = True Then
    
108                 If LVVirus1.ListItems(Y).SubItems(3).Caption <> "0" Then
                        'Kill process
110                     KillProcessById (LVVirus1.ListItems(Y).SubItems(3).Caption)
112                     X = X & " D9a4 ta81t tie61n tri2nh: " & LVVirus1.ListItems(Y).SubItems(3).Caption & vbCrLf
                    End If

                    On Error GoTo KhOnGtHeCaChLy2

114                 SetAttr LVVirus1.ListItems(Y).SubItems(1).Caption, vbNormal
116                 Name LVVirus1.ListItems(Y).SubItems(1).Caption As LVVirus1.ListItems(Y).SubItems(1).Caption & ".DaCachLy"
118                 FileCopy LVVirus1.ListItems(Y).SubItems(1).Caption & ".DaCachLy", AppPath & "VungCachLy\" & GetFileName(LVVirus1.ListItems(Y).SubItems(1).Caption & ".DaCachLy")
120                 modScanVirus.DeleteFile LVVirus1.ListItems(Y).SubItems(1).Caption & ".DaCachLy"
122                 X = X & " D9a4 Ca1ch Ly: " & LVVirus1.ListItems(Y).SubItems(1).Caption & vbCrLf

                End If

124         Next Y

126         If Not X = "" Then
128             DelAllChecked LVVirus1
130             UniMsgBox X, vbOKOnly + vbInformation, "D9a4 Ca1ch Ly Virus", Me.hWnd
            Else
132             UniMsgBox "Kho6ng co1 Virus na2o d9e63 ca1ch ly.", vbOKOnly + vbInformation, "Tho6ng Ba1o", Me.hWnd
            End If

        End If 'unimsgbox

        Exit Sub

KhOnGtHeCaChLy2:
134     X = X & " Kho6ng THe63 Ca1ch Ly: " & LVVirus1.ListItems(Y).SubItems(1).Caption & vbCrLf

136     Resume Next

        '<EhFooter>
        Exit Sub

cmdFSCachLy_Click_Err:
        WriteErr Time & "-" & Date & " - " & Err.Description & " - " & "in PerfectAntivirus2009.frmMain.cmdFSCachLy_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub cmdFSKillVirus_Click()

        '<EhHeader>
        On Error GoTo cmdFSKillVirus_Click_Err

        '</EhHeader>
100     If UniMsgBox("Ba5n co1 cha81c cha81n muo61n xo1a ca1c Virus d9a4 d9a1nh da61u kho6ng?", vbYesNo + vbQuestion, "Die65t Virus") = vbYes Then

102         DoEvents
104         lblStatus.Caption = "D9ang xo1a..."

            Dim Y

            Dim X As String

106         For Y = 1 To LVVirus1.ListItems.Count

108             If LVVirus1.ListItems(Y).Checked = True And FileExists(LVVirus1.ListItems(Y).SubItems(1).Caption) = True Then

110                 DoEvents

112                 If LVVirus1.ListItems(Y).SubItems(3).Caption <> "0" Then
                        'Kill process
114                     KillProcessById (LVVirus1.ListItems(Y).SubItems(3).Caption)
116                     X = X & " D9a4 ta81t tie61n tri2nh: " & LVVirus1.ListItems(Y).SubItems(3).Caption & vbCrLf
                    End If

118                 DoEvents

120                 If LVVirus1.ListItems(Y).SubItems(4).Caption <> "---" Then

                        'Delete Registry Key
                        Dim a

                        Dim u

                        Dim b

                        Dim c

122                     DoEvents
124                     a = Left(LVVirus1.ListItems(Y).SubItems(4).Caption, Len(LVVirus1.ListItems(Y).SubItems(4).Caption) - InStrRev(StrReverse(LVVirus1.ListItems(Y).SubItems(4).Caption), "-"))
126                     b = Right(LVVirus1.ListItems(Y).SubItems(4).Caption, Len(LVVirus1.ListItems(Y).SubItems(4).Caption) - InStrRev(LVVirus1.ListItems(Y).SubItems(4).Caption, ":"))
128                     c = Mid(LVVirus1.ListItems(Y).SubItems(4).Caption, Len(LVVirus1.ListItems(Y).SubItems(4).Caption) - InStrRev(StrReverse(LVVirus1.ListItems(Y).SubItems(4).Caption), "-") + 2, (InStrRev(LVVirus1.ListItems(Y).SubItems(4).Caption, ":")) - (Len(LVVirus1.ListItems(Y).SubItems(4).Caption) - InStrRev(StrReverse(LVVirus1.ListItems(Y).SubItems(4).Caption), "-")) - 2)

130                     If UCase(a) = "HKEY_CURRENT_USER" Then u = &H80000001
132                     If UCase(a) = "HKEY_LOCAL_MACHINE" Then u = &H80000002
134                     DeleteValue u, c, b
136                     X = X & " D9a4 xo1a Key: " & a & "\" & c & ":" & b & vbCrLf
                    End If
        
                    'HKEY_CURRENT_USER = &H80000001
                    'HKEY_LOCAL_MACHINE = &H80000002
                    On Error GoTo KhOnGtHeXoAbO2

138                 DoEvents
140                 SetAttr LVVirus1.ListItems(Y).SubItems(1).Caption, vbNormal
142                 modScanVirus.DeleteFile LVVirus1.ListItems(Y).SubItems(1).Caption
144                 X = X & " D9a4 Xo1a Bo3: " & LVVirus1.ListItems(Y).SubItems(1).Caption & vbCrLf

                End If

146         Next Y

148         If Not X = "" Then
150             DelAllChecked LVVirus1
152             UniMsgBox X, vbOKOnly + vbInformation, "Nhu74ng vie65c d9a4 la2m.", Me.hWnd

            Else
154             UniMsgBox "Kho6ng co1 Virus na2o d9e63 die65t!", vbOKOnly + vbCritical, "Tho6ng Ba1o", Me.hWnd
            End If

156         lblStatus.Caption = "Sa84n sa2ng"
        End If ' Unimsgbox "ban co chac chan ko?"

        Exit Sub

KhOnGtHeXoAbO2:
158     X = X & " Kho6ng The63 Xo1a Bo3: " & LVVirus1.ListItems(Y).SubItems(1).Caption & vbCrLf

160     Resume Next

        '<EhFooter>
        Exit Sub

cmdFSKillVirus_Click_Err:
        WriteErr Time & "-" & Date & " - " & Err.Description & " - " & "in PerfectAntivirus2009.frmMain.cmdFSKillVirus_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub cmdFSReport_Click()

        '<EhHeader>
        On Error GoTo cmdFSReport_Click_Err

        '</EhHeader>

100     frmVirusReport.ShowReport xTotalFile, modLietKeValue.xTotalStartUp, modScanVirus.xTotalProcess, LVVirus1.ListItems.Count, IIf(HoanThanhFull, "D9a4 hoa2n tha2nh", "Chu7a hoa2n tha2nh"), Me.lblScanTime(0).Caption & ":" & Me.lblScanTime(1).Caption & ":" & Me.lblScanTime(2).Caption

        '<EhFooter>
        Exit Sub

cmdFSReport_Click_Err:
        WriteErr Time & "-" & Date & " - " & Err.Description & " - " & "in PerfectAntivirus2009.frmMain.cmdFSReport_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub cmdFSStart_Click()

        '<EhHeader>
        On Error GoTo cmdFSStart_Click_Err

        '</EhHeader>

100     DelAllLV LVVirus1

102     xTimeSC = 0

104     HoanThanhFull = True
106     xFSStopScan = False
108     cmdFSCachLy.Enabled = False
110     cmdFSKillVirus.Enabled = False
112     cmdFSReport.Enabled = False
114     cmdFSStop.Enabled = True
116     cmdFSStart.Enabled = False
118     cmdFSCachLy.Enabled = False
120     fmMain.Enabled = False
        '0000000000
122     frmMenu.quetvadiet.Visible = False
124     frmMenu.ngang0.Visible = False
126     frmMenu.tudongbaove.Visible = False
128     frmMenu.tienichhethong.Visible = False
130     frmMenu.caidatcauhinh.Visible = False
132     frmMenu.tacgia.Visible = False
        '00000000000
134     cmdSettingFullScan.Enabled = False
136     TimeFull = Time & " - " & Day(Date) & "/" & Month(Date) & "/" & Year(Date)

138     Tray1.ToolTipText = "D9ang que1t toa2n bo65 he65 tho61ng..."

140     xTotalFile = 0

142     lblStatus2.Caption = "D9ang Que1t..."

144     tmrStartFullScan.Enabled = True
146     lblStatus.Caption = "Chua63n bi5 que1t..."
148     lblStatus2.Caption = "Vui lo2ng kho6ng cha5y the6m u71ng du5ng na2o trong tho72i gian que1t..."

150     xTime = 0
152     tmrScanTime.Enabled = True

        '<EhFooter>
        Exit Sub

cmdFSStart_Click_Err:
        WriteErr Time & "-" & Date & " - " & Err.Description & " - " & "in PerfectAntivirus2009.frmMain.cmdFSStart_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub cmdFSStop_Click()

        '<EhHeader>
        On Error GoTo cmdFSStop_Click_Err

        '</EhHeader>
100     If UniMsgBox("Chu7o7ng tri2nh d9ang que1t Virus, ba5n co1 cha81c cha81n muo61n du72ng la5i kho6ng?", vbYesNo + vbQuestion, "Tho6ng Ba1o") = vbYes Then
102         cmdFSStop.Enabled = False
104         xFSStopScan = True
106         HoanThanhFull = False
        End If

        '<EhFooter>
        Exit Sub

cmdFSStop_Click_Err:
        WriteErr Time & "-" & Date & " - " & Err.Description & " - " & "in PerfectAntivirus2009.frmMain.cmdFSStop_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub cmdFullScan_Click()

        '<EhHeader>
        On Error GoTo cmdFullScan_Click_Err

        '</EhHeader>

100     HideAllFM
102     fm(0).Visible = True

        '<EhFooter>
        Exit Sub

cmdFullScan_Click_Err:
        WriteErr Time & "-" & Date & " - " & Err.Description & " - " & "in PerfectAntivirus2009.frmMain.cmdFullScan_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub cmdGotoGopY_Click()

        '<EhHeader>
        On Error GoTo cmdGotoGopY_Click_Err

        '</EhHeader>

100     ShellExecute Me.hWnd, vbNullString, "http://qts.come.vn", vbNullString, "", 1

        '<EhFooter>
        Exit Sub

cmdGotoGopY_Click_Err:
        WriteErr Time & "-" & Date & " - " & Err.Description & " - " & "in PerfectAntivirus2009.frmMain.cmdGotoGopY_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub cmdGoToHelp_Click()

        '<EhHeader>
        On Error GoTo cmdGoToHelp_Click_Err

        '</EhHeader>

100     ShellExecute Me.hWnd, vbNullString, AppPath & "Help\HuongDan.html", vbNullString, "", 1

        '<EhFooter>
        Exit Sub

cmdGoToHelp_Click_Err:
        WriteErr Time & "-" & Date & " - " & Err.Description & " - " & "in PerfectAntivirus2009.frmMain.cmdGoToHelp_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub cmdKiemTraHeThong_Click()

        '<EhHeader>
        On Error GoTo cmdKiemTraHeThong_Click_Err

        '</EhHeader>
100     If FileExists(AppPath & "PAVSysReport.exe") = True Then
102         If FileLen(AppPath & "PAVSysReport.exe") = 35328 Then
104             Shell AppPath & "PAVSysReport.exe syscheck", vbNormalFocus
            End If

        Else
106         UniMsgBox "Kho6ng ti2m tha61y File PAVSysReport.exe.", vbOKOnly + vbCritical, "Error!"
        End If

        '<EhFooter>
        Exit Sub

cmdKiemTraHeThong_Click_Err:
        WriteErr Time & "-" & Date & " - " & Err.Description & " - " & "in PerfectAntivirus2009.frmMain.cmdKiemTraHeThong_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub cmdQuanly_Click()

        '<EhHeader>
        On Error GoTo cmdQuanly_Click_Err

        '</EhHeader>
100     If FileExists(AppPath & "PAV2009Manager.exe") = True Then
102         If FileLen(AppPath & "PAV2009Manager.exe") = 53760 Then
104             Shell AppPath & "PAV2009Manager.exe quangtrung", vbNormalFocus
            End If

        Else
106         UniMsgBox "Kho6ng ti2m tha61y File PAV2009Manager.exe.", vbOKOnly + vbCritical, "Error!"
        End If

        '<EhFooter>
        Exit Sub

cmdQuanly_Click_Err:
        WriteErr Time & "-" & Date & " - " & Err.Description & " - " & "in PerfectAntivirus2009.frmMain.cmdQuanly_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub cmdQuanLyFile_Click()

        '<EhHeader>
        On Error GoTo cmdQuanLyFile_Click_Err

        '</EhHeader>
100     If FileExists(AppPath & "PAVExplorer.exe") = True Then
102         If FileLen(AppPath & "PAVExplorer.exe") = 82432 Then
104             Shell AppPath & "PAVExplorer.exe explorer", vbNormalFocus
            End If

        Else
106         UniMsgBox "Kho6ng ti2m tha61y File PAVExplorer.exe.", vbOKOnly + vbCritical, "Error!"
        End If

        '<EhFooter>
        Exit Sub

cmdQuanLyFile_Click_Err:
        WriteErr Time & "-" & Date & " - " & Err.Description & " - " & "in PerfectAntivirus2009.frmMain.cmdQuanLyFile_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub cmdRebotDel_Click()

        '<EhHeader>
        On Error GoTo cmdRebotDel_Click_Err

        '</EhHeader>
        Dim IValues

        Dim strKeyPath

        Dim MultValueName

        Dim strComputer

100     If List1.ListCount = 0 Then Exit Sub

102     strKeyPath = "SYSTEM\CurrentControlSet\Control\Session Manager"
104     MultValueName = "PendingFileRenameOperations"
106     strComputer = "."
108     IValues = "1"

        On Error Resume Next

110     IValues = Array("\??\" & List1.List(0), vbNullString, "\??\" & List1.List(1), vbNullString, "\??\" & List1.List(2), vbNullString, "\??\" & List1.List(3), vbNullString, "\??\" & List1.List(4), vbNullString, "\??\" & List1.List(5), vbNullString, "\??\" & List1.List(6), vbNullString, "\??\" & List1.List(7), vbNullString, "\??\" & List1.List(8), vbNullString, "\??\" & List1.List(9), vbNullString)

        Dim oReg

112     Set oReg = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")
114     oReg.CreateKey HKEY_LOCAL_MACHINE, strKeyPath
116     oReg.SetMultiStringValue HKEY_LOCAL_MACHINE, strKeyPath, MultValueName, IValues

118     Set IValues = Nothing
120     Set strKeyPath = Nothing
122     Set MultValueName = Nothing
124     Set strComputer = Nothing
126     Set oReg = Nothing
128     UniMsgBox "D9a4 the6m va2o danh sa1ch xo1a khi kho73i d9o65ng." & vbCrLf & " Ba5n pha3i kho73i d9o65ng la5i ma1y thi2 mo71i co1 ta1c du5ng!", vbOKOnly + vbInformation, "OK"

        '<EhFooter>
        Exit Sub

cmdRebotDel_Click_Err:
        WriteErr Time & "-" & Date & " - " & Err.Description & " - " & "in PerfectAntivirus2009.frmMain.cmdRebotDel_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub cmdSettingFullScan_Click()

        '<EhHeader>
        On Error GoTo cmdSettingFullScan_Click_Err

        '</EhHeader>

100     HideAllFM
102     fm(3).Visible = True

        '<EhFooter>
        Exit Sub

cmdSettingFullScan_Click_Err:
        WriteErr Time & "-" & Date & " - " & Err.Description & " - " & "in PerfectAntivirus2009.frmMain.cmdSettingFullScan_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Public Sub HideAllFM()

        '<EhHeader>
        On Error GoTo HideAllFM_Err

        '</EhHeader>
        Dim I

100     For I = 0 To fm.Count - 1
102         fm(I).Visible = False
104         fm(I).Left = 3000
106         fm(I).Top = 1320
108     Next I
    
110     mf.Visible = False

        '<EhFooter>
        Exit Sub

HideAllFM_Err:
        WriteErr Time & "-" & Date & " - " & Err.Description & " - " & "in PerfectAntivirus2009.frmMain.HideAllFM " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub cmdStartUp_Click()

        '<EhHeader>
        On Error GoTo cmdStartUp_Click_Err

        '</EhHeader>
100     If FileExists(AppPath & "PAVStartUp.exe") = True Then
102         If FileLen(AppPath & "PAVStartUp.exe") = 138240 Then
104             Shell AppPath & "PAVStartUp.exe startup", vbNormalFocus
            End If

        Else
106         UniMsgBox "Kho6ng ti2m tha61y File PAVStartUp.exe.", vbOKOnly + vbCritical, "Error!"
        End If

        '<EhFooter>
        Exit Sub

cmdStartUp_Click_Err:
        WriteErr Time & "-" & Date & " - " & Err.Description & " - " & "in PerfectAntivirus2009.frmMain.cmdStartUp_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub cmdTangToc_Click()

        '<EhHeader>
        On Error GoTo cmdTangToc_Click_Err

        '</EhHeader>
100     If chkTangToc(0).Value = True Then
102         SaveString HKEY_CURRENT_USER, "Control Panel\Desktop", "MenuShowDelay", "400"
        End If

104     If chkTangToc(7).Value = True Then
106         SaveString HKEY_CURRENT_USER, "Control Panel\Desktop", "AutoEndTasks", "1"
108         SaveString HKEY_CURRENT_USER, "Control Panel\Desktop", "WaitToKillAppTimeout", "3500"
110         SaveString HKEY_CURRENT_USER, "Control Panel\Desktop", "HungAppTimeout", "5000"
112         SaveString HKEY_CURRENT_USER, "Control Panel\Desktop", "WaitToKillServiceTimeout", "500"
114         SaveString HKEY_LOCAL_MACHINE, "SYSTEM\CurrentControlSet\SetControl", "WaitToKillServiceTimeout", "500"
116         SaveString HKEY_LOCAL_MACHINE, "SYSTEM\ControlSet001\Control", "WaitToKillServiceTimeout", "500"
        End If

118     If chkTangToc(8).Value = True Then
120         SaveString HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Dfrg\BootOptimizeFunction", "Enable", "N"
        End If

122     UniMsgBox " D9a4 thu75c hie65n xong ca1c thao ta1c ta8ng to61c ma1y ti1nh." & vbCrLf & " Ha4y kho73i d9o65ng la5i ma1y d9e63 ca1c thao ta1c ta8ng to61c co1 hie65u lu75c.", vbOKOnly + vbInformation, "Xong!"

        '<EhFooter>
        Exit Sub

cmdTangToc_Click_Err:
        WriteErr Time & "-" & Date & " - " & Err.Description & " - " & "in PerfectAntivirus2009.frmMain.cmdTangToc_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub cmdTTCheckAll_Click()

        '<EhHeader>
        On Error GoTo cmdTTCheckAll_Click_Err

        '</EhHeader>

100     chkTangToc(0).Value = True
102     chkTangToc(7).Value = True
104     chkTangToc(8).Value = True
    
        '<EhFooter>
        Exit Sub

cmdTTCheckAll_Click_Err:
        WriteErr Time & "-" & Date & " - " & Err.Description & " - " & "in PerfectAntivirus2009.frmMain.cmdTTCheckAll_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub cmdUpdateOff_Click()

        '<EhHeader>
        On Error GoTo cmdUpdateOff_Click_Err

        '</EhHeader>

100     Shell AppPath & "PAVUpdate.exe off", vbNormalFocus

        '<EhFooter>
        Exit Sub

cmdUpdateOff_Click_Err:
        WriteErr Time & "-" & Date & " - " & Err.Description & " - " & "in PerfectAntivirus2009.frmMain.cmdUpdateOff_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub cmdUpdateOnline_Click()

        '<EhHeader>
        On Error GoTo cmdUpdateOnline_Click_Err

        '</EhHeader>

100     Shell AppPath & "PAVUpdate.exe on", vbNormalFocus

        '<EhFooter>
        Exit Sub

cmdUpdateOnline_Click_Err:
        WriteErr Time & "-" & Date & " - " & Err.Description & " - " & "in PerfectAntivirus2009.frmMain.cmdUpdateOnline_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub cmdViewErr_Click()
On Error Resume Next
If FileExists(AppPath & "err.txt") = False Then
    WriteFileUni AppPath & "Err.txt", " "
End If
If UniMsgBox(" {Ne61u be6n du7o71i kho6ng co1 gi2 co1 nghi4a la2 va64n chu7a co1 lo64i na2o xa3y ra}" & vbCrLf & ReadFileUni(AppPath & "Err.txt") & vbCrLf & vbCrLf & " No65i dung ca1c lo64i d9u7o75c lu7u o73 File Err.txt trong thu7 mu5c ca2i d9a85t cu3a chu7o7ng tri2nh, ne61u pha1t hie65n lo64i ha4y gu73i file d9o1 cho ta1c gia3." & vbCrLf & " Nha61n va2o nu1t 'Co1' d9e63 mo73 no65i dung File Err.txt", vbYesNo + vbInformation, " Nhu74ng lo64i xa3y ra trong qua1 tri2nh su73 du5ng cu3a ba5n.") = vbYes Then
    Shell "notepad " & AppPath & "Err.txt", vbNormalFocus
End If
End Sub

Private Sub csBack_Click()

        '<EhHeader>
        On Error GoTo csBack_Click_Err

        '</EhHeader>

100     ff.Visible = True
102     csBack.Enabled = False
104     csCachLy.Enabled = False
106     csKill.Enabled = False

        '<EhFooter>
        Exit Sub

csBack_Click_Err:
        WriteErr Time & "-" & Date & " - " & Err.Description & " - " & "in PerfectAntivirus2009.frmMain.csBack_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub CSbat_Click()

        '<EhHeader>
        On Error GoTo CSbat_Click_Err

        '</EhHeader>

100     VSbat.Value = CSbat.Value

        '<EhFooter>
        Exit Sub

CSbat_Click_Err:
        WriteErr Time & "-" & Date & " - " & Err.Description & " - " & "in PerfectAntivirus2009.frmMain.CSbat_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub csCachLy_Click()

        '<EhHeader>
        On Error GoTo csCachLy_Click_Err

        '</EhHeader>
100     If UniMsgBox("Ba5n co1 muo61n ca1ch ly ca1c Virus d9a4 cho5n kho6ng?", vbYesNo + vbQuestion, "Tho6ng Ba1o") = vbYes Then

            On Error GoTo HeHeChOqUa1

102         MkDir AppPath & "VungCachLy\"
HeHeChOqUa1:

            Dim Y, j

            Dim X As String

104         For Y = 1 To LVVirus2.ListItems.Count

106             If LVVirus2.ListItems(Y).Checked = True And FileExists(LVVirus2.ListItems(Y).SubItems(1).Caption) = True Then
    
108                 If LVVirus2.ListItems(Y).SubItems(3).Caption <> "0" Then
                        'Kill process
110                     KillProcessById (LVVirus2.ListItems(Y).SubItems(3).Caption)
112                     X = X & " D9a4 ta81t tie61n tri2nh: " & LVVirus2.ListItems(Y).SubItems(3).Caption & vbCrLf
                    End If

                    On Error GoTo KhOnGtHeCaChLy

114                 Set fss = Nothing
116                 SetAttr LVVirus2.ListItems(Y).SubItems(1).Caption, vbNormal
118                 Name LVVirus2.ListItems(Y).SubItems(1).Caption As LVVirus2.ListItems(Y).SubItems(1).Caption & ".DaCachLy"
120                 FileCopy LVVirus2.ListItems(Y).SubItems(1).Caption & ".DaCachLy", AppPath & "VungCachLy\" & GetFileName(LVVirus2.ListItems(Y).SubItems(1).Caption & ".DaCachLy")
122                 modScanVirus.DeleteFile LVVirus2.ListItems(Y).SubItems(1).Caption & ".DaCachLy"
124                 X = X & " D9a4 Ca1ch Ly: " & LVVirus2.ListItems(Y).SubItems(1).Caption & vbCrLf

                End If

126         Next Y

128         If Not X = "" Then
130             DelAllChecked LVVirus2
132             UniMsgBox X, vbOKOnly + vbInformation, "D9a4 Ca1ch Ly Virus", Me.hWnd
            Else
134             UniMsgBox "Kho6ng co1 Virus na2o d9e63 ca1ch ly.", vbOKOnly + vbCritical, "Tho6ng Ba1o", Me.hWnd
            End If
        End If

        Exit Sub

KhOnGtHeCaChLy:
136     X = X & " Kho6ng The63 Ca1ch Ly: " & LVVirus2.ListItems(Y).SubItems(1).Caption & vbCrLf

138     Resume Next

        '<EhFooter>
        Exit Sub

csCachLy_Click_Err:
        WriteErr Time & "-" & Date & " - " & Err.Description & " - " & "in PerfectAntivirus2009.frmMain.csCachLy_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub CScom_Click()

        '<EhHeader>
        On Error GoTo CScom_Click_Err

        '</EhHeader>

100     VScom.Value = CScom.Value

        '<EhFooter>
        Exit Sub

CScom_Click_Err:
        WriteErr Time & "-" & Date & " - " & Err.Description & " - " & "in PerfectAntivirus2009.frmMain.CScom_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub CSFolderView1_ChangeAfter(ByVal OldPath As String)

        '<EhHeader>
        On Error GoTo CSFolderView1_ChangeAfter_Err

        '</EhHeader>

100     cslblPath.Caption = CSFolderView1.Path

        '<EhFooter>
        Exit Sub

CSFolderView1_ChangeAfter_Err:
        WriteErr Time & "-" & Date & " - " & Err.Description & " - " & "in PerfectAntivirus2009.frmMain.CSFolderView1_ChangeAfter " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub csKill_Click()

        '<EhHeader>
        On Error GoTo csKill_Click_Err

        '</EhHeader>
100     If UniMsgBox("Ba5n co1 cha81c cha81n muo61n xo1a ca1c Virus d9a4 d9a1nh da61u kho6ng?", vbYesNo + vbQuestion, "Die65t Virus") = vbYes Then

102         DoEvents

            Dim Y

            Dim X As String

104         For Y = 1 To LVVirus2.ListItems.Count

106             If LVVirus2.ListItems(Y).Checked = True And FileExists(LVVirus2.ListItems(Y).SubItems(1).Caption) = True Then

108                 DoEvents

110                 If LVVirus2.ListItems(Y).SubItems(3).Caption <> "0" Then
                        'Kill process
112                     KillProcessById (LVVirus2.ListItems(Y).SubItems(3).Caption)
114                     X = X & " D9a4 ta81t tie61n tri2nh: " & LVVirus2.ListItems(Y).SubItems(3).Caption & vbCrLf
                    End If

116                 DoEvents

118                 If LVVirus2.ListItems(Y).SubItems(4).Caption <> "---" Then

                        'Delete Registry Key
                        Dim a

                        Dim u

                        Dim b

                        Dim c

120                     DoEvents
122                     a = Left(LVVirus2.ListItems(Y).SubItems(4).Caption, Len(LVVirus2.ListItems(Y).SubItems(4).Caption) - InStrRev(StrReverse(LVVirus2.ListItems(Y).SubItems(4).Caption), "-"))
124                     b = Right(LVVirus2.ListItems(Y).SubItems(4).Caption, Len(LVVirus2.ListItems(Y).SubItems(4).Caption) - InStrRev(LVVirus2.ListItems(Y).SubItems(4).Caption, ":"))
126                     c = Mid(LVVirus2.ListItems(Y).SubItems(4).Caption, Len(LVVirus2.ListItems(Y).SubItems(4).Caption) - InStrRev(StrReverse(LVVirus2.ListItems(Y).SubItems(4).Caption), "-") + 2, (InStrRev(LVVirus2.ListItems(Y).SubItems(4).Caption, ":")) - (Len(LVVirus2.ListItems(Y).SubItems(4).Caption) - InStrRev(StrReverse(LVVirus2.ListItems(Y).SubItems(4).Caption), "-")) - 2)

128                     If UCase(a) = "HKEY_CURRENT_USER" Then u = &H80000001
130                     If UCase(a) = "HKEY_LOCAL_MACHINE" Then u = &H80000002

132                     DoEvents
134                     DeleteValue u, c, b
136                     X = X & " D9a4 xo1a Key: " & a & "\" & c & ":" & b & vbCrLf
                    End If
        
                    'HKEY_CURRENT_USER = &H80000001
                    'HKEY_LOCAL_MACHINE = &H80000002
138                 DoEvents

                    On Error GoTo KhOnGtHeXoAbO

140                 SetAttr LVVirus2.ListItems(Y).SubItems(1).Caption, vbNormal
142                 modScanVirus.DeleteFile LVVirus2.ListItems(Y).SubItems(1).Caption
144                 X = X & " D9a4 Xo1a Bo3: " & LVVirus2.ListItems(Y).SubItems(1).Caption & vbCrLf

                End If

146         Next Y

148         If Not X = "" Then
150             DelAllChecked LVVirus2
152             UniMsgBox X, vbOKOnly + vbInformation, "Nhu74ng vie65c d9a4 la2m.", Me.hWnd

            Else
154             UniMsgBox "Kho6ng co1 Virus na2o d9e63 die65t.", vbOKOnly + vbCritical, "Tho6ng Ba1o", Me.hWnd
            End If

        End If ' Unim

        Exit Sub

KhOnGtHeXoAbO:
156     X = X & " Kho6ng The63 Xo1a Bo3: " & LVVirus2.ListItems(Y).SubItems(1).Caption & vbCrLf

158     Resume Next

        '<EhFooter>
        Exit Sub

csKill_Click_Err:
        WriteErr Time & "-" & Date & " - " & Err.Description & " - " & "in PerfectAntivirus2009.frmMain.csKill_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub CSprocess_Click()

        '<EhHeader>
        On Error GoTo CSprocess_Click_Err

        '</EhHeader>

100     Me.VSScanProcess.Value = CSprocess.Value

        '<EhFooter>
        Exit Sub

CSprocess_Click_Err:
        WriteErr Time & "-" & Date & " - " & Err.Description & " - " & "in PerfectAntivirus2009.frmMain.CSprocess_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub csStart_Click()

        '<EhHeader>
        On Error GoTo csStart_Click_Err

        '</EhHeader>
100     If cslblPath.Caption = "" Then
102         UniMsgBox "Ba5n chu7a cho5n no7i d9e63 que1t!", vbOKOnly + vbCritical, "Tho6ng ba1o"

            Exit Sub

        End If

104     TimeCus = Time & " - " & Day(Date) & "/" & Month(Date) & "/" & Year(Date)
106     ff.Visible = False
108     xFSStopScan2 = False
110     Me.csStart.Enabled = False
112     Me.csStop.Enabled = True
114     Me.csCachLy.Enabled = False
116     Me.csKill.Enabled = False
118     csBack.Enabled = False
120     fmMain.Enabled = False
        '0000000000
122     frmMenu.quetvadiet.Visible = False
124     frmMenu.ngang0.Visible = False
126     frmMenu.tudongbaove.Visible = False
128     frmMenu.tienichhethong.Visible = False
130     frmMenu.caidatcauhinh.Visible = False
132     frmMenu.tacgia.Visible = False
        '00000000000
134     cslblStatus.AutoUnicode = False
136     HoanThanhCus = True
138     DelAllLV LVVirus2

140     If Me.VSScanProcess.Value = True Then
142         cslblStatus2.Caption = "D9ang que1t ca1c chu7o7ng tri2nh cha5y trong bo65 nho71..."
144         modScanVirus.xScanProcess2
        End If

146     If Me.VSScanStartUp.Value = True Then
148         cslblStatus2.Caption = "D9ang que1t ca1c chu7o7ng tri2nh kho73i d9o65ng..."

150         cslblStatus.Caption = "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Run"
152         GetKeyValue2 "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Run"

154         cslblStatus.Caption = "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\RunOnce"
156         GetKeyValue2 "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\RunOnce"

158         cslblStatus.Caption = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Run"
160         GetKeyValue2 "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Run"

162         cslblStatus.Caption = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\RunOnce"
164         GetKeyValue2 "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\RunOnce"

166         cslblStatus.Caption = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\RunOnceEx"
168         GetKeyValue2 "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\RunOnceEx"
        End If

170     xCustomScan = True
172     cslblStatus2.Caption = "D9ang que1t ca1c File..."

174     SearchFile cslblPath.Caption, "*.exe"

176     If VSbat.Value = True Then SearchFile cslblPath.Caption, "*.bat"
178     If VScom.Value = True Then SearchFile cslblPath.Caption, "*.com"
180     xCustomScan = False

182     cslblStatus.AutoUnicode = True

184     cslblStatus2.Caption = "Sa84n Sa2ng"

186     cslblStatus.Caption = "D9a4 Que1t Xong! Ti2m Tha61y: " & LVVirus2.ListItems.Count & " Ta65p Tin Virus"

188     Me.csStart.Enabled = True
190     Me.csStop.Enabled = False
192     Me.csCachLy.Enabled = True
194     Me.csKill.Enabled = True
196     csBack.Enabled = True
198     fmMain.Enabled = True
        '0000000000
200     frmMenu.quetvadiet.Visible = True
202     frmMenu.ngang0.Visible = True
204     frmMenu.tudongbaove.Visible = True
206     frmMenu.tienichhethong.Visible = True
208     frmMenu.caidatcauhinh.Visible = True
210     frmMenu.tacgia.Visible = True
        '00000000000
212     PLaySound AppPath & "Sound\ScanDone.wav"

214     With LVVirusEvents

            Dim I

216         I = .ListItems.Count + 1
218         .ListItems.Add I, , TimeCus
220         .ListItems(I).SubItems(1).Caption = ToUnicode("Que1t tu2y cho5n (" & Me.cslblPath.Caption & ")")
222         .ListItems(I).SubItems(2).Caption = "-"
224         .ListItems(I).SubItems(3).Caption = LVVirus2.ListItems.Count
226         .ListItems(I).SubItems(4).Caption = IIf(HoanThanhCus, ToUnicode("Hoa2n tha2nh"), ToUnicode("Chu7a hoa2n tha2nh"))
228         .AutoUnicode = True
        End With

        '<EhFooter>
        Exit Sub

csStart_Click_Err:
        WriteErr Time & "-" & Date & " - " & Err.Description & " - " & "in PerfectAntivirus2009.frmMain.csStart_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub CSStartUp_Click()

        '<EhHeader>
        On Error GoTo CSStartUp_Click_Err

        '</EhHeader>

100     Me.VSScanStartUp.Value = CSStartUp.Value

        '<EhFooter>
        Exit Sub

CSStartUp_Click_Err:
        WriteErr Time & "-" & Date & " - " & Err.Description & " - " & "in PerfectAntivirus2009.frmMain.CSStartUp_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub csStop_Click()

        '<EhHeader>
        On Error GoTo csStop_Click_Err

        '</EhHeader>
100     If UniMsgBox("Chu7o7ng tri2nh d9ang que1t Virus, ba5n co1 cha81c cha81n muo61n du72ng la5i kho6ng?", vbYesNo + vbQuestion, "Tho6ng Ba1o") = vbYes Then
102         csStop.Enabled = False
104         xFSStopScan2 = True
106         HoanThanhCus = False

        End If

        '<EhFooter>
        Exit Sub

csStop_Click_Err:
        WriteErr Time & "-" & Date & " - " & Err.Description & " - " & "in PerfectAntivirus2009.frmMain.csStop_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Form_Load()

        '<EhHeader>
        On Error GoTo Form_Load_Err

        '</EhHeader>

        '---> Check File Is Running [First of first]

        '<--- Check File Is Running [First of first]

        '---> Check Exists Database [First]
100     If basMain.CheckFiles = False Then
102         UniMsgBox " Chu7o7ng tri2nh kho6ng the63 cha5y ba6y gio72 vi2 thie61u 1 so61 File." & vbCrLf & " Ha4y ca2i d9a85t la5i chu7o7ng tri2nh d9e63 co1 d9a62y d9u3 ca1c File." & vbCrLf & " Chu7o7ng tri2nh se4 thoa1t ngay ba6y gio72.", vbOKOnly + vbCritical

104         End

        End If

        On Error GoTo HeHeChOqUa

106     MkDir AppPath & "VungCachLy\"
HeHeChOqUa:
        '<--- Check Exists Database [First]

        '---> Load Setting [Second]
108     Me.VSbat.Value = ReadIniFile(AppPath & "Setting.ini", "ScanVirus", Me.VSbat.Name, True)
110     Me.VScmd.Value = ReadIniFile(AppPath & "Setting.ini", "ScanVirus", Me.VScmd.Name, True)
112     Me.VScom.Value = ReadIniFile(AppPath & "Setting.ini", "ScanVirus", Me.VScom.Name, True)
114     Me.VSdll.Value = ReadIniFile(AppPath & "Setting.ini", "ScanVirus", Me.VSdll.Name, True)
116     Me.VSscr.Value = ReadIniFile(AppPath & "Setting.ini", "ScanVirus", Me.VSscr.Name, True)

118     Me.VSScanProcess.Value = ReadIniFile(AppPath & "Setting.ini", "ScanVirus", Me.VSScanProcess.Name, True)
120     Me.VSScanStartUp.Value = ReadIniFile(AppPath & "Setting.ini", "ScanVirus", Me.VSScanStartUp.Name, True)

122     Me.CSbat.Value = ReadIniFile(AppPath & "Setting.ini", "ScanVirus", Me.VSbat.Name, True)
124     Me.CScom.Value = ReadIniFile(AppPath & "Setting.ini", "ScanVirus", Me.VScom.Name, True)

126     Me.CSprocess.Value = ReadIniFile(AppPath & "Setting.ini", "ScanVirus", Me.VSScanProcess.Name, True)
128     Me.CSStartUp.Value = ReadIniFile(AppPath & "Setting.ini", "ScanVirus", Me.VSScanStartUp.Name, True)

130     Me.VSDontScanSize.Value = ReadIniFile(AppPath & "Setting.ini", "ScanVirus", Me.VSDontScanSize.Name, True)
132     Me.VSLimitSize.Value = ReadIniFile(AppPath & "Setting.ini", "ScanVirus", Me.VSLimitSize.Name, 4)
134     Me.atsAlwaysScanFolder.Value = ReadIniFile(AppPath & "Setting.ini", "AutoScan", "ScanFolder", True)
136     Me.atsScanEXE.Value = ReadIniFile(AppPath & "Setting.ini", "AutoScan", "ScanEXE", False)
138     Me.atsAutoScanUSB.Value = ReadIniFile(AppPath & "Setting.ini", "AutoScan", "ScanUSB", True)
140     Me.atsScanKeylogger.Value = ReadIniFile(AppPath & "Setting.ini", "AutoScan", "ScanKeylogger", True)

142     Me.chkAutoStart.Value = ReadIniFile(AppPath & "Setting.ini", "Setting", "AutoStart", True)
144     Me.chkShowFlash.Value = ReadIniFile(AppPath & "Setting.ini", "Setting", "FlashScreen", True)
146     Me.chkAutoUpdate.Value = ReadIniFile(AppPath & "Setting.ini", "Setting", "AutoUpdate", True)
148     Me.chkAutoAddAutorun.Value = ReadIniFile(AppPath & "Setting.ini", "AutoProtect", "AutoAddVirus", True)

150     If ReadIniFile(AppPath & "Setting.ini", "Setting", "MiniScan", True) = True Then
152         Me.chkUseFastScan.Value = True
154         SaveString HKEY_CLASSES_ROOT, "Folder\shell\[PAV 2009] Quét Virus\command", "", ChrW(34) & AppPath & "PAVMiniScan.exe" & ChrW(34) & " " & ChrW(34) & "%1" & ChrW(34)
        Else
156         Me.chkUseFastScan.Value = False
158         DeleteKey HKEY_CLASSES_ROOT, "Folder\shell\[PAV 2009] Quét Virus\command"
160         DeleteKey HKEY_CLASSES_ROOT, "Folder\shell\[PAV 2009] Quét Virus"
        End If

162     If ReadIniFile(AppPath & "Setting.ini", "AutoProtect", "Registry", True) = False Then
164         Me.atptmrREG.Enabled = False
166         Me.atplbREG.Caption = "D9ang ta81t"
168         Me.atpcmdREG.Caption = "Mo73 chu71c na8ng na2y"
        Else
170         Me.atptmrREG.Enabled = True
172         Me.atplbREG.Caption = "D9ang mo73"
174         Me.atpcmdREG.Caption = "Ta81t chu71c na8ng na2y"
        End If

176     If ReadIniFile(AppPath & "Setting.ini", "AutoProtect", "Autorun", True) = False Then
178         Me.atptmrAutorun.Enabled = False
180         Me.atplblStaAutorun.Caption = "D9ang ta81t"
182         Me.atpcmdAutorun.Caption = "Mo73 chu71c na8ng na2y"
        Else
184         Me.atptmrAutorun.Enabled = True
186         Me.atplblStaAutorun.Caption = "D9ang mo73"
188         Me.atpcmdAutorun.Caption = "Ta81t chu71c na8ng na2y"
        End If

190     If Me.chkAutoUpdate.Value = True Then
192         Shell AppPath & "PAVUpdate.exe"
        End If

194     If Me.chkAutoStart.Value = True Then
196         SaveString HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Run", "PAV2009", AppPath & "PAV2009.exe /task"
        End If

        '<--- Load Setting [Second]

        '---> Set Tray TooltipText
198     Tray1.ToolTipText = "Perfect Antivirus 2009 - Ma1y ti1nh cu3a ba5n d9ang o73 ti2nh tra5ng to61t nha61t!"
        '<--- Set Tray TooltipText

        '---> Set for "Chuc Nang"

200     With TreeChucNang
202         .Initialize
204         .InitializeImageList 24, 24
206         .HasButtons = True
208         .SingleExpand = False
        
            Dim j

210         For j = 0 To Me.icoFull.Count - 1
212             .AddIcon Me.icoFull(j).Picture
214         Next j
        
216         .AddNode , , "a", "Que1t & Die65t Virus", 0, 0
218         .AddNode "a", , "0", "Que1t toa2n bo65", 9, 9
220         .AddNode "a", , "1", "Que1t tu2y cho5n", 3, 3
222         .AddNode "a", , "3", "Ca61u hi2nh que1t", 4, 4
224         .AddNode "a", , "4", "Nha65t ky1 que1t", 5, 5

226         .AddNode , , "b", "Tu75 d9o65ng ba3o ve65", 1, 1
228         .AddNode "b", , "5", "Ba3o ve65 Registry", 6, 6
230         .AddNode "b", , "6", "Ba3o ve65 Autorun", 7, 7
232         .AddNode "b", , "2", "Tu75 d9o65ng que1t", 8, 8

234         .AddNode , , "c", "Tie65n i1ch he65 tho61ng", 2, 2
236         .AddNode "c", , "7", "Qua3n Ly1 Tie61n Tri2nh", 10, 10
238         .AddNode "c", , "8", "Qua3n Ly1 Ta65p Tin", 11, 11
240         .AddNode "c", , "9", "Qua3n Ly1 Kho73i D9o65ng", 12, 12
242         .AddNode "c", , "10", "Kie63m Tra He65 Tho61ng", 13, 13
244         .AddNode "c", , "13", "Tie65n I1ch", 17, 17
246         .AddNode , , "d", "Perfect AV 2009", 14, 14
248         .AddNode "d", , "11", "Ca61u Hi2nh Chung", 15, 15
250         .AddNode "d", , "12", "Gio71i Thie65u", 16, 16
            
252         .Expand .GetKeyNode("a"), True
254         .Expand .GetKeyNode("b"), True
256         .Expand .GetKeyNode("c"), True
258         .Expand .GetKeyNode("d"), True
        
        End With

        '---> Set for "Chuc Nang"

        '---> Properties Setting
260     xFSStopScan = False
        '<--- Properties Setting

        '---> Connect Database
262     modScanVirus.ConnectDB
        '<--- Connect Database

        '---> FullScan
264     xCustomScan = False
266     HoanThanhFull = False

268     LVVirus1.View = eViewDetails
270     LVVirus1.GridLines = True
272     LVVirus1.HeaderButtons = False
274     LVVirus1.CheckBoxes = True

276     LVVirus1.Columns.Add , , "Virus Name", , 2000
278     LVVirus1.Columns.Add , , "Path", , 3500
280     LVVirus1.Columns.Add , , "Size", , 1000
282     LVVirus1.Columns.Add , , "Process ID", , 1000
284     LVVirus1.Columns.Add , , "Start Up Key", , 5000
286     LVVirus1.Refresh
        '<--- FullScan

        '---> Custom Scan
288     HoanThanhCus = False
290     LVVirus2.View = eViewDetails
292     LVVirus2.GridLines = True
294     LVVirus2.HeaderButtons = False
296     LVVirus2.CheckBoxes = True

298     LVVirus2.Columns.Add , , "Virus Name", , 2000
300     LVVirus2.Columns.Add , , "Path", , 3500
302     LVVirus2.Columns.Add , , "Size", , 1000
304     LVVirus2.Columns.Add , , "Process ID", , 1000
306     LVVirus2.Columns.Add , , "Start Up Key", , 5000
308     LVVirus2.Refresh
        '<--- Custom Scan

        '---> Nhat ky' virus
310     With LVVirusEvents
312         .View = eViewDetails
314         .GridLines = True
316         .HeaderButtons = False
318         .CheckBoxes = True
320         .AutoUnicode = True
322         .Columns.Add , , "Tho72i gian", , 2200
324         .Columns.Add , , "Kie63u que1t", , 2000
326         .Columns.Add , , "So61 File", , 1000
328         .Columns.Add , , "So61 Virus", , 1000
330         .Columns.Add , , "Ke61t qua3", , 2000

            Dim k

332         k = ReadIniFile(AppPath & "VirusScanLog.log", "Other", "Total", 0)

            Dim la
    
334         For la = 1 To k
336             .ListItems.Add la, , ReadIniFile(AppPath & "VirusScanLog.log", la, "TimeQuet", "")
338             .ListItems(la).SubItems(1).Caption = UTF82Unicode(ReadIniFile(AppPath & "VirusScanLog.log", la, "KieuQuet", ""))
340             .ListItems(la).SubItems(2).Caption = ReadIniFile(AppPath & "VirusScanLog.log", la, "SoFile", "")
342             .ListItems(la).SubItems(3).Caption = ReadIniFile(AppPath & "VirusScanLog.log", la, "SoVirus", "")
344             .ListItems(la).SubItems(4).Caption = UTF82Unicode(ReadIniFile(AppPath & "VirusScanLog.log", la, "KetQua", ""))
        
346         Next la

        End With

        '<--- Nhat ky' virus

        '---> LV REG
348     With atpLVREG
350         .View = eViewDetails
352         .GridLines = True
354         .HeaderButtons = False
356         .CheckBoxes = True
358         .AutoUnicode = True
360         .Columns.Add , , "Te6n chu71c na8ng", , 2300
362         .Columns.Add , , "Kho1a go61c", , 2000
364         .Columns.Add , , "D9u7o72ng da64n kho1a", , 5000
366         .Columns.Add , , "Te6n kho1a"
368         .Columns.Add , , "Gia1 tri5 ma85c d9i5nh"
    
            Dim F

370         For F = 1 To ReadIniFile(AppPath & "RegProtect.dat", "Other", "Total", 0)
372             .ListItems.Add F, , UTF82Unicode(ReadIniFile(AppPath & "RegProtect.dat", F, "TenChucNang", ""))
374             .ListItems(F).SubItems(1).Caption = ReadIniFile(AppPath & "RegProtect.dat", F, "KeyGoc", "")
376             .ListItems(F).SubItems(2).Caption = ReadIniFile(AppPath & "RegProtect.dat", F, "KeyPath", "")
378             .ListItems(F).SubItems(3).Caption = ReadIniFile(AppPath & "RegProtect.dat", F, "KeyName", "")
380             .ListItems(F).SubItems(4).Caption = ReadIniFile(AppPath & "RegProtect.dat", F, "KeyData", "")
382         Next F

        End With

        '<--- LV REG

        '---> Hide all fm
384     HideAllFM
        '<--- Hide all fm

        '---> Load auto scan

386     If atsAlwaysScanFolder.Value = True Then
388         Load zfrmAutoScanFolder
390         atsAlwaysScanFolder.Caption = "Tu75 d9o65ng que1t va2 ca3nh ba1o Virus trong Thu7 mu5c d9ang d9u7o75c mo73. [D9ang Ba65t]"
        Else
392         atsAlwaysScanFolder.Caption = "Tu75 d9o65ng que1t va2 ca3nh ba1o Virus trong Thu7 mu5c d9ang d9u7o75c mo73. [D9ang Ta81t]"
        End If

394     If atsScanEXE.Value = True Then
396         atsScanEXE.Caption = "Tu75 d9o65ng que1t va2 ca3nh ba1o Virus cho nhu74ng File sa81p d9u7o75c mo73. [D9ang Ba65t]"
398         SaveString HKEY_CLASSES_ROOT, "exefile\shell\open\command", "", ChrW(34) & AppPath & "Check_Virus.exe" & ChrW(34) & " " & ChrW(34) & "%1" & ChrW(34)
400         SaveString HKEY_CLASSES_ROOT, "batfile\shell\open\command", "", ChrW(34) & AppPath & "Check_Virus.exe" & ChrW(34) & " " & ChrW(34) & "%1" & ChrW(34)
402         SaveString HKEY_CLASSES_ROOT, "cmdfile\shell\open\command", "", ChrW(34) & AppPath & "Check_Virus.exe" & ChrW(34) & " " & ChrW(34) & "%1" & ChrW(34)
404         SaveString HKEY_CLASSES_ROOT, "comfile\shell\open\command", "", ChrW(34) & AppPath & "Check_Virus.exe" & ChrW(34) & " " & ChrW(34) & "%1" & ChrW(34)
        Else
406         atsScanEXE.Caption = "Tu75 d9o65ng que1t va2 ca3nh ba1o Virus cho nhu74ng File sa81p d9u7o75c mo73. [D9ang Ta81t]"
408         SaveString HKEY_CLASSES_ROOT, "exefile\shell\open\command", "", ChrW(34) & "%1" & ChrW(34) & " %*"
410         SaveString HKEY_CLASSES_ROOT, "batfile\shell\open\command", "", ChrW(34) & "%1" & ChrW(34) & " %*"
412         SaveString HKEY_CLASSES_ROOT, "cmdfile\shell\open\command", "", ChrW(34) & "%1" & ChrW(34) & " %*"
414         SaveString HKEY_CLASSES_ROOT, "comfile\shell\open\command", "", ChrW(34) & "%1" & ChrW(34) & " %*"
        End If

416     If atsAutoScanUSB.Value = True Then
418         atsAutoScanUSB.Caption = "Tu75 d9o65ng que1t Virus cho USB khi pha1t hie65n USB ke61t no61i va2o ma1y ti1nh. [D9ang Ba65t]"
420         Load zfrmScanUSB
        Else
422         atsAutoScanUSB.Caption = "Tu75 d9o65ng que1t Virus cho USB khi pha1t hie65n USB ke61t no61i va2o ma1y ti1nh. [D9ang Ta81t]"
        End If

424     If atsScanKeylogger.Value = True Then
426         Load zfrmAntiKey
428         atsScanKeylogger.Caption = "Tu75 d9o65ng pha1t hie65n va2 ca3nh ba1o Keylogger. [D9ang Ba65t]"
        Else
430         atsScanKeylogger.Caption = "Tu75 d9o65ng pha1t hie65n va2 ca3nh ba1o Keylogger. [D9ang Ta81t]"
        End If

        '<--- Load auto scan

        '---> Load main form
        '/////
432     mf.Visible = True
434     HienThiMain
        '////
        '<--- load main form

        'Linh tinh
436     mf.Left = 3000
438     mf.Top = 1440
        'Linh tinh

440     If xTask = False Then frmFlash.Timer2.Enabled = True

        '<EhFooter>
        Exit Sub

Form_Load_Err:
        WriteErr Time & "-" & Date & " - " & Err.Description & " - " & "in PerfectAntivirus2009.frmMain.Form_Load " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Form_Unload(Cancel As Integer)

        '<EhHeader>
        On Error GoTo Form_Unload_Err

        '</EhHeader>

100     Cancel = 1

102     frmMain.Visible = False
104     App.TaskVisible = False

        '<EhFooter>
        Exit Sub

Form_Unload_Err:
        WriteErr Time & "-" & Date & " - " & Err.Description & " - " & "in PerfectAntivirus2009.frmMain.Form_Unload " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub xStartFullScanNow()

        '<EhHeader>
        On Error GoTo xStartFullScanNow_Err

        '</EhHeader>

100     If Me.VSScanProcess.Value = True Then
102         lblStatus2.Caption = "D9ang que1t ca1c chu7o7ng tri2nh cha5y trong bo65 nho71..."

104         modScanVirus.xScanProcess

        End If

106     If Me.VSScanStartUp.Value = True Then

108         lblStatus2.Caption = "D9ang que1t ca1c chu7o7ng tri2nh kho73i d9o65ng..."
    
110         lblStatus.Caption = "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Run"
112         GetKeyValue "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Run"
    
114         lblStatus.Caption = "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\RunOnce"
116         GetKeyValue "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\RunOnce"
    
118         lblStatus.Caption = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Run"
120         GetKeyValue "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Run"
    
122         lblStatus.Caption = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\RunOnce"
124         GetKeyValue "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\RunOnce"
    
126         lblStatus.Caption = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\RunOnceEx"
128         GetKeyValue "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\RunOnceEx"
        End If

        Dim Str

        Dim str2

        Dim FSO  As New FileSystemObject

        Dim drv  As Drive

        Dim drvs As Drives

        'On Error Resume Next
130     Set drvs = FSO.Drives

132     For Each drv In drvs

134         If xFSStopScan = True Then GoTo SsKkIiPp
136         If UCase(drv.DriveLetter) <> "A" Then

138             DoEvents
140             lblStatus2.Caption = "D9ang que1t File EXE..."
142             SearchFile drv.DriveLetter & ":\", "*.exe"
144             lblStatus2.Caption = "D9ang que1t File BAT..."

146             If Me.VSbat.Value = True Then SearchFile drv.DriveLetter & ":\", "*.bat"
148             lblStatus2.Caption = "D9ang que1t File CMD..."

150             If Me.VScmd.Value = True Then SearchFile drv.DriveLetter & ":\", "*.cmd"
152             lblStatus2.Caption = "D9ang que1t File COM..."

154             If Me.VScom.Value = True Then SearchFile drv.DriveLetter & ":\", "*.com"
156             lblStatus2.Caption = "D9ang que1t File SCR..."

158             If Me.VSscr.Value = True Then SearchFile drv.DriveLetter & ":\", "*.scr"
160             lblStatus2.Caption = "D9ang que1t File DLL..."

162             If Me.VSdll.Value = True Then SearchFile drv.DriveLetter & ":\", "*.dll"
                
                'Kill drv.DriveLetter & ":\autorun.inf"
                'SearchFile "C:\", "*.exe"

            End If

        Next

SsKkIiPp:
164     Set FSO = Nothing
166     Set drv = Nothing
168     Set drvs = Nothing

170     DoEvents

172     lblStatus.AutoUnicode = True
174     lblStatus.Caption = "D9a4 Que1t Xong! Ti2m Tha61y: " & LVVirus1.ListItems.Count & " Ta65p Tin Virus Trong To63ng So61 " & xTotalFile & " Ta65p Tin D9a4 Ti2m"
176     lblStatus2.Caption = "D9a4 que1t xong!"
178     cmdFSCachLy.Enabled = True
180     cmdFSKillVirus.Enabled = True
182     cmdFSReport.Enabled = True
184     cmdFSStop.Enabled = False
186     cmdFSStart.Enabled = True
188     tmrScanTime.Enabled = False
190     fmMain.Enabled = True
        '''''''
192     frmMenu.quetvadiet.Visible = True
194     frmMenu.ngang0.Visible = True
196     frmMenu.tudongbaove.Visible = True
198     frmMenu.tienichhethong.Visible = True
200     frmMenu.caidatcauhinh.Visible = True
202     frmMenu.tacgia.Visible = True
        '''''''
204     cmdSettingFullScan.Enabled = True
206     PLaySound AppPath & "Sound\ScanDone.wav"
208     Tray1.ToolTipText = "Perfect Antivirus 2009 - Ma1y ti1nh cu3a ba5n d9ang o73 ti2nh tra5ng to61t nha61t!"

210     If frmMain.Visible = False Then
212         frmMessenger.zShowMessenger "D9a4 que1t xong", "Chu7o7ng tri2nh d9a4 que1t Virus xong! Ke61t qua3: Ti2m Tha61y: " & LVVirus1.ListItems.Count & " Ta65p Tin Virus Trong To63ng So61 " & xTotalFile & " Ta65p Tin D9a4 Ti2m. Ha4y ba65t chu7o7ng tri2nh le6n va2 xu73 ly1 chu1ng.", 5000, xTrang
        End If

214     With LVVirusEvents

            Dim I

216         I = .ListItems.Count + 1
218         .ListItems.Add I, , TimeFull
220         .ListItems(I).SubItems(1).Caption = ToUnicode("Que1t toa2n bo65")
222         .ListItems(I).SubItems(2).Caption = xTotalFile
224         .ListItems(I).SubItems(3).Caption = LVVirus1.ListItems.Count
226         .ListItems(I).SubItems(4).Caption = IIf(HoanThanhFull, ToUnicode("D9a4 hoa2n tha2nh"), ToUnicode("Chu7a hoa2n tha2nh"))
228         .AutoUnicode = True
        End With

        '<EhFooter>
        Exit Sub

xStartFullScanNow_Err:
        WriteErr Time & "-" & Date & " - " & Err.Description & " - " & "in PerfectAntivirus2009.frmMain.xStartFullScanNow " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub tmrScanTime_Timer()

    '<EhHeader>
    On Error Resume Next

    '</EhHeader>

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

    '<EhHeader>
    On Error Resume Next

    '</EhHeader>

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

Private Sub Tray1_TrayClick(Button As UniControls.stMouseEvent)

        '<EhHeader>
        On Error GoTo Tray1_TrayClick_Err

        '</EhHeader>
100     If Button = stRightButtonDown Then
102         frmMenu.baoveregistry.Checked = Me.atptmrREG.Enabled
104         frmMenu.baoveautorun.Checked = Me.atptmrAutorun.Enabled
    
106         frmMenu.ats(1).Checked = Me.atsAlwaysScanFolder.Value
108         frmMenu.ats(0).Checked = Me.atsScanEXE.Value
110         frmMenu.ats(2).Checked = Me.atsAutoScanUSB.Value
112         frmMenu.ats(3).Checked = Me.atsScanKeylogger.Value
    
114         frmMenu.UniMenu1.InitUnicodeMenu frmMenu.hWnd
116         PopupMenu frmMenu.m
118     ElseIf Button = stLeftButtonDoubleClick Then
120         frmMain.Visible = True
122         frmMain.Show
124         frmMain.WindowState = 0
126         App.TaskVisible = True
        End If

        '<EhFooter>
        Exit Sub

Tray1_TrayClick_Err:
        WriteErr Time & "-" & Date & " - " & Err.Description & " - " & "in PerfectAntivirus2009.frmMain.Tray1_TrayClick " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub TreeChucNang_NodeClick(ByVal hNode As Long)

        '<EhHeader>
        On Error GoTo TreeChucNang_NodeClick_Err

        '</EhHeader>

100     HideAllFM

102     If IsNumeric(TreeChucNang.GetNodeKey(hNode)) = True Then
104         fm(TreeChucNang.GetNodeKey(hNode)).Visible = True
        Else
106         mf.Visible = True
108         HienThiMain
        End If

        '<EhFooter>
        Exit Sub

TreeChucNang_NodeClick_Err:
        WriteErr Time & "-" & Date & " - " & Err.Description & " - " & "in PerfectAntivirus2009.frmMain.TreeChucNang_NodeClick " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub VSbat_Click()

        '<EhHeader>
        On Error GoTo VSbat_Click_Err

        '</EhHeader>

100     CSbat.Value = VSbat.Value

        '<EhFooter>
        Exit Sub

VSbat_Click_Err:
        WriteErr Time & "-" & Date & " - " & Err.Description & " - " & "in PerfectAntivirus2009.frmMain.VSbat_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub VScom_Click()

        '<EhHeader>
        On Error GoTo VScom_Click_Err

        '</EhHeader>

100     CScom.Value = VScom.Value

        '<EhFooter>
        Exit Sub

VScom_Click_Err:
        WriteErr Time & "-" & Date & " - " & Err.Description & " - " & "in PerfectAntivirus2009.frmMain.VScom_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub VSScanProcess_Click()

        '<EhHeader>
        On Error GoTo VSScanProcess_Click_Err

        '</EhHeader>

100     CSprocess.Value = VSScanProcess.Value

        '<EhFooter>
        Exit Sub

VSScanProcess_Click_Err:
        WriteErr Time & "-" & Date & " - " & Err.Description & " - " & "in PerfectAntivirus2009.frmMain.VSScanProcess_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub VSScanStartUp_Click()

        '<EhHeader>
        On Error GoTo VSScanStartUp_Click_Err

        '</EhHeader>

100     CSStartUp.Value = VSScanStartUp.Value

        '<EhFooter>
        Exit Sub

VSScanStartUp_Click_Err:
        WriteErr Time & "-" & Date & " - " & Err.Description & " - " & "in PerfectAntivirus2009.frmMain.VSScanStartUp_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub HienThiMain()

        '/////
        '<EhHeader>
        On Error GoTo HienThiMain_Err

        '</EhHeader>
100     With Me
102         .xlblComputerName.Caption = ""
104         .xlblProcess.Caption = ""
106         .xlblRAM.Caption = ""
108         .xlblTinhTrang.Caption = ""
110         .xlblUserName.Caption = ""
        
112         .xlblComputerName.Caption = GetComputer

114         DoEvents

            Dim mPro

116         mPro = 0

            Dim ColItems

            Dim ObjItem

118         Set ColItems = GetObject("winmgmts:\root\CIMV2").ExecQuery("SELECT * FROM Win32_Process")

120         DoEvents

122         For Each ObjItem In ColItems

124             If IsNull(ObjItem.ExecutablePath) = False Then mPro = mPro + 1

126             DoEvents
            Next

128         Set ColItems = Nothing
130         Set ObjItem = Nothing
132         .xlblProcess.Caption = mPro
134         .xlblRAM.Caption = GetRAMTotal
136         .xlblUserName.Caption = Environ$("USERNAME")
138         .xlblTinhTrang.Caption = CheckComputerHeal
140         .xProcessRAM.Value = GetMemoryInfo
        
        End With

142     If Me.atptmrREG.Enabled = True Then PicOnOff(0).Picture = picOn.Picture Else PicOnOff(0).Picture = PicOff.Picture
144     If Me.atptmrAutorun.Enabled = True Then PicOnOff(1).Picture = picOn.Picture Else PicOnOff(1).Picture = PicOff.Picture
146     If Me.atsScanEXE.Value = True Then PicOnOff(2).Picture = picOn.Picture Else PicOnOff(2).Picture = PicOff.Picture
148     If Me.atsAlwaysScanFolder.Value = True Then PicOnOff(3).Picture = picOn.Picture Else PicOnOff(3).Picture = PicOff.Picture
150     If Me.atsScanKeylogger.Value = True Then PicOnOff(4).Picture = picOn.Picture Else PicOnOff(4).Picture = PicOff.Picture
152     If Me.atsAutoScanUSB.Value = True Then PicOnOff(5).Picture = picOn.Picture Else PicOnOff(5).Picture = PicOff.Picture
154     If Me.chkAutoStart.Value = True Then PicOnOff(6).Picture = picOn.Picture Else PicOnOff(6).Picture = PicOff.Picture
156     If Me.chkAutoUpdate.Value = True Then PicOnOff(7).Picture = picOn.Picture Else PicOnOff(7).Picture = PicOff.Picture

        '////
        '<EhFooter>
        Exit Sub

HienThiMain_Err:
        WriteErr Time & "-" & Date & " - " & Err.Description & " - " & "in PerfectAntivirus2009.frmMain.HienThiMain " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub
