VERSION 5.00
Object = "{E8FDD05C-3067-4198-8AEC-1A013A46ABDD}#1.0#0"; "FVUnicodeControl.ocx"
Object = "{6DC1AB90-DCD1-47A1-AB36-924FFD67ADBF}#1.0#0"; "UniControls_v2.0.ocx"
Begin VB.Form frmRegistry 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PAV 2009 - Registry Editor"
   ClientHeight    =   5265
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10305
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmRegistry.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5265
   ScaleWidth      =   10305
   StartUpPosition =   2  'CenterScreen
   Begin UniControls.UniLabel UniLabel5 
      Height          =   255
      Left            =   240
      Top             =   4680
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   450
      Caption         =   "* Mo65t so61 chu71c na8ng ca62n Log Off hoa85c kho73i d9o65ng la5i ma1y thi2 mo71i co1 ta1c du5ng!"
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
   Begin UniControls.UniLabel UniLabel4 
      Height          =   255
      Left            =   240
      Top             =   4440
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   450
      Caption         =   "* Ca1c chu71c na8ng kho6ng d9u7o75c d9a1nh da61u se4 bi5 vo6 hie65u ho1a khi ba5n nha61n nu1t thu75c hie65n!"
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
   Begin UniControls.UniLabel UniLabel1 
      Height          =   375
      Left            =   3240
      Top             =   240
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   661
      Alignment       =   1
      Caption         =   "Tinh chi3nh Registry"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   255
   End
   Begin FVUnicodeControl.FVistaUniFrame FMOTHER 
      Height          =   2175
      Left            =   6840
      TabIndex        =   0
      Top             =   960
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   3836
      Alignment       =   0
      BackColor       =   -2147483643
      ForeColor       =   16711680
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Ca1c Ti1nh Na8ng Kha1c"
      AutoUnicode     =   -1  'True
      Begin FVUnicodeControl.FVistaUniCheckbox chkWrite 
         Height          =   195
         Left            =   240
         TabIndex        =   1
         Top             =   1800
         Width           =   1890
         _ExtentX        =   3334
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
         Caption         =   "Cho phe1p ghi va2o USB"
         ForeColor       =   32768
      End
      Begin FVUnicodeControl.FVistaUniCheckbox chkUSB 
         Height          =   195
         Left            =   240
         TabIndex        =   2
         Top             =   1440
         Width           =   1485
         _ExtentX        =   2619
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
         Caption         =   "Cho phe1p O63 USB"
         ForeColor       =   32768
      End
      Begin FVUnicodeControl.FVistaUniCheckbox chkAUTORUN 
         Height          =   195
         Left            =   240
         TabIndex        =   3
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
         Caption         =   "Kho6ng cha5y Autorun tu72 USB"
         ForeColor       =   32768
      End
      Begin FVUnicodeControl.FVistaUniCheckbox chkHIDDEN 
         Height          =   195
         Left            =   240
         TabIndex        =   4
         Top             =   720
         Width           =   1890
         _ExtentX        =   3334
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
         Caption         =   "Kho6ng Hie63n Thi5 File A63n"
         ForeColor       =   32768
      End
      Begin FVUnicodeControl.FVistaUniCheckbox chkEXE 
         Height          =   195
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Width           =   1530
         _ExtentX        =   2699
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
         Caption         =   "Hie63n Thi5 D9uo6i File"
         ForeColor       =   32768
      End
   End
   Begin FVUnicodeControl.FVistaUniFrame FMChucNang 
      Height          =   3135
      Left            =   3480
      TabIndex        =   6
      Top             =   960
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   5530
      Alignment       =   0
      BackColor       =   -2147483643
      ForeColor       =   16711680
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "A63n/Hie65n Ca1c Ti1nh Na8ng"
      AutoUnicode     =   -1  'True
      Begin FVUnicodeControl.FVistaUniCheckbox chkDOC 
         Height          =   195
         Left            =   240
         TabIndex        =   7
         Top             =   2760
         Width           =   1635
         _ExtentX        =   2884
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
         Caption         =   "Hie65n Folder Option"
         ForeColor       =   32768
      End
      Begin FVUnicodeControl.FVistaUniCheckbox chkPRO 
         Height          =   195
         Left            =   240
         TabIndex        =   8
         Top             =   2520
         Width           =   2505
         _ExtentX        =   4419
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
         Caption         =   "Hie65n All Programs (Start Menu)"
         ForeColor       =   32768
      End
      Begin FVUnicodeControl.FVistaUniCheckbox chkTURNOFF 
         Height          =   195
         Left            =   240
         TabIndex        =   9
         Top             =   360
         Width           =   2535
         _ExtentX        =   4471
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
         Caption         =   "Hie65n Nu1t Turn Off (Start Menu)"
         ForeColor       =   32768
      End
      Begin FVUnicodeControl.FVistaUniCheckbox chkLOGOFF 
         Height          =   195
         Left            =   240
         TabIndex        =   10
         Top             =   600
         Width           =   2460
         _ExtentX        =   4339
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
         Caption         =   "Hie65n Nu1t Log Off (Start Menu)"
         ForeColor       =   32768
      End
      Begin FVUnicodeControl.FVistaUniCheckbox chkSearch 
         Height          =   195
         Left            =   240
         TabIndex        =   11
         Top             =   840
         Width           =   2640
         _ExtentX        =   4657
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
         Caption         =   "Hie65n Search Engine (Start Menu)"
         ForeColor       =   32768
      End
      Begin FVUnicodeControl.FVistaUniCheckbox chkHelp 
         Height          =   195
         Left            =   240
         TabIndex        =   12
         Top             =   1080
         Width           =   1905
         _ExtentX        =   3360
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
         Caption         =   "Hie65n Help and Support"
         ForeColor       =   32768
      End
      Begin FVUnicodeControl.FVistaUniCheckbox chkRUN 
         Height          =   195
         Left            =   240
         TabIndex        =   13
         Top             =   1320
         Width           =   2265
         _ExtentX        =   3995
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
         Caption         =   "Hie65n Tri2nh Le65nh D9o7n Run..."
         ForeColor       =   32768
      End
      Begin FVUnicodeControl.FVistaUniCheckbox chkCPITEM 
         Height          =   195
         Left            =   240
         TabIndex        =   14
         Top             =   1560
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
         Caption         =   "Hie65n Control Panel Items"
         ForeColor       =   32768
      End
      Begin FVUnicodeControl.FVistaUniCheckbox chkTRAYCLOCK 
         Height          =   195
         Left            =   240
         TabIndex        =   15
         Top             =   1800
         Width           =   2100
         _ExtentX        =   3704
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
         Caption         =   "Hie65n Tray Icons && Clock"
         ForeColor       =   32768
      End
      Begin FVUnicodeControl.FVistaUniCheckbox chkCOMPUTER 
         Height          =   195
         Left            =   240
         TabIndex        =   16
         Top             =   2040
         Width           =   2145
         _ExtentX        =   3784
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
         Caption         =   "Hie65n Computer Properties"
         ForeColor       =   32768
      End
      Begin FVUnicodeControl.FVistaUniCheckbox chkCPA 
         Height          =   195
         Left            =   240
         TabIndex        =   17
         Top             =   2280
         Width           =   1620
         _ExtentX        =   2858
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
         Caption         =   "Hie65n Control Panel"
         ForeColor       =   32768
      End
   End
   Begin FVUnicodeControl.FVistaUniFrame FMMoKhoaTinhNang 
      Height          =   3135
      Left            =   120
      TabIndex        =   18
      Top             =   960
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   5530
      Alignment       =   0
      BackColor       =   -2147483643
      ForeColor       =   16711680
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Kho1a/Mo73 Ca1c Ti1nh Na8ng"
      AutoUnicode     =   -1  'True
      Begin FVUnicodeControl.FVistaUniCheckbox chkWin 
         Height          =   195
         Left            =   120
         TabIndex        =   19
         Top             =   2760
         Width           =   2655
         _ExtentX        =   4683
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
         Caption         =   "Cho phe1p Phi1m Windows + [Key]"
         ForeColor       =   32768
      End
      Begin FVUnicodeControl.FVistaUniCheckbox chkTaskbar 
         Height          =   195
         Left            =   120
         TabIndex        =   20
         Top             =   2520
         Width           =   2910
         _ExtentX        =   5133
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
         Caption         =   "Cho phe1p Ca2i D9a85t Taskbar Va2 Folder"
         ForeColor       =   32768
      End
      Begin FVUnicodeControl.FVistaUniCheckbox chkDESKTOP 
         Height          =   195
         Left            =   120
         TabIndex        =   21
         Top             =   2280
         Width           =   2235
         _ExtentX        =   3942
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
         Caption         =   "Hie63n Thi5 Icon Tre6n Desktop"
         ForeColor       =   32768
      End
      Begin FVUnicodeControl.FVistaUniCheckbox chkFILEMENU 
         Height          =   195
         Left            =   120
         TabIndex        =   22
         Top             =   2040
         Width           =   2100
         _ExtentX        =   3704
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
         Caption         =   "Hie65n File Menu (Explorer)"
         ForeColor       =   32768
      End
      Begin FVUnicodeControl.FVistaUniCheckbox chkRIGHT 
         Height          =   195
         Left            =   120
         TabIndex        =   23
         Top             =   1800
         Width           =   1875
         _ExtentX        =   3307
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
         Caption         =   "Hie65n Menu Chuo65t Pha3i"
         ForeColor       =   32768
      End
      Begin FVUnicodeControl.FVistaUniCheckbox chkTRAY 
         Height          =   195
         Left            =   120
         TabIndex        =   24
         Top             =   1560
         Width           =   2055
         _ExtentX        =   3625
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
         Caption         =   "Hie65n Tray Context Menu"
         ForeColor       =   32768
      End
      Begin FVUnicodeControl.FVistaUniCheckbox chkIEHOME 
         Height          =   195
         Left            =   120
         TabIndex        =   25
         Top             =   1320
         Width           =   2430
         _ExtentX        =   4286
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
         Caption         =   "Hie65n Thay D9o63i IE Home Pages"
         ForeColor       =   32768
      End
      Begin FVUnicodeControl.FVistaUniCheckbox chkCP 
         Height          =   195
         Left            =   120
         TabIndex        =   26
         Top             =   1080
         Width           =   1620
         _ExtentX        =   2858
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
         Caption         =   "Hie65n Control Panel"
         ForeColor       =   32768
      End
      Begin FVUnicodeControl.FVistaUniCheckbox chkCMD 
         Height          =   195
         Left            =   120
         TabIndex        =   27
         Top             =   840
         Width           =   2415
         _ExtentX        =   4260
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
         Caption         =   "Hie65n Command Prompt (CMD)"
         ForeColor       =   32768
      End
      Begin FVUnicodeControl.FVistaUniCheckbox chkREG 
         Height          =   195
         Left            =   120
         TabIndex        =   28
         Top             =   600
         Width           =   1725
         _ExtentX        =   3043
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
         Caption         =   "Hie65n Registry Editor"
         ForeColor       =   32768
      End
      Begin FVUnicodeControl.FVistaUniCheckbox chkTask 
         Height          =   195
         Left            =   120
         TabIndex        =   29
         Top             =   360
         Width           =   2355
         _ExtentX        =   4154
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
         Caption         =   "Hie65n Windows Task Manager"
         ForeColor       =   32768
      End
   End
   Begin UniControls.UniLabel UniLabel2 
      Height          =   255
      Left            =   240
      Top             =   4200
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   450
      Caption         =   "* Ma1y ti1nh cu3a ba5n se4 o73 ti2nh tra5ng to61t nha65t ne61u ta61t ca3 ca1c mu5c d9e62u d9u7o75c d9a1nh da61u!"
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
   Begin UniControls.UniLabel UniLabel3 
      Height          =   255
      Left            =   240
      Top             =   4920
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   450
      Caption         =   "* Nha61n va2o nu1t Lu7u La5i sau khi d9a4 thie61t la65p xong!"
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
   Begin FVUnicodeControl.FVistaUniButton cmdSaveRegistry 
      Height          =   375
      Left            =   7920
      TabIndex        =   30
      Top             =   3360
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      BackColor       =   14737632
      ButtonStyle     =   3
      Caption         =   "Lu7u la5i"
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
End
Attribute VB_Name = "frmRegistry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdSaveRegistry_Click()

    If chkDOC.Value = False Then
        SaveDWORD HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoFolderOptions", 1
    Else
        SaveDWORD HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoFolderOptions", 0
    End If

    If chkPRO.Value = False Then
        SaveDWORD HKEY_CURRENT_USER, "SOFTWARE\Microsoft\Windows\CurrentVersion\policies\Explorer", "NoStartMenuMorePrograms", 1
    Else
        SaveDWORD HKEY_CURRENT_USER, "SOFTWARE\Microsoft\Windows\CurrentVersion\policies\Explorer", "NoStartMenuMorePrograms", 0
    End If

    If chkAUTORUN.Value = False Then
        SaveDWORD HKEY_CURRENT_USER, "SOFTWARE\Microsoft\Windows\CurrentVersion\policies\Explorer", "NoDriveTypeAutoRun", 44
    Else
        SaveDWORD HKEY_CURRENT_USER, "SOFTWARE\Microsoft\Windows\CurrentVersion\policies\Explorer", "NoDriveTypeAutoRun", 0
    End If

    If chkCMD.Value = False Then
        SaveDWORD HKEY_CURRENT_USER, "Software\Policies\Microsoft\Windows\System", "DisableCMD", 1
    Else
        DeleteValue HKEY_CURRENT_USER, "Software\Policies\Microsoft\Windows\System", "DisableCMD"
    End If

    If chkCOMPUTER.Value = False Then
        SaveDWORD HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoPropertiesMyComputer", 1
    Else
        DeleteValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoPropertiesMyComputer"
    End If

    If chkCP.Value = False Then
        SaveDWORD HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoControlPanel", 1
    Else
        DeleteValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoControlPanel"
    End If

    If chkCPA.Value = False Then
        SaveDWORD HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "DisallowCpl", 1
        SaveDWORD HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "RestrictCpl", 1

    Else
        DeleteValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "DisallowCpl"
        DeleteValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "RestrictCpl"

    End If

    If chkCPITEM.Value = False Then
        SaveDWORD HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\System", "NoDispCpl", 1
    Else
        DeleteValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\System", "NoDispCpl"
    End If

    If chkDESKTOP.Value = False Then
        SaveDWORD HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoDesktop", 1
    Else
        DeleteValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoDesktop"
    End If

    If chkEXE.Value = False Then
        SaveDWORD HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced", "HideFileExt", 1
    Else
        SaveDWORD HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced", "HideFileExt", 0
    End If

    If chkFILEMENU.Value = False Then
        SaveDWORD HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "FileMenu", 1
    Else
        DeleteValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "FileMenu"
    End If

    If chkHelp.Value = False Then
        SaveDWORD HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoSMHelp", 1
    Else
        DeleteValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoSMHelp"
    End If

    If chkHIDDEN.Value = False Then
        SaveDWORD HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced", "Hidden", 1
        SaveDWORD HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced", "ShowSuperHidden", 1
        SaveDWORD HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced", "SuperHidden", 1

    Else
        SaveDWORD HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced", "Hidden", 0
        SaveDWORD HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced", "ShowSuperHidden", 0
        SaveDWORD HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced", "SuperHidden", 0
    End If

    If chkIEHOME.Value = False Then
        SaveDWORD HKEY_CURRENT_USER, "Software\Policies\Microsoft\Internet Explorer\Control Panel", "HomePage", 1
    Else
        DeleteValue HKEY_CURRENT_USER, "Software\Policies\Microsoft\Internet Explorer\Control Panel", "HomePage"
    End If

    If chkLOGOFF.Value = False Then
        SaveDWORD HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoLogOff", 1
    Else
        DeleteValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoLogOff"
    End If

    If chkREG.Value = False Then
        SaveDWORD HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\System", "DisableRegistryTools", 1
    Else
        DeleteValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\System", "DisableRegistryTools"
    End If

    If chkRIGHT.Value = False Then
        SaveDWORD HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoViewContextMenu", 1
    Else
        DeleteValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoViewContextMenu"
    End If

    If chkRUN.Value = False Then
        SaveDWORD HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoRun", 1
    Else
        DeleteValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoRun"
    End If

    If chkSearch.Value = False Then
        SaveDWORD HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoFind", 1
    Else
        DeleteValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoFind"
    End If

    If chkTask.Value = False Then
        SaveDWORD HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\System", "DisableTaskMgr", 1
    Else
        DeleteValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\System", "DisableTaskMgr"
    End If

    If chkTaskbar.Value = False Then
        SaveDWORD HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoSetTaskbar", 1
        SaveDWORD HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoSetFolders", 1

    Else
        DeleteValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoSetTaskbar"
        DeleteValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoSetFolders"

    End If

    If chkTRAY.Value = False Then
        SaveDWORD HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoTrayContextMenu", 1
    Else
        DeleteValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoTrayContextMenu"
    End If

    If chkTRAYCLOCK.Value = False Then
        SaveDWORD HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "HideClock", 1
        SaveDWORD HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoTrayItemsDisplay", 1

    Else
        DeleteValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "HideClock"
        DeleteValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoTrayItemsDisplay"

    End If

    If chkTURNOFF.Value = False Then
        SaveDWORD HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoClose", 1
    Else
        DeleteValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoClose"
    End If

    If chkUSB.Value = False Then
        SaveDWORD HKEY_LOCAL_MACHINE, "SYSTEM\ControlSet001\Services\USBSTOR", "Start", 4
    Else
        SaveDWORD HKEY_LOCAL_MACHINE, "SYSTEM\ControlSet001\Services\USBSTOR", "Start", 3
    End If

    If chkWrite.Value = False Then
        SaveDWORD HKEY_LOCAL_MACHINE, "SYSTEM\ControlSet001\Control\StorageDevicePolicies", "WriteProtect", 1
    Else
        SaveDWORD HKEY_LOCAL_MACHINE, "SYSTEM\ControlSet001\Control\StorageDevicePolicies", "WriteProtect", 0
    End If

    If chkWin.Value = False Then
        SaveDWORD HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoWinKeys", 1
    Else
        SaveDWORD HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoWinKeys", 0
    End If

    UniMsgBox "D9a4 thie61t la65p xong!", vbOKOnly, "OK!", Me.hwnd

End Sub

Private Sub Form_Load()

    If GetDWORD(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\System", "DisableTaskMgr") = 1 Then chkTask.Value = False Else chkTask.Value = True
    If GetDWORD(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\System", "DisableRegistryTools") = 1 Then chkREG.Value = False Else chkREG.Value = True
    If GetDWORD(HKEY_CURRENT_USER, "Software\Policies\Microsoft\Windows\System", "DisableCMD") = 1 Then chkCMD.Value = False Else chkCMD.Value = True
    If GetDWORD(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoControlPanel") = 1 Then chkCP.Value = False Else chkCP.Value = True
    If GetDWORD(HKEY_CURRENT_USER, "Software\Policies\Microsoft\Internet Explorer\Control Panel", "HomePage") = 1 Then chkIEHOME.Value = False Else chkIEHOME.Value = True
    If GetDWORD(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoTrayContextMenu") = 1 Then chkTRAY.Value = False Else chkTRAY.Value = True
    If GetDWORD(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoViewContextMenu") = 1 Then chkRIGHT.Value = False Else chkRIGHT.Value = True
    If GetDWORD(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "FileMenu") = 1 Then chkFILEMENU.Value = False Else chkFILEMENU.Value = True
    If GetDWORD(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoDesktop") = 1 Then chkDESKTOP.Value = False Else chkDESKTOP.Value = True
    If GetDWORD(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoSetFolders") = 1 Then chkTaskbar.Value = False Else chkTaskbar.Value = True
    If GetDWORD(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoSetTaskbar") = 1 Then chkTaskbar.Value = False Else chkTaskbar.Value = True
    If GetDWORD(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoClose") = 1 Then chkTURNOFF.Value = False Else chkTURNOFF.Value = True
    If GetDWORD(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoLogOff") = 1 Then chkLOGOFF.Value = False Else chkLOGOFF.Value = True
    If GetDWORD(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoRun") = 1 Then chkRUN.Value = False Else chkRUN.Value = True
    If GetDWORD(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoFind") = 1 Then chkSearch.Value = False Else chkSearch.Value = True
    If GetDWORD(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoSMHelp") = 1 Then chkHelp.Value = False Else chkHelp.Value = True
    If GetDWORD(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\System", "NoDispCpl") = 1 Then chkCPITEM.Value = False Else chkCPITEM.Value = True
    If GetDWORD(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoTrayItemsDisplay") = 1 Then chkTRAYCLOCK.Value = False Else chkTRAYCLOCK.Value = True
    If GetDWORD(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "HideClock") = 1 Then chkTRAYCLOCK.Value = False Else chkTRAYCLOCK.Value = True
    If GetDWORD(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoPropertiesMyComputer") = 1 Then chkCOMPUTER.Value = False Else chkCOMPUTER.Value = True
    If GetDWORD(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "RestrictCpl") = 1 Then chkCPA.Value = False Else chkCPA.Value = True
    If GetDWORD(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "DisallowCpl") = 1 Then chkCPA.Value = False Else chkCPA.Value = True
    If GetDWORD(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced", "HideFileExt") = 1 Then chkEXE.Value = False Else chkEXE.Value = True
    If GetDWORD(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced", "ShowSuperHidden") = 0 Then chkHIDDEN.Value = False Else chkHIDDEN.Value = True
    If GetDWORD(HKEY_CURRENT_USER, "SOFTWARE\Microsoft\Windows\CurrentVersion\policies\Explorer", "NoDriveTypeAutoRun") = 44 Then chkHIDDEN.Value = False Else chkHIDDEN.Value = True
    If GetDWORD(HKEY_LOCAL_MACHINE, "SYSTEM\ControlSet001\Services\USBSTOR", "Start") = 4 Then chkUSB.Value = False Else chkUSB.Value = True
    If GetDWORD(HKEY_LOCAL_MACHINE, "SYSTEM\ControlSet001\Control\StorageDevicePolicies", "WriteProtect") = 1 Then chkWrite.Value = False Else chkWrite.Value = True
    If GetDWORD(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoWinKeys") = 1 Then chkWin.Value = False Else chkWin.Value = True
    If GetDWORD(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoStartMenuMorePrograms") = 1 Then chkPRO.Value = False Else chkPRO.Value = True
    If GetDWORD(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoFolderOptions") = 1 Then chkDOC.Value = False Else chkDOC.Value = True

End Sub
