VERSION 5.00
Object = "{D1ECD258-D329-4388-AB83-DEC261A66B86}#1.0#0"; "UniControls_v2.0.ocx"
Begin VB.Form frmVirusReport 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Virus Report"
   ClientHeight    =   3750
   ClientLeft      =   165
   ClientTop       =   465
   ClientWidth     =   4695
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmVirusReport.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3750
   ScaleWidth      =   4695
   StartUpPosition =   2  'CenterScreen
   Begin UniControls.UniLabel lblTime 
      Height          =   255
      Left            =   1560
      Top             =   2760
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   450
      Alignment       =   1
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
   Begin UniControls.UniLabel lblHoanThanh 
      Height          =   255
      Left            =   1920
      Top             =   2400
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   450
      Alignment       =   1
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
      ForeColor       =   255
   End
   Begin UniControls.UniLabel lblTotalVirus 
      Height          =   255
      Left            =   1920
      Top             =   2040
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   450
      Alignment       =   1
      Caption         =   "0"
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
   Begin UniControls.UniLabel lblTotalProcess 
      Height          =   255
      Left            =   2880
      Top             =   1680
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   450
      Alignment       =   1
      Caption         =   "0"
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
   Begin UniControls.UniLabel lblTotalStartUp 
      Height          =   255
      Left            =   2880
      Top             =   1320
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   450
      Alignment       =   1
      Caption         =   "0"
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
   Begin UniControls.UniLabel lblTotalFile 
      Height          =   255
      Left            =   1920
      Top             =   960
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   450
      Alignment       =   1
      Caption         =   "0"
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
   Begin UniControls.UniLabel UniLabel8 
      Height          =   255
      Left            =   3840
      Top             =   960
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   450
      Caption         =   "File"
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
   Begin UniControls.UniButton cmdVirusReportOK 
      Height          =   375
      Left            =   1680
      TabIndex        =   0
      Top             =   3240
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      Icon            =   "frmVirusReport.frx":0A02
      Style           =   2
      Caption         =   "Cha61p nha65n"
      IconAlign       =   3
      iNonThemeStyle  =   2
      MaskColor       =   16711935
      FontColor       =   16711680
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      RoundedBordersByTheme=   0   'False
      ShowFocusRectangle=   0   'False
   End
   Begin UniControls.UniLabel UniLabel7 
      Height          =   255
      Left            =   240
      Top             =   2760
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   450
      Caption         =   "Tho72i gian que1t:"
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
   Begin UniControls.UniLabel UniLabel6 
      Height          =   255
      Left            =   240
      Top             =   2400
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   450
      Caption         =   "Qua1 Tri2nh Que1t:"
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
      Left            =   240
      Top             =   1680
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   450
      Caption         =   "To63ng so61 ca1c tie61n tri2nh d9ang cha5y:"
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
      Left            =   240
      Top             =   1320
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   450
      Caption         =   "To63ng so61 chu7o7ng tri2nh kho73i d9o65ng:"
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
      Left            =   240
      Top             =   2040
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   450
      Caption         =   "So61 Virus pha1t hie65n:"
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
      Left            =   240
      Top             =   960
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   450
      Caption         =   "To63ng so61 File d9a4 que1t:"
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
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   661
      Alignment       =   1
      Caption         =   "Ke61t qua3 que1t"
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
   Begin UniControls.UniLabel UniLabel9 
      Height          =   255
      Left            =   3840
      Top             =   1320
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   450
      Caption         =   "File"
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
      Left            =   3840
      Top             =   1680
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   450
      Caption         =   "File"
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
      Height          =   255
      Left            =   3840
      Top             =   2040
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   450
      Caption         =   "File"
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
   Begin UniControls.UniLabel UniLabel12 
      Height          =   255
      Left            =   3840
      Top             =   2400
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   450
      Caption         =   "File"
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
Attribute VB_Name = "frmVirusReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdVirusReportOK_Click()
Unload Me
End Sub

Public Sub ShowReport(TotalFile, StartUpFile, ProcessFile, Virus, HoanThanh, Time)
With Me
    .lblTime.Caption = Time
    .lblTotalFile.Caption = TotalFile
    .lblHoanThanh.Caption = HoanThanh
    .lblTotalProcess.Caption = ProcessFile
    .lblTotalStartUp.Caption = StartUpFile
    .lblTotalVirus.Caption = Virus
    .Show
End With
End Sub

