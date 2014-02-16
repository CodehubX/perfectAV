VERSION 5.00
Object = "{E8FDD05C-3067-4198-8AEC-1A013A46ABDD}#1.0#0"; "FVUnicodeControl.ocx"
Object = "{6DC1AB90-DCD1-47A1-AB36-924FFD67ADBF}#1.0#0"; "UniControls_v2.0.ocx"
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Perfect Antivirus 2009"
   ClientHeight    =   5280
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5430
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5280
   ScaleWidth      =   5430
   StartUpPosition =   2  'CenterScreen
   Begin FVUnicodeControl.FVistaUniButton FVistaUniButton7 
      Height          =   495
      Left            =   1560
      TabIndex        =   11
      Top             =   4680
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   873
      BackColor       =   16751432
      Caption         =   "Thoa1t che61 d9o65 kha63n ca61p"
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
   Begin FVUnicodeControl.FVistaUniButton FVistaUniButton3 
      Height          =   615
      Left            =   2880
      TabIndex        =   7
      Top             =   3960
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   1085
      BackColor       =   -2147483633
      ButtonStyle     =   3
      Caption         =   "Mo73 Registry Editor"
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
   Begin FVUnicodeControl.FVistaUniButton FVistaUniButton2 
      Height          =   615
      Left            =   720
      TabIndex        =   6
      Top             =   3960
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   1085
      BackColor       =   -2147483633
      ButtonStyle     =   3
      Caption         =   "Mo73 Task Manager"
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
   Begin UniControls.UniAniPictureBox UniAniPictureBox1 
      Height          =   855
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   1508
      Picture         =   "Form1.frx":058A
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
      Height          =   615
      Left            =   1080
      TabIndex        =   0
      Top             =   240
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   1085
      Caption         =   "Che61 d9o65 hoa5t d9o65ng kha63n ca61p"
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
   Begin FVUnicodeControl.FVistaUniButton Cmd 
      Height          =   615
      Left            =   1440
      TabIndex        =   2
      Top             =   1080
      Width           =   2415
      _ExtentX        =   4260
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
   Begin FVUnicodeControl.FVistaUniButton WE 
      Height          =   615
      Left            =   120
      TabIndex        =   3
      Top             =   2040
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   1085
      BackColor       =   14737632
      ButtonShape     =   3
      ButtonStyle     =   3
      Caption         =   "Phu5c ho62i du74 lie65u"
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
   Begin FVUnicodeControl.FVistaUniButton SDA 
      Height          =   615
      Left            =   120
      TabIndex        =   4
      Top             =   2640
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   1085
      BackColor       =   14737632
      ButtonShape     =   3
      ButtonStyle     =   3
      Caption         =   "Tinh chi3nh Registry"
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
   Begin FVUnicodeControl.FVistaUniButton FVistaUniButton1 
      Height          =   615
      Left            =   120
      TabIndex        =   5
      Top             =   3240
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   1085
      BackColor       =   14737632
      ButtonShape     =   3
      ButtonStyle     =   3
      Caption         =   "Ta8ng to61c ma1y ti1nh"
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
   Begin FVUnicodeControl.FVistaUniButton FVistaUniButton4 
      Height          =   615
      Left            =   2880
      TabIndex        =   8
      Top             =   2040
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   1085
      BackColor       =   14737632
      ButtonShape     =   3
      ButtonStyle     =   3
      Caption         =   "Qua3n ly1 kho73i d9o65ng"
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
   Begin FVUnicodeControl.FVistaUniButton FVistaUniButton5 
      Height          =   615
      Left            =   2880
      TabIndex        =   9
      Top             =   2640
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   1085
      BackColor       =   14737632
      ButtonShape     =   3
      ButtonStyle     =   3
      Caption         =   "Kie63m tra he65 tho61ng"
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
   Begin FVUnicodeControl.FVistaUniButton FVistaUniButton6 
      Height          =   615
      Left            =   2880
      TabIndex        =   10
      Top             =   3240
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   1085
      BackColor       =   14737632
      ButtonShape     =   3
      ButtonStyle     =   3
      Caption         =   "Que1t ma64u virus"
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
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Cmd_Click()
frmScan.Show
xScanning = True
End Sub

Private Sub FVistaUniButton1_Click()
frmTangToc.Show
End Sub

Private Sub FVistaUniButton2_Click()
Shell "taskmgr", vbNormalFocus
End Sub

Private Sub FVistaUniButton3_Click()
Shell "regedit", vbNormalFocus
End Sub

Private Sub FVistaUniButton4_Click()
If modMain.FileExists(AppPath & "PSM.exe") = False Then
    UniMsgBox "Kho6ng ti2m tha61y file PSM.exe!", vbOKOnly + vbCritical, "Error!!!", Me.hwnd
Else
    Shell AppPath & "PSM.exe", vbNormalFocus
End If
End Sub

Private Sub FVistaUniButton5_Click()
If modMain.FileExists(AppPath & "PSR.exe") = False Then
    UniMsgBox "Kho6ng ti2m tha61y file PSR.exe!", vbOKOnly + vbCritical, "Error!!!", Me.hwnd
Else
    Shell AppPath & "PSR.exe", vbNormalFocus
End If
End Sub

Private Sub FVistaUniButton6_Click()
If modMain.FileExists(AppPath & "VRA.exe") = False Then
    UniMsgBox "Kho6ng ti2m tha61y file VRA.exe!", vbOKOnly + vbCritical, "Error!!!", Me.hwnd
Else
    Shell AppPath & "VRA.exe", vbNormalFocus
End If
End Sub

Private Sub FVistaUniButton7_Click()
Unload Me
End Sub

Private Sub SDA_Click()
frmRegistry.Show
End Sub

Private Sub WE_Click()
frmPhucHoiDuLieu.Show

End Sub
