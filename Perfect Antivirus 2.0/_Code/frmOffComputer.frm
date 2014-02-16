VERSION 5.00
Object = "{E8FDD05C-3067-4198-8AEC-1A013A46ABDD}#1.0#0"; "FVUnicodeControl.ocx"
Object = "{6DC1AB90-DCD1-47A1-AB36-924FFD67ADBF}#1.0#0"; "UniControls_v2.0.ocx"
Begin VB.Form frmOffComputer 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Turn Off Computer"
   ClientHeight    =   4095
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmOffComputer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4095
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   3600
      Top             =   1800
   End
   Begin VB.Timer Timer2 
      Interval        =   500
      Left            =   480
      Top             =   2160
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   4200
      Top             =   1680
   End
   Begin FVUnicodeControl.FVistaUniButton cmdCancel 
      Height          =   375
      Left            =   1440
      TabIndex        =   0
      Top             =   3600
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      BackColor       =   -2147483633
      ButtonStyle     =   3
      Caption         =   "Kho6ng ta81t ma1y nu74a"
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
      Height          =   495
      Left            =   240
      Top             =   3000
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   873
      Caption         =   "Ne61u ba5n co1 o73 d9a6y va2 kho6ng muo61n ta81t ma1y ti1nh nu74a, ha4y nha61n va2o nu1t be6n d9u7o71i"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin UniControls.UniLabel UniLabel4 
      Height          =   495
      Left            =   3120
      Top             =   2280
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      Caption         =   "Gia6y"
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
   Begin UniControls.UniLabel Label1 
      Height          =   1215
      Left            =   1560
      Top             =   1680
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   2143
      Alignment       =   1
      Caption         =   "10"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   255
   End
   Begin UniControls.UniLabel UniLabel2 
      Height          =   375
      Left            =   120
      Top             =   1200
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   661
      Alignment       =   1
      Caption         =   "Ma1y ti1nh cu3a ba5n se4 tu75 d9o65ng ta81t"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin UniControls.UniLabel UniLabel1 
      Height          =   735
      Left            =   120
      Top             =   240
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   1296
      Alignment       =   1
      Caption         =   "Ba5n d9a4 ki1ch hoa5t che61 d9o65 ta81t ma1y sau khi que1t xong."
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   0
   End
End
Attribute VB_Name = "frmOffComputer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SetWindowPos& Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, _
ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
Private Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long

Dim xTime As Integer

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub Form_Load()
xTime = 10
SetForegroundWindow Me.hwnd
SetWindowPos Me.hwnd, -1, 0, 0, 0, 0, 3
End Sub

Private Sub Timer1_Timer()
xTime = xTime - 1
Label1.Caption = xTime
If xTime < 1 Then
    Timer1.Enabled = False
    cmdCancel.Enabled = False
    Timer3.Enabled = True
End If
End Sub

Private Sub Timer2_Timer()
SetForegroundWindow Me.hwnd
End Sub

Private Sub Timer3_Timer()
Shell "shutdown -s -t 00"
End Sub
