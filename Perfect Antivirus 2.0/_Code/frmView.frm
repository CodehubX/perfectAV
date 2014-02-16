VERSION 5.00
Object = "{E8FDD05C-3067-4198-8AEC-1A013A46ABDD}#1.0#0"; "FVUnicodeControl.ocx"
Begin VB.Form frmView 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3225
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4650
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   3225
   ScaleWidth      =   4650
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   3840
      Top             =   2280
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   3000
      Top             =   2040
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   3360
      Top             =   1440
   End
   Begin FVUnicodeControl.FVistaUniButton Mai 
      Height          =   375
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   661
      BackColor       =   12632319
      ButtonStyle     =   2
      Caption         =   "Hie63n Thi5"
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
   Begin FVUnicodeControl.FVistaUniButton Mai 
      Height          =   375
      Index           =   6
      Left            =   0
      TabIndex        =   1
      Top             =   2430
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   661
      BackColor       =   14737632
      ButtonStyle     =   2
      Caption         =   "Ca2i D9a85t Chung"
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
   Begin FVUnicodeControl.FVistaUniButton Mai 
      Height          =   375
      Index           =   1
      Left            =   0
      TabIndex        =   2
      Top             =   400
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   661
      BackColor       =   14737632
      ButtonStyle     =   2
      Caption         =   "Que1t Virus"
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
   Begin FVUnicodeControl.FVistaUniButton Mai 
      Height          =   375
      Index           =   2
      Left            =   0
      TabIndex        =   3
      Top             =   810
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   661
      BackColor       =   14737632
      ButtonStyle     =   2
      Caption         =   "Ca61u Hi2nh Que1t"
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
   Begin FVUnicodeControl.FVistaUniButton Mai 
      Height          =   375
      Index           =   3
      Left            =   0
      TabIndex        =   4
      Top             =   1210
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   661
      BackColor       =   14737632
      ButtonStyle     =   2
      Caption         =   "Ba3o Ve65 Ma1y Ti1nh"
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
   Begin FVUnicodeControl.FVistaUniButton Mai 
      Height          =   375
      Index           =   4
      Left            =   0
      TabIndex        =   5
      Top             =   1620
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   661
      BackColor       =   14737632
      ButtonStyle     =   2
      Caption         =   "To61i U7u He65 Tho61ng"
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
   Begin FVUnicodeControl.FVistaUniButton Mai 
      Height          =   375
      Index           =   5
      Left            =   0
      TabIndex        =   6
      Top             =   2030
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   661
      BackColor       =   14737632
      ButtonStyle     =   2
      Caption         =   "Co6ng Cu5 Ho64 Tro75"
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
   Begin FVUnicodeControl.FVistaUniButton Mai 
      Height          =   375
      Index           =   7
      Left            =   0
      TabIndex        =   7
      Top             =   2830
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   661
      BackColor       =   8421504
      ButtonStyle     =   2
      Caption         =   "Thoa1t"
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
Attribute VB_Name = "frmView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim xGiaToc As Integer
Private Sub Form_Load()
xGiaToc = 30
Me.Left = Screen.Width
Me.Top = Screen.Height - Me.Height - 470
Me.Height = 3225
Me.Width = Mai(0).Width
End Sub


Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Timer2.Enabled = False
Timer3.Enabled = False
End Sub

Private Sub Mai_Click(Index As Integer)
If Index = 0 Then

    If frmMain.xScanning = True Then
        frmScan.Show
        frmScan.Timer6.Enabled = False
    Else
        frmMain.Show
        App.TaskVisible = True
    End If

ElseIf Index = 7 Then
    frmMain.xThoat = 0
    Unload frmMain
    Unload frmRTP
    Unload frmScan
    Unload frmAutorun
    Unload frmTangToc
    Unload frmRegistry
    Unload frmPhucHoiDuLieu
    Unload frmFOLDERSCAN
    Unload frmMenu
    Unload frmProtect
    Unload frmOffComputer
    Unload Me
Else
    If frmMain.xScanning = True Then
        frmScan.Show
        frmScan.Timer6.Enabled = False
    Else
        frmMain.Show
        App.TaskVisible = True
    End If
    frmMain.ShowMenu Index
End If
End Sub

Private Sub Mai_MouseEnter(Index As Integer)
Timer2.Enabled = False
Timer3.Enabled = False
End Sub

Private Sub Mai_MouseLeave(Index As Integer)
If Timer1.Enabled = False Then
Timer2.Enabled = True
End If
End Sub

Private Sub Timer1_Timer()
xGiaToc = xGiaToc - 1
Me.Left = Me.Left - 5 * xGiaToc
If Me.Left < Screen.Width - Me.Width Then
    Timer1.Enabled = False
    Me.Left = Screen.Width - Me.Width
    Timer2.Enabled = True
End If
End Sub

Private Sub Timer2_Timer()
xGiaToc = 0
Timer3.Enabled = True
Timer2.Enabled = False
End Sub

Private Sub Timer3_Timer()
xGiaToc = xGiaToc + 1
Me.Left = Me.Left + 5 * xGiaToc
If Me.Left > Screen.Width Then
    Timer3.Enabled = False
    Unload Me
End If
End Sub
