VERSION 5.00
Object = "{D1ECD258-D329-4388-AB83-DEC261A66B86}#1.0#0"; "UniControls_v2.0.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4350
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6810
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
   ScaleHeight     =   4350
   ScaleWidth      =   6810
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin UniControls.UniFrame F1 
      Height          =   3735
      Left            =   0
      Top             =   0
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   6588
      MaskColor       =   16711935
      Style           =   2
      Caption         =   "PAV 2009 - Ba3o ve65 tu75 d9o65ng"
      TextColor       =   16711680
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ThemeColor      =   4
      Begin UniControls.UniButton cmdNoRun 
         Height          =   735
         Left            =   2160
         TabIndex        =   2
         Top             =   2760
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   1296
         Icon            =   "Form1.frx":169B2
         Style           =   1
         Caption         =   "Kho6ng mo73 File na2y"
         IconAlign       =   5
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
      Begin UniControls.UniButton cmdNoRunAndKill 
         Height          =   735
         Left            =   4200
         TabIndex        =   1
         Top             =   2760
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   1296
         Icon            =   "Form1.frx":16F4C
         Style           =   1
         Caption         =   "Kho6ng mo73 + xo1a File na2y"
         IconAlign       =   5
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
      Begin UniControls.UniButton cmdRun 
         Height          =   735
         Left            =   120
         TabIndex        =   0
         Top             =   2760
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   1296
         Icon            =   "Form1.frx":174E6
         Style           =   1
         Caption         =   "Tie61p tu5c mo73 File na2y"
         IconAlign       =   5
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
      Begin UniControls.UniLabel UniLabel4 
         Height          =   375
         Left            =   120
         Top             =   2160
         Width           =   6015
         _ExtentX        =   10610
         _ExtentY        =   661
         Alignment       =   1
         BackStyle       =   0
         Caption         =   "Ba5n va64n muo61n cha5y File na2y chu71?"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   16711680
      End
      Begin UniControls.UniLabel lblFileName 
         Height          =   255
         Left            =   4320
         Top             =   1200
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   450
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
         ForeColor       =   16711680
      End
      Begin UniControls.UniLabel lblVirusName 
         Height          =   255
         Left            =   3720
         Top             =   1560
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   450
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
         ForeColor       =   16711680
      End
      Begin UniControls.UniLabel UniLabel3 
         Height          =   255
         Left            =   240
         Top             =   1560
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   450
         BackStyle       =   0
         Caption         =   "Chu7o7ng tri2nh pha1t hie65n File na2y la2 Virus:"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   192
      End
      Begin UniControls.UniLabel UniLabel2 
         Height          =   255
         Left            =   240
         Top             =   1200
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   450
         BackStyle       =   0
         Caption         =   "Ba5n (hoa85c 1 chu7o7ng tri2nh na2o d9o1) sa81p cha5y File:"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   192
      End
      Begin UniControls.UniLabel UniLabel1 
         Height          =   495
         Left            =   720
         Top             =   480
         Width           =   4815
         _ExtentX        =   8493
         _ExtentY        =   873
         Alignment       =   1
         BackStyle       =   0
         Caption         =   "Pha1t hie65n Virus !!!"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   255
         VisitedLinkColor=   255
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SetWindowPos& Lib "User32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, _
ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
Private Declare Function SetForegroundWindow Lib "User32" (ByVal hWnd As Long) As Long
Private Declare Function Beep Lib "kernel32" (ByVal dwFreq As Long, ByVal dwDuration As Long) As Long

Private Declare Function SendMessage Lib "User32" Alias "SendMessageA" (ByVal hWnd As Long, _
ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Sub ReleaseCapture Lib "User32" ()
Const WM_NCLBUTTONDOWN = &HA1
Const HTCAPTION = 2


Dim Comd As String
Dim o As String

Private Sub cmdNoRun_Click()
End
End Sub

Private Sub cmdNoRunAndKill_Click()
SetAttr o, vbNormal
DeleteFile o
End
End Sub

Private Sub cmdRun_Click()
If UniMsgBox("Ba5n cha81c cha81n?", vbYesNo, "Are You Sure?", Me.hWnd) = vbYes Then
    Shell Comd, vbNormalFocus
    End
End If
End Sub


Private Sub F1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 1 Then
   Call ReleaseCapture
   Call SendMessage(Me.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
   End If
End Sub

Private Sub Form_Load()
Me.Visible = False
App.TaskVisible = False
'////////////

Comd = Command()
'////////////
'Shell Comd, vbNormalFocus

If Comd = "" Then
UniMsgBox "File na2y d9u7o75c PAV 2009 su73 du5ng d9e4 kie63m tra Virus cho ca1c File tru7o71c khi chu1ng d9u7o75c mo73." & vbCrLf & "D9e63 ki1ch hoa5t chu71c na8ng na2y, ba5n va2o PAV 2009 -> Tu75 d9o65ng que1t -> Tu75 d9o65ng que1t va2 ca3nh ba1o Virus cho nhu74ng File sa81p d9u7o75c mo73.", vbOKOnly, "PAV 2009 - Tu75 d9o65ng ba3o ve65"
End
End If
modScanVirus.ConnectDB
Dim AX As String

o = Mid(Comd, 2, Len(Comd) - 2)
AX = modScanVirus.CheckVirus(o)

    If AX <> "No" Then
        '///////////// Cai dat cho Form truoc /////////////

        '---> Set Form
        Me.Height = F1.Height
        Me.Width = F1.Width
        Me.Show
        Beep 3000, 500
        SetWindowPos Me.hWnd, -1, 0, 0, 0, 0, 3
        '<--- Set Form
        '///////////// Cai dat cho Form truoc /////////////
        
        
        
        '////////// Tiep theo la nhap ten Virus vao /////////
        lblFileName.Caption = GetFileName(o)
        lblVirusName.Caption = AX
        '////////// Tiep theo la nhap ten Virus vao /////////
    Else
        Shell Comd, vbNormalFocus
        End
    End If
End Sub
