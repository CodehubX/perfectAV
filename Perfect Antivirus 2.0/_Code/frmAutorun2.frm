VERSION 5.00
Object = "{6DC1AB90-DCD1-47A1-AB36-924FFD67ADBF}#1.0#0"; "UniControls_v2.0.ocx"
Begin VB.Form frmAutorun2 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Perfect Antivirus 2009"
   ClientHeight    =   1830
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3270
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAutorun2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1830
   ScaleWidth      =   3270
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer7 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   720
      Top             =   480
   End
   Begin VB.Timer Timer6 
      Interval        =   1
      Left            =   240
      Top             =   840
   End
   Begin VB.Timer Timer5 
      Interval        =   500
      Left            =   2520
      Top             =   1200
   End
   Begin VB.Timer Timer4 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   2760
      Top             =   240
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   2880
      Top             =   120
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   2640
      Top             =   0
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   2880
      Top             =   0
   End
   Begin UniControls.UniLabel LBL1 
      Height          =   255
      Left            =   120
      Top             =   1440
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   450
      BackStyle       =   0
      Caption         =   "D9ang kie63m tra..."
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
   Begin UniControls.UniLabel UniLabel1 
      Height          =   495
      Left            =   240
      Top             =   120
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   873
      BackStyle       =   0
      Caption         =   "Pha1t hie65n co1 thie61t bi5 truy ca65p va2o ma1y ti1nh..."
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
   Begin VB.Image KKK 
      Height          =   720
      Left            =   1320
      Picture         =   "frmAutorun2.frx":058A
      Top             =   600
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Image IMA 
      Height          =   720
      Left            =   1440
      Picture         =   "frmAutorun2.frx":1454
      Top             =   720
      Width           =   720
   End
   Begin VB.Image Oma 
      Height          =   720
      Left            =   1080
      Picture         =   "frmAutorun2.frx":7CA6
      Top             =   360
      Width           =   720
   End
End
Attribute VB_Name = "frmAutorun2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SetWindowPos& Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, _
ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
Private Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long

Dim xGiaToc As Integer
Dim xDem As Integer


Private Sub Form_Load()
On Error Resume Next
SetForegroundWindow Me.hwnd
SetWindowPos Me.hwnd, -1, 0, 0, 0, 0, 3
Me.Left = Screen.Width - Me.Width
Me.Top = Screen.Height - 400
xGiaToc = 200
xDem = 0
End Sub

Private Sub Timer1_Timer()
On Error Resume Next
IMA.Top = IMA.Top - 30
If IMA.Top < Oma.Top Then
    Timer1.Enabled = False
    Timer2.Enabled = True
End If
End Sub

Private Sub Timer2_Timer()
On Error Resume Next
IMA.Left = IMA.Left - 30
If IMA.Left < Oma.Left - Oma.Width / 3 Then
    Timer2.Enabled = False
    Timer3.Enabled = True
End If
End Sub

Private Sub Timer3_Timer()
On Error Resume Next
IMA.Top = IMA.Top + 30
If IMA.Top > Oma.Top + Oma.Height / 2 Then
    Timer3.Enabled = False
    Timer4.Enabled = True
End If
End Sub
Private Sub Timer4_Timer()
On Error Resume Next
IMA.Left = IMA.Left + 30
If IMA.Left > Oma.Left + Oma.Width / 2 Then
    Timer4.Enabled = False
    Timer1.Enabled = True
End If
End Sub

Private Sub Timer5_Timer()
On Error Resume Next
xDem = xDem + 1
If xDem = 6 Then
IMA.Visible = False
Oma.Visible = False
KKK.Visible = True
LBL1.Caption = "Kie63m tra xong!"
End If
If xDem = 7 Then
xGiaToc = 200
Timer7.Enabled = True
End If
End Sub

Private Sub Timer6_Timer()
On Error Resume Next
xGiaToc = xGiaToc - 8
Me.Top = Me.Top - xGiaToc
If Me.Top < Screen.Height - Me.Height - 400 Then
If Me.Top <> Screen.Height - Me.Height - 400 Then Me.Top = Screen.Height - Me.Height - 400
Timer6.Enabled = False
End If
'Me.Left = Screen.Width - Me.Width
'Me.Top = Screen.Height - Me.Height - 400
End Sub

Private Sub Timer7_Timer()
On Error Resume Next
xGiaToc = xGiaToc - 5
Me.Top = Me.Top + xGiaToc
If Me.Top > Screen.Height Then
Unload Me
End If
End Sub
