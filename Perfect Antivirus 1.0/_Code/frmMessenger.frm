VERSION 5.00
Object = "{D1ECD258-D329-4388-AB83-DEC261A66B86}#1.0#0"; "UniControls_v2.0.ocx"
Begin VB.Form frmMessenger 
   BackColor       =   &H0080FFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4800
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7200
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMessenger.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4800
   ScaleWidth      =   7200
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Left            =   3480
      Top             =   4080
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   2760
      Top             =   3600
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   3960
      Top             =   3600
   End
   Begin UniControls.UniFrame F1 
      Height          =   3255
      Left            =   0
      Top             =   0
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   5741
      MaskColor       =   16711935
      Caption         =   "PAV 2009 - Tho6ng Ba1o"
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
      IconSize        =   48
      ThemeColor      =   4
      Begin UniControls.UniButton cmdCloseMessenger 
         Height          =   375
         Left            =   4200
         TabIndex        =   0
         Top             =   0
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         Icon            =   "frmMessenger.frx":000C
         Style           =   2
         IconAlign       =   2
         iNonThemeStyle  =   2
         BackColor       =   -2147483633
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
      Begin UniControls.UniLabel lblText 
         Height          =   975
         Left            =   120
         Top             =   1080
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   1720
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
      End
      Begin UniControls.UniLabel lblTitle 
         Height          =   375
         Left            =   120
         Top             =   600
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   661
         Alignment       =   1
         BackStyle       =   0
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   255
      End
      Begin VB.Image PicI 
         Height          =   480
         Index           =   2
         Left            =   120
         Top             =   480
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image PicI 
         Height          =   480
         Index           =   1
         Left            =   120
         Picture         =   "frmMessenger.frx":05A6
         Top             =   480
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image PicI 
         Height          =   480
         Index           =   0
         Left            =   120
         Top             =   480
         Visible         =   0   'False
         Width           =   480
      End
   End
End
Attribute VB_Name = "frmMessenger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Declare Function SetWindowPos& _
                Lib "user32" (ByVal hWnd As Long, _
                              ByVal hWndInsertAfter As Long, _
                              ByVal X As Long, _
                              ByVal Y As Long, _
                              ByVal cx As Long, _
                              ByVal cy As Long, _
                              ByVal wFlags As Long)

Private Declare Function SetForegroundWindow Lib "user32" (ByVal hWnd As Long) As Long

Private Declare Function SendMessage _
                Lib "user32" _
                Alias "SendMessageA" (ByVal hWnd As Long, _
                                      ByVal wMsg As Long, _
                                      ByVal wParam As Long, _
                                      lParam As Any) As Long

Private Declare Sub ReleaseCapture Lib "user32" ()

Const WM_NCLBUTTONDOWN = &HA1

Const HTCAPTION = 2

Public Enum xIcon

    xTrang = 0
    xvang = 1
    xdo = 2

End Enum

Dim sNotClick As Boolean

Public Sub zShowMessenger(xTitle, xText, xTime, sIcon As xIcon)

        '<EhHeader>
        On Error GoTo zShowMessenger_Err

        '</EhHeader>

        Dim a As New frmMessenger

100     With a
102         PLaySound AppPath & "Sound\Mes.wav"
104         App.TaskVisible = False
106         sNotClick = False
108         .lblTitle.Caption = xTitle
110         .lblText.Caption = xText
112         .lblText.Height = 195 + TextHeight(.lblText.Caption) * (Len(.lblText.Caption) / (.lblText.Width / TextHeight(.lblText.Caption)) / 2)
114         .F1.Height = .lblText.Top + .lblText.Height + 200
116         .Height = 0
118         .Width = F1.Width
120         .Top = Screen.Height - frmMain.xStart - 450
122         frmMain.xStart = frmMain.xStart + .F1.Height
124         .Left = Screen.Width - .Width
126         .Timer2.Interval = xTime
128         .Timer1.Enabled = True
130         .Show
132         .PicI(sIcon).Visible = True
   
134         SetWindowPos .hWnd, -1, 0, 0, 0, 0, 3
        End With

        '<EhFooter>
        Exit Sub

zShowMessenger_Err:
        WriteErr Time & "-" & Date & " - " & Err.Description & " - " & "in PerfectAntivirus2009.frmMessenger.zShowMessenger " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub cmdCloseMessenger_Click()

        '<EhHeader>
        On Error GoTo cmdCloseMessenger_Click_Err

        '</EhHeader>

100     sNotClick = True
102     Timer3.Enabled = True
104     Timer2.Enabled = False

        '<EhFooter>
        Exit Sub

cmdCloseMessenger_Click_Err:
        WriteErr Time & "-" & Date & " - " & Err.Description & " - " & "in PerfectAntivirus2009.frmMessenger.cmdCloseMessenger_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub F1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

        '<EhHeader>
        On Error GoTo F1_MouseMove_Err

        '</EhHeader>
100     If Timer2.Enabled = False And sNotClick = False Then
102         Timer3.Enabled = False
104         Timer1.Enabled = True
        End If

106     If Button = 1 Then
108         Call ReleaseCapture
110         Call SendMessage(Me.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
        End If

        '<EhFooter>
        Exit Sub

F1_MouseMove_Err:
        WriteErr Time & "-" & Date & " - " & Err.Description & " - " & "in PerfectAntivirus2009.frmMessenger.F1_MouseMove " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Timer1_Timer()

    '<EhHeader>
    On Error Resume Next

    '</EhHeader>
    With Me
        .Height = .Height + 50
        .Top = .Top - 50

        If .Height > .F1.Height Then
            Timer1.Enabled = False
            Timer2.Enabled = True
        End If

    End With

End Sub

Private Sub Timer2_Timer()

    '<EhHeader>
    On Error Resume Next

    '</EhHeader>

    Timer3.Enabled = True
    Timer2.Enabled = False
End Sub

Private Sub Timer3_Timer()

    '<EhHeader>
    On Error Resume Next

    '</EhHeader>
    With Me
        .Height = .Height - 50
        .Top = .Top + 50

        If .Height < 60 Then
            frmMain.xStart = frmMain.xStart - F1.Height
            Unload Me
            Timer3.Enabled = False
        End If

    End With

End Sub

