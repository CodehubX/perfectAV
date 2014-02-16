VERSION 5.00
Begin VB.Form frmFlash 
   BackColor       =   &H00FF0000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   1575
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7455
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
   ScaleHeight     =   1575
   ScaleWidth      =   7455
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer2 
      Left            =   2880
      Top             =   2520
   End
   Begin VB.Timer Timer1 
      Left            =   1080
      Top             =   2520
   End
   Begin VB.Image I 
      Height          =   1335
      Left            =   120
      Picture         =   "frmFlash.frx":0000
      Top             =   120
      Width           =   7215
   End
End
Attribute VB_Name = "frmFlash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Declare Function GetWindowLong _
                Lib "user32" _
                Alias "GetWindowLongA" (ByVal hWnd As Long, _
                                        ByVal nIndex As Long) As Long

Private Declare Function SetWindowLong _
                Lib "user32" _
                Alias "SetWindowLongA" (ByVal hWnd As Long, _
                                        ByVal nIndex As Long, _
                                        ByVal dwNewLong As Long) As Long

Private Declare Function SetLayeredWindowAttributes _
                Lib "user32" (ByVal hWnd As Long, _
                              ByVal crKey As Long, _
                              ByVal bAlpha As Byte, _
                              ByVal dwFlags As Long) As Long

Private Const GWL_EXSTYLE = (-20)

Private Const LWA_ALPHA = &H2

Private Const WS_EX_LAYERED = &H80000

Dim m_lAlpha

Private Sub Form_Load()

        '<EhHeader>
        On Error GoTo Form_Load_Err

        '</EhHeader>

100     App.TaskVisible = False
102     I.Top = 40
104     I.Left = 40
106     Me.Height = I.Height + 80
108     Me.Width = I.Width + 80

        Dim lStyle As Long

110     lStyle = GetWindowLong(Me.hWnd, GWL_EXSTYLE)
112     lStyle = lStyle Or WS_EX_LAYERED
114     SetWindowLong Me.hWnd, GWL_EXSTYLE, lStyle
116     SetLayeredWindowAttributes Me.hWnd, 0, 0, LWA_ALPHA
118     Timer1.Interval = 50
120     Timer2.Interval = 50
122     Timer2.Enabled = False
124     Timer1.Enabled = True

        '<EhFooter>
        Exit Sub

Form_Load_Err:
        WriteErr Time & "-" & Date & " - " & Err.Description & " - " & "in PerfectAntivirus2009.frmFlash.Form_Load " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Timer1_Timer()

    '<EhHeader>
    On Error Resume Next

    '</EhHeader>

    m_lAlpha = m_lAlpha + 15

    If (m_lAlpha > 255) Then
        m_lAlpha = 255
        Timer1.Enabled = False
        frmMain.xTask = False
        Load frmMain
    Else
        SetLayeredWindowAttributes Me.hWnd, 0, m_lAlpha, LWA_ALPHA
    End If

End Sub

Private Sub Timer2_Timer()

    '<EhHeader>
    On Error Resume Next

    '</EhHeader>

    m_lAlpha = m_lAlpha - 15

    If (m_lAlpha < 0) Then
        m_lAlpha = 0
        Unload Me
        frmMain.Show
    Else
        SetLayeredWindowAttributes Me.hWnd, 0, m_lAlpha, LWA_ALPHA
    End If

End Sub

