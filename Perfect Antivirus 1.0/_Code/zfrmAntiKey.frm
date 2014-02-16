VERSION 5.00
Object = "{D1ECD258-D329-4388-AB83-DEC261A66B86}#1.0#0"; "UniControls_v2.0.ocx"
Begin VB.Form zfrmAntiKey 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PAV 2009 - Thong Bao !!!"
   ClientHeight    =   3615
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   4800
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "zfrmAntiKey.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3615
   ScaleWidth      =   4800
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox lstBQ 
      Height          =   1815
      Left            =   5160
      TabIndex        =   6
      Top             =   1680
      Width           =   1335
   End
   Begin VB.Timer tmrScan 
      Interval        =   5000
      Left            =   4080
      Top             =   480
   End
   Begin UniControls.UniFrame F1 
      Height          =   3615
      Left            =   0
      Top             =   0
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   6376
      MaskColor       =   16711935
      Style           =   2
      Caption         =   "PAV 2009 - Ca3nh ba1o Virus"
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
      Icon            =   "zfrmAntiKey.frx":57E2
      IconSize        =   32
      ThemeColor      =   4
      Begin UniControls.UniCheckBox chkBoQua 
         Height          =   195
         Left            =   240
         TabIndex        =   5
         Top             =   2880
         Width           =   3630
         _ExtentX        =   6403
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
         Caption         =   "Kho6ng bao gio72 ca3nh ba1o chu7o7ng tri2nh na2y nu74a"
         BackStyle       =   0
         ForeColor       =   0
         ShowFocusRectangle=   0   'False
      End
      Begin UniControls.UniLabel lblSta 
         Height          =   255
         Left            =   1560
         Top             =   3240
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   450
         Alignment       =   1
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
      End
      Begin UniControls.UniLabel UniLabel4 
         Height          =   255
         Left            =   240
         Top             =   2280
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   450
         BackStyle       =   0
         Caption         =   "Chi3 so61 tie61n tri2nh:"
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
      Begin UniControls.UniTextBox txtPID 
         Height          =   270
         Left            =   240
         TabIndex        =   4
         Top             =   2520
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   476
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   0
         BackColor       =   12648447
         Text            =   ""
         Locked          =   -1  'True
         BorderStyle     =   2
      End
      Begin UniControls.UniButton cmdClose 
         Height          =   375
         Left            =   3360
         TabIndex        =   3
         Top             =   3120
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         Icon            =   "zfrmAntiKey.frx":5D74
         Style           =   2
         Caption         =   "Bo3 Qua"
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
      Begin UniControls.UniButton cmdKill 
         Height          =   375
         Left            =   240
         TabIndex        =   2
         Top             =   3120
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         Icon            =   "zfrmAntiKey.frx":6786
         Style           =   2
         Caption         =   "Die65t"
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
      Begin UniControls.UniTextBox txtName 
         Height          =   270
         Left            =   240
         TabIndex        =   1
         Top             =   1920
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   476
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
         Text            =   ""
         Locked          =   -1  'True
         BorderStyle     =   2
      End
      Begin UniControls.UniLabel UniLabel3 
         Height          =   255
         Left            =   240
         Top             =   1680
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   450
         BackStyle       =   0
         Caption         =   "Te6n Virus:"
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
      Begin UniControls.UniTextBox txtPath 
         Height          =   270
         Left            =   240
         TabIndex        =   0
         Top             =   1320
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   476
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
         BackColor       =   12648447
         Text            =   ""
         Locked          =   -1  'True
         BorderStyle     =   2
      End
      Begin UniControls.UniLabel UniLabel2 
         Height          =   255
         Left            =   240
         Top             =   1080
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   450
         BackStyle       =   0
         Caption         =   "Chu7o7ng tri2nh pha1t hie65n Virus d9ang cha5y ta5i:"
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
         Height          =   495
         Left            =   120
         Top             =   480
         Width           =   4575
         _ExtentX        =   8070
         _ExtentY        =   873
         Alignment       =   1
         BackStyle       =   0
         Caption         =   "Pha1t hie65n Virus!"
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
   End
End
Attribute VB_Name = "zfrmAntiKey"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Declare Function SendMessage _
                Lib "user32" _
                Alias "SendMessageA" (ByVal hWnd As Long, _
                                      ByVal wMsg As Long, _
                                      ByVal wParam As Long, _
                                      lParam As Any) As Long

Private Declare Function BringWindowToTop Lib "user32" (ByVal hWnd As Long) As Long

Private Declare Function SetWindowPos& _
                Lib "user32" (ByVal hWnd As Long, _
                              ByVal hWndInsertAfter As Long, _
                              ByVal X As Long, _
                              ByVal Y As Long, _
                              ByVal cx As Long, _
                              ByVal cy As Long, _
                              ByVal wFlags As Long)

Private Declare Function SetForegroundWindow Lib "user32" (ByVal hWnd As Long) As Long

Private Declare Sub ReleaseCapture Lib "user32" ()

Const WM_NCLBUTTONDOWN = &HA1

Const HTCAPTION = 2

Private Sub cmdClose_Click()

        '<EhHeader>
        On Error GoTo cmdClose_Click_Err

        '</EhHeader>

100     basProcess.SuspendResumeProcess Me.txtPID.Text, False

102     If Me.chkBoQua.Value = True Then lstBQ.AddItem Me.txtPID.Text
104     tmrScan.Enabled = True
106     Me.Hide

        '<EhFooter>
        Exit Sub

cmdClose_Click_Err:
        WriteErr Time & "-" & Date & " - " & Err.Description & " - " & "in PerfectAntivirus2009.zfrmAntiKey.cmdClose_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub cmdKill_Click()

        '<EhHeader>
        On Error GoTo cmdKill_Click_Err

        '</EhHeader>

100     KillProcessById Me.txtPID.Text
102     Shell "taskkill /pid " & Me.txtPID.Text, vbHide

104     If FileExists(Me.txtPath.Text) = True Then
106         SetAttr Me.txtPath.Text, vbNormal
108         modScanVirus.DeleteFile Me.txtPath.Text
        End If

110     Shell "explorer.exe " & GetFolderPath(Me.txtPath.Text), vbNormalFocus
112     zfrmAutoScanFolder.Timer1.Enabled = False
114     zfrmAutoScanFolder.Timer1.Enabled = True

116     If FileExists(Me.txtPath.Text) = False Then lblSta.Caption = "D9a4 die65t xong!" Else lblSta.Caption = "Chu7a die65t d9u7o75c!"
118     cmdClose.Caption = "D9o1ng"
120     tmrScan.Enabled = True
122     Me.Hide

        '<EhFooter>
        Exit Sub

cmdKill_Click_Err:
        WriteErr Time & "-" & Date & " - " & Err.Description & " - " & "in PerfectAntivirus2009.zfrmAntiKey.cmdKill_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Form_Load()

        '<EhHeader>
        On Error GoTo Form_Load_Err

        '</EhHeader>

100     App.TaskVisible = False

        '<EhFooter>
        Exit Sub

Form_Load_Err:
        WriteErr Time & "-" & Date & " - " & Err.Description & " - " & "in PerfectAntivirus2009.zfrmAntiKey.Form_Load " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

        '<EhHeader>
        On Error GoTo Form_QueryUnload_Err

        '</EhHeader>
100     Select Case UnloadMode

            Case 0
                'unload form X button
102             Cancel = True
104             basProcess.SuspendResumeProcess Me.txtPID.Text, False
106             tmrScan.Enabled = True
108             Me.Hide

110         Case 1
                'unload by code

        End Select

        '<EhFooter>
        Exit Sub

Form_QueryUnload_Err:
        WriteErr Time & "-" & Date & " - " & Err.Description & " - " & "in PerfectAntivirus2009.zfrmAntiKey.Form_QueryUnload " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub tmrScan_Timer()

    'On Error Resume Next
    '<EhHeader>
    On Error Resume Next

    '</EhHeader>
    Dim ColItems

    Dim ObjItem

    Set ColItems = GetObject("winmgmts:\root\CIMV2").ExecQuery("SELECT * FROM Win32_Process")

    For Each ObjItem In ColItems

        Dim ih

        For ih = 0 To lstBQ.ListCount - 1

            If ObjItem.ProcessID = lstBQ.List(ih) Then GoTo DaDaNhDaUbOqUa
        Next ih

        If ObjItem.Caption <> "System Idle Process" And ObjItem.Caption <> "System" Then

            'frmMain.lblStatus.Caption = ObjItem.ExecutablePath
            If FileExists(ObjItem.ExecutablePath) = True Then
                If FileLen(ObjItem.ExecutablePath) < 1000000 Then

                    Dim AX As String

                    AX = CheckVirus(ObjItem.ExecutablePath)

                    If AX <> "No" Then
                        Me.tmrScan.Enabled = False
                        App.TaskVisible = False
                        Me.Show
                        PLaySound AppPath & "Sound\Ring.wav"
                        BringWindowToTop Me.hWnd
                        Me.txtName.Text = AX
                        Me.txtPath.Text = ObjItem.ExecutablePath
                        Me.txtPID.Text = ObjItem.ProcessID
                        Me.lblSta.Caption = ""
                        basProcess.SuspendResumeProcess Me.txtPID.Text, True
                        Me.chkBoQua.Value = False
                        GoTo ThOaTkHoIfOr
                    End If
                End If
            End If
        End If 'ObjItem.Caption <> "System Idle Process" And ObjItem.Caption <> "System"

DaDaNhDaUbOqUa:
    Next

ThOaTkHoIfOr:
    Set ColItems = Nothing
    Set ObjItem = Nothing

End Sub
