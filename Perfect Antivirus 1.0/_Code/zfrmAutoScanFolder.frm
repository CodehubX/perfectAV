VERSION 5.00
Object = "{D1ECD258-D329-4388-AB83-DEC261A66B86}#1.0#0"; "UniControls_v2.0.ocx"
Begin VB.Form zfrmAutoScanFolder 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PAV 2009 - Thong Bao !!!"
   ClientHeight    =   5295
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   6960
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "zfrmAutoScanFolder.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5295
   ScaleWidth      =   6960
   StartUpPosition =   2  'CenterScreen
   Begin UniControls.UniFrame F1 
      Height          =   4335
      Left            =   0
      Top             =   0
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   7646
      MaskColor       =   16711935
      Style           =   2
      Caption         =   "PAV 2009 - Tu75 d9o65ng ba3o ve65"
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
      Icon            =   "zfrmAutoScanFolder.frx":57E2
      ThemeColor      =   4
      Begin UniControls.UniLabel lblSta 
         Height          =   255
         Left            =   1800
         Top             =   3960
         Width           =   3375
         _ExtentX        =   5953
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
      Begin UniControls.UniButton UniButton2 
         Height          =   375
         Left            =   5280
         TabIndex        =   6
         Top             =   3840
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         Icon            =   "zfrmAutoScanFolder.frx":5D7C
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
      Begin UniControls.UniButton UniButton1 
         Height          =   375
         Left            =   240
         TabIndex        =   5
         Top             =   3840
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         Icon            =   "zfrmAutoScanFolder.frx":678E
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
      Begin UniControls.UniListView LVVirus3 
         Height          =   1935
         Left            =   120
         TabIndex        =   4
         Top             =   1680
         Width           =   6735
         _ExtentX        =   11880
         _ExtentY        =   3413
         AutoUnicode     =   0   'False
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
         MultiSelect     =   -1  'True
         LabelEdit       =   0   'False
         FullRowSelect   =   -1  'True
         AutoArrange     =   0   'False
         HeaderButtons   =   0   'False
         HeaderTrackSelect=   0   'False
         HideSelection   =   0   'False
         InfoTips        =   0   'False
      End
      Begin UniControls.UniTextBox txtPathFolder 
         Height          =   270
         Left            =   1680
         TabIndex        =   3
         Top             =   1200
         Width           =   5175
         _ExtentX        =   9128
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
      Begin UniControls.UniLabel UniLabel2 
         Height          =   255
         Left            =   120
         Top             =   1200
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   450
         BackStyle       =   0
         Caption         =   "Pha1t Hie65n Virus Ta5i:"
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
         Top             =   600
         Width           =   6735
         _ExtentX        =   11880
         _ExtentY        =   873
         Alignment       =   1
         BackStyle       =   0
         Caption         =   "Pha1t Hie65n Co1 Virus!"
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
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   240
      MultiLine       =   -1  'True
      TabIndex        =   2
      Text            =   "zfrmAutoScanFolder.frx":6D28
      Top             =   5040
      Width           =   615
   End
   Begin VB.ListBox List1 
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   4800
      Width           =   495
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   3000
      Top             =   4920
   End
   Begin VB.FileListBox File1 
      Height          =   285
      Left            =   240
      TabIndex        =   0
      Top             =   4560
      Width           =   855
   End
End
Attribute VB_Name = "zfrmAutoScanFolder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const WM_GETTEXT = &HD

Private Const WM_GETTEXTLENGTH = &HE

Private Declare Function SendMessage _
                Lib "user32" _
                Alias "SendMessageA" (ByVal hWnd As Long, _
                                      ByVal wMsg As Long, _
                                      ByVal wParam As Long, _
                                      lParam As Any) As Long

Private Declare Function GetClassName _
                Lib "user32" _
                Alias "GetClassNameA" (ByVal hWnd As Long, _
                                       ByVal lpClassName As String, _
                                       ByVal nMaxCount As Long) As Long

Private Declare Function BringWindowToTop Lib "user32" (ByVal hWnd As Long) As Long

Private Declare Function GetForegroundWindow Lib "user32" () As Long

Private Declare Function FindWindowEx _
                Lib "user32" _
                Alias "FindWindowExA" (ByVal hWnd1 As Long, _
                                       ByVal hWnd2 As Long, _
                                       ByVal lpsz1 As String, _
                                       ByVal lpsz2 As String) As Long

Private Declare Function FindWindow _
                Lib "user32" _
                Alias "FindWindowA" (ByVal lpClassName As String, _
                                     ByVal lpWindowName As String) As Long

Dim h1 As Long, h2 As Long, h3 As Long, h4 As Long, h5 As Long, h6 As Long

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

Private Sub F1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

        '<EhHeader>
        On Error GoTo F1_MouseMove_Err

        '</EhHeader>
100     If Button = 1 Then
102         Call ReleaseCapture
104         Call SendMessage(Me.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
        End If

        '<EhFooter>
        Exit Sub

F1_MouseMove_Err:
        WriteErr Time & "-" & Date & " - " & Err.Description & " - " & "in PerfectAntivirus2009.zfrmAutoScanFolder.F1_MouseMove " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Form_Load()

        '<EhHeader>
        On Error GoTo Form_Load_Err

        '</EhHeader>

        '---> Setting File1 dir box
100     File1.Pattern = "*.exe;*.bat;*.com;*.cmd;*.pif;*.dll"
102     File1.System = True
104     File1.ReadOnly = True
106     File1.Hidden = True
108     File1.Archive = True
        '<--- Setting File1 dir box

        '---> List Virus
110     LVVirus3.View = eViewDetails
112     LVVirus3.GridLines = True
114     LVVirus3.HeaderButtons = False
116     LVVirus3.CheckBoxes = True
118     LVVirus3.AutoUnicode = False
120     LVVirus3.Columns.Add , , "Virus Name", , 2000
122     LVVirus3.Columns.Add , , "Path", , 3500
124     LVVirus3.Columns.Add , , "Process", , 900
126     LVVirus3.Refresh
        '<--- Virus

        '<EhFooter>
        Exit Sub

Form_Load_Err:
        WriteErr Time & "-" & Date & " - " & Err.Description & " - " & "in PerfectAntivirus2009.zfrmAutoScanFolder.Form_Load " & "at line " & Erl

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
104             Timer1.Enabled = True
106             Me.Hide

108         Case 1
                'unload by code

        End Select

        '<EhFooter>
        Exit Sub

Form_QueryUnload_Err:
        WriteErr Time & "-" & Date & " - " & Err.Description & " - " & "in PerfectAntivirus2009.zfrmAutoScanFolder.Form_QueryUnload " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Text1_Change()

        '<EhHeader>
        On Error GoTo Text1_Change_Err

        '</EhHeader>

100     LVVirus3.ListItems.Clear

        Dim j

102     For j = 0 To List1.ListCount - 1

            'MsgBox List1.List(j)
104         If FileLen(List1.List(j)) < 10000000 Then ' filelen

                Dim AX As String

106             AX = modScanVirus.CheckVirus(List1.List(j))

108             If AX <> "No" Then
        
                    '///////////// Cai dat cho Form truoc /////////////
        
                    '---> Set Form
110                 Me.Show
112                 PLaySound AppPath & "Sound\Found.wav"
114                 BringWindowToTop Me.hWnd
116                 UniButton2.Caption = "Bo3 Qua"
118                 Timer1.Enabled = False
120                 lblSta.Caption = ""
                    '<--- Set Form
        
                    '///////////// Cai dat cho Form truoc /////////////
        
                    '////////// Tiep theo la nhap ten Virus vao /////////
        
122                 txtPathFolder.Text = ""
124                 txtPathFolder.Text = GetFolderPath(List1.List(j))

                    Dim I

126                 I = LVVirus3.ListItems.Count + 1
128                 LVVirus3.ListItems.Add I, , AX
130                 LVVirus3.ListItems(I).SubItems(1).Caption = List1.List(j)
132                 LVVirus3.ListItems(I).SubItems(2).Caption = basProcess.CheckProcess(List1.List(j))
134                 LVVirus3.ListItems(I).Checked = True
                    '////////// Tiep theo la nhap ten Virus vao /////////
                End If
            End If 'filelen

136     Next j

        '<EhFooter>
        Exit Sub

Text1_Change_Err:
        WriteErr Time & "-" & Date & " - " & Err.Description & " - " & "in PerfectAntivirus2009.zfrmAutoScanFolder.Text1_Change " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Timer1_Timer()

    '<EhHeader>
    On Error Resume Next

    '</EhHeader>

    h6 = 0
    h1 = GetForegroundWindow
    h2 = FindWindowEx(h1, ByVal 0&, "WorkerW", vbNullString)
    h3 = FindWindowEx(h2, ByVal 0&, "ReBarWindow32", vbNullString)
    h4 = FindWindowEx(h3, ByVal 0&, "ComboBoxEx32", vbNullString)
    h5 = FindWindowEx(h4, ByVal 0&, "ComboBox", vbNullString)
    h6 = FindWindowEx(h5, ByVal 0&, "Edit", vbNullString)

    'Text1.Text = h6  'L?y handle c?a d?i tu?ng

    'L?y n?i dung hi?n th? c?a d?i tu?ng
    Dim Length As Long

    Dim result As Long

    Dim strtmp As String

    Length = SendMessage(h6, WM_GETTEXTLENGTH, ByVal 0, ByVal 0) + 1
    strtmp = Space(Length)
    result = SendMessage(h6, WM_GETTEXT, ByVal Length, ByVal strtmp)

    Dim s  As Variant

    Dim st As String

    s = Split(strtmp, vbNullChar)

    If Right(s(0), 1) <> "\" Then s(0) = s(0) & "\"

    List1.Clear

    Dim a As String

    On Error Resume Next

    File1.Path = s(0)
    File1.Refresh

    If File1.ListCount < 100 Then

        Dim h

        For h = 0 To File1.ListCount - 1

            If FileExists(s(0) & File1.List(h)) = True Then
                List1.AddItem s(0) & File1.List(h)
                a = a & s(0) & File1.List(h) & vbCrLf
            End If

        Next h

        Text1.Text = a
    End If

End Sub

Private Sub UniButton1_Click()

        '<EhHeader>
        On Error GoTo UniButton1_Click_Err

        '</EhHeader>

        'If UniMsgBox("Ba5n co1 cha81c cha81n muo61n xo1a ca1c Virus d9a4 d9a1nh da61u kho6ng?", vbYesNo, "Die65t Virus") = vbYes Then
        Dim Y

        Dim X As String

100     For Y = 1 To LVVirus3.ListItems.Count

102         If LVVirus3.ListItems(Y).Checked = True And FileExists(LVVirus3.ListItems(Y).SubItems(1).Caption) = True Then

104             DoEvents

106             If LVVirus3.ListItems(Y).SubItems(2).Caption <> "0" Then
                    'Kill process
108                 KillProcessById (LVVirus3.ListItems(Y).SubItems(2).Caption)
110                 X = X & " D9a4 ta81t tie61n tri2nh: " & LVVirus3.ListItems(Y).SubItems(2).Caption & vbCrLf
                End If
        
112             DoEvents
114             SetAttr LVVirus3.ListItems(Y).SubItems(1).Caption, vbNormal
116             modScanVirus.DeleteFile LVVirus3.ListItems(Y).SubItems(1).Caption
118             X = X & " D9a4 Xo1a Bo3: " & LVVirus3.ListItems(Y).SubItems(1).Caption & vbCrLf
        
            End If

120     Next Y

122     If Not X = "" Then
124         DelAllChecked LVVirus3
            'UniMsgBox X, vbOKOnly, "Nhu74ng vie65c d9a4 la2m.", Me.hWnd
126         lblSta.Caption = "D9a4 die65t xong ca1c Virus d9a4 d9a1nh d9a61u."
        Else
128         lblSta.Caption = "Kho6ng co1 Virus d9e63 die65t."
            'UniMsgBox "Kho6ng co1 Virus na2o d9e63 die65t!", vbOKOnly, "Tho6ng Ba1o", Me.hWnd
        End If

130     UniButton2.Caption = "D9o1ng"
        'End If ' Unimsgbox "ban co chac chan ko?"

        '<EhFooter>
        Exit Sub

UniButton1_Click_Err:
        WriteErr Time & "-" & Date & " - " & Err.Description & " - " & "in PerfectAntivirus2009.zfrmAutoScanFolder.UniButton1_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub UniButton2_Click()

        '<EhHeader>
        On Error GoTo UniButton2_Click_Err

        '</EhHeader>

100     Timer1.Enabled = True
102     Me.Hide

        '<EhFooter>
        Exit Sub

UniButton2_Click_Err:
        WriteErr Time & "-" & Date & " - " & Err.Description & " - " & "in PerfectAntivirus2009.zfrmAutoScanFolder.UniButton2_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

