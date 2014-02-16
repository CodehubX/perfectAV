VERSION 5.00
Object = "{E8FDD05C-3067-4198-8AEC-1A013A46ABDD}#1.0#0"; "FVUnicodeControl.ocx"
Object = "{6DC1AB90-DCD1-47A1-AB36-924FFD67ADBF}#1.0#0"; "UniControls_v2.0.ocx"
Begin VB.Form frmRTP 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PAV - Real Time Protection"
   ClientHeight    =   2190
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6015
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
   Icon            =   "frmRTP.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmRTP.frx":058A
   ScaleHeight     =   2190
   ScaleWidth      =   6015
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrRTP 
      Interval        =   1000
      Left            =   960
      Top             =   3240
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1440
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   3360
      Width           =   3975
   End
   Begin VB.FileListBox File1 
      Height          =   285
      Hidden          =   -1  'True
      Left            =   120
      System          =   -1  'True
      TabIndex        =   4
      Top             =   120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   1560
      Top             =   240
   End
   Begin UniControls.UniTextBox txtVirusName 
      Height          =   270
      Left            =   2160
      TabIndex        =   3
      Top             =   720
      Width           =   3735
      _ExtentX        =   6588
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
      Text            =   "Virus.User.123.exe"
      Locked          =   -1  'True
      BorderStyle     =   2
   End
   Begin UniControls.UniLabel UniLabel2 
      Height          =   255
      Left            =   1800
      Top             =   1320
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   450
      BackStyle       =   0
      Caption         =   "Ba5n co1 muo61n die65t Virus na2y ngay ba6y gio72 kho6ng?"
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
   Begin FVUnicodeControl.FVistaUniButton cmdBack 
      Height          =   375
      Left            =   5160
      TabIndex        =   2
      Top             =   1680
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   661
      BackColor       =   14737632
      ButtonStyle     =   3
      Caption         =   "Kho6ng"
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
   Begin FVUnicodeControl.FVistaUniButton cmdKill 
      Height          =   375
      Left            =   2280
      TabIndex        =   1
      Top             =   1680
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   661
      BackColor       =   14737632
      ButtonStyle     =   3
      Caption         =   "Co1"
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
      Left            =   2160
      TabIndex        =   0
      Top             =   960
      Width           =   3735
      _ExtentX        =   6588
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
      Text            =   "C:\WINDOWS\system32\userinit.exe"
      Locked          =   -1  'True
      BorderStyle     =   2
   End
   Begin UniControls.UniLabel UniLabel1 
      Height          =   495
      Left            =   2160
      Top             =   120
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   873
      Alignment       =   1
      Caption         =   "Phát hie65n Virus ta5i"
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
Attribute VB_Name = "frmRTP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Const WM_GETTEXT = &HD
Private Const WM_GETTEXTLENGTH = &HE


Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function GetForegroundWindow Lib "user32" () As Long
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Dim h1 As Long, h2 As Long, h3 As Long, h4 As Long, h5 As Long, h6 As Long
Private Declare Function SetWindowPos& Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, _
ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
Private Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long


Private Sub cmdBack_Click()
HideMeDown
End Sub

Private Sub cmdKill_Click()
On Error Resume Next
Timer1.Enabled = False
tXoaFile txtPath.Text
If modMain.FileExists(txtPath.Text) = False Then
    UniMsgBox "D9a4 xo1a xong!", vbOKOnly, "OK", Me.hwnd
    
    '=======
    Dim xPath As String
    xPath = FixPath(modMain.GetFolderPath(txtPath.Text))
    File1.Path = xPath
    File1.Refresh
    Dim Jk As Integer
    For Jk = 0 To File1.ListCount - 1
        Dim AJ As String
        AJ = RTPCheckVirus(FixPath(modMain.GetFolderPath(txtPath.Text)) & File1.List(Jk))
        If AJ <> "No" Then
            GoTo TiMrAvIrUs
        End If
    Next Jk
    '======
    
    HideMeDown
Else
    UniMsgBox "Kho6ng xo1a d9u7o75c!", vbOKOnly, "!", Me.hwnd
End If

Exit Sub
TiMrAvIrUs:
ShowMeOn
txtVirusName.Text = AJ
txtPath.Text = xPath & File1.List(Jk)
End Sub

Private Sub Timer1_Timer()
On Error Resume Next
   SetForegroundWindow Me.hwnd
   Me.Caption = "PAV - Real Time Protection"
End Sub

Private Sub Form_Load()
On Error Resume Next
Me.Visible = False
RTPConnectDB

File1.Normal = True
File1.ReadOnly = True
File1.System = True
File1.Hidden = True
File1.Pattern = "*.exe;*.bat;*.cmd;*.com;*.dll;*.ocx;*.pif;*.scr"
End Sub

Public Sub ShowMeOn()
On Error Resume Next
    Me.Show
    SetWindowPos Me.hwnd, -1, 0, 0, 0, 0, 3
    Timer1.Interval = 500
    Timer1.Enabled = True
End Sub
Public Sub HideMeDown()
On Error Resume Next
    Me.Hide
    Timer1.Enabled = False
End Sub

Private Sub tmrRTP_Timer()
On Error GoTo GaPlOiThItHoAtRa
h6 = 0
h1 = GetForegroundWindow
h2 = FindWindowEx(h1, ByVal 0&, "WorkerW", vbNullString)
h3 = FindWindowEx(h2, ByVal 0&, "ReBarWindow32", vbNullString)
h4 = FindWindowEx(h3, ByVal 0&, "ComboBoxEx32", vbNullString)
h5 = FindWindowEx(h4, ByVal 0&, "ComboBox", vbNullString)
h6 = FindWindowEx(h5, ByVal 0&, "Edit", vbNullString)

Dim Length As Long
Dim result As Long
Dim strTMP As String
Length = SendMessage(h6, WM_GETTEXTLENGTH, ByVal 0, ByVal 0) + 1
strTMP = Space(Length)
result = SendMessage(h6, WM_GETTEXT, ByVal Length, ByVal strTMP)
Dim s As Variant
Dim st As String
s = Split(strTMP, vbNullChar)

Dim sKQ As String
sKQ = s(0)
If DirExists(sKQ) = True Then
    Text1.Text = FixPath(sKQ)
    Dim xPath As String
    xPath = FixPath(sKQ)
    File1.Path = xPath
    File1.Refresh
    If File1.ListCount < 50 Then
        Dim Jk As Integer
        For Jk = 0 To File1.ListCount - 1
            Dim AJ As String
            AJ = RTPCheckVirus(FixPath(sKQ) & File1.List(Jk))
            If AJ <> "No" Then
                GoTo TiMrAvIrUs
            End If
        Next Jk
    End If 'file1.listcount < 50
End If

Exit Sub
TiMrAvIrUs:
ShowMeOn
txtVirusName.Text = AJ
txtPath.Text = xPath & File1.List(Jk)
Exit Sub
GaPlOiThItHoAtRa:
End Sub
