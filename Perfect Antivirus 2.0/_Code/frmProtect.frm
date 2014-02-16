VERSION 5.00
Object = "{6DC1AB90-DCD1-47A1-AB36-924FFD67ADBF}#1.0#0"; "UniControls_v2.0.ocx"
Begin VB.Form frmProtect 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "PAV2009 - Protect"
   ClientHeight    =   1245
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3180
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
   Icon            =   "frmProtect.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1245
   ScaleWidth      =   3180
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      Height          =   840
      Left            =   3360
      TabIndex        =   0
      Top             =   600
      Width           =   735
   End
   Begin UniControls.UniLabel lblPath 
      Height          =   375
      Left            =   720
      Top             =   480
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   661
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
   Begin UniControls.UniLabel UniLabel1 
      Height          =   255
      Left            =   120
      Top             =   120
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   450
      BackStyle       =   0
      Caption         =   "Pha1t hie65n ta65p tin gia3 ma5o ta5i:"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   255
   End
   Begin VB.Timer Timer1 
      Interval        =   5000
      Left            =   960
      Top             =   0
   End
   Begin UniControls.UniLabel LBL1 
      Height          =   255
      Left            =   120
      Top             =   960
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   450
      BackStyle       =   0
      Caption         =   "D9a4 die65t va2 the6m va2o danh sa1ch virus."
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
   End
   Begin VB.Timer Timer5 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   600
      Top             =   840
   End
   Begin VB.Timer Timer6 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   2160
      Top             =   720
   End
   Begin VB.Timer Timer7 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   2520
      Top             =   0
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   0
      Picture         =   "frmProtect.frx":058A
      Top             =   240
      Width           =   720
   End
End
Attribute VB_Name = "frmProtect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SetWindowPos& Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, _
ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
Private Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long

Dim xGiaToc As Integer
Dim xDem As Integer
Dim xVirusFound As Boolean

Private Sub Form_Load()
xVirusFound = False
Me.Visible = False
SetForegroundWindow Me.hwnd
SetWindowPos Me.hwnd, -1, 0, 0, 0, 0, 3

List1.AddItem "C:\WINDOWS\userinit.exe"
List1.AddItem "C:\WINDOWS\system.exe"
List1.AddItem "C:\WINDOWS\svchost.exe"
List1.AddItem "C:\WINDOWS\system32\system.exe"

End Sub



Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
SaveString HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon", "Userinit", Environ("windir") & "\system32\userinit.exe,"
SaveString HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon", "Shell", "Explorer.exe"
SaveString HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon", "UIHost", "logonui.exe"
End Sub

Private Sub Timer1_Timer()
Dim Jh As Integer
For Jh = 0 To List1.ListCount - 1
If modMain.FileExists(List1.List(Jh)) Then
    xVirusFound = True
    'LBL1.Caption = "Ba5n ha4y que1t Virus ngay ba65y gio72!"
    lblPath.Caption = List1.List(Jh)
    
    ShowMe
    If FileLen(List1.List(Jh)) > 5 Then
        If CheckVirus(List1.List(Jh)) = "No" Then
            Dim xVirusName As String
            Dim xMd5 As String
            xVirusName = "Virus." & modMain.GetFileName(List1.List(Jh)) & "(Auto)"
            xMd5 = GetMD5(List1.List(Jh))
            AddVirus xVirusName, xMd5
            modReadWrite.WriteFileUni AppPath & "UserData\" & File2Str(xVirusName), xMd5
            GetListData
        End If
    End If
    basProcess.SuspendResumeProcess CheckProcess(List1.List(Jh)), True
    tXoaFile List1.List(Jh)
    xVirusFound = False
    Exit For
End If
Next Jh
End Sub

Private Sub Timer5_Timer()
xDem = xDem + 1
If xDem = 6 Then
'LBL1.Caption = "Xong!"
End If
If xDem = 7 Then
xGiaToc = 30
Timer7.Enabled = True
End If
End Sub

Private Sub Timer6_Timer()
xGiaToc = xGiaToc - 1
Me.Top = Me.Top - 5 * xGiaToc
If Me.Top < Screen.Height - Me.Height - 400 Then
If Me.Top <> Screen.Height - Me.Height - 400 Then Me.Top = Screen.Height - Me.Height - 400
Timer6.Enabled = False
End If
End Sub

Private Sub Timer7_Timer()
xGiaToc = xGiaToc - 1
Me.Top = Me.Top + 5 * xGiaToc
If Me.Top > Screen.Height Then
If xVirusFound = False Then
    HideMe
Else
    Unload Me
End If
End If
End Sub

Sub ShowMe()
Me.Left = Screen.Width - Me.Width
Me.Top = Screen.Height - 400
Me.Show
xGiaToc = 30
xDem = 0
Timer6.Enabled = True
Timer5.Enabled = True
Timer7.Enabled = False
End Sub
Sub HideMe()
Me.Visible = False
Timer6.Enabled = False
Timer7.Enabled = False
Timer5.Enabled = False
End Sub


