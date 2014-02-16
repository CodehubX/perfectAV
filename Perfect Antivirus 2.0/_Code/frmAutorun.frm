VERSION 5.00
Object = "{E8FDD05C-3067-4198-8AEC-1A013A46ABDD}#1.0#0"; "FVUnicodeControl.ocx"
Object = "{6DC1AB90-DCD1-47A1-AB36-924FFD67ADBF}#1.0#0"; "UniControls_v2.0.ocx"
Object = "{6FBA474E-43AC-11CE-9A0E-00AA0062BB4C}#1.0#0"; "SYSINFO.OCX"
Begin VB.Form frmAutorun 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PAV 2009 - Autorun Protect"
   ClientHeight    =   2220
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6495
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAutorun.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmAutorun.frx":058A
   ScaleHeight     =   2220
   ScaleWidth      =   6495
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin SysInfoLib.SysInfo SysInfo1 
      Left            =   240
      Top             =   480
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin FVUnicodeControl.FVistaUniCheckbox ChkXoaKhongCanHoi 
      Height          =   195
      Left            =   4440
      TabIndex        =   3
      Top             =   1440
      Width           =   1605
      _ExtentX        =   2831
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
      Caption         =   "Xo1a kho6ng ca62n ho3i"
      ForeColor       =   0
      ShowFocusRectangle=   0   'False
   End
   Begin VB.Timer Timer2 
      Interval        =   2000
      Left            =   2040
      Top             =   0
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   1320
      Top             =   120
   End
   Begin FVUnicodeControl.FVistaUniButton cmdKhongXoa 
      Height          =   375
      Left            =   4920
      TabIndex        =   2
      Top             =   1680
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      BackColor       =   14737632
      ButtonStyle     =   3
      Caption         =   "Kho6ng xo1a"
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
   Begin FVUnicodeControl.FVistaUniButton cmdXoa 
      Height          =   375
      Left            =   2160
      TabIndex        =   1
      Top             =   1680
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      BackColor       =   14737632
      ButtonStyle     =   3
      Caption         =   "Xo1a"
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
   Begin UniControls.UniTextBox txtPathAutorun 
      Height          =   270
      Left            =   2040
      TabIndex        =   0
      Top             =   1080
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   476
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
      Text            =   "C:\autorun.inf"
      Locked          =   -1  'True
      BorderStyle     =   2
   End
   Begin UniControls.UniLabel UniLabel2 
      Height          =   255
      Left            =   1560
      Top             =   600
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   450
      Alignment       =   1
      Caption         =   "Chu7o7ng tri2nh pha1t hie65n co1 Autorun trong o63 d9i4a."
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin UniControls.UniLabel UniLabel1 
      Height          =   375
      Left            =   1800
      Top             =   120
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   661
      Alignment       =   1
      Caption         =   "Pha1t hie65n Autorun !!!"
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
   Begin UniControls.UniLabel lblXoaXong 
      Height          =   255
      Left            =   2040
      Top             =   1440
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   450
      Alignment       =   1
      Caption         =   "D9a4 xo1a xong!"
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
End
Attribute VB_Name = "frmAutorun"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SetWindowPos& Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, _
ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
Private Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
Dim xOkCount As Integer
Private Sub ChkXoaKhongCanHoi_Click()
On Error Resume Next
SaveSetting "PAV2009", "AutorunProtect", "AutoDelete", ChkXoaKhongCanHoi.Value
End Sub

Private Sub cmdKhongXoa_Click()
On Error Resume Next
HideMeDown
xOkCount = xOkCount + 1
End Sub

Private Sub cmdXoa_Click()
On Error Resume Next
Timer1.Enabled = False
tXoaFile txtPathAutorun.Text
If modMain.FileExists(txtPathAutorun.Text) = False Then
    UniMsgBox "D9a4 xo1a xong!", vbOKOnly, "OK!", Me.hwnd
    Timer2_Timer
Else
    UniMsgBox "Kho6ng xo1a d9u7o75c!", vbOKOnly, "Error!", Me.hwnd
End If

HideMeDown
End Sub

Private Sub Form_Load()
On Error Resume Next
xOkCount = 0
Me.Visible = False
End Sub


Private Sub SysInfo1_DeviceArrival(ByVal DeviceType As Long, ByVal DeviceID As Long, ByVal DeviceName As String, ByVal DeviceData As Long)
Timer2_Timer
frmAutorun2.Show
End Sub

Private Sub Timer1_Timer()
On Error Resume Next
SetForegroundWindow Me.hwnd
End Sub

Private Sub Timer2_Timer()
On Error Resume Next
    Dim Str
    Dim str2
    Dim FSO  As New FileSystemObject
    Dim drv  As Drive
    Dim drvs As Drives
    DoEvents
    Set drvs = FSO.Drives
    For Each drv In drvs
        If UCase(drv.DriveLetter) <> "A" And drv.DriveType <> CDRom Then
            If modMain.FileExists(drv.DriveLetter & ":\autorun.inf") = True Then
                ShowMeOn
                txtPathAutorun.Text = drv.DriveLetter & ":\autorun.inf"
                If GetSetting("PAV2009", "AutorunProtect", "AutoDelete", False) = True Then
                    tXoaFile txtPathAutorun.Text
                    cmdXoa.Visible = False
                    cmdKhongXoa.Caption = "OK"
                    lblXoaXong.Visible = True
                End If
            End If
        End If
    Next
    Set FSO = Nothing
    Set drv = Nothing
    Set drvs = Nothing
If xOkCount > 3 Then Unload Me
End Sub
Public Function GetOpenAutorun(sAutorunFile) As String
On Error Resume Next
On Error Resume Next
Dim xStart1
Dim xEnd1
Dim xAutoFile
Dim x1 As String
Dim FSO
 Set FSO = CreateObject("Scripting.FileSystemObject").OpenTextFile(sAutorunFile, 1, , -2)
 xAutoFile = FSO.ReadAll
 xAutoFile = DelAllSpace(xAutoFile)
 Set FSO = Nothing
 xStart1 = (Len(xAutoFile) - InStrRev(LCase(StrReverse(xAutoFile)), StrReverse("open=")) + 2)
 xEnd1 = (Len(xAutoFile) - InStrRev(StrReverse(xAutoFile), Chr(13), Len(xAutoFile) - xStart1) + 1)
 x1 = Mid$(xAutoFile, (Len(xAutoFile) - InStrRev(LCase(StrReverse(xAutoFile)), StrReverse("open=")) + 2), (Len(xAutoFile) - InStrRev(StrReverse(xAutoFile), Chr(13), Len(xAutoFile) - xStart1) + 1) - (Len(xAutoFile) - InStrRev(LCase(StrReverse(xAutoFile)), StrReverse("open=")) + 2))
 GetOpenAutorun = x1
End Function

Public Function GetShellOpenAutorun(sAutorunFile) As String
On Error Resume Next
On Error Resume Next
Dim xStart1
Dim xEnd1
Dim xAutoFile
Dim x1 As String
Dim FSO
 Set FSO = CreateObject("Scripting.FileSystemObject").OpenTextFile(sAutorunFile, 1, , -2)
 xAutoFile = FSO.ReadAll
 xAutoFile = DelAllSpace(xAutoFile)
 Set FSO = Nothing
 xStart1 = (Len(xAutoFile) - InStrRev(LCase(StrReverse(xAutoFile)), StrReverse("shell\open\command=")) + 2)
 xEnd1 = (Len(xAutoFile) - InStrRev(StrReverse(xAutoFile), Chr(13), Len(xAutoFile) - xStart1) + 1)
 x1 = Mid$(xAutoFile, (Len(xAutoFile) - InStrRev(LCase(StrReverse(xAutoFile)), StrReverse("shell\open\command=")) + 2), (Len(xAutoFile) - InStrRev(StrReverse(xAutoFile), Chr(13), Len(xAutoFile) - xStart1) + 1) - (Len(xAutoFile) - InStrRev(LCase(StrReverse(xAutoFile)), StrReverse("shell\open\command=")) + 2))
 GetShellOpenAutorun = x1
End Function
Public Function DelAllSpace(Str) As String
On Error Resume Next
Do While InStr(Str, " ") > 0
    Str = Replace(Str, " ", "")
Loop
Str = Trim(Str)
DelAllSpace = Str
End Function

Public Sub ShowMeOn()
On Error Resume Next
    Me.Show
    SetWindowPos Me.hwnd, -1, 0, 0, 0, 0, 3
    Timer1.Interval = 500
    Timer1.Enabled = True
    'SaveSetting "PAV2009", "AutorunProtect", "AutoDelete", ChkXoaKhongCanHoi.Value
    ChkXoaKhongCanHoi.Value = GetSetting("PAV2009", "AutorunProtect", "AutoDelete", False)
End Sub
Public Sub HideMeDown()
On Error Resume Next
cmdXoa.Visible = True
cmdKhongXoa.Caption = "Kho6ng xo1a"
lblXoaXong.Visible = False
    Me.Hide
    Timer1.Enabled = False
End Sub
