VERSION 5.00
Object = "{E8FDD05C-3067-4198-8AEC-1A013A46ABDD}#1.0#0"; "FVUnicodeControl.ocx"
Object = "{6DC1AB90-DCD1-47A1-AB36-924FFD67ADBF}#1.0#0"; "UniControls_v2.0.ocx"
Begin VB.Form frmAddVirus 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Update New Virus"
   ClientHeight    =   2970
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5775
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAddVirus.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2970
   ScaleWidth      =   5775
   StartUpPosition =   2  'CenterScreen
   Begin PAVUpdate.Downloader DL1 
      Height          =   255
      Left            =   5760
      TabIndex        =   3
      Top             =   2040
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   450
   End
   Begin UniControls.UniTrayIcon Tray1 
      Left            =   4800
      Top             =   120
      _ExtentX        =   1376
      _ExtentY        =   1376
      TooltipText     =   "[PAV 2009] D9ang ca65p nha65t..."
      Icon            =   "frmAddVirus.frx":628A
   End
   Begin FVUnicodeControl.FVistaUniLabel FVistaUniLabel1 
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   240
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   873
      Alignment       =   1
      Caption         =   "Ca65p nha65t ma64u Virus"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin UniControls.UniLabel LBL 
      Height          =   255
      Index           =   2
      Left            =   1560
      Top             =   1800
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   450
      Caption         =   "12 ma64u"
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
   Begin UniControls.UniLabel LBL 
      Height          =   255
      Index           =   1
      Left            =   1560
      Top             =   1560
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   450
      Caption         =   "123 KB"
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
   Begin UniControls.UniLabel LBL 
      Height          =   255
      Index           =   0
      Left            =   1560
      Top             =   1320
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   450
      Caption         =   "1.0.1"
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
   Begin FVUnicodeControl.FVistaUniButton cmdCancel 
      Height          =   375
      Left            =   4200
      TabIndex        =   1
      Top             =   2520
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      BackColor       =   14737632
      ButtonStyle     =   3
      Caption         =   "D9o1ng"
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
   Begin FVUnicodeControl.FVistaUniButton cmdUpdate 
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   2520
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      BackColor       =   -2147483629
      ButtonStyle     =   3
      Caption         =   "Ca65p nha65t ngay!"
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
   Begin FVUnicodeControl.FVistaUniProgressbar Pro1 
      Height          =   225
      Left            =   360
      Top             =   2160
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   397
      Max             =   100
      Value           =   0
      TStyle          =   1
      Min             =   0
      Style           =   1
      Text            =   "ZProgessbar"
      Align           =   1
   End
   Begin UniControls.UniLabel UniLabel4 
      Height          =   255
      Left            =   240
      Top             =   1800
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   450
      Alignment       =   2
      Caption         =   "So61 ma64u Virus:"
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
   Begin UniControls.UniLabel UniLabel3 
      Height          =   255
      Left            =   360
      Top             =   1560
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   450
      Alignment       =   2
      Caption         =   "Dung lu7o75ng:"
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
      Left            =   240
      Top             =   1320
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   450
      Alignment       =   2
      Caption         =   "Phie6n ba3n:"
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
   Begin UniControls.UniLabel UniLabel2 
      Height          =   255
      Left            =   360
      Top             =   960
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   450
      Caption         =   "Tho6ng tin:"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmAddVirus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function DeleteFile Lib "kernel32" Alias "DeleteFileA" (ByVal lpFileName As String) As Long

Public xLinkDownLoad As String
Private Sub cmdCancel_Click()
Unload frmMain
Unload Me
End Sub

Private Sub cmdUpdate_Click()
'MsgBox GetUrlSource("http://phanmemvn.net/files/update.txt")
cmdUpdate.Enabled = False
cmdCancel.Enabled = False
On Error GoTo KhOnGxOaFiLe
SetAttr AppPath & "UpdateFile.exe", vbNormal
DeleteFile AppPath & "UpdateFile.exe"
KhOnGxOaFiLe:
DL1.DownloadFile xLinkDownLoad, AppPath & "UpdateFile.exe"
cmdCancel.Enabled = True
End Sub



Private Sub DL1_Complete(URL As String)
On Error Resume Next
Shell AppPath & "UpdateFile.exe", vbNormalFocus
UniMsgBox "D9a4 ca65p nha65t xong!", vbOKOnly, "OK!"
End Sub

Private Sub DL1_Progress(ByteDownloaded As Long, FileSize As Long, URL As String)
On Error Resume Next
If FileSize <> 0 Then
Pro1.Value = (Int(ByteDownloaded * 100 / FileSize))
End If
End Sub

Private Sub Form_Load()
Tray1.Remove
frmMain.Timer1.Enabled = False
End Sub

Private Sub Form_Resize()
If Me.WindowState = 1 Then
    Me.Hide
    App.TaskVisible = False
    Tray1.Create "[PAV 2009] D9ang ta3i ba3n ca65p nha65t..."
    Tray1.BalloonTip "D9ang ta3i ba3n ca65p nha65t..." & vbCrLf & "Nha61n va2o d9a6y d9e63 xem chi tie61t!", btsInfo, "Tho6ng ba1o!"
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
If frmMain.FileExists(AppPath & "UpdateInfo.txt") = True Then
Kill AppPath & "UpdateInfo.txt"
Kill AppPath & "Mess.txt"
End If

Unload frmMain
End Sub

Private Sub Tray1_BalloonClick(ClickType As UniControls.stBalloonClickType)
If ClickType <> stbXClick Then
Me.WindowState = 0
Me.Show
App.TaskVisible = True
End If
End Sub

Private Sub Tray1_TrayClick(Button As UniControls.stMouseEvent)
Me.WindowState = 0
Me.Show
App.TaskVisible = True
End Sub
