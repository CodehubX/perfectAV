VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Offline Creator"
   ClientHeight    =   3090
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmMain.frx":1CFA
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   2160
      Top             =   2520
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "Download"
      Enabled         =   0   'False
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   2640
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3240
      TabIndex        =   1
      Top             =   2640
      Width           =   1215
   End
   Begin PAVOfflineCreator.Downloader DL1 
      Height          =   135
      Left            =   4080
      TabIndex        =   0
      Top             =   3120
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   238
   End
   Begin VB.Label LBL 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   2
      Left            =   1680
      TabIndex        =   6
      Top             =   1800
      Width           =   2535
   End
   Begin VB.Label LBL 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   1680
      TabIndex        =   5
      Top             =   1560
      Width           =   2535
   End
   Begin VB.Label LBL 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   1680
      TabIndex        =   4
      Top             =   1320
      Width           =   2535
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "0%"
      Height          =   255
      Left            =   1440
      TabIndex        =   3
      Top             =   2760
      Width           =   975
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function DeleteFile Lib "kernel32" Alias "DeleteFileA" (ByVal lpFileName As String) As Long

Dim xLinkDownLoad As String
Private Sub CheckForUpdate()
'xStrUpdate = GetUrlSource("http://phanmemvn.net/files/update.txt")

Me.LBL(0).Caption = ReadIniFile(AppPath & "UpdateInfo.txt", "Update", "Version", 0)
Me.LBL(1).Caption = ReadIniFile(AppPath & "UpdateInfo.txt", "Update", "Size", 0)
Me.LBL(2).Caption = ReadIniFile(AppPath & "UpdateInfo.txt", "Update", "Number", 0)
xLinkDownLoad = ReadIniFile(AppPath & "UpdateInfo.txt", "Update", "Link", "")
If xLinkDownLoad <> "" Then
    cmdUpdate.Enabled = True
End If
If FileExists(AppPath & "UpdateInfo.txt") = True Then
DeleteFile AppPath & "UpdateInfo.txt"
End If
End Sub

Public Function FileExists(sFile As String) As Boolean
On Error Resume Next
FileExists = ((GetAttr(sFile) And vbDirectory) = 0)
End Function






'================================================

Private Sub cmdCancel_Click()
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

On Error Resume Next
DL1.DownloadFile xLinkDownLoad, AppPath & "UpdateFile.exe"
cmdCancel.Enabled = True
End Sub



Private Sub DL1_Complete(URL As String)
On Error Resume Next
If URL = xLinkDownLoad Then
MsgBox "Download Complete!" & vbCrLf & AppPath & "UpdateFile.exe"
End If
End Sub

Private Sub DL1_Progress(ByteDownloaded As Long, FileSize As Long, URL As String)
On Error Resume Next
If FileSize <> 0 Then
Label1.Caption = (Int(ByteDownloaded * 100 / FileSize)) & " %"
End If
End Sub


Private Sub Form_Load()
DL1.DownloadFile "http://phanmemvn.net/files/update.txt", AppPath & "UpdateInfo.txt"
End Sub

Private Sub Form_Unload(Cancel As Integer)
If frmMain.FileExists(AppPath & "UpdateInfo.txt") = True Then
DeleteFile AppPath & "UpdateInfo.txt"
End If
End Sub

Private Sub Timer1_Timer()
CheckForUpdate
Timer1.Enabled = False
End Sub
