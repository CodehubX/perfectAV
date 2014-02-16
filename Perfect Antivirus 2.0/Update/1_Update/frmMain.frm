VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PAV Update"
   ClientHeight    =   3645
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5850
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3645
   ScaleWidth      =   5850
   StartUpPosition =   3  'Windows Default
   Begin PAVUpdate.Downloader FDL2 
      Height          =   615
      Left            =   3000
      TabIndex        =   1
      Top             =   2640
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   1085
   End
   Begin PAVUpdate.Downloader FDL1 
      Height          =   735
      Left            =   1560
      TabIndex        =   0
      Top             =   840
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   1296
   End
   Begin VB.Timer Timer1 
      Interval        =   5000
      Left            =   1080
      Top             =   1560
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim xClose As Boolean
Private Sub CheckForUpdate()
On Error GoTo ThOaTvIlOi
'xStrUpdate = GetUrlSource("http://phanmemvn.net/files/update.txt")
FDL1.DownloadFile "http://phanmemvn.net/files/update.txt", AppPath & "UpdateInfo.txt"

ThOaTvIlOi:
End Sub

Public Function FileExists(sFile As String) As Boolean
On Error Resume Next
FileExists = ((GetAttr(sFile) And vbDirectory) = 0)
End Function

Private Sub FDL1_Complete(URL As String)

Dim xVer As Long
xVer = IIf(IsNumeric(GetSetting("PAV2009", "Update", "Version", 0)), GetSetting("PAV2009", "Update", "Version", 0), 0)

If ReadIniFile(AppPath & "UpdateInfo.txt", "Update", "Version", "NoData") = "NoData" Then Exit Sub
    If ReadIniFile(AppPath & "UpdateInfo.txt", "Update", "Version", 0) > xVer Then
        xVer = ReadIniFile(AppPath & "UpdateInfo.txt", "Update", "Version", 0)
        SaveSetting "PAV2009", "Update", "Version", xVer
        frmAddVirus.Show
        frmAddVirus.LBL(0).Caption = xVer
        frmAddVirus.LBL(1).Caption = ReadIniFile(AppPath & "UpdateInfo.txt", "Update", "Size", 0)
        frmAddVirus.LBL(2).Caption = ReadIniFile(AppPath & "UpdateInfo.txt", "Update", "Number", 0)
        frmAddVirus.xLinkDownLoad = ReadIniFile(AppPath & "UpdateInfo.txt", "Update", "Link", 0)
        xClose = False
    End If
FDL2.DownloadFile "http://phanmemvn.net/files/mess.txt", AppPath & "Mess.txt"
End Sub

Private Sub FDL2_Complete(URL As String)
Dim zVer As Long
zVer = IIf(IsNumeric(GetSetting("PAV2009", "Update", "MESS", 0)), GetSetting("PAV2009", "Update", "MESS", 0), 0)
    If ReadIniFile(AppPath & "Mess.txt", "Mess", "VerMess", 0) > zVer Then
        zVer = ReadIniFile(AppPath & "Mess.txt", "Mess", "VerMess", 0)
        SaveSetting "PAV2009", "Update", "MESS", zVer
        frmMess.Show
        frmMess.txtMess.Text = UTF82Unicode(ReadIniFile(AppPath & "Mess.txt", "Mess", "NoiDung", ""))
        xClose = False
    End If
    If FileExists(AppPath & "UpdateInfo.txt") = True Then
    Kill AppPath & "UpdateInfo.txt"
    Kill AppPath & "Mess.txt"
    End If
    If xClose = True Then
    Unload frmAddVirus
    Unload frmMess
    Unload frmMain
    End If
End Sub

Private Sub Form_Load()
If App.PrevInstance = True Then End
Me.Visible = False
App.TaskVisible = False
xClose = True
End Sub

Private Sub Timer1_Timer()
If CheckInternet = True Then
CheckForUpdate
End If
End Sub
