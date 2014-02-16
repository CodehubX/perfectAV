VERSION 5.00
Object = "{E8FDD05C-3067-4198-8AEC-1A013A46ABDD}#1.0#0"; "FVUnicodeControl.ocx"
Object = "{6DC1AB90-DCD1-47A1-AB36-924FFD67ADBF}#1.0#0"; "UniControls_v2.0.ocx"
Begin VB.Form frmPhucHoiDuLieu 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Phuc hoi du lieu"
   ClientHeight    =   2235
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6090
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPhucHoiDuLieu.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2235
   ScaleWidth      =   6090
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin FVUnicodeControl.FVistaUniButton cmdStartPhucHoi 
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   1080
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      BackColor       =   14737632
      ButtonStyle     =   3
      Caption         =   "Ba81t d9a62u"
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
   Begin FVUnicodeControl.FVistaUniButton cmdChonDiaDiem 
      Height          =   255
      Left            =   5040
      TabIndex        =   1
      Top             =   720
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   450
      BackColor       =   14737632
      ButtonStyle     =   3
      Caption         =   "Cho5n"
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
      Top             =   360
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   450
      Caption         =   "Cho5n d9i5a d9ie63m ca62n phu5c ho62i:"
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
   Begin UniControls.UniTextBox TxtPathPhucHoi 
      Height          =   270
      Left            =   240
      TabIndex        =   0
      Top             =   720
      Width           =   4695
      _ExtentX        =   8281
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
      ForeColor       =   192
      Text            =   "C:\Program Files\"
      Locked          =   -1  'True
      BorderStyle     =   2
   End
   Begin FVUnicodeControl.FVistaUniButton cmdStopPhucHoi 
      Height          =   375
      Left            =   1320
      TabIndex        =   3
      Top             =   1080
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      BackColor       =   14737632
      ButtonStyle     =   3
      Caption         =   "Du72ng la5i"
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
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   240
      TabIndex        =   4
      Top             =   1560
      Width           =   5655
   End
End
Attribute VB_Name = "frmPhucHoiDuLieu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim XStop As Boolean
Private Sub cmdChonDiaDiem_Click()
Me.TxtPathPhucHoi.Text = ChonThuMuc(Me)
End Sub

Sub ShowFolderList(folderspec)
On Error Resume Next
  Dim MyFSO, NameFld, ffTmp, ffCollec
  Set MyFSO = CreateObject("Scripting.FileSystemObject")
  Set NameFld = MyFSO.GetFolder(folderspec)
  Set ffCollec = NameFld.SubFolders

  For Each ffTmp In ffCollec 'tìm các thu muc con
     Label1.Caption = ffTmp.Path: DoEvents
    If UCase(FixPath(ffTmp.Path)) <> "C:\WINDOWS\" And UCase(FixPath(ffTmp.Path)) <> "C:\" Then
        SetAttr ffTmp.Path, vbNormal
    End If
            'Làm tùy ý tai Ðây
    If XStop = True Then Exit Sub
     ShowFolderList ffTmp
  Next
  Set MyFSO = Nothing 'Giai phóng tài nguyên
End Sub
Sub ShowFileList(folderspec)
  On Error Resume Next
  Dim MyFSO, NameFld, ffTmp, ffCollec
  Set MyFSO = CreateObject("Scripting.FileSystemObject")
  Set NameFld = MyFSO.GetFolder(folderspec)
  Set ffCollec = NameFld.Files
    For Each ffTmp In ffCollec
        Label1.Caption = ffTmp.Path: DoEvents
        If UCase(FixPath(modMain.GetFolderPath(ffTmp.Path))) <> "C:\WINDOWS\" And UCase(FixPath(modMain.GetFolderPath(ffTmp.Path))) <> "C:\" And UCase$(modMain.GetFileName(ffTmp.Path)) <> "DESKTOP.INI" Then
            SetAttr ffTmp.Path, vbNormal
        End If
        If UCase$(modMain.GetFileName(ffTmp.Path)) = "DESKTOP.INI" Then
            SetAttr ffTmp.Path, vbSystem + vbHidden
        End If
        If XStop = True Then Exit Sub
        'Làm tùy ý tai Ðây
    Next 'fftmp.name thay vì fftmp.path Ðê gon và Ðep
    Set ffCollec = NameFld.SubFolders
  For Each ffTmp In ffCollec 'tìm các thu muc con
     ShowFileList ffTmp
  Next
  Set MyFSO = Nothing 'Giai phóng tài nguyên
End Sub
Private Sub cmdStartPhucHoi_Click()
cmdStartPhucHoi.Enabled = False
Me.cmdChonDiaDiem.Enabled = False
Me.TxtPathPhucHoi.Enabled = False

XStop = False
ShowFolderList Me.TxtPathPhucHoi.Text
ShowFileList Me.TxtPathPhucHoi.Text
cmdStartPhucHoi.Enabled = True
Me.cmdChonDiaDiem.Enabled = True
Me.TxtPathPhucHoi.Enabled = True

Label1.Caption = "Done!"
End Sub

Private Sub cmdStopPhucHoi_Click()
XStop = True
End Sub

