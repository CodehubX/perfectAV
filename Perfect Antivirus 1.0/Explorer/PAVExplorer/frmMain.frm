VERSION 5.00
Object = "{D1ECD258-D329-4388-AB83-DEC261A66B86}#1.0#0"; "UniControls_v2.0.ocx"
Begin VB.Form frmMain 
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PAV Explorer"
   ClientHeight    =   7215
   ClientLeft      =   45
   ClientTop       =   435
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
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7215
   ScaleWidth      =   7455
   StartUpPosition =   2  'CenterScreen
   Begin UniControls.UniLabel UniLabel5 
      Height          =   255
      Left            =   840
      Top             =   6600
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   450
      Caption         =   "Dung Lu7o75ng:"
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
   Begin UniControls.UniTextBox txtSize 
      Height          =   270
      Left            =   2040
      TabIndex        =   5
      Top             =   6600
      Width           =   5295
      _ExtentX        =   9340
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
      Text            =   ""
      Locked          =   -1  'True
      BorderStyle     =   2
   End
   Begin UniControls.UniTextBox txtAttr 
      Height          =   270
      Left            =   2040
      TabIndex        =   4
      Top             =   6840
      Width           =   5295
      _ExtentX        =   9340
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
      Text            =   ""
      Locked          =   -1  'True
      BorderStyle     =   2
   End
   Begin UniControls.UniLabel UniLabel3 
      Height          =   255
      Left            =   840
      Top             =   6840
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   450
      Caption         =   "Thuo65c Ti1nh:"
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
      Left            =   840
      Top             =   6360
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   450
      Caption         =   "D9i5a Chi3:"
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
   Begin UniControls.UniTextBox txtFilePath 
      Height          =   270
      Left            =   2040
      TabIndex        =   3
      Top             =   6360
      Width           =   5295
      _ExtentX        =   9340
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
      Text            =   ""
      Locked          =   -1  'True
      BorderStyle     =   2
   End
   Begin VB.PictureBox picIcon 
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   120
      ScaleHeight     =   615
      ScaleWidth      =   615
      TabIndex        =   2
      Top             =   6480
      Width           =   615
   End
   Begin UniControls.UniTextBox txtPath 
      Height          =   270
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   7215
      _ExtentX        =   12726
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
      Text            =   "My Computer"
      BorderStyle     =   2
   End
   Begin UniControls.UniLabel UniLabel1 
      Height          =   495
      Left            =   120
      Top             =   120
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   873
      Alignment       =   1
      Caption         =   "Qua3n ly1 File"
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
   Begin PAVExplorer.McListBox lstFolder 
      Height          =   5175
      Left            =   120
      TabIndex        =   0
      Top             =   1080
      Width           =   7215
      _ExtentX        =   7223
      _ExtentY        =   11033
      Picture         =   "frmMain.frx":617A
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   131072
      IconFocus       =   0   'False
      RowHeight       =   16
      BackGradient    =   6
      BackGradientCol =   12648447
      ShowIcon        =   -1  'True
      AutoHideScrollBars=   -1  'True
      Mode            =   4
      ShowSystemFiles =   -1  'True
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
Dim Comd
Comd = Command()
If Comd <> "explorer" Then
    UniMsgBox "D9a6y la2 ta65p tin ho64 tro75 qua3n ly1 File cu3a chu7o7ng tri2nh PAV 2009." & vbCrLf & " D9e63 su73 du5ng chu71c na8ng na2y, ba5n va2o mu5c 'Tie65n I1ch He65 Tho61ng' -> 'Qua3n Ly1 File'"
    End
End If


lstFolder.Mode = Mode_FileBrowser
'lstFolder.Path = ""
lstFolder.ShowHiddenFiles = True
lstFolder.ShowSystemFiles = True
lstFolder.ShowIcon = True
lstFolder.AutoHideScrollBars = False

End Sub




Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub lstFolder_DbClick()
If FileExists(txtPath.Text) = True Then
    If UniMsgBox("Ba5n muo61ng mo73 File " & GetFileName(txtPath.Text) & " kho6ng?", vbYesNo) = vbYes Then
        ShellExecute Me.hwnd, vbNullString, txtPath.Text, vbNullString, "", 1
    End If
End If
End Sub



Private Sub lstFolder_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.picIcon.Cls
Me.txtFilePath.Text = ""
Me.txtAttr.Text = ""
Me.txtSize.Text = ""
If FileExists(lstFolder.Text) = True Then
    Dim FSO As FileSystemObject
    GetLargeIcon lstFolder.Text, picIcon
    txtFilePath.Text = lstFolder.Text
    txtAttr.Text = basFile.GetAttribute(lstFolder.Text)
    txtSize.Text = FileLen(lstFolder.Text) \ 1024 & " KB (" & FileLen(lstFolder.Text) & " Bytes)"
End If
End Sub

Private Sub lstFolder_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
txtPath.Text = lstFolder.Text



If Button = 2 Then
    If frmMenu.sClipBoard = "" Then frmMenu.paste.Enabled = False Else frmMenu.paste.Enabled = True
    PopupMenu frmMenu.ttf
End If
End Sub

Private Sub txtExt_Change()
If Left(txtExt.Text, 2) <> "*." Then
txtExt.Text = "*." & txtExt.Text
txtExt.SelStart = Len(txtExt.Text)
End If
End Sub

Public Function FileExists(sFile) As Boolean
On Error Resume Next
FileExists = ((GetAttr(sFile) And vbDirectory) = 0)
End Function


