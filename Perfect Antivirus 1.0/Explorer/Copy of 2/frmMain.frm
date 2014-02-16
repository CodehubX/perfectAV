VERSION 5.00
Object = "{D1ECD258-D329-4388-AB83-DEC261A66B86}#1.0#0"; "UniControls_v2.0.ocx"
Begin VB.Form frmMain 
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PAV Explorer"
   ClientHeight    =   7440
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9930
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
   ScaleHeight     =   7440
   ScaleWidth      =   9930
   StartUpPosition =   2  'CenterScreen
   Begin UniControls.UniFrame fm2 
      Height          =   2415
      Left            =   4320
      Top             =   4920
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   4260
      MaskColor       =   16711935
      FrameColor      =   -2147483629
      Style           =   0
      Caption         =   "Tu2y cho5n"
      TextColor       =   13579779
      Alignment       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin UniControls.UniTextBox txtExt 
         Height          =   270
         Left            =   1680
         TabIndex        =   14
         Top             =   360
         Width           =   1695
         _ExtentX        =   2990
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
         Text            =   "*.*"
         BorderStyle     =   2
      End
      Begin UniControls.UniLabel UniLabel12 
         Height          =   255
         Left            =   240
         Top             =   360
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         Caption         =   "Kie63u File Hie63n Thi5:"
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
      Begin UniControls.UniCheckBox chkShowHidden 
         Height          =   195
         Left            =   240
         TabIndex        =   13
         Top             =   840
         Width           =   1965
         _ExtentX        =   3466
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
         Caption         =   "Hie63n thi5 File, Folder a63n."
         ForeColor       =   0
      End
   End
   Begin UniControls.UniFrame fm 
      Height          =   4095
      Left            =   4320
      Top             =   720
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   7223
      MaskColor       =   16711935
      FrameColor      =   -2147483629
      Style           =   0
      Caption         =   "Tho6ng tin"
      TextColor       =   13579779
      Alignment       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin UniControls.UniTextBox txtFileExt 
         Height          =   270
         Left            =   2040
         TabIndex        =   4
         Top             =   1200
         Width           =   3255
         _ExtentX        =   5741
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
      Begin UniControls.UniLabel UniLabel10 
         Height          =   255
         Left            =   120
         Top             =   3120
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   450
         Caption         =   "Phie6n Ba3n:"
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
      Begin UniControls.UniLabel UniLabel9 
         Height          =   255
         Left            =   120
         Top             =   2880
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   450
         Caption         =   "Mie6u Ta3:"
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
      Begin UniControls.UniLabel UniLabel8 
         Height          =   255
         Left            =   120
         Top             =   2640
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         Caption         =   "Nha2 Pha1t Ha2nh:"
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
      Begin UniControls.UniLabel UniLabel7 
         Height          =   255
         Left            =   120
         Top             =   2400
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
      Begin UniControls.UniLabel UniLabel6 
         Height          =   255
         Left            =   120
         Top             =   2160
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   450
         Caption         =   "Nga2y Kho73i Ta5o:"
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
         Left            =   120
         Top             =   1920
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   450
         Caption         =   "Ki1ch Thu7o71c:"
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
      Begin UniControls.UniTextBox txtFileName 
         Height          =   375
         Left            =   960
         TabIndex        =   3
         Top             =   360
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   661
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   0
         Text            =   ""
         BorderStyle     =   2
      End
      Begin UniControls.UniLabel UniLabel5 
         Height          =   255
         Left            =   120
         Top             =   1680
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
      Begin UniControls.UniLabel UniLabel4 
         Height          =   255
         Left            =   120
         Top             =   1440
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   450
         Caption         =   "Mo73 Ba82ng Chu7o7ng Tri2nh:"
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
         Left            =   120
         Top             =   1200
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   450
         Caption         =   "Kie63u File:"
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
      Begin VB.PictureBox picIcon 
         BackColor       =   &H80000009&
         BorderStyle     =   0  'None
         Height          =   615
         Left            =   120
         ScaleHeight     =   615
         ScaleWidth      =   615
         TabIndex        =   2
         Top             =   240
         Width           =   615
      End
      Begin UniControls.UniTextBox txtOpenWith 
         Height          =   270
         Left            =   2040
         TabIndex        =   5
         Top             =   1440
         Width           =   3255
         _ExtentX        =   5741
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
      Begin UniControls.UniTextBox txtFilePath 
         Height          =   270
         Left            =   2040
         TabIndex        =   6
         Top             =   1680
         Width           =   3255
         _ExtentX        =   5741
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
      Begin UniControls.UniTextBox txtFileSize 
         Height          =   270
         Left            =   2040
         TabIndex        =   7
         Top             =   1920
         Width           =   3255
         _ExtentX        =   5741
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
      Begin UniControls.UniTextBox txtDateCreate 
         Height          =   270
         Left            =   2040
         TabIndex        =   8
         Top             =   2160
         Width           =   3255
         _ExtentX        =   5741
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
         TabIndex        =   9
         Top             =   2400
         Width           =   3255
         _ExtentX        =   5741
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
      Begin UniControls.UniTextBox txtCompany 
         Height          =   270
         Left            =   2040
         TabIndex        =   10
         Top             =   2640
         Width           =   3255
         _ExtentX        =   5741
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
      Begin UniControls.UniTextBox txtDescription 
         Height          =   270
         Left            =   2040
         TabIndex        =   11
         Top             =   2880
         Width           =   3255
         _ExtentX        =   5741
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
      Begin UniControls.UniTextBox txtVersion 
         Height          =   270
         Left            =   2040
         TabIndex        =   12
         Top             =   3120
         Width           =   3255
         _ExtentX        =   5741
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
      Begin VB.Line Line2 
         BorderColor     =   &H00B99D7F&
         X1              =   120
         X2              =   5400
         Y1              =   1080
         Y2              =   1080
      End
   End
   Begin UniControls.UniTextBox txtPath 
      Height          =   270
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   4095
      _ExtentX        =   7223
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
      Width           =   9735
      _ExtentX        =   17171
      _ExtentY        =   873
      Alignment       =   1
      Caption         =   "Perfect AV - Explorer"
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
      Height          =   6255
      Left            =   120
      TabIndex        =   0
      Top             =   1080
      Width           =   4095
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
