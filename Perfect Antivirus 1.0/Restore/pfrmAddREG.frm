VERSION 5.00
Object = "{D1ECD258-D329-4388-AB83-DEC261A66B86}#1.0#0"; "UniControls_v2.0.ocx"
Begin VB.Form pfrmAddREG 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add"
   ClientHeight    =   3600
   ClientLeft      =   165
   ClientTop       =   465
   ClientWidth     =   5415
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "pfrmAddREG.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3600
   ScaleWidth      =   5415
   StartUpPosition =   2  'CenterScreen
   Begin UniControls.UniLabel UniLabel7 
      Height          =   495
      Left            =   360
      Top             =   600
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   873
      Caption         =   "Ha4y chi3nh su73a tho6ng tin cu3a chu71c na8ng cho phu2 ho75p, sau d9o1 nha61n nu1t The6m Va2o."
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   33023
   End
   Begin UniControls.UniButton cmdAddREGBack 
      Height          =   375
      Left            =   4920
      TabIndex        =   6
      Top             =   3120
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      Icon            =   "pfrmAddREG.frx":058A
      Style           =   2
      IconAlign       =   2
      iNonThemeStyle  =   2
      BackColor       =   -2147483643
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
   Begin UniControls.UniTextBox txtChucNang 
      Height          =   270
      Left            =   1800
      TabIndex        =   5
      Top             =   1200
      Width           =   3495
      _ExtentX        =   6165
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
      Text            =   "Windows Task Manager"
      BorderStyle     =   2
   End
   Begin UniControls.UniLabel UniLabel4 
      Height          =   255
      Left            =   120
      Top             =   1200
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   450
      Caption         =   "Te6n chu71c na8ng:"
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
   Begin UniControls.UniButton cmdAddREG 
      Height          =   375
      Left            =   1800
      TabIndex        =   4
      Top             =   3120
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      Style           =   2
      Caption         =   "The6m va2o"
      IconAlign       =   3
      iNonThemeStyle  =   2
      BackColor       =   -2147483643
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
   Begin UniControls.UniTextBox txtkeyData 
      Height          =   270
      Left            =   1800
      TabIndex        =   3
      Top             =   2640
      Width           =   3495
      _ExtentX        =   6165
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
      Text            =   "0"
      BorderStyle     =   2
      NumberOnly      =   -1  'True
   End
   Begin UniControls.UniTextBox txtKeyName 
      Height          =   270
      Left            =   1800
      TabIndex        =   2
      Top             =   2280
      Width           =   3495
      _ExtentX        =   6165
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
      Text            =   "DisableTaskMgr"
      BorderStyle     =   2
   End
   Begin UniControls.UniTextBox txtKeyPath 
      Height          =   270
      Left            =   1800
      TabIndex        =   1
      Top             =   1920
      Width           =   3495
      _ExtentX        =   6165
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
      Text            =   "Software\Microsoft\Windows\CurrentVersion\Policies\System"
      BorderStyle     =   2
   End
   Begin UniControls.UniComboBox cbKeyGoc 
      Height          =   330
      Left            =   1800
      TabIndex        =   0
      Top             =   1560
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   582
      Style           =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ExtendedUI      =   0   'False
      DropDownWidth   =   0
   End
   Begin UniControls.UniLabel UniLabel6 
      Height          =   255
      Left            =   120
      Top             =   2640
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   450
      Caption         =   "Gia1 tri5 ma85c d9i5nh:"
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
   Begin UniControls.UniLabel UniLabel5 
      Height          =   255
      Left            =   120
      Top             =   2280
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   450
      Caption         =   "Te6n kho1a:"
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
      Top             =   1920
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   450
      Caption         =   "D9u7o72ng da64n kho1a:"
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
      Top             =   1560
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   450
      Caption         =   "Kho1a go61c:"
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
      Height          =   375
      Left            =   120
      Top             =   120
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   661
      Alignment       =   1
      Caption         =   "The6m va2o chu71c na8ng ca62n ba3o ve65"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16711680
   End
End
Attribute VB_Name = "pfrmAddREG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmdAddREG_Click()
If Me.txtChucNang.Text <> "" And Me.txtkeyData.Text <> "" And Me.txtKeyPath.Text <> "" And Me.txtKeyName.Text <> "" Then
With frmMain.atpLVREG
Dim i
i = .ListItems.Count + 1
    .ListItems.Add i, , Me.txtChucNang.Text
    .ListItems(i).SubItems(1).Caption = Me.cbKeyGoc.Text
    .ListItems(i).SubItems(2).Caption = Me.txtKeyPath.Text
    .ListItems(i).SubItems(3).Caption = Me.txtKeyName.Text
    .ListItems(i).SubItems(4).Caption = Me.txtkeyData.Text

End With
frmMain.SaveREG
    Me.txtChucNang.Text = ""
    Me.cbKeyGoc.ListIndex = 1
    Me.txtKeyPath.Text = ""
    Me.txtKeyName.Text = ""
    Me.txtkeyData.Text = "0"
    
UniMsgBox "D9a4 the6m va2o danh sa1ch chu71c na8ng ca62n ba3o ve65!", vbOKOnly, "OK!"

Else
UniMsgBox "Ba5n chu71a nha65o d9u3 tho6ng tin!", vbOKOnly, "!"
End If
End Sub

Private Sub cmdAddREGBack_Click()
Unload Me
End Sub

Private Sub Form_Load()
cbKeyGoc.AutoUnicode = False
cbKeyGoc.AddItem "HKEY_CLASSES_ROOT"
cbKeyGoc.AddItem "HKEY_CURRENT_USER"
cbKeyGoc.AddItem "HKEY_LOCAL_MACHINE"
cbKeyGoc.AddItem "HKEY_USERS"
cbKeyGoc.AddItem "HKEY_CURRENT_CONFIG"
cbKeyGoc.ListIndex = 1
End Sub

