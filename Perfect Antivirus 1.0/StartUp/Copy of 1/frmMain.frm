VERSION 5.00
Object = "{D1ECD258-D329-4388-AB83-DEC261A66B86}#1.0#0"; "UniControls_v2.0.ocx"
Begin VB.Form frmMain 
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PAV Start Up"
   ClientHeight    =   5535
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9330
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
   ScaleHeight     =   5535
   ScaleWidth      =   9330
   StartUpPosition =   2  'CenterScreen
   Begin VB.FileListBox File1 
      Height          =   195
      Left            =   120
      TabIndex        =   8
      Top             =   4860
      Visible         =   0   'False
      Width           =   105
   End
   Begin UniControls.UniListView LV 
      Height          =   3255
      Left            =   3120
      TabIndex        =   7
      Top             =   720
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   5741
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MultiSelect     =   -1  'True
      LabelEdit       =   0   'False
      AutoArrange     =   0   'False
      BorderStyle     =   2
      HeaderButtons   =   0   'False
      HeaderTrackSelect=   0   'False
      HideSelection   =   0   'False
      InfoTips        =   0   'False
   End
   Begin VB.PictureBox Ico 
      Height          =   375
      Index           =   5
      Left            =   360
      Picture         =   "frmMain.frx":169B2
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   6
      Top             =   3720
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox Ico 
      Height          =   375
      Index           =   4
      Left            =   360
      Picture         =   "frmMain.frx":16F3C
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   5
      Top             =   3240
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox Ico 
      Height          =   375
      Index           =   3
      Left            =   720
      Picture         =   "frmMain.frx":174C6
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   4
      Top             =   2760
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox Ico 
      Height          =   375
      Index           =   2
      Left            =   480
      Picture         =   "frmMain.frx":17A50
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   3
      Top             =   2280
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox Ico 
      Height          =   375
      Index           =   1
      Left            =   360
      Picture         =   "frmMain.frx":17FDA
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   2
      Top             =   1800
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox Ico 
      Height          =   375
      Index           =   0
      Left            =   120
      Picture         =   "frmMain.frx":18564
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   1
      Top             =   1320
      Visible         =   0   'False
      Width           =   375
   End
   Begin UniControls.UniTreeView Tree1 
      Height          =   3255
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   5741
   End
   Begin UniControls.UniLabel UniLabel1 
      Height          =   495
      Left            =   120
      Top             =   120
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   873
      Alignment       =   1
      Caption         =   "Qua3n ly1 kho73i d9o65ng"
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
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

    File1.System = True
    File1.Hidden = True
    File1.ReadOnly = True
    File1.Archive = True
    File1.Pattern = "*.*"
    
    
With Tree1
    .Initialize
    .InitializeImageList
    .AddIcon Ico(0).Picture
    .AddIcon Ico(1).Picture
    .AddIcon Ico(2).Picture
    .AddIcon Ico(3).Picture
    .AddIcon Ico(4).Picture
    .AddIcon Ico(5).Picture
    
    
    .AddNode , , "a", "Ta61t ca3 ca1c tri2nh kho73i d9o65ng", 0, 0
        .AddNode "a", , "aReg", "Tu72 Registry", 1, 1
            .AddNode "aReg", , "aRegAdmin", "Ta61t ca3 ngu7o72i du2ng", 2, 2
                .AddNode "aRegAdmin", , "aRegAdmin0", "Cha5y", 3, 3
                .AddNode "aRegAdmin", , "aRegAdmin1", "Cha5y 1 La62n", 3, 3
            .AddNode "aReg", , "aRegUser", Environ$("USERNAME"), 2, 2
                .AddNode "aRegUser", , "aRegUser0", "Cha5y", 3, 3
                .AddNode "aRegUser", , "aRegUser1", "Cha5y 1 La62n", 3, 3
        .AddNode "a", , "aFol", "Tu72 Thu7 Mu5c Kho73i D9o65ng", 4, 4
            .AddNode "aFol", , "aFolAdmin", "Ta61t ca3 ngu7o72i du2ng", 2, 2
            .AddNode "aFol", , "aFolUser", Environ$("USERNAME"), 2, 2
        .AddNode "a", , "aSys", "Kho1a He65 Tho61ng", 5, 5
        
        
    .Expand .GetKeyNode(a), True
End With


With LV
    .View = eViewDetails
    .FullRowSelect = True
    .GridLines = True
    .AutoUnicode = False
    .CheckBoxes = True
    .MultiSelect = False
    
    .Columns.Add , , ToUnicode("Te6n Chu7o7ng Tri2nh"), , 2000
    .Columns.Add , , ToUnicode("D9i5a Chi3"), , 4000
    .Columns.Add , , ToUnicode("D9u7o72ng Da64n Key"), , 4000

End With



GetKeyValue "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Run"
GetKeyValue "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\RunOnce"
GetKeyValue "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Run"
GetKeyValue "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\RunOnce"
GetKeyValue "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\RunOnceEx"

GetSystemKey
GetFolderStartUp 1
GetFolderStartUp 2

End Sub

