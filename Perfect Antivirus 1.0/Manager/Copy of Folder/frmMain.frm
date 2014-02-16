VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PAV 2009 Manager"
   ClientHeight    =   5715
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8475
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5715
   ScaleWidth      =   8475
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   960
      Top             =   3960
   End
   Begin VB.ListBox lstPro 
      Height          =   2985
      Left            =   6960
      TabIndex        =   3
      Top             =   720
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   1920
      TabIndex        =   2
      Top             =   0
      Width           =   1095
   End
   Begin MSComctlLib.ImageList Ima 
      Left            =   3000
      Top             =   4680
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin VB.PictureBox Picture1 
      Height          =   255
      Left            =   360
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   0
      Top             =   3960
      Width           =   255
   End
   Begin MSComctlLib.ListView LV1 
      Height          =   3135
      Left            =   480
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   600
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   5530
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      SmallIcons      =   "Ima"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Image Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Path"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "PID"
         Object.Width           =   2540
      EndProperty
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
GetProcess LV1, Ima, Picture1
End Sub

Private Sub Timer_Timer()
  Dim i As Integer
  Dim theloop As Long
  Dim proc As PROCESSENTRY32
  Dim snap As Long
    i = 0
   snap = CreateToolhelpSnapshot(TH32CS_SNAPall, 0)
   proc.dwSize = Len(proc)
   theloop = ProcessFirst(snap, proc)
   While theloop <> 0

      theloop = ProcessNext(snap, proc)
      If ProcessPathByPID(proc.th32ProcessID) <> "SYSTEM" Then
              i = i + 1
              If i > LV.ListItems.Count Then GoTo KetThuc
            If LV.ListItems(i).SubItems(1) <> ProcessPathByPID(proc.th32ProcessID) Then GoTo KetThuc
      End If
   Wend
   CloseHandle snap
Exit Sub
KetThuc:
    GetProcess LV, Ima, Pic
End Sub

Private Sub Form_Load()
ThietLap LV1, Ima, Picture1
End Sub
