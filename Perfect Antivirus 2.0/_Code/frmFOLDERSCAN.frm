VERSION 5.00
Begin VB.Form frmFOLDERSCAN 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PAV_FOLDER_SCAN_VIRUS"
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
   Icon            =   "frmFOLDERSCAN.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdScanFolder 
      Caption         =   "ScanVirusForFolderNow"
      Height          =   345
      Left            =   240
      TabIndex        =   0
      Top             =   1080
      Width           =   2175
   End
End
Attribute VB_Name = "frmFOLDERSCAN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdScanFolder_Click()
On Error Resume Next
frmScan.Show
frmScan.lstPath.AddItem cmdScanFolder.Caption
frmScan.RefereshListPath
End Sub

Private Sub Form_Load()
On Error Resume Next
Me.Hide
Me.Visible = False
End Sub
