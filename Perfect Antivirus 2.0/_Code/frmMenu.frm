VERSION 5.00
Object = "{6DC1AB90-DCD1-47A1-AB36-924FFD67ADBF}#1.0#0"; "UniControls_v2.0.ocx"
Begin VB.Form frmMenu 
   Caption         =   "Form1"
   ClientHeight    =   2460
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   3195
   Icon            =   "frmMenu.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2460
   ScaleWidth      =   3195
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin UniControls.UniMenu UniMenu1 
      Left            =   960
      Top             =   1680
      _ExtentX        =   741
      _ExtentY        =   767
   End
   Begin UniControls.UniTrayIcon Tray 
      Left            =   0
      Top             =   0
      _ExtentX        =   529
      _ExtentY        =   529
      TooltipText     =   "Perfect Antivirus 2009"
      Icon            =   "frmMenu.frx":57E2
   End
   Begin VB.Menu scanvi 
      Caption         =   "ScanVirus"
      Begin VB.Menu Xcheckall 
         Caption         =   "D9a1nh da61u/Bo3 d9a1nh da61u"
      End
      Begin VB.Menu ngang0 
         Caption         =   "-"
      End
      Begin VB.Menu Goto 
         Caption         =   "D9i d9e61n thu7 mu5c go61c"
      End
      Begin VB.Menu properties 
         Caption         =   "Xem thuo65c ti1nh"
      End
      Begin VB.Menu Copy 
         Caption         =   "Copy ra ngoa2i Desktop"
      End
      Begin VB.Menu ngang1 
         Caption         =   "-"
      End
      Begin VB.Menu del 
         Caption         =   "Xo1a Virus d9ang cho5n"
      End
      Begin VB.Menu cachly 
         Caption         =   "Ca1ch ly Virus d9ang cho5n"
      End
      Begin VB.Menu add 
         Caption         =   "The6m va2o danh sa1ch tin tu7o73ng"
      End
      Begin VB.Menu ngang2 
         Caption         =   "-"
      End
      Begin VB.Menu report 
         Caption         =   "Xem ba3n ba1o ca1o"
      End
   End
End
Attribute VB_Name = "frmMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Copy_Click()
On Error Resume Next
FileCopy frmScan.LV1.ListItems(frmScan.LV1.SelectedItem.Index).SubItems(1).Caption, "C:\Documents and Settings\" & Environ("USERNAME") & "\Desktop\" & modMain.GetFileName(frmScan.LV1.ListItems(frmScan.LV1.SelectedItem.Index).SubItems(1).Caption)
UniMsgBox "D9a4 copy ra Desktop cu3a ba5n!", vbOKOnly, "OK!", Me.hwnd
End Sub

Private Sub del_Click()
On Error Resume Next
If UniMsgBox("Ba5n co1 muo61n xo1a file na2y kho6ng?" & vbCrLf & frmScan.LV1.ListItems(frmScan.LV1.SelectedItem.Index).SubItems(1).Caption, vbYesNo + vbQuestion, "Delete") = vbYes Then
    tXoaFile frmScan.LV1.ListItems(frmScan.LV1.SelectedItem.Index).SubItems(1).Caption
    If modMain.FileExists(frmScan.LV1.ListItems(frmScan.LV1.SelectedItem.Index).SubItems(1).Caption) = False Then
        frmScan.LV1.ListItems.Remove (frmScan.LV1.SelectedItem.Index)
    End If
End If
End Sub

Private Sub exit_Click()
frmMain.xThoat = 0
Unload frmMain
Unload frmRTP
Unload frmScan
Unload frmAutorun
Unload frmTangToc
Unload frmRegistry
Unload frmPhucHoiDuLieu
Unload frmFOLDERSCAN
Unload frmMenu
Unload Me
End Sub

Private Sub Form_Load()
UniMenu1.InitUnicodeMenu Me.hwnd
End Sub

Private Sub Goto_Click()
On Error Resume Next
Shell "explorer " & modMain.GetFolderPath(frmScan.LV1.ListItems(frmScan.LV1.SelectedItem.Index).SubItems(1).Caption), vbNormalFocus
End Sub

Private Sub Mai_Click(Index As Integer)
XShow_Click
frmMain.ShowMenu Index
End Sub

Private Sub properties_Click()
On Error Resume Next
ShowProps frmScan.LV1.ListItems(frmScan.LV1.SelectedItem.Index).SubItems(1).Caption, frmScan.hwnd
End Sub

Private Sub report_Click()
On Error Resume Next
Shell "notepad " & frmScan.xNowReport, vbNormalFocus
End Sub

Private Sub XShow_Click()
    If frmMain.xScanning = True Then
        frmScan.Show
        frmScan.Timer6.Enabled = False
    Else
        frmMain.Show
        App.TaskVisible = True
    End If
End Sub

Private Sub Tray_TrayClick(Button As UniControls.stMouseEvent)
If Button = stLeftButtonDoubleClick Then
    If frmMain.xScanning = True Then
        frmScan.Show
        frmScan.Timer6.Enabled = False
    Else
        frmMain.Show
        App.TaskVisible = True
    End If
End If
If Button = stRightButtonClick Then
    frmView.Show
End If
End Sub


Private Sub Xcheckall_Click()
On Error Resume Next
frmScan.CheckAll
End Sub

