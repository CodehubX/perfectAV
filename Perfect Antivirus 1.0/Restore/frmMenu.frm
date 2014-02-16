VERSION 5.00
Object = "{D1ECD258-D329-4388-AB83-DEC261A66B86}#1.0#0"; "UniControls_v2.0.ocx"
Begin VB.Form frmMenu 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   210
   ClientTop       =   780
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin UniControls.UniMenu UniMenu1 
      Left            =   2400
      Top             =   1800
      _ExtentX        =   741
      _ExtentY        =   767
   End
   Begin VB.Menu m 
      Caption         =   "Main"
      Begin VB.Menu sho 
         Caption         =   "&Hie63n Thi5"
      End
      Begin VB.Menu ngang 
         Caption         =   "-"
      End
      Begin VB.Menu exi 
         Caption         =   "&Thoa1t"
      End
   End
End
Attribute VB_Name = "frmMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub exi_Click()
If UniMsgBox("Ba5n muo61n ta81t chu7o7ng tri2nh PAV 2009?", vbYesNo) = vbNo Then Exit Sub

'===> Save Scan Virus Setting
If FileExists(AppPath & "Setting.ini") = True Then modScanVirus.DeleteFile (AppPath & "Setting.ini")
WriteIniFile AppPath & "Setting.ini", "ScanVirus", frmMain.VSbat.Name, frmMain.VSbat.Value
WriteIniFile AppPath & "Setting.ini", "ScanVirus", frmMain.VScmd.Name, frmMain.VScmd.Value
WriteIniFile AppPath & "Setting.ini", "ScanVirus", frmMain.VScom.Name, frmMain.VScom.Value
WriteIniFile AppPath & "Setting.ini", "ScanVirus", frmMain.VSdll.Name, frmMain.VSdll.Value
WriteIniFile AppPath & "Setting.ini", "ScanVirus", frmMain.VSscr.Name, frmMain.VSscr.Value

WriteIniFile AppPath & "Setting.ini", "ScanVirus", frmMain.VSScanProcess.Name, frmMain.VSScanProcess.Value
WriteIniFile AppPath & "Setting.ini", "ScanVirus", frmMain.VSScanStartUp.Name, frmMain.VSScanStartUp.Value
WriteIniFile AppPath & "Setting.ini", "ScanVirus", frmMain.VSDontScanSize.Name, frmMain.VSDontScanSize.Value
WriteIniFile AppPath & "Setting.ini", "ScanVirus", frmMain.VSLimitSize.Name, frmMain.VSLimitSize.Value


WriteIniFile AppPath & "Setting.ini", "AutoScan", "ScanFolder", frmMain.atsAlwaysScanFolder.Value
WriteIniFile AppPath & "Setting.ini", "AutoScan", "ScanEXE", frmMain.atsScanEXE.Value
WriteIniFile AppPath & "Setting.ini", "AutoScan", "ScanUSB", frmMain.atsAutoScanUSB.Value
WriteIniFile AppPath & "Setting.ini", "AutoScan", "ScanKeylogger", frmMain.atsScanKeylogger.Value

WriteIniFile AppPath & "Setting.ini", "AutoProtect", "Registry", frmMain.atptmrREG.Enabled
WriteIniFile AppPath & "Setting.ini", "AutoProtect", "Autorun", frmMain.atptmrAutorun.Enabled


'<=== Save Scan Virus Setting



'===> Virus Events
If FileExists(AppPath & "VirusScanLog.log") = True Then modScanVirus.DeleteFile (AppPath & "VirusScanLog.log")
With frmMain.LVVirusEvents
    Dim l
    For l = 1 To .ListItems.Count
        WriteIniFile AppPath & "VirusScanLog.log", l, "TimeQuet", .ListItems(l).Text
        WriteIniFile AppPath & "VirusScanLog.log", l, "KieuQuet", Unicode2UTF8(.ListItems(l).SubItems(1).Caption)
        WriteIniFile AppPath & "VirusScanLog.log", l, "SoFile", .ListItems(l).SubItems(2).Caption
        WriteIniFile AppPath & "VirusScanLog.log", l, "SoVirus", .ListItems(l).SubItems(3).Caption
        WriteIniFile AppPath & "VirusScanLog.log", l, "KetQua", Unicode2UTF8(.ListItems(l).SubItems(4).Caption)
    Next l
    WriteIniFile AppPath & "VirusScanLog.log", "Other", "Total", .ListItems.Count
    
End With

'<=== Virus Events

End
End Sub

Private Sub Form_Load()
UniMenu1.InitUnicodeMenu
End Sub

Private Sub sho_Click()
frmMain.Visible = True
frmMain.Show
frmMain.WindowState = 0
App.TaskVisible = True
End Sub
