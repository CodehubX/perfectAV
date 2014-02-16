VERSION 5.00
Object = "{D1ECD258-D329-4388-AB83-DEC261A66B86}#1.0#0"; "UniControls_v2.0.ocx"
Begin VB.Form frmMenu 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   4680
   Icon            =   "frmMenu.frx":0000
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
      Begin VB.Menu ngang0 
         Caption         =   "-"
      End
      Begin VB.Menu help 
         Caption         =   "Xem Hu7o71ng Da64n"
      End
      Begin VB.Menu quetvadiet 
         Caption         =   "Que1t && Die65t Virus"
         Begin VB.Menu scan 
            Caption         =   "Que1t Toa2n Bo65"
            Index           =   0
         End
         Begin VB.Menu scan 
            Caption         =   "Que1t Tu2y Cho5n"
            Index           =   1
         End
         Begin VB.Menu scan 
            Caption         =   "Ca61u Hi2nh Que1t"
            Index           =   3
         End
         Begin VB.Menu scan 
            Caption         =   "Nha65t Ky1 Que1t"
            Index           =   4
         End
      End
      Begin VB.Menu tudongbaove 
         Caption         =   "Tu75 D9o65ng Ba3o Ve65"
         Begin VB.Menu baoveregistry 
            Caption         =   "Ba3o ve65 Registry"
         End
         Begin VB.Menu baoveautorun 
            Caption         =   "Ba3o Ve65 Autorun"
         End
         Begin VB.Menu tudongquet 
            Caption         =   "Tu75 D9o65ng Que1t"
            Begin VB.Menu ats 
               Caption         =   "Kie63m Tra Nhu74ng File Sa81p Mo73"
               Index           =   0
            End
            Begin VB.Menu ats 
               Caption         =   "Ca3nh Ba1o Virus Trong Thu7 Mu5c Hie65n Ha2nh"
               Index           =   1
            End
            Begin VB.Menu ats 
               Caption         =   "Tu75 D9o65ng Que1t USB"
               Index           =   2
            End
            Begin VB.Menu ats 
               Caption         =   "Tu75 D9o65ng Pha1t Hie65n Keylogger"
               Index           =   3
            End
         End
      End
      Begin VB.Menu tienichhethong 
         Caption         =   "Tie65n I1ch He65 Tho61ng"
         Begin VB.Menu quanlytientrinh 
            Caption         =   "Qua3n Ly1 Tie61n Tri2nh"
         End
         Begin VB.Menu quanlyfile 
            Caption         =   "Qua3n Ly1 Ta65p Tin"
         End
         Begin VB.Menu quanlykhoidong 
            Caption         =   "Qua3n Ly1 Kho73i D9o65ng"
         End
         Begin VB.Menu kiemtrahethong 
            Caption         =   "Kie63m Tra He65 Tho61ng"
         End
         Begin VB.Menu tangtocmaytinh 
            Caption         =   "Tie65n I1ch"
         End
      End
      Begin VB.Menu caidatcauhinh 
         Caption         =   "Ca2i d9a85t ca61u hi2nh"
      End
      Begin VB.Menu tacgia 
         Caption         =   "Ta1c Gia3"
      End
      Begin VB.Menu ngang1 
         Caption         =   "-"
      End
      Begin VB.Menu exi 
         Caption         =   "Ta81t Chu7o7ng Tri2nh PAV 2009"
      End
      Begin VB.Menu ngang2 
         Caption         =   "-"
      End
      Begin VB.Menu hidemenu 
         Caption         =   "A63n Menu Na2y D9i"
      End
   End
End
Attribute VB_Name = "frmMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub ats_Click(index As Integer)

        '    frmMenu.ats(1).Checked = Me.atsAlwaysScanFolder.Value
        '    frmMenu.ats(0).Checked = Me.atsScanEXE.Value
        '    frmMenu.ats(2).Checked = Me.atsAutoScanUSB.Value
        '    frmMenu.ats(3).Checked = Me.atsScanKeylogger.Value
        '<EhHeader>
        On Error GoTo ats_Click_Err

        '</EhHeader>
100     If index = 0 Then

102         frmMain.atsScanEXE.Value = Not frmMain.atsScanEXE.Value

104         With frmMain.atsScanEXE

106             If .Value = True Then
108                 .Caption = "Tu75 d9o65ng que1t va2 ca3nh ba1o Virus cho nhu74ng File sa81p d9u7o75c mo73. [D9ang Ba65t]"
110                 WriteIniFile AppPath & "Setting.ini", "AutoScan", "ScanEXE", .Value
112                 SaveString HKEY_CLASSES_ROOT, "exefile\shell\open\command", "", ChrW(34) & AppPath & "Check_Virus.exe" & ChrW(34) & " " & ChrW(34) & "%1" & ChrW(34)
114                 SaveString HKEY_CLASSES_ROOT, "batfile\shell\open\command", "", ChrW(34) & AppPath & "Check_Virus.exe" & ChrW(34) & " " & ChrW(34) & "%1" & ChrW(34)
116                 SaveString HKEY_CLASSES_ROOT, "cmdfile\shell\open\command", "", ChrW(34) & AppPath & "Check_Virus.exe" & ChrW(34) & " " & ChrW(34) & "%1" & ChrW(34)
118                 SaveString HKEY_CLASSES_ROOT, "comfile\shell\open\command", "", ChrW(34) & AppPath & "Check_Virus.exe" & ChrW(34) & " " & ChrW(34) & "%1" & ChrW(34)
    
                Else
    
120                 .Caption = "Tu75 d9o65ng que1t va2 ca3nh ba1o Virus cho nhu74ng File sa81p d9u7o75c mo73. [D9ang Ta81t]"
122                 WriteIniFile AppPath & "Setting.ini", "AutoScan", "ScanEXE", .Value
124                 SaveString HKEY_CLASSES_ROOT, "exefile\shell\open\command", "", ChrW(34) & "%1" & ChrW(34) & " %*"
126                 SaveString HKEY_CLASSES_ROOT, "batfile\shell\open\command", "", ChrW(34) & "%1" & ChrW(34) & " %*"
128                 SaveString HKEY_CLASSES_ROOT, "cmdfile\shell\open\command", "", ChrW(34) & "%1" & ChrW(34) & " %*"
130                 SaveString HKEY_CLASSES_ROOT, "comfile\shell\open\command", "", ChrW(34) & "%1" & ChrW(34) & " %*"
    
                End If

            End With
    
132     ElseIf index = 1 Then
134         frmMain.atsAlwaysScanFolder.Value = Not frmMain.atsAlwaysScanFolder.Value

136         With frmMain.atsAlwaysScanFolder

138             If .Value = True Then
140                 .Caption = "Tu75 d9o65ng que1t va2 ca3nh ba1o Virus trong Thu7 mu5c d9ang d9u7o75c mo73. [D9ang Ba65t]"
142                 WriteIniFile AppPath & "Setting.ini", "AutoScan", "ScanFolder", .Value
144                 Load zfrmAutoScanFolder
                Else
146                 .Caption = "Tu75 d9o65ng que1t va2 ca3nh ba1o Virus trong Thu7 mu5c d9ang d9u7o75c mo73. [D9ang Ta81t]"
148                 WriteIniFile AppPath & "Setting.ini", "AutoScan", "ScanFolder", .Value
150                 Unload zfrmAutoScanFolder
    
                End If

            End With

152     ElseIf index = 2 Then
154         frmMain.atsAutoScanUSB.Value = Not frmMain.atsAutoScanUSB.Value

156         With frmMain.atsAutoScanUSB

158             If .Value = True Then
160                 .Caption = "Tu75 d9o65ng que1t Virus cho USB khi pha1t hie65n USB ke61t no61i va2o ma1y ti1nh. [D9ang Ba65t]"
162                 WriteIniFile AppPath & "Setting.ini", "AutoScan", "ScanUSB", .Value
                    '////
164                 Load zfrmScanUSB
                    '///
                Else
166                 .Caption = "Tu75 d9o65ng que1t Virus cho USB khi pha1t hie65n USB ke61t no61i va2o ma1y ti1nh. [D9ang Ta81t]"
168                 WriteIniFile AppPath & "Setting.ini", "AutoScan", "ScanUSB", .Value
                    '///
170                 Unload zfrmScanUSB
                    '///
    
                End If

            End With

172     ElseIf index = 3 Then
174         frmMain.atsScanKeylogger.Value = Not frmMain.atsScanKeylogger.Value

176         With frmMain.atsScanKeylogger

178             If .Value = True Then
180                 .Caption = "Tu75 d9o65ng pha1t hie65n va2 ca3nh ba1o Keylogger. [D9ang Ba65t]"
182                 WriteIniFile AppPath & "Setting.ini", "AutoScan", "ScanKeylogger", .Value
                    '////
184                 Load zfrmAntiKey
                    '///
                Else
186                 .Caption = "Tu75 d9o65ng pha1t hie65n va2 ca3nh ba1o Keylogger. [D9ang Ta81t]"
188                 WriteIniFile AppPath & "Setting.ini", "AutoScan", "ScanKeylogger", .Value
                    '///
190                 Unload zfrmAntiKey
                    '///
    
                End If

            End With

        End If

192     frmMenu.ats(1).Checked = frmMain.atsAlwaysScanFolder.Value
194     frmMenu.ats(0).Checked = frmMain.atsScanEXE.Value
196     frmMenu.ats(2).Checked = frmMain.atsAutoScanUSB.Value
198     frmMenu.ats(3).Checked = frmMain.atsScanKeylogger.Value

        '<EhFooter>
        Exit Sub

ats_Click_Err:
        WriteErr Time & "-" & Date & " - " & Err.Description & " - " & "in PerfectAntivirus2009.frmMenu.ats_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub baoveautorun_Click()

        '<EhHeader>
        On Error GoTo baoveautorun_Click_Err

        '</EhHeader>
100     If frmMain.atptmrAutorun.Enabled = True Then
102         frmMain.atptmrAutorun.Enabled = False
104         frmMain.atplblStaAutorun.Caption = "D9ang ta81t"
106         frmMain.atpcmdAutorun.Caption = "Mo73 chu71c na8ng na2y"
        Else
108         frmMain.atptmrAutorun.Enabled = True
110         frmMain.atplblStaAutorun.Caption = "D9ang mo73"
112         frmMain.atpcmdAutorun.Caption = "Ta81t chu71c na8ng na2y"
        End If

114     baoveautorun.Checked = frmMain.atptmrAutorun.Enabled

        '<EhFooter>
        Exit Sub

baoveautorun_Click_Err:
        WriteErr Time & "-" & Date & " - " & Err.Description & " - " & "in PerfectAntivirus2009.frmMenu.baoveautorun_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub baoveregistry_Click()

        '<EhHeader>
        On Error GoTo baoveregistry_Click_Err

        '</EhHeader>
100     If frmMain.atptmrREG.Enabled = True Then
102         frmMain.atptmrREG.Enabled = False
104         frmMain.atplbREG.Caption = "D9ang ta81t"
106         frmMain.atpcmdREG.Caption = "Mo73 chu71c na8ng na2y"
        Else
108         frmMain.atptmrREG.Enabled = True
110         frmMain.atplbREG.Caption = "D9ang mo73"
112         frmMain.atpcmdREG.Caption = "Ta81t chu71c na8ng na2y"
        End If

114     baoveregistry.Checked = frmMain.atptmrREG.Enabled

        '<EhFooter>
        Exit Sub

baoveregistry_Click_Err:
        WriteErr Time & "-" & Date & " - " & Err.Description & " - " & "in PerfectAntivirus2009.frmMenu.baoveregistry_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub caidatcauhinh_Click()

        '<EhHeader>
        On Error GoTo caidatcauhinh_Click_Err

        '</EhHeader>

100     sho_Click
102     frmMain.HideAllFM
104     frmMain.fm(11).Visible = True

        '<EhFooter>
        Exit Sub

caidatcauhinh_Click_Err:
        WriteErr Time & "-" & Date & " - " & Err.Description & " - " & "in PerfectAntivirus2009.frmMenu.caidatcauhinh_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub exi_Click()

        '<EhHeader>
        On Error GoTo exi_Click_Err

        '</EhHeader>
100     If UniMsgBox("Ba5n muo61n ta81t chu7o7ng tri2nh PAV 2009?", vbYesNo + vbQuestion) = vbNo Then Exit Sub

        '===> Save Scan Virus Setting
102     If FileExists(AppPath & "Setting.ini") = True Then modScanVirus.DeleteFile (AppPath & "Setting.ini")
104     WriteIniFile AppPath & "Setting.ini", "ScanVirus", frmMain.VSbat.Name, frmMain.VSbat.Value
106     WriteIniFile AppPath & "Setting.ini", "ScanVirus", frmMain.VScmd.Name, frmMain.VScmd.Value
108     WriteIniFile AppPath & "Setting.ini", "ScanVirus", frmMain.VScom.Name, frmMain.VScom.Value
110     WriteIniFile AppPath & "Setting.ini", "ScanVirus", frmMain.VSdll.Name, frmMain.VSdll.Value
112     WriteIniFile AppPath & "Setting.ini", "ScanVirus", frmMain.VSscr.Name, frmMain.VSscr.Value

114     WriteIniFile AppPath & "Setting.ini", "ScanVirus", frmMain.VSScanProcess.Name, frmMain.VSScanProcess.Value
116     WriteIniFile AppPath & "Setting.ini", "ScanVirus", frmMain.VSScanStartUp.Name, frmMain.VSScanStartUp.Value
118     WriteIniFile AppPath & "Setting.ini", "ScanVirus", frmMain.VSDontScanSize.Name, frmMain.VSDontScanSize.Value
120     WriteIniFile AppPath & "Setting.ini", "ScanVirus", frmMain.VSLimitSize.Name, frmMain.VSLimitSize.Value

122     WriteIniFile AppPath & "Setting.ini", "AutoScan", "ScanFolder", frmMain.atsAlwaysScanFolder.Value
124     WriteIniFile AppPath & "Setting.ini", "AutoScan", "ScanEXE", frmMain.atsScanEXE.Value
126     WriteIniFile AppPath & "Setting.ini", "AutoScan", "ScanUSB", frmMain.atsAutoScanUSB.Value
128     WriteIniFile AppPath & "Setting.ini", "AutoScan", "ScanKeylogger", frmMain.atsScanKeylogger.Value

130     WriteIniFile AppPath & "Setting.ini", "AutoProtect", "Registry", frmMain.atptmrREG.Enabled
132     WriteIniFile AppPath & "Setting.ini", "AutoProtect", "Autorun", frmMain.atptmrAutorun.Enabled
134     WriteIniFile AppPath & "Setting.ini", "AutoProtect", "AutoAddVirus", frmMain.chkAutoAddAutorun.Value

136     WriteIniFile AppPath & "Setting.ini", "Setting", "AutoStart", frmMain.chkAutoStart.Value
138     WriteIniFile AppPath & "Setting.ini", "Setting", "FlashScreen", frmMain.chkShowFlash.Value
140     WriteIniFile AppPath & "Setting.ini", "Setting", "AutoUpdate", frmMain.chkAutoUpdate.Value
142     WriteIniFile AppPath & "Setting.ini", "Setting", "MiniScan", frmMain.chkUseFastScan.Value

        '<=== Save Scan Virus Setting

        '===> Virus Events
144     If FileExists(AppPath & "VirusScanLog.log") = True Then modScanVirus.DeleteFile (AppPath & "VirusScanLog.log")

146     With frmMain.LVVirusEvents

            Dim l

148         For l = 1 To .ListItems.Count
150             WriteIniFile AppPath & "VirusScanLog.log", l, "TimeQuet", .ListItems(l).Text
152             WriteIniFile AppPath & "VirusScanLog.log", l, "KieuQuet", Unicode2UTF8(.ListItems(l).SubItems(1).Caption)
154             WriteIniFile AppPath & "VirusScanLog.log", l, "SoFile", .ListItems(l).SubItems(2).Caption
156             WriteIniFile AppPath & "VirusScanLog.log", l, "SoVirus", .ListItems(l).SubItems(3).Caption
158             WriteIniFile AppPath & "VirusScanLog.log", l, "KetQua", Unicode2UTF8(.ListItems(l).SubItems(4).Caption)
160         Next l

162         WriteIniFile AppPath & "VirusScanLog.log", "Other", "Total", .ListItems.Count
    
        End With

        '<=== Virus Events

164     End

        '<EhFooter>
        Exit Sub

exi_Click_Err:
        WriteErr Time & "-" & Date & " - " & Err.Description & " - " & "in PerfectAntivirus2009.frmMenu.exi_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Form_Load()

        '<EhHeader>
        On Error GoTo Form_Load_Err

        '</EhHeader>

100     UniMenu1.InitUnicodeMenu

        '<EhFooter>
        Exit Sub

Form_Load_Err:
        WriteErr Time & "-" & Date & " - " & Err.Description & " - " & "in PerfectAntivirus2009.frmMenu.Form_Load " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub help_Click()

        '<EhHeader>
        On Error GoTo help_Click_Err

        '</EhHeader>

100     ShellExecute Me.hWnd, vbNullString, AppPath & "Help\HuongDan.html", vbNullString, "", 1

        '<EhFooter>
        Exit Sub

help_Click_Err:
        WriteErr Time & "-" & Date & " - " & Err.Description & " - " & "in PerfectAntivirus2009.frmMenu.help_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub hidemenu_Click()

        '<EhHeader>
        On Error GoTo hidemenu_Click_Err

        '</EhHeader>

100     frmMenu.Hide

        '<EhFooter>
        Exit Sub

hidemenu_Click_Err:
        WriteErr Time & "-" & Date & " - " & Err.Description & " - " & "in PerfectAntivirus2009.frmMenu.hidemenu_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub kiemtrahethong_Click()

        '<EhHeader>
        On Error GoTo kiemtrahethong_Click_Err

        '</EhHeader>
100     If FileExists(AppPath & "PAVSysReport.exe") = True Then
102         If FileLen(AppPath & "PAVSysReport.exe") = 35328 Then
104             Shell AppPath & "PAVSysReport.exe syscheck", vbNormalFocus
            End If

        Else
106         UniMsgBox "Kho6ng ti2m tha61y File PAVSysReport.exe.", vbOKOnly + vbCritical, "Error!"
        End If

        '<EhFooter>
        Exit Sub

kiemtrahethong_Click_Err:
        WriteErr Time & "-" & Date & " - " & Err.Description & " - " & "in PerfectAntivirus2009.frmMenu.kiemtrahethong_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub quanlyfile_Click()

        '<EhHeader>
        On Error GoTo quanlyfile_Click_Err

        '</EhHeader>
100     If FileExists(AppPath & "PAVExplorer.exe") = True Then
102         If FileLen(AppPath & "PAVExplorer.exe") = 82432 Then
104             Shell AppPath & "PAVExplorer.exe explorer", vbNormalFocus
            End If

        Else
106         UniMsgBox "Kho6ng ti2m tha61y File PAVExplorer.exe.", vbOKOnly + vbCritical, "Error!"
        End If

        '<EhFooter>
        Exit Sub

quanlyfile_Click_Err:
        WriteErr Time & "-" & Date & " - " & Err.Description & " - " & "in PerfectAntivirus2009.frmMenu.quanlyfile_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub quanlykhoidong_Click()

        '<EhHeader>
        On Error GoTo quanlykhoidong_Click_Err

        '</EhHeader>
100     If FileExists(AppPath & "PAVStartUp.exe") = True Then
102         If FileLen(AppPath & "PAVStartUp.exe") = 138240 Then
104             Shell AppPath & "PAVStartUp.exe startup", vbNormalFocus
            End If

        Else
106         UniMsgBox "Kho6ng ti2m tha61y File PAVStartUp.exe.", vbOKOnly + vbCritical, "Error!"
        End If

        '<EhFooter>
        Exit Sub

quanlykhoidong_Click_Err:
        WriteErr Time & "-" & Date & " - " & Err.Description & " - " & "in PerfectAntivirus2009.frmMenu.quanlykhoidong_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub quanlytientrinh_Click()

        '<EhHeader>
        On Error GoTo quanlytientrinh_Click_Err

        '</EhHeader>
100     If FileExists(AppPath & "PAV2009Manager.exe") = True Then
102         If FileLen(AppPath & "PAV2009Manager.exe") = 53760 Then
104             Shell AppPath & "PAV2009Manager.exe quangtrung", vbNormalFocus
            End If

        Else
106         UniMsgBox "Kho6ng ti2m tha61y File PAV2009Manager.exe.", vbOKOnly + vbCritical, "Error!"
        End If

        '<EhFooter>
        Exit Sub

quanlytientrinh_Click_Err:
        WriteErr Time & "-" & Date & " - " & Err.Description & " - " & "in PerfectAntivirus2009.frmMenu.quanlytientrinh_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub scan_Click(index As Integer)

        '<EhHeader>
        On Error GoTo scan_Click_Err

        '</EhHeader>

100     sho_Click
102     frmMain.HideAllFM
104     frmMain.fm(index).Visible = True

        '<EhFooter>
        Exit Sub

scan_Click_Err:
        WriteErr Time & "-" & Date & " - " & Err.Description & " - " & "in PerfectAntivirus2009.frmMenu.scan_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub sho_Click()

        '<EhHeader>
        On Error GoTo sho_Click_Err

        '</EhHeader>

100     frmMain.Visible = True
102     frmMain.Show
104     frmMain.WindowState = 0
106     App.TaskVisible = True

        '<EhFooter>
        Exit Sub

sho_Click_Err:
        WriteErr Time & "-" & Date & " - " & Err.Description & " - " & "in PerfectAntivirus2009.frmMenu.sho_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub tacgia_Click()

        '<EhHeader>
        On Error GoTo tacgia_Click_Err

        '</EhHeader>

100     sho_Click
102     frmMain.HideAllFM
104     frmMain.fm(12).Visible = True

        '<EhFooter>
        Exit Sub

tacgia_Click_Err:
        WriteErr Time & "-" & Date & " - " & Err.Description & " - " & "in PerfectAntivirus2009.frmMenu.tacgia_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub tangtocmaytinh_Click()

        '<EhHeader>
        On Error GoTo tangtocmaytinh_Click_Err

        '</EhHeader>

100     sho_Click
102     frmMain.HideAllFM
104     frmMain.fm(13).Visible = True

        '<EhFooter>
        Exit Sub

tangtocmaytinh_Click_Err:
        WriteErr Time & "-" & Date & " - " & Err.Description & " - " & "in PerfectAntivirus2009.frmMenu.tangtocmaytinh_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub
