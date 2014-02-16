VERSION 5.00
Object = "{D1ECD258-D329-4388-AB83-DEC261A66B86}#1.0#0"; "UniControls_v2.0.ocx"
Begin VB.Form zfrmScanUSB 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PAV 2009 - Thong Bao !!!"
   ClientHeight    =   5295
   ClientLeft      =   345
   ClientTop       =   510
   ClientWidth     =   7065
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "zfrmScanUSB.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5295
   ScaleWidth      =   7065
   StartUpPosition =   2  'CenterScreen
   Begin UniControls.UniButton cmdUSBKill 
      Height          =   375
      Left            =   3000
      TabIndex        =   2
      Top             =   4800
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      Icon            =   "zfrmScanUSB.frx":57E2
      Style           =   2
      Caption         =   "Die65t"
      IconAlign       =   3
      iNonThemeStyle  =   2
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
   Begin UniControls.UniFrame F1 
      Height          =   5295
      Left            =   0
      Top             =   0
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   9340
      MaskColor       =   16711935
      Style           =   2
      Caption         =   "PAV 2009 - Ba3o ve65 USB"
      TextColor       =   16711680
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Icon            =   "zfrmScanUSB.frx":5D7C
      IconSize        =   32
      ThemeColor      =   4
      Begin UniControls.UniButton cmdStopScan 
         Height          =   375
         Left            =   1680
         TabIndex        =   4
         Top             =   4800
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         Icon            =   "zfrmScanUSB.frx":630E
         Style           =   2
         Caption         =   "Du72ng"
         IconAlign       =   3
         iNonThemeStyle  =   2
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
      Begin UniControls.UniLabel lblUSBStatus 
         Height          =   495
         Left            =   240
         Top             =   4320
         Width           =   6615
         _ExtentX        =   11668
         _ExtentY        =   873
         AutoUnicode     =   0   'False
         BackStyle       =   0
         Caption         =   ""
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
      Begin UniControls.UniButton cmdUSBScan 
         Height          =   375
         Left            =   240
         TabIndex        =   1
         Top             =   4800
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         Icon            =   "zfrmScanUSB.frx":68A8
         Style           =   2
         Caption         =   "Que1t"
         IconAlign       =   3
         iNonThemeStyle  =   2
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
      Begin UniControls.UniListView LVVirus4 
         Height          =   2535
         Left            =   240
         TabIndex        =   0
         Top             =   1680
         Width           =   6615
         _ExtentX        =   11668
         _ExtentY        =   4471
         AutoUnicode     =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   12648447
         MultiSelect     =   -1  'True
         LabelEdit       =   0   'False
         FullRowSelect   =   -1  'True
         AutoArrange     =   0   'False
         HeaderButtons   =   0   'False
         HeaderTrackSelect=   0   'False
         HideSelection   =   0   'False
         InfoTips        =   0   'False
      End
      Begin UniControls.UniLabel lblUSBPath 
         Height          =   255
         Left            =   960
         Top             =   1080
         Width           =   6015
         _ExtentX        =   10610
         _ExtentY        =   450
         BackStyle       =   0
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   255
      End
      Begin UniControls.UniLabel UniLabel4 
         Height          =   255
         Left            =   240
         Top             =   1320
         Width           =   6375
         _ExtentX        =   11245
         _ExtentY        =   450
         BackStyle       =   0
         Caption         =   "Ha4y que1t Virus cho USB d9e63 d9a3m ba3o an toa2n cho ma1y ti1nh cu3a ba5n."
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   16711680
      End
      Begin UniControls.UniLabel UniLabel3 
         Height          =   255
         Left            =   120
         Top             =   1080
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   450
         BackStyle       =   0
         Caption         =   "D9i5a chi3:"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   16711680
      End
      Begin UniControls.UniLabel UniLabel1 
         Height          =   375
         Left            =   120
         Top             =   480
         Width           =   6495
         _ExtentX        =   11456
         _ExtentY        =   661
         Alignment       =   1
         BackStyle       =   0
         Caption         =   "Co1 USB d9ang ke61t no61i va2o ma1y ti1nh!"
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
      Begin UniControls.UniButton cmdUSBCancel 
         Height          =   375
         Left            =   5400
         TabIndex        =   3
         Top             =   4800
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         Icon            =   "zfrmScanUSB.frx":72BA
         Style           =   2
         Caption         =   "Bo3 Qua"
         IconAlign       =   3
         iNonThemeStyle  =   2
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
   End
End
Attribute VB_Name = "zfrmScanUSB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
   
Private Declare Function FindFirstFile _
                Lib "kernel32" _
                Alias "FindFirstFileA" (ByVal lpFileName As String, _
                                        lpFindFileData As WIN32_FIND_DATA) As Long
   
Private Declare Function FindNextFile _
                Lib "kernel32" _
                Alias "FindNextFileA" (ByVal hFindFile As Long, _
                                       lpFindFileData As WIN32_FIND_DATA) As Long

Private Declare Function GetTickCount Lib "kernel32" () As Long

Private Declare Function BringWindowToTop Lib "user32" (ByVal hWnd As Long) As Long

Private Declare Function RegisterDeviceNotification _
                Lib "User32.dll" _
                Alias "RegisterDeviceNotificationA" (ByVal hRecipient As Long, _
                                                     ByRef NotificationFilter As Any, _
                                                     ByVal Flags As Long) As Long

Private Declare Function UnregisterDeviceNotification _
                Lib "User32.dll" (ByVal Handle As Long) As Long
    
Private Declare Function SendMessage _
                Lib "user32" _
                Alias "SendMessageA" (ByVal hWnd As Long, _
                                      ByVal wMsg As Long, _
                                      ByVal wParam As Long, _
                                      lParam As Any) As Long

Private Declare Function SetWindowPos& _
                Lib "user32" (ByVal hWnd As Long, _
                              ByVal hWndInsertAfter As Long, _
                              ByVal X As Long, _
                              ByVal Y As Long, _
                              ByVal cx As Long, _
                              ByVal cy As Long, _
                              ByVal wFlags As Long)

Private Declare Function SetForegroundWindow Lib "user32" (ByVal hWnd As Long) As Long

Private Declare Sub ReleaseCapture Lib "user32" ()

Const WM_NCLBUTTONDOWN = &HA1

Const HTCAPTION = 2
    
Private Type DEV_BROADCAST_DEVICEINTERFACE

    dbcc_size As Long
    dbcc_devicetype As Long
    dbcc_reserved As Long
    dbcc_classguid As Guid
    dbcc_name As Long

End Type

Private lDevNotify                                As Long

Private Const DEVICE_NOTIFY_WINDOW_HANDLE         As Long = &H0

Private Const DBT_DEVTYP_DEVICEINTERFACE          As Long = &H5

Private Const DEVICE_NOTIFY_ALL_INTERFACE_CLASSES As Long = &H4

Option Explicit

Private Const vbDot = 46

Private Const MAXDWORD As Long = &HFFFFFFFF

Private Const MAX_PATH As Long = 260

Private Const INVALID_HANDLE_VALUE = -1

Private Const FILE_ATTRIBUTE_DIRECTORY = &H10

Private Type FILETIME

    dwLowDateTime As Long
    dwHighDateTime As Long

End Type

Private Type WIN32_FIND_DATA

    dwFileAttributes As Long
    ftCreationTime As FILETIME
    ftLastAccessTime As FILETIME
    ftLastWriteTime As FILETIME
    nFileSizeHigh As Long
    nFileSizeLow As Long
    dwReserved0 As Long
    dwReserved1 As Long
    cFileName As String * MAX_PATH
    cAlternate As String * 14

End Type

Private Type FILE_PARAMS

    bRecurse As Boolean
    sFileRoot As String
    sFileNameExt As String
    sResult As String
    sMatches As String
    Count As Long

End Type

Dim xStopScan As Boolean

Private Function SearchFile(Path, FileName)

        '<EhHeader>
        On Error GoTo SearchFile_Err

        '</EhHeader>

        Dim FP     As FILE_PARAMS  'holds search parameters

        Dim tstart As Single   'timer var for this routine only

        Dim tend   As Single     'timer var for this routine only

100     With FP
102         .sFileRoot = Path       'start path
104         .sFileNameExt = FileName    'file type of interest
106         .bRecurse = 1 ' Check1.Value = 1  '1 = recursive search
        End With

108     tstart = GetTickCount()
110     Call SearchForFiles(FP)
112     tend = GetTickCount()

        '<EhFooter>
        Exit Function

SearchFile_Err:
        WriteErr Time & "-" & Date & " - " & Err.Description & " - " & "in PerfectAntivirus2009.zfrmScanUSB.SearchFile " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Function

Private Sub GetFileInformation(FP As FILE_PARAMS)

        '<EhHeader>
        On Error GoTo GetFileInformation_Err

        '</EhHeader>
100     DoEvents

        Dim WFD     As WIN32_FIND_DATA

        Dim hFile   As Long

        Dim sPath   As String

        Dim sRoot   As String

        Dim sTmp    As String

        Dim SKetQua As String

102     sRoot = QualifyPath(FP.sFileRoot)
104     sPath = sRoot & FP.sFileNameExt
106     hFile = FindFirstFile(sPath, WFD)

108     If hFile <> INVALID_HANDLE_VALUE Then

            Do

110             If Not (WFD.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) = FILE_ATTRIBUTE_DIRECTORY Then
112                 FP.Count = FP.Count + 1
114                 sTmp = TrimNull(WFD.cFileName)
116                 SKetQua = sRoot & sTmp

118                 If xStopScan = True Then Exit Sub
120                 If FileLen(SKetQua) < 10000000 And FileLen(SKetQua) > 0 Then
122                     Me.lblUSBStatus.Caption = SKetQua

                        Dim AX As String

124                     AX = modScanVirus.CheckVirus(SKetQua)

126                     If AX <> "No" Then

                            Dim I

128                         I = LVVirus4.ListItems.Count + 1
130                         LVVirus4.ListItems.Add I, , AX
132                         LVVirus4.ListItems(I).SubItems(1).Caption = SKetQua
134                         LVVirus4.ListItems(I).Checked = True
136                         LVVirus4.Refresh
                        End If
                    End If
            
                    'Text1.Text = Text1.Text & SKetQua & vbCrLf
                    '*********************************************
                End If

138         Loop While FindNextFile(hFile, WFD)

140         hFile = FindClose(hFile)
        End If

142     DoEvents

        '<EhFooter>
        Exit Sub

GetFileInformation_Err:
        WriteErr Time & "-" & Date & " - " & Err.Description & " - " & "in PerfectAntivirus2009.zfrmScanUSB.GetFileInformation " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub SearchForFiles(FP As FILE_PARAMS)

        'local working variables
        '<EhHeader>
        On Error GoTo SearchForFiles_Err

        '</EhHeader>
        Dim WFD   As WIN32_FIND_DATA

        Dim hFile As Long

        Dim sPath As String

        Dim sRoot As String

        Dim sTmp  As String

100     sRoot = QualifyPath(FP.sFileRoot)
102     sPath = sRoot & "*.*"
104     hFile = FindFirstFile(sPath, WFD)

106     If hFile <> INVALID_HANDLE_VALUE Then
108         Call GetFileInformation(FP)

            Do

110             If (WFD.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) Then
112                 If FP.bRecurse Then
114                     If Asc(WFD.cFileName) <> vbDot Then
116                         FP.sFileRoot = sRoot & TrimNull(WFD.cFileName)
118                         Call SearchForFiles(FP)
                        End If
                    End If
                End If

120         Loop While FindNextFile(hFile, WFD)

122         hFile = FindClose(hFile)
        End If

        '<EhFooter>
        Exit Sub

SearchForFiles_Err:
        WriteErr Time & "-" & Date & " - " & Err.Description & " - " & "in PerfectAntivirus2009.zfrmScanUSB.SearchForFiles " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Function QualifyPath(sPath As String) As String

        '<EhHeader>
        On Error GoTo QualifyPath_Err

        '</EhHeader>
100     If Right$(sPath, 1) <> "\" Then
102         QualifyPath = sPath & "\"
        Else
104         QualifyPath = sPath
        End If

        '<EhFooter>
        Exit Function

QualifyPath_Err:
        WriteErr Time & "-" & Date & " - " & Err.Description & " - " & "in PerfectAntivirus2009.zfrmScanUSB.QualifyPath " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Function

Private Function TrimNull(startstr As String) As String

        '<EhHeader>
        On Error GoTo TrimNull_Err

        '</EhHeader>
        Dim pos As Integer

100     pos = InStr(startstr, Chr$(0))

102     If pos Then
104         TrimNull = Left$(startstr, pos - 1)

            Exit Function

        End If

106     TrimNull = startstr

        '<EhFooter>
        Exit Function

TrimNull_Err:
        WriteErr Time & "-" & Date & " - " & Err.Description & " - " & "in PerfectAntivirus2009.zfrmScanUSB.TrimNull " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Function

Private Sub cmdStopScan_Click()

        '<EhHeader>
        On Error GoTo cmdStopScan_Click_Err

        '</EhHeader>

100     xStopScan = True

        '<EhFooter>
        Exit Sub

cmdStopScan_Click_Err:
        WriteErr Time & "-" & Date & " - " & Err.Description & " - " & "in PerfectAntivirus2009.zfrmScanUSB.cmdStopScan_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub cmdUSBCancel_Click()

        '<EhHeader>
        On Error GoTo cmdUSBCancel_Click_Err

        '</EhHeader>

100     Me.Hide

        '<EhFooter>
        Exit Sub

cmdUSBCancel_Click_Err:
        WriteErr Time & "-" & Date & " - " & Err.Description & " - " & "in PerfectAntivirus2009.zfrmScanUSB.cmdUSBCancel_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub cmdUSBKill_Click()

        '<EhHeader>
        On Error GoTo cmdUSBKill_Click_Err

        '</EhHeader>
        Dim Y

        Dim X As String

100     For Y = 1 To LVVirus4.ListItems.Count

102         If LVVirus4.ListItems(Y).Checked = True And FileExists(LVVirus4.ListItems(Y).SubItems(1).Caption) = True Then

104             DoEvents
106             SetAttr LVVirus4.ListItems(Y).SubItems(1).Caption, vbNormal
108             modScanVirus.DeleteFile LVVirus4.ListItems(Y).SubItems(1).Caption
110             X = "YES"
            End If
    
112     Next Y

114     If Not X = "" Then
116         DelAllChecked LVVirus4
            'UniMsgBox X, vbOKOnly, "Nhu74ng vie65c d9a4 la2m.", Me.hWnd
118         lblUSBStatus.Caption = "D9a4 die65t xong ca1c Virus d9a4 d9a1nh d9a61u."
        Else
120         lblUSBStatus.Caption = "Kho6ng co1 Virus d9e63 die65t."
            'UniMsgBox "Kho6ng co1 Virus na2o d9e63 die65t!", vbOKOnly, "Tho6ng Ba1o", Me.hWnd
        End If

        '<EhFooter>
        Exit Sub

cmdUSBKill_Click_Err:
        WriteErr Time & "-" & Date & " - " & Err.Description & " - " & "in PerfectAntivirus2009.zfrmScanUSB.cmdUSBKill_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub cmdUSBScan_Click()

        '<EhHeader>
        On Error GoTo cmdUSBScan_Click_Err

        '</EhHeader>

100     Me.cmdUSBCancel.Enabled = False
102     Me.cmdStopScan.Enabled = True
104     Me.cmdUSBScan.Enabled = False
106     Me.lblUSBStatus.Caption = "D9ang que1t Autorun..."

        '///////////////////////////////////////////////////
108     If FileExists(Me.lblUSBPath.Caption & "autorun.inf") = True Then

            On Error GoTo KhOiDiEtAuToRuN

            'Start Diet Autorun
            'frmMessenger.zShowMessenger "Pha1t hie65n Autorun!", "Pha1t hie65n ta65p tin tu75 cha5y (Autorun.inf) ta5i o63 d9i4a [" & drv.DriveLetter & ":\] Chu7o7ng tri2nh se4 xo1a no1 ra kho3i he65 tho61ng va2 thie61t la65p ba3o ve65 cho o63 d9i4a na2y ngay ba6y gio72.)", 5000, xvang
            Dim i0

110         i0 = LVVirus4.ListItems.Count + 1
112         LVVirus4.ListItems.Add i0, , "Autorun"
114         LVVirus4.ListItems(i0).SubItems(1).Caption = lblUSBPath.Caption & "autorun.inf"
116         LVVirus4.ListItems(i0).Checked = True
118         LVVirus4.Refresh

            '/////// Tim Nguon Goc Cua Virus ////////
            Dim Str        As String

            Dim Xx1        As String

            Dim Xx2        As String

            Dim xFileName1 As String

            Dim xFileName2 As String

120         DoEvents
122         Xx1 = lblUSBPath.Caption & GetOpenAutorun(lblUSBPath.Caption & "autorun.inf")
124         Xx2 = lblUSBPath.Caption & GetShellOpenAutorun(lblUSBPath.Caption & "autorun.inf")

126         If Xx1 <> lblUSBPath.Caption And modScanVirus.FileExists(Xx1) = True Then
128             xFileName1 = GetFileName(Xx1)

130             If CheckProcess(xFileName1) <> 0 Then KillProcessById CheckProcess(xFileName1) 'EndTask xFileName1
132             SetAttr Xx1, vbNormal

                'modScanVirus.DeleteFile Xx1
                'frmMessenger.zShowMessenger "Pha1t hie65n Virus!", "D9a4 pha1t hie65n Virus ta5i: [" & drv.DriveLetter & ":\...\" & GetFileName(Xx1) & "]. Ti2nh tra5ng: D9a4 Xo1a", 5000, xvang
                Dim i1

134             i1 = LVVirus4.ListItems.Count + 1
136             LVVirus4.ListItems.Add i1, , GetFileName(Xx1)
138             LVVirus4.ListItems(i1).SubItems(1).Caption = Xx1
140             LVVirus4.ListItems(i1).Checked = True
142             LVVirus4.Refresh
            End If

144         If Xx2 <> lblUSBPath.Caption And modScanVirus.FileExists(Xx2) = True Then
146             xFileName2 = GetFileName(Xx2)

148             If CheckProcess(xFileName2) <> 0 Then KillProcessById CheckProcess(xFileName2) 'EndTask xFileName2
150             SetAttr Xx2, vbNormal

                'modScanVirus.DeleteFile Xx2
                'frmMessenger.zShowMessenger "Pha1t hie65n Virus!", "D9a4 pha1t hie65n Virus ta5i: [" & drv.DriveLetter & ":\...\" & GetFileName(Xx2) & "]. Ti2nh tra5ng: D9a4 Xo1a", 5000, xvang
                Dim i2

152             i2 = LVVirus4.ListItems.Count + 1
154             LVVirus4.ListItems.Add i2, , GetFileName(Xx2)
156             LVVirus4.ListItems(i2).SubItems(1).Caption = Xx2
158             LVVirus4.ListItems(i2).Checked = True
160             LVVirus4.Refresh
            End If

            '////////// End / Tim nguon goc Virus ///////////

            '////// Diet Autorun ////////
162         DoEvents
164         SetAttr lblUSBPath.Caption & "autorun.inf", vbNormal
166         modScanVirus.DeleteFile lblUSBPath.Caption & "autorun.inf"
168         MkDir lblUSBPath.Caption & "autorun.inf"
170         Str = "cmd /c md \\?\" & lblUSBPath.Caption & "autorun.inf\.PAV.2009."
172         Shell Str, vbHide
174         SetAttr lblUSBPath.Caption & "autorun.inf", vbHidden + vbReadOnly + vbSystem
176         FileCopy AppPath & "PAV2009.ico", lblUSBPath.Caption & "autorun.inf\Icon.ico"
178         WriteFileUni lblUSBPath.Caption & "autorun.inf\ThongTin.txt", ToUnicode("Thu7 mu5c na2y la2 thu7 mu5c Autorun gia3, d9u7o75c ta5o ra d9e63 d9a1nh lu72a Virus, nha82m nga8n Virus la6y qua USB." & vbCrLf & "File na2y d9u7o75c ta5o bo73i chu7o7ng tri2nh PAV 2009." & vbCrLf & "Phát ha2nh bo73i: http://qts.come.vn") 'CreateTextFile drv.DriveLetter & ":\autorun.inf\AlwaysProtected.txt", "
180         WriteFileUni lblUSBPath.Caption & "autorun.inf\desktop.ini", "[.ShellClassInfo]" & vbCrLf & "IconFile=" & lblUSBPath.Caption & "autorun.inf\Icon.ico" & vbCrLf & "IconIndex = 0"
182         SetAttr lblUSBPath.Caption & "autorun.inf\desktop.ini", vbHidden + vbSystem + vbReadOnly

184         DoEvents
            '////// End Diet Autorun ////////
KhOiDiEtAuToRuN:
            '////////////////////////////////////
        End If

186     Me.lblUSBStatus.AutoUnicode = False

188     SearchFile Me.lblUSBPath.Caption, "*.exe"
190     SearchFile Me.lblUSBPath.Caption, "*.bat"
192     SearchFile Me.lblUSBPath.Caption, "*.com"
194     SearchFile Me.lblUSBPath.Caption, "*.cmd"
196     PLaySound AppPath & "Sound\ScanDone.wav"
198     Me.lblUSBStatus.AutoUnicode = True
200     Me.lblUSBStatus.Caption = "D9a4 que1t xong! Ti2m tha61y " & LVVirus4.ListItems.Count & " Virus."
202     Me.cmdUSBKill.Enabled = True
204     Me.cmdStopScan.Enabled = False
206     Me.cmdUSBCancel.Enabled = True
208     Me.cmdUSBCancel.Caption = "D9o1ng"

        '<EhFooter>
        Exit Sub

cmdUSBScan_Click_Err:
        WriteErr Time & "-" & Date & " - " & Err.Description & " - " & "in PerfectAntivirus2009.zfrmScanUSB.cmdUSBScan_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub F1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

        '<EhHeader>
        On Error GoTo F1_MouseMove_Err

        '</EhHeader>
100     If Button = 1 Then
102         Call ReleaseCapture
104         Call SendMessage(Me.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
        End If

        '<EhFooter>
        Exit Sub

F1_MouseMove_Err:
        WriteErr Time & "-" & Date & " - " & Err.Description & " - " & "in PerfectAntivirus2009.zfrmScanUSB.F1_MouseMove " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Form_Load()

        '<EhHeader>
        On Error GoTo Form_Load_Err

        '</EhHeader>

100     App.TaskVisible = False
102     xStopScan = False
        '---> List Virus
104     LVVirus4.View = eViewDetails
106     LVVirus4.GridLines = True
108     LVVirus4.HeaderButtons = False
110     LVVirus4.CheckBoxes = True
112     LVVirus4.AutoUnicode = False
114     LVVirus4.Columns.Add , , "Virus Name", , 2200
116     LVVirus4.Columns.Add , , "Path", , 4000
118     LVVirus4.Refresh
        '<--- Virus

        Dim NotifFilter As DEV_BROADCAST_DEVICEINTERFACE

120     With NotifFilter
122         .dbcc_size = Len(NotifFilter)
124         .dbcc_devicetype = DBT_DEVTYP_DEVICEINTERFACE
        End With

126     Call SubClass(Me.hWnd)
128     lDevNotify = RegisterDeviceNotification(Me.hWnd, NotifFilter, DEVICE_NOTIFY_WINDOW_HANDLE Or DEVICE_NOTIFY_ALL_INTERFACE_CLASSES)

        '<EhFooter>
        Exit Sub

Form_Load_Err:
        WriteErr Time & "-" & Date & " - " & Err.Description & " - " & "in PerfectAntivirus2009.zfrmScanUSB.Form_Load " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

        '<EhHeader>
        On Error GoTo Form_QueryUnload_Err

        '</EhHeader>
100     Select Case UnloadMode

            Case 0
                'unload form X button
102             Cancel = True
104             Me.Hide

106         Case 1
                'unload by code
108             Call UnregisterDeviceNotification(lDevNotify)
110             Call UnSubClass
        End Select

        '<EhFooter>
        Exit Sub

Form_QueryUnload_Err:
        WriteErr Time & "-" & Date & " - " & Err.Description & " - " & "in PerfectAntivirus2009.zfrmScanUSB.Form_QueryUnload " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Public Sub zScanUSB(USBLetter)

        '////// Setting for Form
        '<EhHeader>
        On Error GoTo zScanUSB_Err

        '</EhHeader>
100     With Me
102         .lblUSBPath.Caption = USBLetter
104         BringWindowToTop Me.hWnd
106         .Show
108         PLaySound AppPath & "Sound\Found.wav"
110         .cmdStopScan.Enabled = False
112         .cmdUSBKill.Enabled = False
114         .cmdUSBScan.Enabled = True
116         .lblUSBStatus.Caption = ""
            '///// Setting for Form
118         LVVirus4.ListItems.Clear
        End With

        '<EhFooter>
        Exit Sub

zScanUSB_Err:
        WriteErr Time & "-" & Date & " - " & Err.Description & " - " & "in PerfectAntivirus2009.zfrmScanUSB.zScanUSB " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

