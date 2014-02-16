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
Private Declare Function FindClose Lib "kernel32" _
  (ByVal hFindFile As Long) As Long
   
Private Declare Function FindFirstFile Lib "kernel32" _
   Alias "FindFirstFileA" _
  (ByVal lpFileName As String, _
   lpFindFileData As WIN32_FIND_DATA) As Long
   
Private Declare Function FindNextFile Lib "kernel32" _
   Alias "FindNextFileA" _
  (ByVal hFindFile As Long, _
   lpFindFileData As WIN32_FIND_DATA) As Long

Private Declare Function GetTickCount Lib "kernel32" () As Long
Private Declare Function BringWindowToTop Lib "user32" (ByVal hWnd As Long) As Long


Private Declare Function RegisterDeviceNotification Lib "User32.dll" Alias _
    "RegisterDeviceNotificationA" (ByVal hRecipient As Long, _
    ByRef NotificationFilter As Any, ByVal Flags As Long) As Long
Private Declare Function UnregisterDeviceNotification Lib "User32.dll" ( _
    ByVal Handle As Long) As Long
    
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Private Declare Function SetWindowPos& Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, _
ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
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
Private lDevNotify As Long
Private Const DEVICE_NOTIFY_WINDOW_HANDLE As Long = &H0
Private Const DBT_DEVTYP_DEVICEINTERFACE As Long = &H5
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


Private Function SearchFile(Path, Filename)

   Dim FP As FILE_PARAMS  'holds search parameters
   Dim tstart As Single   'timer var for this routine only
   Dim tend As Single     'timer var for this routine only
   With FP
      .sFileRoot = Path       'start path
      .sFileNameExt = Filename    'file type of interest
      .bRecurse = 1 ' Check1.Value = 1  '1 = recursive search
   End With
   tstart = GetTickCount()
   Call SearchForFiles(FP)
   tend = GetTickCount()
End Function


Private Sub GetFileInformation(FP As FILE_PARAMS)
DoEvents
   Dim WFD As WIN32_FIND_DATA
   Dim hFile As Long
   Dim sPath As String
   Dim sRoot As String
   Dim sTmp As String
   Dim SKetQua As String
   sRoot = QualifyPath(FP.sFileRoot)
   sPath = sRoot & FP.sFileNameExt
   hFile = FindFirstFile(sPath, WFD)
   If hFile <> INVALID_HANDLE_VALUE Then
      Do
         If Not (WFD.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) = _
                 FILE_ATTRIBUTE_DIRECTORY Then
            FP.Count = FP.Count + 1
            sTmp = TrimNull(WFD.cFileName)
            SKetQua = sRoot & sTmp
If xStopScan = True Then Exit Sub
If FileLen(SKetQua) < 10000000 And FileLen(SKetQua) > 0 Then
            Me.lblUSBStatus.Caption = SKetQua
            Dim AX As String
            AX = modScanVirus.CheckVirus(SKetQua)
            If AX <> "No" Then
                Dim i
                i = LVVirus4.ListItems.Count + 1
                LVVirus4.ListItems.Add i, , AX
                LVVirus4.ListItems(i).SubItems(1).Caption = SKetQua
                LVVirus4.ListItems(i).Checked = True
                LVVirus4.Refresh
            End If
End If
            
            
            'Text1.Text = Text1.Text & SKetQua & vbCrLf
            '*********************************************
         End If
      Loop While FindNextFile(hFile, WFD)
      hFile = FindClose(hFile)
   End If
DoEvents
End Sub


Private Sub SearchForFiles(FP As FILE_PARAMS)
  'local working variables
   Dim WFD As WIN32_FIND_DATA
   Dim hFile As Long
   Dim sPath As String
   Dim sRoot As String
   Dim sTmp As String
   sRoot = QualifyPath(FP.sFileRoot)
   sPath = sRoot & "*.*"
   hFile = FindFirstFile(sPath, WFD)
   If hFile <> INVALID_HANDLE_VALUE Then
      Call GetFileInformation(FP)
      Do
         If (WFD.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) Then
            If FP.bRecurse Then
               If Asc(WFD.cFileName) <> vbDot Then
                  FP.sFileRoot = sRoot & TrimNull(WFD.cFileName)
                  Call SearchForFiles(FP)
               End If
            End If
         End If
      Loop While FindNextFile(hFile, WFD)
      hFile = FindClose(hFile)
   End If
End Sub


Private Function QualifyPath(sPath As String) As String
   If Right$(sPath, 1) <> "\" Then
      QualifyPath = sPath & "\"
   Else
      QualifyPath = sPath
   End If
End Function


Private Function TrimNull(startstr As String) As String
   Dim pos As Integer
   pos = InStr(startstr, Chr$(0))
   If pos Then
      TrimNull = Left$(startstr, pos - 1)
      Exit Function
   End If
   TrimNull = startstr
End Function

Private Sub cmdStopScan_Click()
xStopScan = True
End Sub

Private Sub cmdUSBCancel_Click()
Me.Hide
End Sub

Private Sub cmdUSBKill_Click()
Dim y
Dim x As String
For y = 1 To LVVirus4.ListItems.Count

    If LVVirus4.ListItems(y).Checked = True And FileExists(LVVirus4.ListItems(y).SubItems(1).Caption) = True Then
    DoEvents
        SetAttr LVVirus4.ListItems(y).SubItems(1).Caption, vbNormal
        modScanVirus.DeleteFile LVVirus4.ListItems(y).SubItems(1).Caption
        x = "YES"
    End If
    
Next y

If Not x = "" Then
DelAllChecked LVVirus4
'UniMsgBox X, vbOKOnly, "Nhu74ng vie65c d9a4 la2m.", Me.hWnd
lblUSBStatus.Caption = "D9a4 die65t xong ca1c Virus d9a4 d9a1nh d9a61u."
Else
lblUSBStatus.Caption = "Kho6ng co1 Virus d9e63 die65t."
'UniMsgBox "Kho6ng co1 Virus na2o d9e63 die65t!", vbOKOnly, "Tho6ng Ba1o", Me.hWnd
End If

End Sub

Private Sub cmdUSBScan_Click()
Me.cmdUSBCancel.Enabled = False
Me.cmdStopScan.Enabled = True
Me.cmdUSBScan.Enabled = False
Me.lblUSBStatus.Caption = "D9ang que1t Autorun..."

'///////////////////////////////////////////////////
On Error GoTo KhOiDiEtAuToRuN
                'Start Diet Autorun
                'frmMessenger.zShowMessenger "Pha1t hie65n Autorun!", "Pha1t hie65n ta65p tin tu75 cha5y (Autorun.inf) ta5i o63 d9i4a [" & drv.DriveLetter & ":\] Chu7o7ng tri2nh se4 xo1a no1 ra kho3i he65 tho61ng va2 thie61t la65p ba3o ve65 cho o63 d9i4a na2y ngay ba6y gio72.)", 5000, xvang
                    Dim i0
                    i0 = LVVirus4.ListItems.Count + 1
                    LVVirus4.ListItems.Add i0, , "Autorun"
                    LVVirus4.ListItems(i0).SubItems(1).Caption = lblUSBPath.Caption & "autorun.inf"
                    LVVirus4.ListItems(i0).Checked = True
                    LVVirus4.Refresh
                '/////// Tim Nguon Goc Cua Virus ////////
                Dim Str As String
                Dim Xx1        As String
                Dim Xx2        As String
                Dim xFileName1 As String
                Dim xFileName2 As String
                DoEvents
                Xx1 = lblUSBPath.Caption & GetOpenAutorun(lblUSBPath.Caption & "autorun.inf")
                Xx2 = lblUSBPath.Caption & GetShellOpenAutorun(lblUSBPath.Caption & "autorun.inf")
                If Xx1 <> lblUSBPath.Caption And modScanVirus.FileExists(Xx1) = True Then
                    xFileName1 = GetFileName(Xx1)
                    If CheckProcess(xFileName1) <> 0 Then KillProcessById CheckProcess(xFileName1) 'EndTask xFileName1
                    SetAttr Xx1, vbNormal
                    'modScanVirus.DeleteFile Xx1
                    'frmMessenger.zShowMessenger "Pha1t hie65n Virus!", "D9a4 pha1t hie65n Virus ta5i: [" & drv.DriveLetter & ":\...\" & GetFileName(Xx1) & "]. Ti2nh tra5ng: D9a4 Xo1a", 5000, xvang
                    Dim i1
                    i1 = LVVirus4.ListItems.Count + 1
                    LVVirus4.ListItems.Add i1, , GetFileName(Xx1)
                    LVVirus4.ListItems(i1).SubItems(1).Caption = Xx1
                    LVVirus4.ListItems(i1).Checked = True
                    LVVirus4.Refresh
                End If

                If Xx2 <> lblUSBPath.Caption And modScanVirus.FileExists(Xx2) = True Then
                    xFileName2 = GetFileName(Xx2)
                    If CheckProcess(xFileName2) <> 0 Then KillProcessById CheckProcess(xFileName2) 'EndTask xFileName2
                    SetAttr Xx2, vbNormal
                    'modScanVirus.DeleteFile Xx2
                    'frmMessenger.zShowMessenger "Pha1t hie65n Virus!", "D9a4 pha1t hie65n Virus ta5i: [" & drv.DriveLetter & ":\...\" & GetFileName(Xx2) & "]. Ti2nh tra5ng: D9a4 Xo1a", 5000, xvang
                    Dim i2
                    i2 = LVVirus4.ListItems.Count + 1
                    LVVirus4.ListItems.Add i2, , GetFileName(Xx2)
                    LVVirus4.ListItems(i2).SubItems(1).Caption = Xx2
                    LVVirus4.ListItems(i2).Checked = True
                    LVVirus4.Refresh
                End If

                '////////// End / Tim nguon goc Virus ///////////

                '////// Diet Autorun ////////
                DoEvents
                SetAttr lblUSBPath.Caption & "autorun.inf", vbNormal
                modScanVirus.DeleteFile lblUSBPath.Caption & "autorun.inf"
                MkDir lblUSBPath.Caption & "autorun.inf"
                Str = "cmd /c md \\?\" & lblUSBPath.Caption & "autorun.inf\.PAV.2009."
                Shell Str, vbHide
                SetAttr lblUSBPath.Caption & "autorun.inf", vbHidden + vbReadOnly + vbSystem
                FileCopy AppPath & "PAV2009.ico", lblUSBPath.Caption & "autorun.inf\Icon.ico"
                WriteFileUni lblUSBPath.Caption & "autorun.inf\ThongTin.txt", ToUnicode("Thu7 mu5c na2y la2 thu7 mu5c Autorun gia3, d9u7o75c ta5o ra d9e63 d9a1nh lu72a Virus, nha82m nga8n Virus la6y qua USB." & vbCrLf & "File na2y d9u7o75c ta5o bo73i chu7o7ng tri2nh PAV 2009." & vbCrLf & "Phát ha2nh bo73i: http://qts.come.vn") 'CreateTextFile drv.DriveLetter & ":\autorun.inf\AlwaysProtected.txt", "
                WriteFileUni lblUSBPath.Caption & "autorun.inf\desktop.ini", "[.ShellClassInfo]" & vbCrLf & "IconFile=" & lblUSBPath.Caption & "autorun.inf\Icon.ico" & vbCrLf & "IconIndex = 0"
                SetAttr lblUSBPath.Caption & "autorun.inf\desktop.ini", vbHidden + vbSystem + vbReadOnly
                DoEvents
                '////// End Diet Autorun ////////
KhOiDiEtAuToRuN:
'////////////////////////////////////







Me.lblUSBStatus.AutoUnicode = False

SearchFile Me.lblUSBPath.Caption, "*.exe"
SearchFile Me.lblUSBPath.Caption, "*.bat"
SearchFile Me.lblUSBPath.Caption, "*.com"
SearchFile Me.lblUSBPath.Caption, "*.cmd"
PLaySound AppPath & "Sound\ScanDone.wav"
Me.lblUSBStatus.AutoUnicode = True
Me.lblUSBStatus.Caption = "D9a4 que1t xong! Ti2m tha61y " & LVVirus4.ListItems.Count & " Virus."
Me.cmdUSBKill.Enabled = True
Me.cmdStopScan.Enabled = False
Me.cmdUSBCancel.Enabled = True
Me.cmdUSBCancel.Caption = "D9o1ng"
End Sub

Private Sub F1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 1 Then
   Call ReleaseCapture
   Call SendMessage(Me.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
   End If
End Sub

Private Sub Form_Load()
App.TaskVisible = False
xStopScan = False
'---> List Virus
        LVVirus4.View = eViewDetails
        LVVirus4.GridLines = True
        LVVirus4.HeaderButtons = False
        LVVirus4.CheckBoxes = True
        LVVirus4.AutoUnicode = False
        LVVirus4.Columns.Add , , "Virus Name", , 2200
        LVVirus4.Columns.Add , , "Path", , 4000
        LVVirus4.Refresh
'<--- Virus

    Dim NotifFilter As DEV_BROADCAST_DEVICEINTERFACE
    With NotifFilter
        .dbcc_size = Len(NotifFilter)
        .dbcc_devicetype = DBT_DEVTYP_DEVICEINTERFACE
    End With
    Call SubClass(Me.hWnd)
    lDevNotify = RegisterDeviceNotification(Me.hWnd, NotifFilter, _
    DEVICE_NOTIFY_WINDOW_HANDLE Or DEVICE_NOTIFY_ALL_INTERFACE_CLASSES)
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Select Case UnloadMode
        Case 0
        'unload form X button
                Cancel = True
                Me.Hide
        Case 1
        'unload by code
                Call UnregisterDeviceNotification(lDevNotify)
                Call UnSubClass
    End Select
End Sub

Public Sub zScanUSB(USBLetter)
'////// Setting for Form
With Me
.lblUSBPath.Caption = USBLetter
BringWindowToTop Me.hWnd
.Show
PLaySound AppPath & "Sound\Found.wav"
.cmdStopScan.Enabled = False
.cmdUSBKill.Enabled = False
.cmdUSBScan.Enabled = True
.lblUSBStatus.Caption = ""
'///// Setting for Form
LVVirus4.ListItems.Clear
End With
End Sub

