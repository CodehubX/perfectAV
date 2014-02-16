VERSION 5.00
Object = "{D1ECD258-D329-4388-AB83-DEC261A66B86}#1.0#0"; "UniControls_v2.0.ocx"
Begin VB.Form frmMain 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PAV 2009 - Mini Folder Scan"
   ClientHeight    =   4950
   ClientLeft      =   405
   ClientTop       =   525
   ClientWidth     =   7575
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
   ScaleHeight     =   4950
   ScaleWidth      =   7575
   StartUpPosition =   2  'CenterScreen
   Begin UniControls.UniLabel lblStatus 
      Height          =   375
      Left            =   120
      Top             =   3960
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   661
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
   Begin UniControls.UniOptionButton optAll 
      Height          =   195
      Left            =   1440
      TabIndex        =   7
      Top             =   720
      Width           =   2775
      _ExtentX        =   4895
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
      Value           =   -1  'True
      Caption         =   "Ta61t ca3 ca1c loa5i File (*.*) [Se4 cha65m]"
      ForeColor       =   0
   End
   Begin UniControls.UniOptionButton optCustom 
      Height          =   195
      Left            =   1440
      TabIndex        =   6
      Top             =   960
      Width           =   4215
      _ExtentX        =   7435
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
      Caption         =   "1 So61 loa5i File (*.exe, *.bat, *.cmd, *.com, *.dll, *.scr)"
      ForeColor       =   0
   End
   Begin UniControls.UniLabel UniLabel2 
      Height          =   255
      Left            =   120
      Top             =   720
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   450
      Caption         =   "Kie63u File Que1t:"
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
   Begin UniControls.UniButton cmdExit 
      Height          =   375
      Left            =   6960
      TabIndex        =   5
      Top             =   4440
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      Icon            =   "frmMain.frx":15162
      Style           =   2
      IconAlign       =   2
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
   Begin UniControls.UniButton cmdKill 
      Height          =   375
      Left            =   5280
      TabIndex        =   4
      Top             =   4440
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      Icon            =   "frmMain.frx":15B74
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
   Begin UniControls.UniButton cmdCachLy 
      Height          =   375
      Left            =   3600
      TabIndex        =   3
      Top             =   4440
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      Icon            =   "frmMain.frx":1610E
      Style           =   2
      Caption         =   "Ca1ch Ly"
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
   Begin UniControls.UniButton cmdStop 
      Height          =   375
      Left            =   1920
      TabIndex        =   2
      Top             =   4440
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      Icon            =   "frmMain.frx":16B20
      Style           =   2
      Caption         =   "Du72ng"
      IconAlign       =   3
      iNonThemeStyle  =   2
      Enabled         =   0   'False
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
   Begin UniControls.UniButton cmdScan 
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   4440
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      Icon            =   "frmMain.frx":170BA
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
   Begin UniControls.UniListView LV 
      Height          =   2535
      Left            =   120
      TabIndex        =   0
      Top             =   1320
      Width           =   7335
      _ExtentX        =   12938
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
      AutoArrange     =   0   'False
      HeaderButtons   =   0   'False
      HeaderTrackSelect=   0   'False
      HideSelection   =   0   'False
      InfoTips        =   0   'False
   End
   Begin UniControls.UniLabel UniLabel1 
      Height          =   495
      Left            =   120
      Top             =   120
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   873
      Alignment       =   1
      Caption         =   "D9ang que1t Virus cho:"
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
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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



Dim xStop As Boolean
Dim Comd As String
Dim AppScan As String
Private Function SearchFile(Path, FileName)
If xStop = True Then Exit Function
   Dim FP As FILE_PARAMS  'holds search parameters
   Dim tstart As Single   'timer var for this routine only
   Dim tend As Single     'timer var for this routine only
   With FP
      .sFileRoot = Path       'start path
      .sFileNameExt = FileName    'file type of interest
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
            
            If FileExists(SKetQua) = True Then
                If FileLen(SKetQua) > 0 And FileLen(SKetQua) < 10000000 Then
                    lblStatus.Caption = SKetQua
                    Dim AX As String
                    AX = modScanVirus.CheckVirus(SKetQua)
                    If AX <> "No" Then
                        With LV
                            Dim i
                            i = .ListItems.Count + 1
                            .ListItems.Add i, , AX
                            .ListItems(i).SubItems(1).Caption = SKetQua
                            .ListItems(i).SubItems(2).Caption = FileLen(SKetQua) & " Bytes"
                            .ListItems(i).SubItems(3).Caption = CheckProcess(SKetQua)
                            .ListItems(i).Checked = True
                    'List1.AddItem SKetQua
                        End With
                    End If
                End If
            End If
         End If
        If xStop = True Then Exit Sub
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


Private Sub cmdCachLy_Click()
On Error Resume Next
If UniMsgBox("Ba5n co1 muo61n ca1ch ly ca1c Virus d9a4 cho5n kho6ng?", vbYesNo, "Tho6ng Ba1o") = vbYes Then

MkDir AppPath & "VungCachLy"

Dim y, j
Dim x As String
For y = 1 To LV.ListItems.Count
    If LV.ListItems(y).Checked = True And FileExists(LV.ListItems(y).SubItems(1).Caption) = True Then
    
        If LV.ListItems(y).SubItems(3).Caption <> "0" Then
        'Kill process
            KillProcessById (LV.ListItems(y).SubItems(3).Caption)
            x = x & " D9a4 ta81t tie61n tri2nh: " & LV.ListItems(y).SubItems(3).Caption & vbCrLf
        End If
        Set fss = Nothing
        SetAttr LV.ListItems(y).SubItems(1).Caption, vbNormal
        Name LV.ListItems(y).SubItems(1).Caption As LV.ListItems(y).SubItems(1).Caption & ".DaCachLy"
        FileCopy LV.ListItems(y).SubItems(1).Caption & ".DaCachLy", AppPath & "VungCachLy\" & GetFileName(LV.ListItems(y).SubItems(1).Caption & ".DaCachLy")
        modScanVirus.DeleteFile LV.ListItems(y).SubItems(1).Caption & ".DaCachLy"
        x = x & " D9a4 Ca1ch Ly: " & LV.ListItems(y).SubItems(1).Caption & vbCrLf
    End If
Next y
If Not x = "" Then
DelAllChecked LV
UniMsgBox x, vbOKOnly, "D9a4 Ca1ch Ly Virus", Me.hWnd
Else
UniMsgBox "Kho6ng co1 Virus na2o d9e63 ca1ch ly.", vbOKOnly, "Tho6ng Ba1o", Me.hWnd
End If
End If
End Sub

Private Sub cmdExit_Click()
End
End Sub

Private Sub cmdKill_Click()
If UniMsgBox("Ba5n co1 cha81c cha81n muo61n xo1a ca1c Virus d9a4 d9a1nh da61u kho6ng?", vbYesNo, "Die65t Virus") = vbYes Then
DoEvents
Me.lblStatus.Caption = "D9ang xo1a..."
Dim y
Dim x As String
For y = 1 To LV.ListItems.Count
    If LV.ListItems(y).Checked = True And FileExists(LV.ListItems(y).SubItems(1).Caption) = True Then
    DoEvents
        If LV.ListItems(y).SubItems(3).Caption <> "0" Then
        'Kill process
            KillProcessById (LV.ListItems(y).SubItems(3).Caption)
            x = x & " D9a4 ta81t tie61n tri2nh: " & LV.ListItems(y).SubItems(3).Caption & vbCrLf
        End If
        DoEvents

        SetAttr LV.ListItems(y).SubItems(1).Caption, vbNormal
        modScanVirus.DeleteFile LV.ListItems(y).SubItems(1).Caption
        x = x & " D9a4 Xo1a Bo3: " & LV.ListItems(y).SubItems(1).Caption & vbCrLf
        
        End If
Next y

If Not x = "" Then
DelAllChecked LV
UniMsgBox x, vbOKOnly, "Nhu74ng vie65c d9a4 la2m.", Me.hWnd
Else
UniMsgBox "Kho6ng co1 Virus na2o d9e63 die65t.", vbOKOnly, "Tho6ng Ba1o", Me.hWnd
End If

End If ' U
End Sub

Private Sub cmdScan_Click()
Me.optAll.Enabled = False
Me.optCustom.Enabled = False
Me.cmdStop.Enabled = True
Me.cmdScan.Enabled = False
Me.cmdExit.Enabled = False
Me.cmdKill.Enabled = False
Me.cmdCachLy.Enabled = False
xStop = False
'Start scan
DelAllLV LV
lblStatus.Caption = "Ba81t d9a62u que1t"


If optAll.Value = True Then
    SearchFile AppScan, "*.*"
Else
    SearchFile AppScan, "*.exe"
    If xStop = True Then GoTo DaQuEtXoNg
    SearchFile AppScan, "*.bat"
    If xStop = True Then GoTo DaQuEtXoNg
    SearchFile AppScan, "*.cmd"
    If xStop = True Then GoTo DaQuEtXoNg
    SearchFile AppScan, "*.com"
    If xStop = True Then GoTo DaQuEtXoNg
    SearchFile AppScan, "*.dll"
    If xStop = True Then GoTo DaQuEtXoNg
    SearchFile AppScan, "*.scr"
End If
DaQuEtXoNg:
lblStatus.Caption = "D9a4 die65t xong! Ti2m tha61y: " & LV.ListItems.Count & " Virus."
'End scan
PLaySound AppPath & "Sound\ScanDone.wav"
Me.optAll.Enabled = True
Me.optCustom.Enabled = True
Me.cmdScan.Enabled = True
Me.cmdStop.Enabled = False
Me.cmdExit.Enabled = True
Me.cmdCachLy.Enabled = True
Me.cmdKill.Enabled = True
End Sub



Private Sub cmdStop_Click()
xStop = True
End Sub

Private Sub Form_Load()
'GoTo Test
Comd = Command()

If Comd = "" Or Comd = ChrW(34) & ChrW(34) Then
UniMsgBox " File na2y d9u7o75c su73 du5ng d9e63 que1t Virus cho thu7 mu5c." & vbCrLf & " D9e63 su73 du5ng chu71c na8ng na2y, ba5n Click chuo65t pha3i va2o thu7 mu5c ca62n que1t -> Cho5n 'Que1t Virus Ba82ng PAV 2009'", vbOKOnly, "Tho6ng ba1o"
End
End If

AppScan = Mid(Comd, 2, Len(Comd) - 2)
If Right(AppScan, 1) <> "\" Then AppScan = AppScan & "\"
'MsgBox Comd

'Test:
'///////// Start ////////////
'AppScan = "F:\" 'Day chi la vi du

modScanVirus.ConnectDB '---> Connect DB
UniLabel1.Caption = "Que1t Virus: " & AppScan '---> Set label

'---> Set for lv
DelAllLV LV
LV.View = eViewDetails
LV.GridLines = True
LV.HeaderButtons = True
LV.CheckBoxes = True

LV.Columns.Add , , "Virus Name", , 1700
LV.Columns.Add , , "Path", , 3250
LV.Columns.Add , , "Size", , 1000
LV.Columns.Add , , "Process ID", , 1000
LV.Refresh




End Sub

Private Sub Form_Unload(Cancel As Integer)
Cancel = 1
End Sub

Private Sub UniButton1_Click()
SearchFile AppScan, "*.*"
End Sub
