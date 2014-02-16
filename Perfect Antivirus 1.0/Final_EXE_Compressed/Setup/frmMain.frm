VERSION 5.00
Object = "{D1ECD258-D329-4388-AB83-DEC261A66B86}#1.0#0"; "UniControls_v2.0.ocx"
Begin VB.Form frmMain 
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Setup Perfect Antivirus 2009"
   ClientHeight    =   4950
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6360
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
   ScaleWidth      =   6360
   StartUpPosition =   2  'CenterScreen
   Begin UniControls.UniLabel UniLabel4 
      Height          =   495
      Left            =   240
      Top             =   2280
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   873
      Alignment       =   1
      Caption         =   "Ha4y d9o5c qua hu7o71ng da64n su73 du5ng 1 la62n d9e63 co1 the63 su73 du5ng hie65u qua3 chu7o7ng tri2nh"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   255
   End
   Begin UniControls.UniButton cmdExit 
      Height          =   375
      Left            =   5160
      TabIndex        =   9
      Top             =   4440
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      Icon            =   "frmMain.frx":6852
      Style           =   2
      Caption         =   "D9o1ng"
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
   Begin UniControls.UniLabel lblStatus 
      Height          =   255
      Left            =   1080
      Top             =   3480
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   450
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
      ForeColor       =   16711680
   End
   Begin UniControls.UniLabel UniLabel3 
      Height          =   255
      Left            =   240
      Top             =   3480
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   450
      BackStyle       =   0
      Caption         =   "Extracting:"
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
   Begin UniControls.UniTextBox strFileList 
      Height          =   1455
      Left            =   4200
      TabIndex        =   8
      Top             =   5160
      Visible         =   0   'False
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   2566
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
      Text            =   "UniTextBox1"
      MultiLine       =   -1  'True
   End
   Begin UniControls.ProgressBar Bar1 
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   3720
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   450
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BrushStyle      =   0
      Color           =   8454016
      Max             =   80
      ShowText        =   -1  'True
      Value           =   0
   End
   Begin UniControls.UniButton cmdStart 
      Height          =   735
      Left            =   1920
      TabIndex        =   6
      Top             =   4080
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   1296
      Icon            =   "frmMain.frx":6DEC
      Style           =   2
      Caption         =   "Ba81t D9a62u Ca2i D9a85t"
      IconSize        =   32
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
   Begin UniControls.UniCheckBox chkHelp 
      Height          =   195
      Left            =   240
      TabIndex        =   5
      Top             =   3120
      Width           =   2805
      _ExtentX        =   4948
      _ExtentY        =   344
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Value           =   -1  'True
      Caption         =   "Mo73 hu7o71ng da64n sau khi ca2i d9a85t."
      ForeColor       =   33023
      ShowFocusRectangle=   0   'False
   End
   Begin UniControls.UniCheckBox chkRun 
      Height          =   195
      Left            =   240
      TabIndex        =   4
      Top             =   2880
      Width           =   3150
      _ExtentX        =   5556
      _ExtentY        =   344
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Value           =   -1  'True
      Caption         =   "Cha5y chu7o7ng tri2nh sau khi ca2i d9a85t."
      ForeColor       =   33023
      ShowFocusRectangle=   0   'False
   End
   Begin UniControls.UniCheckBox chkDesktop 
      Height          =   195
      Left            =   240
      TabIndex        =   3
      Top             =   1920
      Width           =   2340
      _ExtentX        =   4128
      _ExtentY        =   344
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Value           =   -1  'True
      Caption         =   "Ta5o Shortcut o73 Desktop."
      ForeColor       =   16711680
      ShowFocusRectangle=   0   'False
   End
   Begin UniControls.UniCheckBox chkStartMenu 
      Height          =   195
      Left            =   240
      TabIndex        =   2
      Top             =   1680
      Width           =   2580
      _ExtentX        =   4551
      _ExtentY        =   344
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Value           =   -1  'True
      Caption         =   "Ta5o Shortcut o73 Start Menu."
      ForeColor       =   16711680
      ShowFocusRectangle=   0   'False
   End
   Begin UniControls.UniButton cmdSelectPath 
      Height          =   255
      Left            =   5640
      TabIndex        =   1
      Top             =   1200
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   450
      Icon            =   "frmMain.frx":76C6
      Style           =   2
      Caption         =   "..."
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
   Begin UniControls.UniTextBox txtPath 
      Height          =   270
      Left            =   1680
      TabIndex        =   0
      Top             =   1200
      Width           =   3855
      _ExtentX        =   6800
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
      Text            =   ""
      Locked          =   -1  'True
      Enabled         =   0   'False
      BorderStyle     =   2
   End
   Begin UniControls.UniLabel UniLabel2 
      Height          =   255
      Left            =   120
      Top             =   1200
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   450
      Caption         =   "Thu7 mu5c ca2i d9a85t:"
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
      Height          =   495
      Left            =   1200
      Top             =   360
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   873
      Alignment       =   1
      BackStyle       =   0
      Caption         =   "Ca2i d9a85t Perfect Antivirus 2009"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16711680
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   480
      Picture         =   "frmMain.frx":76E2
      Top             =   240
      Width           =   480
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Files List
'-----------------------
'101 DATA.PAV
'102 Data.Str
'103 RegProtect.dat
'----------------------
'104 scrrun.dll
'105 UniControls_v2.0.ocx
'-------------------------
'106 Check_Virus.exe
'107 PAV2009.exe
'108 PAV2009.ico
'109 PAV2009Manager.exe
'110 PAVExplorer.exe
'111 PAVMiniScan.exe
'112 PAVStartUp.exe
'113 PAVSysReport.exe
'114 PAVUpdate.exe
'-------------------------
'115 HuongDan.html
'116 1.GIF
'117 2.GIF
'118 3.GIF
'119 4.GIF
'120 5.GIF
'121 6.GIF
'122 7.GIF
'123 8.GIF
'124 9.GIF
'125 10.GIF
'126 11.GIF
'127 12.GIF
'128 13.GIF
'129 14.GIF
'130 15.GIF
'131 16.GIF
'-------------
'132 Deleted.wav
'133 Found.wav
'134 Mes.wav
'135 Ring.wav
'136 ScanDone.wav
'-----------------
'137 RemovePAV2009.exe
Dim UnFileList As String

Private Type BrowseInfo

    hWndOwner As Long
    pIDLRoot As Long
    pszDisplayName As Long
    lpszTitle As Long
    ulFlags As Long
    lpfnCallback As Long
    lParam As Long
    iImage As Long

End Type

Const BIF_RETURNONLYFSDIRS = 1

Const MAX_PATH = 260

Private Declare Function DeleteFile _
                Lib "kernel32" _
                Alias "DeleteFileA" (ByVal lpFileName As String) As Long

Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal hMem As Long)

Private Declare Function lstrcat _
                Lib "kernel32" _
                Alias "lstrcatA" (ByVal lpString1 As String, _
                                  ByVal lpString2 As String) As Long

Private Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseInfo) As Long

Private Declare Function SHGetPathFromIDList _
                Lib "shell32" (ByVal pidList As Long, _
                               ByVal lpBuffer As String) As Long

Private Declare Function ShellExecute _
                Lib "shell32.dll" _
                Alias "ShellExecuteA" (ByVal hWnd As Long, _
                                       ByVal lpOperation As String, _
                                       ByVal lpFile As String, _
                                       ByVal lpParameters As String, _
                                       ByVal lpDirectory As String, _
                                       ByVal nShowCmd As Long) As Long

Private Sub cmdExit_Click()

    End

End Sub

Private Sub cmdSelectPath_Click()

    Dim iNull As Integer, lpIDList As Long, lResult As Long

    Dim sPath As String, udtBI As BrowseInfo

    With udtBI
        .hWndOwner = Me.hWnd
        .lpszTitle = lstrcat("", "Select Folder")
        .ulFlags = BIF_RETURNONLYFSDIRS
    End With

    lpIDList = SHBrowseForFolder(udtBI)

    If lpIDList Then
        sPath = String$(MAX_PATH, 0)
        SHGetPathFromIDList lpIDList, sPath
        CoTaskMemFree lpIDList
        iNull = InStr(sPath, vbNullChar)

        If iNull Then
            sPath = Left$(sPath, iNull - 1)
        End If
    End If

    If sPath <> "" Then
        If Right(sPath, 1) <> "\" Then sPath = sPath & "\"
        txtPath.Text = sPath
    Else
        txtPath.Text = Environ$("PROGRAMFILES") & "\Perfect Antivirus 2009\"
    End If

End Sub

Private Sub cmdStart_Click()
    UnFileList = ""
    AddFileU "S"
    Me.cmdStart.Enabled = False
    Me.cmdExit.Enabled = False
    Me.cmdSelectPath.Enabled = False
    Me.chkDesktop.Enabled = False
    Me.chkHelp.Enabled = False
    Me.chkRun.Enabled = False
    Me.chkStartMenu.Enabled = False
    Me.UniLabel2.Enabled = False
    'Dim ocxDir$
    'ocxDir = txtPath.Text & "\DATA.PAV"
    'If (FileExists(ocxDir) = False) Then
    'Dim bytResourceData() As Byte
    'bytResourceData = LoadResData(101, "CUSTOM")
    'Open ocxDir For Binary Shared As #1
    'Put #1, 1, bytResourceData
    'Close #1
    'End If
    CreateFolder txtPath.Text

    DoEvents
    WriteFileUni Environ$("WINDIR") & "\FileList.lst", strFileList.Text

    DoEvents

    Dim ocxDir$

    Dim sFileName As String

    Dim sFileNum  As Integer

    Dim InputData As String

    Open Environ$("WINDIR") & "\FileList.lst" For Input As #1

    DoEvents
    Do While Not EOF(1)
        DoEvents
        Bar1.Value = Bar1.Value + 2

        DoEvents
        Line Input #1, InputData
        sFileName = Split(InputData, "|", , vbBinaryCompare)(1)
        'MsgBox sFileNum & "|" & sFileName
        If UCase(sFileName) <> "S" Then
            
            
            DoEvents
            sFileNum = Split(InputData, "|", , vbBinaryCompare)(0)

            '------------------------------------
            DoEvents

            If sFileNum = 104 Or sFileNum = 105 Then
                ocxDir = Environ$("WINDIR") & "\System32\" & sFileName

                DoEvents
            ElseIf sFileNum > 114 And sFileNum < 132 Then
                CreateFolder txtPath.Text & "Help\"
                ocxDir = txtPath.Text & "Help\" & sFileName
            ElseIf sFileNum > 131 And sFileNum < 137 Then
                DoEvents
                CreateFolder txtPath.Text & "Sound\"
                ocxDir = txtPath.Text & "Sound\" & sFileName
            Else
                ocxDir = txtPath.Text & sFileName

                DoEvents
            End If
        
            DoEvents

            If (FileExists(ocxDir) = True) Then
                SetAttr ocxDir, vbNormal
                DeleteFile ocxDir
            End If
        
            DoEvents

            Dim bytResourceData() As Byte

            bytResourceData = LoadResData(sFileNum, "CUSTOM")

            DoEvents
            lblStatus.Caption = ocxDir

            If FileExists(ocxDir) = False Then

                DoEvents
                Open ocxDir For Binary Shared As #2
                Put #2, 1, bytResourceData
                Close #2
                
            End If
            
            If Right$(ocxDir, 3) <> "dll" And Right$(ocxDir, 3) <> "ocx" And Right(UCase(ocxDir), Len("REMOVEPAV2009.EXE")) <> "REMOVEPAV2009.EXE" Then AddFileU ocxDir
            
            DoEvents

            If sFileNum = 104 Or sFileNum = 105 Then
                Shell "regsvr32 /s " & ocxDir, vbHide

                DoEvents
            End If

            '------------------------------------
        End If

        DoEvents
    Loop

    DoEvents
    Close #1

    If Me.chkDesktop.Value = True Then
        lblStatus.Caption = ToUnicode("Ta5o Shortcut o74 Desktop...")
        CreateShorcut txtPath.Text & "PAV2009.exe", "C:\Documents and Settings\All Users\Desktop\Perfect Antivirus 2009.lnk"
        AddFileU "C:\Documents and Settings\All Users\Desktop\Perfect Antivirus 2009.lnk"
    End If

    If Me.chkStartMenu.Value = True Then
        lblStatus.Caption = ToUnicode("Ta5o Shortcut o74 Desktop...")
        CreateFolder "C:\Documents and Settings\All Users\Start Menu\Programs\Perfect Antivirus 2009\"
        CreateShorcut txtPath.Text & "PAV2009.exe", "C:\Documents and Settings\All Users\Start Menu\Programs\Perfect Antivirus 2009\PAV 2009.lnk"
        CreateShorcut txtPath.Text & "Help\HuongDan.html", "C:\Documents and Settings\All Users\Start Menu\Programs\Perfect Antivirus 2009\Help.lnk"
        CreateShorcut txtPath.Text & "RemovePAV2009.exe", "C:\Documents and Settings\All Users\Start Menu\Programs\Perfect Antivirus 2009\RemovePAV2009.lnk"
        AddFileU "C:\Documents and Settings\All Users\Start Menu\Programs\Perfect Antivirus 2009\PAV 2009.lnk"
        AddFileU "C:\Documents and Settings\All Users\Start Menu\Programs\Perfect Antivirus 2009\Help.lnk"
    End If
    WriteFileUni txtPath.Text & "FileList.usl", UnFileList
    lblStatus.Caption = "Hoàn Thành!"
    DeleteFile Environ$("WINDIR") & "\FileList.lst"
    UniMsgBox "D9a4 ca2i d9a85t xong!", vbOKOnly + vbInformation, "OK"

    If Me.chkRun.Value = True Then Shell txtPath.Text & "PAV2009.exe", vbNormalFocus
    If Me.chkHelp.Value = True Then ShellExecute Me.hWnd, vbNullString, txtPath.Text & "Help\HuongDan.html", vbNullString, "", 1
    cmdExit_Click
End Sub

'Private Sub Command1_Click()

'Dim sFileName As String
'Dim sFileNum As String
'Dim InputData As String
'Open AppPath & "FileList.lst" For Input As #1
'Do While Not EOF(1)
'    Line Input #1, InputData
'    sFileNum = Split(InputData, "|", , vbBinaryCompare)(0)
'    sFileName = Split(InputData, "|", , vbBinaryCompare)(1)
'    If UCase(sFileName) <> "S" Then
'        If MsgBox("|" & sFileNum & "| - |" & sFileName & "|", vbYesNo) = vbNo Then Exit Sub
'    End If
'Loop
'    Close #1
'End Sub

Private Sub Form_Load()
    txtPath.Text = Environ$("PROGRAMFILES") & "\Perfect Antivirus 2009\"

    strFileList.Text = "S S" & vbCrLf & "101 DATA.PAV" & vbCrLf & "102 Data.Str" & vbCrLf & "103 RegProtect.dat" & vbCrLf & "104 scrrun.dll" & vbCrLf & "105 UniControls_v2.0.ocx" & vbCrLf & "106 Check_Virus.exe" & vbCrLf & "107 PAV2009.exe" & vbCrLf & "108 PAV2009.ico" & vbCrLf & "109 PAV2009Manager.exe" & vbCrLf & "110 PAVExplorer.exe" & vbCrLf & "111 PAVMiniScan.exe" & vbCrLf & "112 PAVStartUp.exe" & vbCrLf & "113 PAVSysReport.exe" & vbCrLf & "114 PAVUpdate.exe" & vbCrLf & "115 HuongDan.html" & vbCrLf & "116 1.GIF" & vbCrLf & "117 2.GIF" & vbCrLf & "118 3.GIF" & vbCrLf & "119 4.GIF" & vbCrLf & "120 5.GIF" & vbCrLf & "121 6.GIF" & vbCrLf & "122 7.GIF" & vbCrLf & "123 8.GIF" & vbCrLf & "124 9.GIF" & vbCrLf
    strFileList.Text = strFileList.Text & "" & "125 10.GIF" & vbCrLf & "126 11.GIF" & vbCrLf & "127 12.GIF" & vbCrLf & "128 13.GIF" & vbCrLf & "129 14.GIF" & vbCrLf & "130 15.GIF" & vbCrLf & "131 16.GIF" & vbCrLf & "132 Deleted.wav" & vbCrLf & "133 Found.wav" & vbCrLf & "134 Mes.wav" & vbCrLf & "135 Ring.wav" & vbCrLf & "136 ScanDone.wav" & vbCrLf & "137 RemovePAV2009.exe"
    strFileList.Text = Replace(strFileList.Text, " ", "|")

    Bar1.Value = 0
End Sub

Private Sub CreateFolder(sPath)

    On Error Resume Next

    MkDir sPath
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Cancel = 1
End Sub

Private Sub CreateShorcut(sEXE, sLNK)

    Dim ObjShell

    Dim ObjShortCut

    Set ObjShell = CreateObject("WScript.Shell")
    Set ObjShortCut = ObjShell.createshortcut(sLNK)
    ObjShortCut.TargetPath = sEXE
    ObjShortCut.Save
    Set ObjShell = Nothing
    Set ObjShortCut = Nothing
End Sub
Private Sub AddFileU(sFile)
UnFileList = UnFileList & sFile & vbCrLf
End Sub
