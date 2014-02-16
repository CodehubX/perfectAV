Attribute VB_Name = "mdlScan"

Option Explicit

Private Const MaxLen = 260
Private Const INVALID_HANDLE_VALUE = -1
Private Const FILE_ATTRIBUTE_SYSTEM = &H4
Private Const FILE_ATTRIBUTE_READONLY = &H1
Private Const FILE_ATTRIBUTE_HIDDEN = &H2
Private Const FILE_ATTRIBUTE_DIRECTORY = &H10
Private Const FILE_ATTRIBUTE_ARCHIVE = &H20
Private Const FILE_ATTRIBUTE_NORMAL = &H80
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Declare Function lstrlen Lib _
    "kernel32" Alias "lstrlenA" ( _
    ByVal lpString As String) As Long
Private Declare Function FindFirstFile Lib _
    "kernel32" Alias "FindFirstFileA" _
    (ByVal lpFileName As String, _
    lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindNextFile Lib _
    "kernel32" Alias "FindNextFileA" _
    (ByVal hFindFile As Long, _
    lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindClose Lib _
    "kernel32" (ByVal hFindFile As Long) As Long

    
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
    cFileName As String * MaxLen
    cShortFileName As String * 14
End Type

Dim FileSpec As String, UseFileSpec As Boolean
Dim WFD As WIN32_FIND_DATA, hFindFile As Long
Public nFile As Long
Public nVirus As Long
Public Sub scanvirus(Path As String, xEXT As String, xYzTotal As Long)
    'On Error Resume Next
    'path = FixPath(path)
    Dim dirs As Integer, dirbuff() As String, i As Integer
    Dim sFileName As String
    hFindFile = FindFirstFile(Path & "*", WFD)
    If hFindFile <> INVALID_HANDLE_VALUE Then
        Do
            If (WFD.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) Then
                If Asc(WFD.cFileName) <> 46 Then
                    If (dirs Mod 10) = 0 Then ReDim Preserve dirbuff(dirs + 10)
                    dirs = dirs + 1
                    dirbuff(dirs) = StripNulls(WFD.cFileName)
                End If
            End If
        Loop While FindNextFile(hFindFile, WFD)
        Call FindClose(hFindFile)
    End If
    hFindFile = FindFirstFile(Path & xEXT, WFD)
    If hFindFile <> INVALID_HANDLE_VALUE Then
        Do
        If frmScan.xStopScan = True Then Exit Sub
            sFileName = StripNulls(WFD.cFileName)
                If modMain.FileExists(Path & sFileName) = True Then
                    'frmScan.lblStatus.Caption = ThuGonPath(Path) & sFilename: DoEvents
                    'frmScan.lblStatus.Caption = sFilename: DoEvents
                    nFile = nFile + 1
                    'frmScan.lblProcess.Caption = Abs(Round((nFile * 100) / xYzTotal, 2)) & " % (" & nFile & ")": DoEvents
                    frmScan.lblProcess.Caption = nFile
                    '====================
                    '==  Check Virus   ==
                    '====================
                    Dim Ax As String
                    Ax = CheckVirus(Path & sFileName)
                    If Ax <> "No" Then
                        'frmScan.Caption = frmScan.Caption + 1
                        nVirus = nVirus + 1
                        frmScan.lblVirus.Caption = nVirus
                        Dim Ui As Integer
                        Ui = frmScan.LV1.ListItems.Count + 1
                        frmScan.LV1.ListItems.Add Ui, , Ax
                        frmScan.LV1.ListItems(Ui).SubItems(1).Caption = Path & sFileName
                    End If
                End If
        Loop While FindNextFile(hFindFile, WFD)
        Call FindClose(hFindFile)
    End If
    For i = 1 To dirs
        Call scanvirus(Path & dirbuff(i) & "\", xEXT, xYzTotal)
    Next i
End Sub



Function StripNulls(ByVal sStr As String) As String
    StripNulls = Left$(sStr, lstrlen(sStr))
End Function

Public Sub TongSoFile(Path As String, xEXT As String)
    Dim dirs As Integer, dirbuff() As String, i As Integer
    Dim sFileName As String
    hFindFile = FindFirstFile(Path & "*", WFD)
    If hFindFile <> INVALID_HANDLE_VALUE Then
        Do
            If (WFD.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) Then
                If Asc(WFD.cFileName) <> 46 Then
                    If (dirs Mod 10) = 0 Then ReDim Preserve dirbuff(dirs + 10)
                    dirs = dirs + 1
                    dirbuff(dirs) = StripNulls(WFD.cFileName)
                End If
            End If
        Loop While FindNextFile(hFindFile, WFD)
        Call FindClose(hFindFile)
    End If
    hFindFile = FindFirstFile(Path & xEXT, WFD)
    If hFindFile <> INVALID_HANDLE_VALUE Then
        Do
        DoEvents
            sFileName = StripNulls(WFD.cFileName)
            DoEvents
                If modMain.FileExists(Path & sFileName) = True Then
                '====================
                DoEvents
                    frmScan.xTongSoFile = frmScan.xTongSoFile + 1: DoEvents
                End If
        Loop While FindNextFile(hFindFile, WFD)
        DoEvents
        Call FindClose(hFindFile)
    End If
    For i = 1 To dirs
        Call TongSoFile(Path & dirbuff(i) & "\", xEXT)
    Next i
End Sub

Public Function ThuGonPath(xPath)
On Error Resume Next
If Len(xPath) > 40 Then
    ThuGonPath = Left(xPath, 40) & "...\"
Else
    ThuGonPath = xPath
End If
End Function

Public Function MakeReport()
On Error Resume Next
If frmScan.chk(3).Value = False Then
    frmScan.lblPro(5).ForeColor = &HC0C0C0
    frmScan.PicIcon(5).Picture = frmScan.picOK.Picture
    frmScan.LBL.Caption = "D9a4 que1t xong!"
    Exit Function
End If
Dim xAA As String
xAA = frmScan.LBL.Caption
Dim xRe As String
xRe = ToUnicode(" Ma64u ba1o ca1o la62n que1t Virus" & vbCrLf & vbCrLf _
& " Tho72i gian: " & Date & " - " & Time & vbCrLf _
& " D9i5a chi3 thu7 mu5c ca62n que1t:" & vbCrLf)
Dim Ok As Integer
For Ok = 0 To frmScan.lstPath.ListCount - 1
    DoEvents
    xRe = xRe & " " & frmScan.lstPath.List(Ok) & vbCrLf
Next Ok
xRe = xRe & vbCrLf & ToUnicode(" Loa5i file: " & IIf(frmScan.optEXT(0).Value, "Ta61t ca3 ca1c File (*.*)", frmScan.cboEXT.Text) & vbCrLf & vbCrLf _
& " To63ng so61 File: " & frmScan.xTongSoFile & vbCrLf _
& " So61 File d9a4 que1t: " & mdlScan.nFile & vbCrLf _
& " So61 Virus pha1t hie65n: " & frmScan.lblVirus.Caption & vbCrLf _
& " To63ng tho72i gian que1t: " & frmScan.lblTime.Caption & vbCrLf _
& "   Chi tie61t:" & vbCrLf)
Dim Uk As Integer
For Uk = 1 To frmScan.LV1.ListItems.Count
    DoEvents
    xRe = xRe & "     " & frmScan.LV1.ListItems(Uk).Text & ": " & frmScan.LV1.ListItems(Uk).SubItems(1).Caption & vbCrLf
Next Uk
frmScan.lblPro(5).ForeColor = &HC0C0C0
frmScan.PicIcon(5).Picture = frmScan.picOK.Picture
frmScan.LBL.Caption = "D9a4 que1t xong!"
On Error GoTo KhOnGtAoFoLdEr
MkDir AppPath & "NhatKy\"
KhOnGtAoFoLdEr:
frmScan.xNowReport = AppPath & "NhatKy\" & File2Str("  " & frmScan.lstPath.List(0))
modReadWrite.WriteFileUni frmScan.xNowReport, xRe
Shell "notepad " & frmScan.xNowReport, vbNormalFocus
End Function

Public Sub ScanProcess()
On Error Resume Next
Dim ColItems
Dim ObjItem
Set ColItems = GetObject("winmgmts:\root\CIMV2").ExecQuery("SELECT * FROM Win32_Process")
For Each ObjItem In ColItems
   'frmMain.lblStatus.Caption = ObjItem.ExecutablePath
   If IsNull(ObjItem.ExecutablePath) = False Then
        'List1.AddItem ObjItem.ExecutablePath
        Dim Ax As String
        Ax = CheckVirus(ObjItem.ExecutablePath)
        If Ax <> "No" Then
            'frmScan.Caption = frmScan.Caption + 1
            nVirus = nVirus + 1
            frmScan.lblVirus.Caption = nVirus
            Dim Ui As Integer
            Ui = frmScan.LV1.ListItems.Count + 1
            frmScan.LV1.ListItems.Add Ui, , Ax
            frmScan.LV1.ListItems(Ui).SubItems(1).Caption = ObjItem.ExecutablePath & ToUnicode("[D9a4 d9o1ng ba8ng]")
            basProcess.SuspendResumeProcess ObjItem.ProcessID, True
        End If
   End If
Next
Set ColItems = Nothing
Set ObjItem = Nothing
End Sub

Public Function CheckTinTuong(ByRef xPath2Check$) As Boolean
On Error Resume Next
Dim Oo As Integer
CheckTinTuong = False
frmMain.File1.Path = AppPath & "TinTuong"
frmMain.File1.Refresh
For Oo = 0 To frmMain.File1.ListCount - 1
'MsgBox GetMD5(xPath2Check)
'MsgBox ReadFileUni(AppPath & "TinTuong\" & frmMain.File1.List(Oo))
    If GetMD5(xPath2Check) = Left(frmMain.File1.List(Oo), Len(frmMain.File1.List(Oo)) - 5) Then
        'MsgBox "OK!"
        CheckTinTuong = True
    End If
Next
End Function
Public Sub LoaiBoTinTuong()
On Error Resume Next
BaTdAuLoAiBoTiNtUoNg:
Dim Kl As Integer
For Kl = 1 To frmScan.LV1.ListItems.Count
    If CheckTinTuong(frmScan.LV1.ListItems(Kl).SubItems(1).Caption) = True Then
        frmScan.LV1.ListItems.Remove Kl
        GoTo BaTdAuLoAiBoTiNtUoNg
    End If
Next Kl
End Sub
