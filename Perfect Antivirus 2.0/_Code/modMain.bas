Attribute VB_Name = "modMain"
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long


Public Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Public Const MAX_COMPUTERNAME_LENGTH As Long = 31

Dim sConnType As String * 255
Private Type MEMORYSTATUS
    dwLength As Long
    dwMemoryLoad As Long
    dwTotalPhys As Long
    dwAvailPhys As Long
    dwTotalPageFile As Long
    dwAvailPageFile As Long
    dwTotalVirtual As Long
    dwAvailVirtual As Long
End Type
Private memInfo As MEMORYSTATUS
Dim memoryInfo As MEMORYSTATUS
Dim lastpcent As Single, lastTot As Long
Private Declare Sub GlobalMemoryStatus Lib "kernel32" _
   (lpBuffer As MEMORYSTATUS)
Private Const WM_CLICK = &HF5
Private Const WM_SETTEXT = &HC
Dim WinHwnd As Long
Dim ExHwnd As Long
Public KhanCap As Boolean


Public Function GetRAMTotal() As String
   Call GlobalMemoryStatus(memInfo)
        GetRAMTotal = Round(memInfo.dwTotalPhys / 1024 / 1024, 3) & " MB"
End Function
Function GetMemoryInfo()
  DoEvents
  GlobalMemoryStatus memoryInfo
    Dim Totp1
    Dim Availp1
    Dim pcent
    Dim lastpcent
    Dim lastTot
  Totp1 = Int(memoryInfo.dwTotalPhys / 1044032 * 10 + 0.5) / 10
  Availp1 = Int(memoryInfo.dwAvailPhys / 1044032 * 10 + 0.5) / 10
  pcent = Int(Availp1 / Totp1 * 100)
  lastpcent = pcent
  lastTot = memoryInfo.dwMemoryLoad
  GetMemoryInfo = Format(lastpcent)
End Function

Public Function CheckComputerHeal() As String
On Error Resume Next
Dim sRAM
Dim sKQ
sRAM = GetMemoryInfo
If sRAM < 20 Then
    sKQ = "Ma1y ti1nh d9ang cha5y ra61t cha65m!"
ElseIf sRAM >= 20 And sRAM < 30 Then
    sKQ = "Ma1y ti1nh d9ang cha5y cha65m"
ElseIf sRAM >= 30 And sRAM < 40 Then
    sKQ = "Ma1y ti1nh cha5y o63n d9i5nh,  bi2nh thu7o72ng."
ElseIf sRAM >= 40 And sRAM < 50 Then
    sKQ = "Ma1y ti1nh cha5y bi2nh thu7o72ng"
ElseIf sRAM >= 50 And sRAM < 60 Then
    sKQ = "Ma1y ti1nh cha5y nhanh va2 o63n d9i5nh, ti2nh tra5ng to61t."
ElseIf sRAM >= 60 And sRAM < 70 Then
    sKQ = "Ma1y ti1nh d9ang cha5y ra61t nhanh, to61c d9o65 xu73 ly1 to61t"
ElseIf sRAM >= 70 And sRAM < 80 Then
    sKQ = "Ma1y ti1nh d9ang ra61t to61t"
ElseIf sRAM >= 80 Then
    sKQ = "Ma1y ti1nh cu3a ba5n la2m vie65c 1 ca1ch cho1ng ma85t! Ba5n co1 1 bo65 nho71 RAM tha65t tuye65t vo72i!"
End If
    
frmMain.ProTinhTrang.Value = sRAM
CheckComputerHeal = sKQ
End Function
Function GetComputer()
On Error Resume Next
    Dim dwlen As Long
    Dim strString As String
    dwlen = MAX_COMPUTERNAME_LENGTH + 1
    strString = String(dwlen, "X")
    GetComputerName strString, dwlen
    strString = Left(strString, dwlen)
    GetComputer = strString
End Function
Public Function AppPath()
AppPath = App.Path
If Right(AppPath, 1) <> "\" Then AppPath = AppPath & "\"
End Function
Public Function CoTime(ByVal sGiay As Integer) As String
On Error Resume Next
    Dim ChuoiTime As String
    ChuoiTime = Format(sGiay \ 3600, "00") & ":" '  h
    ChuoiTime = ChuoiTime & Format((sGiay Mod 3600) \ 60, "00") & ":" 'p
    ChuoiTime = ChuoiTime & Format((sGiay Mod 3600) Mod 60, "00") ' s
    CoTime = ChuoiTime
End Function
Public Function FileExists(sFile) As Boolean
On Error Resume Next
FileExists = ((GetAttr(sFile) And vbDirectory) = 0)
End Function

Public Function GetFileName(ByVal sPath As String) As String
GetFileName = Mid(sPath, InStrRev(sPath, "\") + 1)
End Function
Public Function GetFolderPath(ByVal sPath As String) As String
GetFolderPath = Left(sPath, InStrRev(sPath, "\") - 1)
End Function

Public Function FixPath(xPath)
On Error Resume Next
 FixPath = xPath
If Right(FixPath, 1) <> "\" Then FixPath = FixPath & "\"
End Function

Sub Main()
On Error Resume Next
Dim Comd As String
Comd = Command()

KhanCap = False
If Comd = "/logon" Then
    KhanCap = True
    Form1.Show
    GoTo KhanCap
End If

If FileExists(Comd) = True Then
    Dim Ax As String
    ConnectDB
    Ax = CheckVirus(Comd)
    If Ax = "No" Then
        UniMsgBox "File na2y kho6ng co1 Virus!", vbOKOnly, "D9a4 kie63m tra xong!"
    Else
        If UniMsgBox("!!! File na2y bi5 nhie63m Virus: " & Ax & " !!!" & vbCrLf & "Ba5n co1 muo61n die65t no1 ngay ba6y gio72 kho6ng?", vbYesNo, "VIRUS!!!") = vbYes Then
            tXoaFile Comd
        End If
    End If
End
End If
        

If DirExists(FixPath(Comd)) = True Then
    If App.PrevInstance = True Then
        WinHwnd = FindWindow(vbNullString, "PAV_FOLDER_SCAN_VIRUS")
        'MsgBox WinHwnd
        ExHwnd = FindWindowEx(WinHwnd, 0&, vbNullString, "ScanVirusForFolderNow")
        'MsgBox ExHwnd
        If WinHwnd > 0 Then
        Dim xPathScan As String
        xPathScan = FixPath(Comd)
            SendMessage ExHwnd, WM_SETTEXT, 0&, ByVal xPathScan
            SendMessage ExHwnd, WM_CLICK, 0&, 0&
            SendMessage ExHwnd, WM_SETTEXT, 0&, ByVal "ScanVirusForFolderNow"
        End If
        '===========================
    End
    Else
        frmScan.Show
        frmScan.lstPath.AddItem FixPath(Comd)
        frmScan.RefereshListPath
    End If
End If

If Comd = "/task" Then
    Load frmMenu
    If GetSetting("PAV2009", "RealTimeProtection", "OnOff", True) = True Then
        Load frmRTP
    End If
    If GetSetting("PAV2009", "AutorunProtect", "OnOff", True) = True Then
        Load frmAutorun
    End If
    Dim J As Integer
    If GetSetting("PAV2009", "Update", "0", True) = True Then
        If modMain.FileExists(AppPath & "PAVUPDATE.EXE") = True Then
            Shell AppPath & "PAVUPDATE.EXE", vbNormalFocus
        Else
            UniMsgBox "Kho6ng ti2m tha61y file PAVUPDATE.EXE!", vbOKOnly, "Error!"
        End If
    End If
Else
    frmMain.Show
End If

Load frmFOLDERSCAN
'SaveSetting "PAV2009", "Setting", "PhucHoi", CHK0(1).Value
If GetSetting("PAV2009", "Setting", "PhucHoi", True) = True Then RegistryClean
KhanCap:
End Sub
Public Function DirExists(sDir As String) As Boolean
On Error Resume Next
DirExists = ((GetAttr(sDir) And vbDirectory) <> 0)
If sDir = "\" Then DirExists = False
End Function

Public Function GetTotalProcess() As Integer
On Error Resume Next
Dim ColItems
Dim ObjItem
Set ColItems = GetObject("winmgmts:\root\CIMV2").ExecQuery("SELECT * FROM Win32_Process")
For Each ObjItem In ColItems
   'frmMain.lblStatus.Caption = ObjItem.ExecutablePath
   If IsNull(ObjItem.ExecutablePath) = False Then
        'List1.AddItem ObjItem.ExecutablePath
        GetTotalProcess = GetTotalProcess + 1
   End If
Next
Set ColItems = Nothing
Set ObjItem = Nothing
End Function

