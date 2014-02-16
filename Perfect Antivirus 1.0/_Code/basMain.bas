Attribute VB_Name = "basMain"

Public Declare Function ShellExecute _
               Lib "shell32.dll" _
               Alias "ShellExecuteA" (ByVal hWnd As Long, _
                                      ByVal lpOperation As String, _
                                      ByVal lpFile As String, _
                                      ByVal lpParameters As String, _
                                      ByVal lpDirectory As String, _
                                      ByVal nShowCmd As Long) As Long

Public Declare Function GetComputerName _
               Lib "kernel32" _
               Alias "GetComputerNameA" (ByVal lpBuffer As String, _
                                         nSize As Long) As Long

Public Const MAX_COMPUTERNAME_LENGTH As Long = 31

Dim sConnType                        As String * 255

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

Dim memoryInfo  As MEMORYSTATUS

Dim lastpcent   As Single, lastTot As Long

Private Declare Sub GlobalMemoryStatus Lib "kernel32" (lpBuffer As MEMORYSTATUS)

Public Function CheckFiles() As Boolean
CheckFiles = True

End Function

Sub Main()

        '<EhHeader>
        On Error GoTo Main_Err

        '</EhHeader>
100     If App.PrevInstance = True Then End

        Dim Comd

102     Comd = Command()

104     If Comd = "/task" Then
            'Run in Taskbar
106         frmMain.xTask = True
108         Load frmMain
        Else

            'Run Normal
110         If ReadIniFile(AppPath & "Setting.ini", "Setting", "FlashScreen", True) = True Then
112             frmFlash.Show
            Else
114             frmMain.xTask = False
116             Load frmMain
            End If
        End If

        '<EhFooter>
        Exit Sub

Main_Err:
        WriteErr Time & "-" & Date & " - " & Err.Description & " - " & "in PerfectAntivirus2009.basMain.Main " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Function GetComputer()

        '<EhHeader>
        On Error GoTo GetComputer_Err

        '</EhHeader>
        Dim dwlen     As Long

        Dim strString As String

100     dwlen = MAX_COMPUTERNAME_LENGTH + 1
102     strString = String(dwlen, "X")
104     GetComputerName strString, dwlen
106     strString = Left(strString, dwlen)
108     GetComputer = strString

        '<EhFooter>
        Exit Function

GetComputer_Err:
        WriteErr Time & "-" & Date & " - " & Err.Description & " - " & "in PerfectAntivirus2009.basMain.GetComputer " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Function

Public Function GetRAMTotal() As String

        '<EhHeader>
        On Error GoTo GetRAMTotal_Err

        '</EhHeader>

100     Call GlobalMemoryStatus(memInfo)
102     GetRAMTotal = Round(memInfo.dwTotalPhys / 1024 / 1024, 3) & " MB"

        '<EhFooter>
        Exit Function

GetRAMTotal_Err:
        WriteErr Time & "-" & Date & " - " & Err.Description & " - " & "in PerfectAntivirus2009.basMain.GetRAMTotal " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Function

Public Function CheckComputerHeal() As String

        '<EhHeader>
        On Error GoTo CheckComputerHeal_Err

        '</EhHeader>
        Dim sRAM

        Dim sKQ

100     sRAM = GetMemoryInfo

102     If sRAM < 20 Then
104         sKQ = "Ma1y ti1nh d9ang ra61t na85ng va2 lag, co1 le4 ba5n d9ang cha5y ra61t nhie62u chu7o7ng tri2nh!"
106     ElseIf sRAM >= 20 And sRAM < 30 Then
108         sKQ = "Ma1y ti1nh d9ang lag"
110     ElseIf sRAM >= 30 And sRAM < 40 Then
112         sKQ = "Ma1y ti1nh cha5y o63n d9i5nh,  bi2nh thu7o72ng."
114     ElseIf sRAM >= 40 And sRAM < 50 Then
116         sKQ = "Ma1y ti1nh cha5y bi2nh thu7o72ng"
118     ElseIf sRAM >= 50 And sRAM < 60 Then
120         sKQ = "Ma1y ti1nh cha5y nhanh va2 o63n d9i5nh, ti2nh tra5ng to61t."
122     ElseIf sRAM >= 60 And sRAM < 70 Then
124         sKQ = "Ma1y ti1nh d9ang cha5y ra61t nhanh, to61c d9o65 xu73 ly1 to61t"
126     ElseIf sRAM >= 70 And sRAM < 80 Then
128         sKQ = "Ma1y ti1nh d9ang ra61t to61t"
130     ElseIf sRAM >= 80 Then
132         sKQ = "Ma1y ti1nh cu3a ba5n la2m vie65c 1 ca1ch cho1ng ma85t! Ba5n co1 1 bo65 nho71 RAM tha65t tuye65t vo72i!"
        End If
    
134     sKQ = sKQ & vbCrLf & "[Free RAM: " & sRAM & " %]"
136     CheckComputerHeal = sKQ

        '<EhFooter>
        Exit Function

CheckComputerHeal_Err:
        WriteErr Time & "-" & Date & " - " & Err.Description & " - " & "in PerfectAntivirus2009.basMain.CheckComputerHeal " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Function

Function GetMemoryInfo()

        '<EhHeader>
        On Error GoTo GetMemoryInfo_Err

        '</EhHeader>
100     DoEvents

102     GlobalMemoryStatus memoryInfo

        Dim Totp1

        Dim Availp1

        Dim pcent

        Dim lastpcent

        Dim lastTot

104     Totp1 = Int(memoryInfo.dwTotalPhys / 1044032 * 10 + 0.5) / 10
106     Availp1 = Int(memoryInfo.dwAvailPhys / 1044032 * 10 + 0.5) / 10
108     pcent = Int(Availp1 / Totp1 * 100)
110     lastpcent = pcent
112     lastTot = memoryInfo.dwMemoryLoad
114     GetMemoryInfo = Format(lastpcent)

        '<EhFooter>
        Exit Function

GetMemoryInfo_Err:
        WriteErr Time & "-" & Date & " - " & Err.Description & " - " & "in PerfectAntivirus2009.basMain.GetMemoryInfo " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Function

Public Function ReadFileUni(FileName As String) As String

        '<EhHeader>
        On Error GoTo ReadFileUni_Err

        '</EhHeader>

If FileExists(FileName) = False Then GoTo ReadFileUni_Err
        Dim FSO

100     Set FSO = CreateObject("Scripting.FileSystemObject").OpenTextFile(FileName, 1, , -2)
102     ReadFileUni = FSO.ReadAll
104     Set FSO = Nothing

        '<EhFooter>
        Exit Function

ReadFileUni_Err:

        ReadFileUni = ""

        '</EhFooter>

End Function

Public Sub WriteErr(ErrStr)

    On Error Resume Next

    Dim Xa As String

    Xa = ReadFileUni(AppPath & "Err.txt")
    Xa = Xa & vbCrLf & ErrStr
    WriteFileUni AppPath & "Err.txt", Xa

End Sub
