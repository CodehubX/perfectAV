Attribute VB_Name = "basProcess"

Option Explicit

Public Const TOKEN_ADJUST_PRIVILEGES = &H20

Public Const TOKEN_QUERY = &H8

Public Const SE_PRIVILEGE_ENABLED = &H2

Public Const ANYSIZE_ARRAY = 1

Public Const SE_DEBUG_NAME = "SeDebugPrivilege"

Public Declare Function GetCurrentProcess Lib "kernel32" () As Long

Public Declare Function OpenProcessToken _
               Lib "advapi32" (ByVal ProcessHandle As Long, _
                               ByVal DesiredAccess As Long, _
                               TokenHandle As Long) As Long

Public Declare Function LookupPrivilegeValue _
               Lib "advapi32" _
               Alias "LookupPrivilegeValueA" (ByVal lpSystemName As String, _
                                              ByVal lpName As String, _
                                              lpLuid As LUID) As Long

Public Declare Function AdjustTokenPrivileges _
               Lib "advapi32" (ByVal TokenHandle As Long, _
                               ByVal DisableAllPrivileges As Long, _
                               NewState As TOKEN_PRIVILEGES, _
                               ByVal BufferLength As Long, _
                               PreviousState As TOKEN_PRIVILEGES, _
                               ReturnLength As Long) As Long

Public Declare Function GetExitCodeProcess _
               Lib "kernel32" (ByVal hProcess As Long, _
                               lpExitCode As Long) As Long

'To suspend or resume thread
Public Declare Function SuspendThread Lib "kernel32" (ByVal hthread As Long) As Long

Public Declare Function ResumeThread Lib "kernel32" (ByVal hthread As Long) As Long

'To open a thread
Public Declare Function CreateToolhelp32Snapshot _
               Lib "kernel32" (ByVal dwFlags As Long, _
                               ByVal dwProcessId As Long) As Long

Public Declare Function Thread32First _
               Lib "kernel32" (ByVal hObject As Long, _
                               p As THREADENTRY32) As Boolean

Public Declare Function Thread32Next _
               Lib "kernel32" (ByVal hObject As Long, _
                               p As THREADENTRY32) As Boolean

'To open a thread
Public Declare Function OpenThread _
               Lib "kernel32" (ByVal dwDesiredAccess As Long, _
                               ByVal bInheritHandle As Long, _
                               ByVal dwProcessId As Long) As Long

Public Const STANDARD_RIGHTS_REQUIRED = &HF0000

Public Const SYNCHRONIZE = &H100000

'THREAD_SUSPEND_RESUME can be used instead of THREAD_ALL_ACCESS for suspending and resuming a thread
Public Const THREAD_ALL_ACCESS = STANDARD_RIGHTS_REQUIRED Or SYNCHRONIZE Or &H3FF

'Current running process PID or ID
Public Declare Function GetCurrentProcessId Lib "kernel32" () As Long

' This function is not from ToolHelp but you need it to destroy a snapshot
Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Public Declare Function CreateToolhelpSnapshot _
               Lib "kernel32" _
               Alias "CreateToolhelp32Snapshot" (ByVal lFlags As Long, _
                                                 ByVal lProcessID As Long) As Long

Public Declare Function ProcessFirst _
               Lib "kernel32" _
               Alias "Process32First" (ByVal hsnapshot As Long, _
                                       uProcess As PROCESSENTRY32) As Long

Public Declare Function ProcessNext _
               Lib "kernel32" _
               Alias "Process32Next" (ByVal hsnapshot As Long, _
                                      uProcess As PROCESSENTRY32) As Long

Public Declare Function TerminateProcess _
               Lib "kernel32" (ByVal hProcess As Long, _
                               ByVal uExitCode As Long) As Long

Public Declare Function OpenProcess _
               Lib "kernel32" (ByVal dwDesiredAccess As Long, _
                               ByVal bInheritHandle As Long, _
                               ByVal dwProcessId As Long) As Long

Public Const TH32CS_SNAPPROCESS = &H2

Public Const TH32CS_SNAPheaplist = &H1

Public Const TH32CS_SNAPthread = &H4

Public Const TH32CS_SNAPmodule = &H8

Public Const TH32CS_SNAPall = TH32CS_SNAPPROCESS + TH32CS_SNAPheaplist + TH32CS_SNAPthread + TH32CS_SNAPmodule

Public Declare Function EnumProcessModules _
               Lib "psapi.dll" (ByVal hProcess As Long, _
                                ByRef lphModule As Long, _
                                ByVal cb As Long, _
                                ByRef cbNeeded As Long) As Long

Public Declare Function GetModuleFileNameExA _
               Lib "psapi.dll" (ByVal hProcess As Long, _
                                ByVal hModule As Long, _
                                ByVal ModuleName As String, _
                                ByVal nSize As Long) As Long

Public Const MAX_PATH As Integer = 260

'define PROCESSENTRY32 structure

Public Type PROCESSENTRY32

    dwSize As Long
    cntUsage As Long
    th32ProcessID As Long
    th32DefaultHeapID As Long
    th32ModuleID As Long
    cntThreads As Long
    th32ParentProcessID As Long
    pcPriClassBase As Long
    dwFlags As Long
    szExeFile As String * MAX_PATH

End Type

Private Const PROCESS_QUERY_INFORMATION = &H400

Private Const PROCESS_VM_READ = &H10

Private Type THREADENTRY32

    dwSize As Long
    cntUsage As Long
    th32ThreadID As Long
    th32OwnerProcessID As Long
    tpBasePri As Long
    tpDeltaPri As Long
    dwFlags As Long

End Type

Private Type LUID

    LowPart As Long
    HighPart As Long

End Type

Private Type LUID_AND_ATTRIBUTES

    pLuid As LUID
    Attributes As Long

End Type

Private Type TOKEN_PRIVILEGES

    PrivilegeCount As Long
    Privileges(ANYSIZE_ARRAY) As LUID_AND_ATTRIBUTES

End Type

Sub GetDebugPriv()

        '<EhHeader>
        On Error GoTo GetDebugPriv_Err

        '</EhHeader>
        Dim hToken           As Long

        Dim sedebugnameValue As LUID

        Dim tkp              As TOKEN_PRIVILEGES, mNewPriv As TOKEN_PRIVILEGES

100     If OpenProcessToken(GetCurrentProcess(), TOKEN_ADJUST_PRIVILEGES Or TOKEN_QUERY, hToken) Then
102         If LookupPrivilegeValue(vbNullString, SE_DEBUG_NAME, sedebugnameValue) Then
104             tkp.PrivilegeCount = 1
106             tkp.Privileges(0).pLuid = sedebugnameValue
108             tkp.Privileges(0).Attributes = SE_PRIVILEGE_ENABLED
110             AdjustTokenPrivileges hToken, False, tkp, Len(tkp), mNewPriv, Len(mNewPriv)
112             CloseHandle (hToken)
            Else
114             CloseHandle (hToken)

                Exit Sub

            End If
        End If

        '<EhFooter>
        Exit Sub

GetDebugPriv_Err:
        WriteErr Time & "-" & Date & " - " & Err.Description & " - " & "in PerfectAntivirus2009.basProcess.GetDebugPriv " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Function KillProcessById(lPID As Long)

        '<EhHeader>
        On Error GoTo KillProcessById_Err

        '</EhHeader>

        Dim hnd As Long, t1 As Long

100     hnd = OpenProcess(&H1&, 0, lPID)
102     Call GetDebugPriv
104     Call GetExitCodeProcess(hnd, t1)
106     Call TerminateProcess(hnd, t1)
108     Call CloseHandle(hnd)

        '<EhFooter>
        Exit Function

KillProcessById_Err:
        WriteErr Time & "-" & Date & " - " & Err.Description & " - " & "in PerfectAntivirus2009.basProcess.KillProcessById " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Function

Function ProcessPathByPID(pid As Long) As String

        'Return path to the executable from PID
        'http://support.microsoft.com/default.aspx?scid=kb;en-us;187913
        '<EhHeader>
        On Error GoTo ProcessPathByPID_Err

        '</EhHeader>
        Dim cbNeeded          As Long

        Dim Modules(1 To 200) As Long

        Dim Ret               As Long

        Dim ModuleName        As String

        Dim nSize             As Long

        Dim hProcess          As Long

100     hProcess = OpenProcess(PROCESS_QUERY_INFORMATION Or PROCESS_VM_READ, 0, pid)
            
102     If hProcess <> 0 Then
                
104         Ret = EnumProcessModules(hProcess, Modules(1), 200, cbNeeded)
                
106         If Ret <> 0 Then
108             ModuleName = Space(MAX_PATH)
110             nSize = 500
112             Ret = GetModuleFileNameExA(hProcess, Modules(1), ModuleName, nSize)
114             ProcessPathByPID = Left(ModuleName, Ret)
            End If
        End If
          
116     Ret = CloseHandle(hProcess)

118     If ProcessPathByPID = "" Then
120         ProcessPathByPID = "SYSTEM"
        End If

        '<EhFooter>
        Exit Function

ProcessPathByPID_Err:
        WriteErr Time & "-" & Date & " - " & Err.Description & " - " & "in PerfectAntivirus2009.basProcess.ProcessPathByPID " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Function

Public Function SuspendResumeProcess(ByVal procid As Long, _
                                     ByVal suspendresume As Boolean) As Boolean

        '<EhHeader>
        On Error GoTo SuspendResumeProcess_Err

        '</EhHeader>
        Dim hsnapshot As Long

        Dim htthread  As Long

        Dim pthread   As Boolean

        Dim pt        As THREADENTRY32

100     SuspendResumeProcess = False

102     hsnapshot = CreateToolhelp32Snapshot(TH32CS_SNAPthread, 0)

104     pt.dwSize = Len(pt)

106     pthread = Thread32First(hsnapshot, pt)

108     While pthread

110         If pt.th32OwnerProcessID = procid Then
112             htthread = OpenThread(THREAD_ALL_ACCESS, 0, pt.th32ThreadID)

114             If htthread <> 0 Then
116                 If suspendresume Then SuspendThread (htthread) Else ResumeThread (htthread)
118                 CloseHandle htthread
120                 SuspendResumeProcess = True
                End If
            End If

122         pthread = Thread32Next(hsnapshot, pt)

        Wend

124     CloseHandle hsnapshot

        '<EhFooter>
        Exit Function

SuspendResumeProcess_Err:
        WriteErr Time & "-" & Date & " - " & Err.Description & " - " & "in PerfectAntivirus2009.basProcess.SuspendResumeProcess " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Function

Public Function CheckProcess(FilePath As String) As Long

        '<EhHeader>
        On Error GoTo CheckProcess_Err

        '</EhHeader>

100     CheckProcess = 0

        Dim theloop As Long

        Dim proc    As PROCESSENTRY32

        Dim snap    As Long

        Dim strtmp  As String

102     snap = CreateToolhelpSnapshot(TH32CS_SNAPall, 0)
104     proc.dwSize = Len(proc)
106     theloop = ProcessFirst(snap, proc)

108     While theloop <> 0

110         theloop = ProcessNext(snap, proc)
112         strtmp = ProcessPathByPID(proc.th32ProcessID)

114         If strtmp <> "SYSTEM" Then
116             If UCase(strtmp) = UCase(FilePath) Then CheckProcess = proc.th32ProcessID: GoTo KetThuc
            End If

        Wend

118     CloseHandle snap
KetThuc:

        '<EhFooter>
        Exit Function

CheckProcess_Err:
        WriteErr Time & "-" & Date & " - " & Err.Description & " - " & "in PerfectAntivirus2009.basProcess.CheckProcess " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Function

Public Function CheckID(ID As Long) As Long

        '<EhHeader>
        On Error GoTo CheckID_Err

        '</EhHeader>

100     CheckID = 0

        Dim theloop As Long

        Dim proc    As PROCESSENTRY32

        Dim snap    As Long

102     snap = CreateToolhelpSnapshot(TH32CS_SNAPall, 0)
104     proc.dwSize = Len(proc)
106     theloop = ProcessFirst(snap, proc)

108     While theloop <> 0

110         theloop = ProcessNext(snap, proc)

112         If proc.th32ProcessID = ID Then CheckID = ID: GoTo KetThuc

        Wend

114     CloseHandle snap
KetThuc:

        '<EhFooter>
        Exit Function

CheckID_Err:
        WriteErr Time & "-" & Date & " - " & Err.Description & " - " & "in PerfectAntivirus2009.basProcess.CheckID " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Function
