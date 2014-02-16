Attribute VB_Name = "modSaveSetting"
Option Explicit
Declare Function GetPrivateProfileString _
        Lib "kernel32" _
        Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, _
                                          ByVal lpKeyName As Any, _
                                          ByVal lpDefault As String, _
                                          ByVal lpReturnedString As String, _
                                          ByVal nSize As Long, _
                                          ByVal lpFileName As String) As Long
Declare Function WritePrivateProfileString _
        Lib "kernel32" _
        Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, _
                                            ByVal lpKeyName As Any, _
                                            ByVal lpString As Any, _
                                            ByVal lpFileName As String) As Long
Declare Function DeleteFile _
        Lib "kernel32.dll" _
        Alias "DeleteFileA" (ByVal lpFileName As String) As Long
 
'Read file
Function WriteIniFile(ByVal sIniFileName As String, _
                      ByVal sSection As String, _
                      ByVal sItem As String, _
                      ByVal sText As String) As Boolean

        '<EhHeader>
        On Error GoTo WriteIniFile_Err

        '</EhHeader>
        Dim I As Integer

        On Error GoTo sWriteIniFileError

100     I = WritePrivateProfileString(sSection, sItem, sText, sIniFileName)
102     WriteIniFile = True

        Exit Function

sWriteIniFileError:
104     WriteIniFile = False

        '<EhFooter>
        Exit Function

WriteIniFile_Err:
        WriteErr Time & "-" & Date & " - " & Err.Description & " - " & "in PerfectAntivirus2009.modSaveSetting.WriteIniFile " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Function

'Write file
Function ReadIniFile(ByVal sIniFileName As String, _
                     ByVal sSection As String, _
                     ByVal sItem As String, _
                     ByVal sDefault As String) As String

        '<EhHeader>
        On Error GoTo ReadIniFile_Err

        '</EhHeader>
        Dim iRetAmount As Integer

        Dim sTemp      As String

100     sTemp = String$(100, 0)
102     iRetAmount = GetPrivateProfileString(sSection, sItem, sDefault, sTemp, 100, sIniFileName)
104     sTemp = Left$(sTemp, iRetAmount)
106     ReadIniFile = sTemp

        '<EhFooter>
        Exit Function

ReadIniFile_Err:
        WriteErr Time & "-" & Date & " - " & Err.Description & " - " & "in PerfectAntivirus2009.modSaveSetting.ReadIniFile " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Function
