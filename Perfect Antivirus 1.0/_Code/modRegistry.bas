Attribute VB_Name = "modRegistry"
Option Explicit

Public Enum RegistryKeys

    HKEY_CLASSES_ROOT = &H80000000
    HKEY_CURRENT_USER = &H80000001
    HKEY_LOCAL_MACHINE = &H80000002
    HKEY_USERS = &H80000003
    HKEY_CURRENT_CONFIG = &H80000005
    HKEY_DYN_DATA = &H80000006

End Enum

Public Const HKEY_PERFORMANCE_DATA = &H80000004

Public Const ERROR_SUCCESS = 0&

Public Const REG_SZ = 1

Public Const REG_DWORD = 4

Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Declare Function RegCreateKey _
        Lib "advapi32.dll" _
        Alias "RegCreateKeyA" (ByVal hKey As Long, _
                               ByVal lpSubKey As String, _
                               phkResult As Long) As Long
Declare Function RegDeleteKey _
        Lib "advapi32.dll" _
        Alias "RegDeleteKeyA" (ByVal hKey As Long, _
                               ByVal lpSubKey As String) As Long
Declare Function RegDeleteValue _
        Lib "advapi32.dll" _
        Alias "RegDeleteValueA" (ByVal hKey As Long, _
                                 ByVal lpValueName As String) As Long
Declare Function RegOpenKey _
        Lib "advapi32.dll" _
        Alias "RegOpenKeyA" (ByVal hKey As Long, _
                             ByVal lpSubKey As String, _
                             phkResult As Long) As Long
Declare Function RegQueryValueEx _
        Lib "advapi32.dll" _
        Alias "RegQueryValueExA" (ByVal hKey As Long, _
                                  ByVal lpValueName As String, _
                                  ByVal lpReserved As Long, _
                                  lpType As Long, _
                                  lpData As Any, _
                                  lpcbData As Long) As Long
Declare Function RegSetValueEx _
        Lib "advapi32.dll" _
        Alias "RegSetValueExA" (ByVal hKey As Long, _
                                ByVal lpValueName As String, _
                                ByVal Reserved As Long, _
                                ByVal dwType As Long, _
                                lpData As Any, _
                                ByVal cbData As Long) As Long

Public Sub SaveKey(ByVal hKey As RegistryKeys, ByVal strPath As String)

        '<EhHeader>
        On Error GoTo SaveKey_Err

        '</EhHeader>
        On Error Resume Next
  
        Dim KeyHand As Long
  
100     RegCreateKey hKey, strPath, KeyHand
102     RegCloseKey KeyHand
  
        '<EhFooter>
        Exit Sub

SaveKey_Err:
        WriteErr Time & "-" & Date & " - " & Err.Description & " - " & "in PerfectAntivirus2009.modRegistry.SaveKey " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Public Function DeleteKey(ByVal hKey As RegistryKeys, ByVal strKey As String)

        '<EhHeader>
        On Error GoTo DeleteKey_Err

        '</EhHeader>
        On Error Resume Next
  
100     RegDeleteKey hKey, strKey

        '<EhFooter>
        Exit Function

DeleteKey_Err:
        WriteErr Time & "-" & Date & " - " & Err.Description & " - " & "in PerfectAntivirus2009.modRegistry.DeleteKey " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Function

Public Function DeleteValue(ByVal hKey As RegistryKeys, _
                            ByVal strPath As String, _
                            ByVal strValue As String)

        '<EhHeader>
        On Error GoTo DeleteValue_Err

        '</EhHeader>
        On Error Resume Next

        Dim KeyHand As Long
  
100     RegOpenKey hKey, strPath, KeyHand
102     RegDeleteValue KeyHand, strValue
104     RegCloseKey KeyHand

        '<EhFooter>
        Exit Function

DeleteValue_Err:
        WriteErr Time & "-" & Date & " - " & Err.Description & " - " & "in PerfectAntivirus2009.modRegistry.DeleteValue " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Function

Public Function GetString(ByVal hKey As RegistryKeys, _
                          ByVal strPath As String, _
                          ByVal strValue As String) As String

        '<EhHeader>
        On Error GoTo GetString_Err

        '</EhHeader>
        On Error Resume Next

        Dim KeyHand      As Long

        Dim datatype     As Long

        Dim lResult      As Long

        Dim strBuf       As String

        Dim lDataBufSize As Long

        Dim intZeroPos   As Integer

        Dim lValueType   As Long
  
100     RegOpenKey hKey, strPath, KeyHand
102     lResult = RegQueryValueEx(KeyHand, strValue, 0&, lValueType, ByVal 0&, lDataBufSize)

104     If lValueType = REG_SZ Then
106         strBuf = String(lDataBufSize, " ")
108         lResult = RegQueryValueEx(KeyHand, strValue, 0&, 0&, ByVal strBuf, lDataBufSize)

110         If lResult = ERROR_SUCCESS Then
112             intZeroPos = InStr(strBuf, Chr(0))

114             If intZeroPos > 0 Then
116                 GetString = Left(strBuf, intZeroPos - 1)
                Else
118                 GetString = strBuf
                End If
            End If
        End If
    
        '<EhFooter>
        Exit Function

GetString_Err:
        WriteErr Time & "-" & Date & " - " & Err.Description & " - " & "in PerfectAntivirus2009.modRegistry.GetString " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Function

Public Sub SaveString(ByVal hKey As RegistryKeys, _
                      ByVal strPath As String, _
                      ByVal strValue As String, _
                      ByVal strData As String)

        '<EhHeader>
        On Error GoTo SaveString_Err

        '</EhHeader>
        On Error Resume Next

        Dim KeyHand As Long
  
100     RegCreateKey hKey, strPath, KeyHand
102     RegSetValueEx KeyHand, strValue, 0, REG_SZ, ByVal strData, Len(strData)
104     RegCloseKey KeyHand

        '<EhFooter>
        Exit Sub

SaveString_Err:
        WriteErr Time & "-" & Date & " - " & Err.Description & " - " & "in PerfectAntivirus2009.modRegistry.SaveString " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Function GetDWORD(ByVal hKey As RegistryKeys, _
                  ByVal strPath As String, _
                  ByVal strValueName As String) As Long

        '<EhHeader>
        On Error GoTo GetDWORD_Err

        '</EhHeader>
        On Error Resume Next

        Dim lResult      As Long

        Dim lValueType   As Long

        Dim lBuf         As Long

        Dim lDataBufSize As Long

        Dim KeyHand      As Long

100     RegOpenKey hKey, strPath, KeyHand
102     lDataBufSize = 4
104     lResult = RegQueryValueEx(KeyHand, strValueName, 0&, lValueType, lBuf, lDataBufSize)

106     If lResult = ERROR_SUCCESS Then
108         If lValueType = REG_DWORD Then
110             GetDWORD = lBuf
            End If
        End If

112     RegCloseKey KeyHand
    
        '<EhFooter>
        Exit Function

GetDWORD_Err:
        WriteErr Time & "-" & Date & " - " & Err.Description & " - " & "in PerfectAntivirus2009.modRegistry.GetDWORD " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Function

Function SaveDWORD(ByVal hKey As RegistryKeys, _
                   ByVal strPath As String, _
                   ByVal strValueName As String, _
                   ByVal lData As Long)

        '<EhHeader>
        On Error GoTo SaveDWORD_Err

        '</EhHeader>
        On Error Resume Next

        Dim lResult As Long

        Dim KeyHand As Long
   
100     RegCreateKey hKey, strPath, KeyHand
102     lResult = RegSetValueEx(KeyHand, strValueName, 0&, REG_DWORD, lData, 4)
104     RegCloseKey KeyHand
    
        '<EhFooter>
        Exit Function

SaveDWORD_Err:
        WriteErr Time & "-" & Date & " - " & Err.Description & " - " & "in PerfectAntivirus2009.modRegistry.SaveDWORD " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Function
