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
Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpdata As Any, lpcbData As Long) As Long
Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpdata As Any, ByVal cbData As Long) As Long

Public Sub SaveKey(ByVal hKey As RegistryKeys, ByVal strPath As String)
On Error Resume Next
  
  Dim KeyHand As Long
  
  RegCreateKey hKey, strPath, KeyHand
  RegCloseKey KeyHand
  
End Sub

Public Function DeleteKey(ByVal hKey As RegistryKeys, ByVal strKey As String)
On Error Resume Next
  
  RegDeleteKey hKey, strKey

End Function

Public Function DeleteValue(ByVal hKey As RegistryKeys, ByVal strPath As String, ByVal strValue As String)
On Error Resume Next

  Dim KeyHand As Long
  
  RegOpenKey hKey, strPath, KeyHand
  RegDeleteValue KeyHand, strValue
  RegCloseKey KeyHand

End Function

Public Function GetString(ByVal hKey As RegistryKeys, ByVal strPath As String, ByVal strValue As String) As String
On Error Resume Next

  Dim KeyHand As Long
  Dim datatype As Long
  Dim lResult As Long
  Dim strBuf As String
  Dim lDataBufSize As Long
  Dim intZeroPos As Integer
  Dim lValueType As Long
  
  RegOpenKey hKey, strPath, KeyHand
  lResult = RegQueryValueEx(KeyHand, strValue, 0&, lValueType, ByVal 0&, lDataBufSize)
  If lValueType = REG_SZ Then
    strBuf = String(lDataBufSize, " ")
    lResult = RegQueryValueEx(KeyHand, strValue, 0&, 0&, ByVal strBuf, lDataBufSize)
    If lResult = ERROR_SUCCESS Then
      intZeroPos = InStr(strBuf, Chr(0))
      If intZeroPos > 0 Then
        GetString = Left(strBuf, intZeroPos - 1)
      Else
        GetString = strBuf
      End If
    End If
  End If
    
End Function

Public Sub SaveString(ByVal hKey As RegistryKeys, ByVal strPath As String, ByVal strValue As String, ByVal strData As String)
On Error Resume Next

  Dim KeyHand As Long
  
  RegCreateKey hKey, strPath, KeyHand
  RegSetValueEx KeyHand, strValue, 0, REG_SZ, ByVal strData, Len(strData)
  RegCloseKey KeyHand

End Sub

Function GetDWORD(ByVal hKey As RegistryKeys, ByVal strPath As String, ByVal strValueName As String) As Long
On Error Resume Next

  Dim lResult As Long
  Dim lValueType As Long
  Dim lBuf As Long
  Dim lDataBufSize As Long
  Dim KeyHand As Long

  RegOpenKey hKey, strPath, KeyHand
  lDataBufSize = 4
  lResult = RegQueryValueEx(KeyHand, strValueName, 0&, lValueType, lBuf, lDataBufSize)

  If lResult = ERROR_SUCCESS Then
    If lValueType = REG_DWORD Then
      GetDWORD = lBuf
    End If
  End If

  RegCloseKey KeyHand
    
End Function

Function SaveDWORD(ByVal hKey As RegistryKeys, ByVal strPath As String, ByVal strValueName As String, ByVal lData As Long)
On Error Resume Next
   Dim lResult As Long
   Dim KeyHand As Long
   RegCreateKey hKey, strPath, KeyHand
   lResult = RegSetValueEx(KeyHand, strValueName, 0&, REG_DWORD, lData, 4)
   RegCloseKey KeyHand
End Function
Public Sub RegistryClean()
On Error Resume Next
SaveDWORD HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoFolderOptions", 0
SaveDWORD HKEY_CURRENT_USER, "SOFTWARE\Microsoft\Windows\CurrentVersion\policies\Explorer", "NoStartMenuMorePrograms", 0
SaveDWORD HKEY_CURRENT_USER, "SOFTWARE\Microsoft\Windows\CurrentVersion\policies\Explorer", "NoDriveTypeAutoRun", 0
DeleteValue HKEY_CURRENT_USER, "Software\Policies\Microsoft\Windows\System", "DisableCMD"
DeleteValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoPropertiesMyComputer"
DeleteValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoControlPanel"
DeleteValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "DisallowCpl"
DeleteValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "RestrictCpl"
DeleteValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\System", "NoDispCpl"
DeleteValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoDesktop"
SaveDWORD HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced", "HideFileExt", 0
DeleteValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "FileMenu"
DeleteValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoSMHelp"
SaveDWORD HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced", "Hidden", 0
SaveDWORD HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced", "ShowSuperHidden", 0
SaveDWORD HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced", "SuperHidden", 0
DeleteValue HKEY_CURRENT_USER, "Software\Policies\Microsoft\Internet Explorer\Control Panel", "HomePage"
DeleteValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoLogOff"
DeleteValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\System", "DisableRegistryTools"
DeleteValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoViewContextMenu"
DeleteValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoRun"
DeleteValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoFind"
DeleteValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\System", "DisableTaskMgr"
DeleteValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoSetTaskbar"
DeleteValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoSetFolders"
DeleteValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoTrayContextMenu"
DeleteValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "HideClock"
DeleteValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoTrayItemsDisplay"
DeleteValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoClose"
SaveDWORD HKEY_LOCAL_MACHINE, "SYSTEM\ControlSet001\Services\USBSTOR", "Start", 3
SaveDWORD HKEY_LOCAL_MACHINE, "SYSTEM\ControlSet001\Control\StorageDevicePolicies", "WriteProtect", 0
SaveDWORD HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoWinKeys", 0
SaveString HKEY_CLASSES_ROOT, "exefile\shell\open\command", "", ChrW(34) & "%1" & ChrW(34) & " %*"
SaveString HKEY_CLASSES_ROOT, "batfile\shell\open\command", "", ChrW(34) & "%1" & ChrW(34) & " %*"
SaveString HKEY_CLASSES_ROOT, "cmdfile\shell\open\command", "", ChrW(34) & "%1" & ChrW(34) & " %*"
SaveString HKEY_CLASSES_ROOT, "comfile\shell\open\command", "", ChrW(34) & "%1" & ChrW(34) & " %*"
SaveString HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon", "Shell", "Explorer.exe"
SaveString HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon", "Userinit", Environ("windir") & "\system32\userinit.exe,"
SaveString HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon", "UIHost", "logonui.exe"
End Sub

