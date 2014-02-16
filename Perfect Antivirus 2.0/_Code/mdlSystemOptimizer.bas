Attribute VB_Name = "mdlSystemOptimizer"
Public Enum SpecialFolder
    CSIDL_RECENT = &H8
    CSIDL_PROFILER = &H28
    CSIDL_HISTORY = &H22
End Enum
Private Type SHITEMID
    cb As Long
    abID As Byte
End Type

Private Type ITEMIDLIST
    mkid As SHITEMID
End Type
Private Declare Function SHGetSpecialFolderLocation Lib _
    "shell32.dll" (ByVal hWndOwner As Long, _
    ByVal nFolder As Long, _
    pidl As ITEMIDLIST) As Long
Private Declare Function GetSystemDirectory Lib _
    "kernel32.dll" Alias "GetSystemDirectoryA" ( _
    ByVal lpBuffer As String, _
    ByVal nSize As Long) As Long
Private Declare Function ShellExecute Lib _
    "shell32.dll" Alias "ShellExecuteA" ( _
    ByVal hwnd As Long, _
    ByVal lpOperation As String, _
    ByVal lpFile As String, _
    ByVal lpParameters As String, _
    ByVal lpDirectory As String, _
    ByVal nShowCmd As Long) As Long

Private Declare Function GetWindowsDirectory Lib _
    "kernel32.dll" Alias "GetWindowsDirectoryA" ( _
    ByVal lpBuffer As String, _
    ByVal nSize As Long) As Long
Private Declare Function SHEmptyRecycleBin Lib "shell32.dll" Alias "SHEmptyRecycleBinA" (ByVal hwnd As Long, ByVal pszRootPath As String, ByVal dwFlags As Long) As Long

Private Const SHERB_NORMAL = &H0 'Normal execution

Private Const SHERB_NOCONFIRMATION = &H1 'execute without confirmation

Private Const SHERB_NOPROGRESSUI = &H2 'execute without progress window

Private Const SHERB_NOSOUND = &H4 'execute without sound

Private Const SHERB_NOALL = (SHERB_NOCONFIRMATION And SHERB_NOPROGRESSUI And SHERB_NOSOUND)
Dim RetVal As Long

Public Sub EmpRecBin()
    RetVal = SHEmptyRecycleBin(0&, vbNullString, SHERB_NORMAL)
End Sub



Public Sub ClearJunkFile()
    On Error Resume Next
    Kill GetWindowsPath & "Prefetch\*.*"
    Kill GetWindowsPath & "Temp\*.*"
    Kill GetSpecialFolder(CSIDL_RECENT) & "\*.*"
    Kill GetSpecialFolder(CSIDL_HISTORY) & "\*.*"
    Kill GetSpecialFolder(CSIDL_PROFILER) & "\Cookies\*.*"
    Kill GetSpecialFolder(CSIDL_PROFILER) & "\Local Settings\Temp\*.*"
    Kill GetSpecialFolder(CSIDL_PROFILER) & _
        "\Local Settings\Temporary Internet Files\*.*"
End Sub
Public Function GetSpecialFolder(FolderType As SpecialFolder) As String
    Dim r As Long, sPath As String
    Dim IDL As ITEMIDLIST
    r = SHGetSpecialFolderLocation(100, FolderType, IDL)
    sPath = Space$(512)
    r = SHGetPathFromIDList(ByVal IDL.mkid.cb, ByVal sPath)
    GetSpecialFolder = Left$(sPath, InStr(1, sPath, Chr$(0)) - 1)
End Function

Public Function GetWindowsPath() As String
    Dim lpBuffer As String * 255
    Dim nSize As Long
    nSize = GetWindowsDirectory(lpBuffer, 255)
    GetWindowsPath = Left(lpBuffer, nSize) & "\"
End Function

Public Function GetSystem32Path() As String
    Dim lpBuffer As String * 255
    Dim nSize As Long
    nSize = GetSystemDirectory(lpBuffer, 255)
    GetSystem32Path = Left(lpBuffer, nSize) & "\"
End Function

