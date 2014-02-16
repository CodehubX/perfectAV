Attribute VB_Name = "basFile"
Option Explicit

'##############################################################################################
'Purpose: Used for File System operations
'Author:  Richard Mewett ©2004

'Credits:
'The GetFolder() code was sourced from VB.NET (Brad Martinez & Randy Birch)
'##############################################################################################
Public Const INVALID_HANDLE_VALUE = -1
Public Const FILE_ATTRIBUTE_ARCHIVE = &H20
Public Const FILE_ATTRIBUTE_DIRECTORY = &H10
Public Const FILE_ATTRIBUTE_HIDDEN = &H2
Public Const FILE_ATTRIBUTE_NORMAL = &H80
Public Const FILE_ATTRIBUTE_READONLY = &H1
Public Const FILE_ATTRIBUTE_SYSTEM = &H4
Public Const FILE_ATTRIBUTE_TEMPORARY = &H100
Public Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long

Public Const BIF_RETURNONLYFSDIRS = &H1
Public Const MAX_PATH = 260
Public Type FILETIME
    dwLowDateTime     As Long
    dwHighDateTime    As Long
End Type
Public Type WIN32_FIND_DATA
    dwFileAttributes  As Long
    ftCreationTime    As FILETIME
    ftLastAccessTime  As FILETIME
    ftLastWriteTime   As FILETIME
    nFileSizeHigh     As Long
    nFileSizeLow      As Long
    dwReserved0       As Long
    dwReserved1       As Long
    cFileName         As String * MAX_PATH
    cAlternate        As String * 14
End Type
Public Type SHELLEXECUTEINFO
    cbSize As Long
    fMask As Long
    hwnd As Long
    lpVerb As String
    lpFile As String
    lpParameters As String
    lpDirectory As String
    nShow As Long
    hInstApp As Long
    lpIDList As Long
    lpClass As String
    hkeyClass As Long
    dwHotKey As Long
    hIcon As Long
    hProcess As Long
End Type
Public Const SEE_MASK_INVOKEIDLIST = &HC
Public Const SEE_MASK_NOCLOSEPROCESS = &H40
Public Const SEE_MASK_FLAG_NO_UI = &H400
Declare Function ShellExecuteEX Lib "shell32.dll" Alias "ShellExecuteEx" (SEI As SHELLEXECUTEINFO) As Long

'Get icon


Private Declare Function SHGetFileInfo Lib "shell32.dll" Alias "SHGetFileInfoA" (ByVal pszPath As String, ByVal dwFileAttributes As Long, psfi As SHFILEINFO, ByVal cbSizeFileInfo As Long, ByVal uFlags As Long) As Long
Private Declare Function ImageList_Draw Lib "comctl32.dll" (ByVal himl&, ByVal i&, ByVal hdcDest&, ByVal x&, ByVal Y&, ByVal flags&) As Long

Private Type SHFILEINFO
    hIcon As Long
    iIcon As Long
    dwAttributes As Long
    szDisplayName As String * 260
    szTypeName As String * 80
End Type

Private Const SHGFI_DISPLAYNAME = &H200
Private Const SHGFI_EXETYPE = &H2000
Private Const SHGFI_SYSICONINDEX = &H4000
Private Const SHGFI_LARGEICON = &H0
Private Const SHGFI_SMALLICON = &H1
Private Const ILD_TRANSPARENT = &H1
Private Const SHGFI_SHELLICONSIZE = &H4
Private Const SHGFI_TYPENAME = &H400
Private Const BASIC_SHGFI_FLAGS = SHGFI_TYPENAME _
Or SHGFI_SHELLICONSIZE Or SHGFI_SYSICONINDEX _
Or SHGFI_DISPLAYNAME Or SHGFI_EXETYPE

Private SIconInfo As SHFILEINFO

Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Public Declare Function GetFileAttributes Lib "kernel32" Alias "GetFileAttributesA" (ByVal lpFileName As String) As Long

Function GetAttribute(ByVal sFilePath As String) As String
    Select Case GetFileAttributes(sFilePath)
        Case 1: GetAttribute = "Read Only"
        Case 2: GetAttribute = "Hidden"
        Case 3: GetAttribute = "Read Only + Hidden"
        Case 4: GetAttribute = "System"
        Case 5: GetAttribute = "Read Only + System"
        Case 6: GetAttribute = "Hidden + System"
        Case 7: GetAttribute = "Read Only + Hidden + System"
        '-------------------------------------------------'
        Case 32: GetAttribute = "Archive"
        Case 33: GetAttribute = "Read Only + Archive"
        Case 34: GetAttribute = "Hidden + Archive"
        Case 35: GetAttribute = "Read Only + Hidden + Archive"
        Case 36: GetAttribute = "System + Archive"
        Case 37: GetAttribute = "Read Only + System + Archive"
        Case 38: GetAttribute = "HSA"
        Case 39: GetAttribute = "Read Only + Hidden + System + Archive"
        '-------------------------------------------------'
        Case 128: GetAttribute = "Normal"
        '-------------------------------------------------'
        Case Else: GetAttribute = "N/A"
    End Select
End Function

'Dimensionalize SIconInfo as SHFILEINFO type structure
Public Sub GetIcon(icPath$, pDisp As PictureBox)
pDisp.Cls
Dim hImgSmall&: hImgSmall = SHGetFileInfo(icPath$, 0&, SIconInfo, Len(SIconInfo), BASIC_SHGFI_FLAGS Or SHGFI_SMALLICON)
'call SHGetFileInfo to return a handle to the icon associated with the specified file
 ImageList_Draw hImgSmall, SIconInfo.iIcon, pDisp.hDC, 0, 0, ILD_TRANSPARENT
 'Draw the icon to the specified picturebox control
End Sub
Public Sub GetLargeIcon(icPath$, pDisp As PictureBox)
Dim hImgLrg&: hImgLrg = SHGetFileInfo(icPath$, 0&, SIconInfo, Len(SIconInfo), BASIC_SHGFI_FLAGS Or SHGFI_LARGEICON)
'call SHGetFileInfo to return a handle to the icon associated with the specified file
 ImageList_Draw hImgLrg, SIconInfo.iIcon, pDisp.hDC, 0, 0, ILD_TRANSPARENT
 'Draw the icon to the specified picturebox control
End Sub
Public Sub ShowProperties(sFileName As String, hwndOwner As Long)
    '##############################################################################################
    'Displays the Properties of the specified file
    '##############################################################################################
    
    Dim SEI As SHELLEXECUTEINFO
    
    With SEI
        .cbSize = Len(SEI)
        .fMask = SEE_MASK_NOCLOSEPROCESS Or SEE_MASK_INVOKEIDLIST Or SEE_MASK_FLAG_NO_UI
        .hwnd = hwndOwner
        .lpVerb = "properties"
        .lpFile = sFileName
        .lpParameters = vbNullChar
        .lpDirectory = vbNullChar
        .nShow = 0
        .hInstApp = 0
        .lpIDList = 0
    End With
    
    Call ShellExecuteEX(SEI)
End Sub


Public Function FolderExists(sFolder) As Boolean
    Dim WFD As WIN32_FIND_DATA
    Dim lResult As Long
    
    lResult = FindFirstFile(sFolder, WFD)
    If lResult <> INVALID_HANDLE_VALUE Then
        If (WFD.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) = FILE_ATTRIBUTE_DIRECTORY Then
            FolderExists = True
        Else
            FolderExists = False
        End If
    End If
End Function

Public Function FileExists(sFile) As Boolean
On Error Resume Next
FileExists = ((GetAttr(sFile) And vbDirectory) = 0)
End Function


Public Function GetExt(FilePathName As String) As String
On Error Resume Next
    GetExt = Right(FilePathName, InStr(1, StrReverse(FilePathName), ".", vbBinaryCompare) - 1)
End Function
Public Function GetFileName(ByVal sPath As String) As String
On Error Resume Next
GetFileName = Mid(sPath, InStrRev(sPath, "\") + 1)
End Function
Public Function GetFolderPath(ByVal sPath As String) As String
On Error Resume Next
GetFolderPath = Left(sPath, InStrRev(sPath, "\") - 1)
End Function


