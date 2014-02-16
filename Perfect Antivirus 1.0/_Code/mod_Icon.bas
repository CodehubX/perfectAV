Attribute VB_Name = "mod_Icon"

Option Explicit

Public Declare Function SetPixel _
               Lib "gdi32" (ByVal hDC As Long, _
                            ByVal X As Long, _
                            ByVal Y As Long, _
                            ByVal crColor As Long) As Long

Public Declare Function GetPixel _
               Lib "gdi32" (ByVal hDC As Long, _
                            ByVal X As Long, _
                            ByVal Y As Long) As Long
'Private Declare Function StretchBlt Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long

Public Declare Function SHGetFileInfo _
               Lib "shell32.dll" _
               Alias "SHGetFileInfoA" (ByVal pszPath As String, _
                                       ByVal dwFileAttributes As Long, _
                                       psfi As typSHFILEINFO, _
                                       ByVal cbSizeFileInfo As Long, _
                                       ByVal uFlags As Long) As Long

Public Declare Function ImageList_Draw _
               Lib "comctl32.dll" (ByVal himl&, _
                                   ByVal I&, _
                                   ByVal hDCDest&, _
                                   ByVal X&, _
                                   ByVal Y&, _
                                   ByVal Flags&) As Long
 
Public Type typSHFILEINFO

    hIcon As Long
    iIcon As Long
    dwAttributes As Long
    szDisplayName As String * 260
    szTypeName As String * 80

End Type
 
Public Const SHGFI_DISPLAYNAME = &H200

Public Const SHGFI_EXETYPE = &H2000

Public Const SHGFI_SYSICONINDEX = &H4000

Public Const SHGFI_SHELLICONSIZE = &H4

Public Const SHGFI_TYPENAME = &H400

Public Const SHGFI_LARGEICON = &H0

Public Const SHGFI_SMALLICON = &H1

Public Const ILD_TRANSPARENT = &H1

Public Const Flags = SHGFI_TYPENAME Or SHGFI_SHELLICONSIZE Or SHGFI_SYSICONINDEX Or SHGFI_DISPLAYNAME Or SHGFI_EXETYPE
 
Dim FileInfo As typSHFILEINFO

Dim dXM(3)   As Long, DYM(3) As Long

Dim isStart  As Boolean

Function GetIconFromFile(FileName As String, PictureBox As PictureBox) As Long

        '<EhHeader>
        On Error GoTo GetIconFromFile_Err

        '</EhHeader>
 
        Dim SmallIcon As Long
   
        Dim IconIndex As Integer

        Dim PixelsXY

100     PixelsXY = 32
       
102     SmallIcon = SHGetFileInfo(FileName, 0&, FileInfo, Len(FileInfo), Flags Or SHGFI_LARGEICON)
   
104     If SmallIcon <> 0 Then

106         With PictureBox
108             .Height = 15 * PixelsXY
110             .Width = 15 * PixelsXY
112             .ScaleHeight = 15 * PixelsXY
114             .ScaleWidth = 15 * PixelsXY
116             .Picture = LoadPicture("")
118             .AutoRedraw = True
       
120             SmallIcon = ImageList_Draw(SmallIcon, FileInfo.iIcon, PictureBox.hDC, 0, 0, ILD_TRANSPARENT)
122             .Refresh
            End With
     
        End If

        '<EhFooter>
        Exit Function

GetIconFromFile_Err:
        WriteErr Time & "-" & Date & " - " & Err.Description & " - " & "in PerfectAntivirus2009.mod_Icon.GetIconFromFile " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Function
