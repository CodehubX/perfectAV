Attribute VB_Name = "modCachLy"
Public Sub CachLyFile(xFile)

On Error GoTo KhOnGtAoFoLdErNuA
MkDir AppPath & "CachLy"
MkDir AppPath & "CachLy\File"
MkDir AppPath & "CachLy\KhongTimThay"
KhOnGtAoFoLdErNuA:

If modMain.FileExists(xFile) = True Then
FileCopy xFile, AppPath & "CachLy\File\" & File2Str(xFile)
tXoaFile xFile
    Open AppPath & "CachLy\File\" & File2Str(xFile) For Binary As #1
        Seek #1, 1
        Put #1, , "P"
    Close #1
End If
End Sub

Public Sub PhucHoiCachLy(xFile)
On Error Resume Next
CreateFol
xFile = Path2File(xFile)
'MsgBox xFile
If modMain.FileExists(xFile) = True Then
Dim xOldPath As String
xOldPath = PathFromFile(Str2File(modMain.GetFileName(xFile)))
'MsgBox xOldPath
On Error GoTo ChuyenKhongTimThay
BaTdAuCoPy:
    FileCopy xFile, xOldPath
    tXoaFile xFile
    Open xOldPath For Binary As #1
        Seek #1, 1
        Put #1, , "M"
    Close #1
End If
Exit Sub
ChuyenKhongTimThay:
Dim OkI As String
OkI = modMain.GetFileName(xOldPath)
xOldPath = AppPath & "CachLy\KhongTimThay\" & OkI
GoTo BaTdAuCoPy
End Sub

Public Function File2Str(xString)
On Error Resume Next
File2Str = Replace(Replace(Replace(Date & "," & DoiTime & "," & xString, ":", "&"), "/", "'"), "\", "^")
End Function
Public Function Str2File(xString)
On Error Resume Next
Str2File = Replace(Replace(Replace(Replace(xString, "&", ":"), "'", "/"), "^", "\"), ",", "  **  ")
End Function
Public Function PathFromFile(xString)
On Error Resume Next
PathFromFile = Mid(xString, InStrRev(xString, "  **  ") + Len("  **  "), Len(xString) - InStrRev(xString, "  **  ") + 1)
End Function
Public Function Path2File(xString)
On Error Resume Next
Path2File = Mid(xString, InStrRev(xString, " --- ") + Len(" --- "), Len(xString) - InStrRev(xString, " --- ") + 1)
End Function

'" --- "
Public Sub GetListCachLy()
On Error Resume Next
    MkDir AppPath & "CachLy"
    MkDir AppPath & "CachLy\File"
    MkDir AppPath & "CachLy\KhongTimThay"
'On Error GoTo BoQuA
frmMain.lstCachLy.Clear
frmMain.File1.Path = AppPath & "CachLy\File\"
frmMain.File1.Refresh
Dim Au As Integer
For Au = 0 To frmMain.File1.ListCount - 1
    frmMain.lstCachLy.AddItem Str2File(modMain.GetFileName(frmMain.File1.List(Au))) & " --- " & AppPath & "CachLy\File\" & frmMain.File1.List(Au)
Next Au
'BoQuA:
End Sub
Public Function GetFileNhatKy(xStr)
On Error Resume Next
GetFileNhatKy = Replace(Replace(Replace(Replace(xStr, ":", "&"), "  **  ", ","), "\", "^"), "/", "'")
End Function
Public Sub GetListNhatKy()
'9&22&28 PM,8'22'2009,  F&^
'9:22:28 PM  **  8/22/2009  **   F:\
On Error Resume Next
    MkDir AppPath & "NhatKy"

'On Error GoTo BoQuA
frmMain.lstNhatKy.Clear
frmMain.File1.Path = AppPath & "NhatKy\"
frmMain.File1.Refresh
Dim AuX As Integer
For AuX = 0 To frmMain.File1.ListCount - 1
    frmMain.lstNhatKy.AddItem Str2File(frmMain.File1.List(AuX))
Next AuX
'BoQuA:
End Sub
Public Function DoiTime()
On Error Resume Next
DoiTime = IIf(Len(Hour(Time)) = 1, "0" & Hour(Time), Hour(Time)) & ":" & IIf(Len(Minute(Time)) = 1, "0" & Minute(Time), Minute(Time)) & ":" & IIf(Len(Second(Time)) = 1, "0" & Second(Time), Second(Time)) & " " & Right(Time, 2)
End Function

Public Sub GetListData()
On Error Resume Next
    MkDir AppPath & "UserData"

frmMain.lstUserData.Clear
frmMain.File1.Path = AppPath & "UserData\"
frmMain.File1.Refresh
Dim AuX As Integer
For AuX = 0 To frmMain.File1.ListCount - 1
    frmMain.lstUserData.AddItem Str2File(frmMain.File1.List(AuX)) ' & " - " & GetMD5(AppPath & "UserData\" & Str2File(frmMain.File1.List(AuX)))
Next AuX

End Sub
Public Sub CreateFolderUserData()
On Error Resume Next
    MkDir AppPath & "UserData"
End Sub

Public Sub CreateFol()
On Error Resume Next
MkDir AppPath & "CachLy"
MkDir AppPath & "CachLy\File"
MkDir AppPath & "CachLy\KhongTimThay"
End Sub
Public Sub CreateFolDd()
On Error Resume Next
MkDir AppPath & "TinTuong\"
End Sub

Public Sub GetListTinTuong()
On Error Resume Next
CreateFolDd
frmMain.lstTinTuong.Clear
frmMain.File1.Path = AppPath & "TinTuong\"
frmMain.File1.Refresh
Dim AuX As Integer
For AuX = 0 To frmMain.File1.ListCount - 1
    frmMain.lstTinTuong.AddItem ReadFileUni(AppPath & "TinTuong\" & frmMain.File1.List(AuX))
Next AuX

End Sub
