Attribute VB_Name = "modThaoTac"
Private Declare Function DeleteFile Lib "kernel32" Alias "DeleteFileA" (ByVal lpFileName As String) As Long


Public Function tXoaFile(xFile) As Boolean
On Error Resume Next
If modMain.FileExists(xFile) = True Then
    SetAttr xFile, vbNormal
    DeleteFile xFile
    If modMain.FileExists(xFile) = True Then
        KillProcessById (CheckProcess(xFile))
        Sleep 10
        SetAttr xFile, vbNormal
        DeleteFile xFile
        tXoaFile = Not modMain.FileExists(xFile)
    Else
        tXoaFile = True
    End If
Else
    tXoaFile = False
End If
End Function

