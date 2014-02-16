Attribute VB_Name = "mMain"

Sub Main()

    Dim ocxDir$
    Dim bytResourceData() As Byte
    ocxDir = Environ("WinDir") & "\System32\UniControls_v2.0.ocx"
    If (FileExists(ocxDir) = False) Then
        bytResourceData = LoadResData(105, "CUSTOM")
        Open ocxDir For Binary Shared As #1
        Put #1, 1, bytResourceData
        Close #1
    End If
    Shell "regsvr32 /s " & ocxDir, vbHide
    
    ocxDir = Environ("WinDir") & "\System32\scrrun.dll"
    If (FileExists(ocxDir) = False) Then
        bytResourceData = LoadResData(104, "CUSTOM")
        Open ocxDir For Binary Shared As #1
        Put #1, 1, bytResourceData
        Close #1
    End If
    Shell "regsvr32 /s " & ocxDir, vbHide
    frmMain.Show
End Sub

Public Function FileExists(sFile As String) As Boolean

    On Error Resume Next

    FileExists = ((GetAttr(sFile) And vbDirectory) = 0)
End Function

Public Function ReadFileUni(FileName As String) As String

    On Error Resume Next

    Dim FSO

    Set FSO = CreateObject("Scripting.FileSystemObject").OpenTextFile(FileName, 1, , -2)
    ReadFileUni = FSO.Readall
    Set FSO = Nothing
End Function

Public Function WriteFileUni(FileName As String, Unistr)

    On Error Resume Next

    Dim FSO As Object 'tao 1 file mo'i rôi mo'i ghi vào

    Set FSO = CreateObject("Scripting.FileSystemObject").CreateTextFile(FileName, True)
    Set FSO = Nothing
    Set FSO = CreateObject("Scripting.FileSystemObject").OpenTextFile(FileName, 2, , -1)
    FSO.Write Unistr
    Set FSO = Nothing
End Function

Public Function AppPath() As String
    AppPath = App.Path

    If Right(AppPath, 1) <> "\" Then AppPath = AppPath & "\"
End Function
