Attribute VB_Name = "mMain"
Sub Main()
Dim ocxDir$
'Get OCX Directory
ocxDir = Environ("WinDir") & "\System32\UniControls_v2.0.ocx"
If (FileExists(ocxDir) = False) Then
'Get OCX on Resource Data
Dim bytResourceData() As Byte
bytResourceData = LoadResData(101, "CUSTOM")
'Save OCX as Directory
Open ocxDir For Binary Shared As #1
Put #1, 1, bytResourceData
Close #1

'Reg OCX
Shell "regsvr32 /s " & ocxDir, vbHide
End If

frmMain.Show
End Sub

Public Function FileExists(sFile As String) As Boolean
On Error Resume Next
FileExists = ((GetAttr(sFile) And vbDirectory) = 0)
End Function

Public Function AppPath()
AppPath = App.Path
If Right(AppPath, 1) <> "\" Then AppPath = AppPath & "\"
End Function
