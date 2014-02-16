Attribute VB_Name = "basTest"
Public Function ReadFile(FileName As String) As String
On Error Resume Next
Dim FSO
   Set FSO = CreateObject("Scripting.FileSystemObject").OpenTextFile(FileName, 1, , -2)
   ReadFile = FSO.Readall
   Set FSO = Nothing
End Function
Public Function WriteFile(FileName As String, Unistr As String)
On Error Resume Next
Dim FSO As Object 'tao 1 file mo'i rôi mo'i ghi vào
      Set FSO = CreateObject("Scripting.FileSystemObject").CreateTextFile(FileName, True)
      Set FSO = Nothing
      Set FSO = CreateObject("Scripting.FileSystemObject").OpenTextFile(FileName, 2, , -1)
          FSO.Write Unistr
      Set FSO = Nothing
End Function

Public Function AppPath() As String
Dim x
x = App.Path
If Right(x, 1) <> "\" Then x = x & "\"
AppPath = x
End Function
