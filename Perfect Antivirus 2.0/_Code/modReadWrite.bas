Attribute VB_Name = "modReadWrite"
Public Function ReadFileUni(FileName)
On Error Resume Next
Dim FSO
   Set FSO = CreateObject("Scripting.FileSystemObject").OpenTextFile(FileName, 1, , -2)
   ReadFileUni = FSO.ReadAll
   Set FSO = Nothing
End Function

Public Function WriteFileUni(FileName, Unistr)
On Error Resume Next
Dim FSO As Object 'tao 1 file mo'i rôi mo'i ghi vào
      Set FSO = CreateObject("Scripting.FileSystemObject").CreateTextFile(FileName, True)
      Set FSO = Nothing
      Set FSO = CreateObject("Scripting.FileSystemObject").OpenTextFile(FileName, 2, , -1)
          FSO.Write Unistr
      Set FSO = Nothing
End Function
