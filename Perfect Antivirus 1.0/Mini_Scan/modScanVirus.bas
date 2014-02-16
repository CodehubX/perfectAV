Attribute VB_Name = "modScanVirus"
Dim db As Database
Dim rs As Recordset
Dim WS As Workspace
Dim Max As Long
Dim mData As Recordset
Public xTotalProcess

Public Declare Function DeleteFile Lib "kernel32" Alias "DeleteFileA" (ByVal lpFileName As String) As Long

Public Function CheckVirus(xFile) As String
DoEvents
Dim MD5Cod
Dim InputData As String
Dim sData As String
Dim Str1
Dim str2
Dim sFile As String
sFile = xFile

'Check Virus by MD5 code
'MD5Cod = HashFile(sFile, MD5)
DoEvents
MD5Cod = GetMD5(sFile)

If UCase(MD5Cod) <> "D41D8CD98F00B204E9800998ECF8427E" Then
    Set mData = db.OpenRecordset("SELECT * FROM " & "Data" & " WHERE " & "MD5Code" & "='" & UCase(MD5Cod) & "'")
    If mData.RecordCount > 0 Then
       If IsNull(mData.Fields("VirusName")) = False Then CheckVirus = mData.Fields("VirusName")
       GoTo eNdKeTtHuC
    Else
        CheckVirus = "No"
    End If

Else
    CheckVirus = "No"
End If
'If MsgBox(MD5Cod & " - " & CheckVirus, vbYesNo) = vbYes Then End

'Check Virus By String
If FileExists(sFile) = True Then
    Open sFile For Binary As #1
        sData = Space(LOF(1))
        Get #1, , sData
    Close #1
    Set mData = db.OpenRecordset("STRI", dbOpenTable)
    mData.MoveFirst
    Dim ix
    ix = 1
BaTdAuTiMkIeM:
    
    Dim strSTR As String
    strSTR = mData.Fields("STRING")
    'List1.AddItem strSTR
    If InStr(1, sData, strSTR) <> 0 Then
        CheckVirus = mData.Fields("VirusName")
    Else
        mData.MoveNext
        ix = ix + 1
        If ix > mData.RecordCount Then
            CheckVirus = "No"
        Else
            GoTo BaTdAuTiMkIeM
        End If
    End If

End If
eNdKeTtHuC:
Close #1
End Function

Public Sub ConnectDB()
If FileExists(AppPath & "Data.PAV") = True Then

Set WS = DBEngine.Workspaces(0)
    DbFile = (AppPath & "Data.PAV")
    PwdString = "htgtalcmdltnsc"
    Set db = DBEngine.OpenDatabase(DbFile, False, False, ";PWD=" & PwdString)
Else
    UniMsgBox "Kho6ng ti2m tha61y du74 lie65u cu3a chu7o7ng tri2nh", vbOKOnly, "Tho6ng ba1o"
    End
End If
End Sub
Public Function GetFileName(ByVal sPath As String) As String
GetFileName = Mid(sPath, InStrRev(sPath, "\") + 1)
End Function

Public Function GetFolderPath(ByVal sPath As String) As String
GetFolderPath = Left(sPath, InStrRev(sPath, "\") - 1)
End Function


Public Function AppPath() As String
Dim x As String
x = App.Path
If Right(x, 1) <> "\" Then x = x & "\"
AppPath = x
End Function

Public Function FileExists(sFile) As Boolean
On Error Resume Next
FileExists = ((GetAttr(sFile) And vbDirectory) = 0)
End Function



Public Sub DelAllLV(LV As UniListView)
Dim k
For k = 1 To LV.ListItems.Count
If k > LV.ListItems.Count Then Exit Sub
LV.ListItems.Remove k
k = k - 1
Next k
End Sub


Public Sub DelAllChecked(LV As UniListView)
Dim k
For k = 1 To LV.ListItems.Count
If k > LV.ListItems.Count Then Exit Sub
If LV.ListItems(k).Checked = True Then
LV.ListItems.Remove k
k = k - 1
End If
Next k
End Sub

Public Function DelAllSpace(Str) As String
Do While InStr(Str, " ") > 0
    Str = Replace(Str, " ", "")
Loop
Str = Trim(Str)
DelAllSpace = Str
End Function
