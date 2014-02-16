Attribute VB_Name = "modScanVirus"
Dim db As Database
Dim rs As Recordset
Dim WS As Workspace
Dim mData As Recordset
Dim Rdb As Database
Dim RWS As Workspace
Dim RData As Recordset
Public xTotalProcess

Public Declare Function DeleteFile Lib "kernel32" Alias "DeleteFileA" (ByVal lpFileName As String) As Long

Public Function CheckVirus(xFile) As String
On Error GoTo ThoatRaViError
If FileLen(xFile) > 5000000 Then GoTo ThoatRaViError
Dim MD5Cod
Dim InputData As String
Dim sData As String
Dim Str1
Dim str2
Dim sFile As String
sFile = xFile
'Check Virus by MD5 code
'MD5Cod = HashFile(sFile, MD5)
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
On Error GoTo ThoatRaViError
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
DoEvents
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

Exit Function
ThoatRaViError:
CheckVirus = "No"
End Function

Public Sub ConnectDB()
On Error GoTo Connect_ERr
    Set WS = DBEngine.Workspaces(0)
    DbFile = (AppPath & "Data.PAV")
    PwdString = "htgtalcmdltnsc"
    Set db = DBEngine.OpenDatabase(DbFile, False, False, ";PWD=" & PwdString)
Exit Sub
Connect_ERr:
UniMsgBox "Kho6ng the63 ke61t no61i d9e61n CSDL!", vbOKOnly, "Error!"
End
End Sub
Public Sub RTPConnectDB()
On Error GoTo Connect_ERr
    Set RWS = DBEngine.Workspaces(0)
    Set Rdb = DBEngine.OpenDatabase(AppPath & "Data.PAV", False, False, ";PWD=" & "htgtalcmdltnsc")
Exit Sub
Connect_ERr:
Unload frmRTP
End Sub
Public Function GetFileName(ByVal sPath As String) As String
GetFileName = Mid(sPath, InStrRev(sPath, "\") + 1)
End Function

Public Function GetFolderPath(ByVal sPath As String) As String
GetFolderPath = Left(sPath, InStrRev(sPath, "\") - 1)
End Function

Public Function GetFolderCha(ByVal sPath As String) As String
On Error Resume Next
GetFolderCha = Mid(sPath, (InStrRev(sPath, "\", InStrRev(sPath, "\") - 1)) + 1, ((InStrRev(sPath, "\") - 1) - InStrRev(sPath, "\", InStrRev(sPath, "\") - 1)))
End Function




Public Function FileExists(sFile) As Boolean
On Error Resume Next
FileExists = ((GetAttr(sFile) And vbDirectory) = 0)
End Function


Public Function DelAllSpace(Str) As String
On Error Resume Next
Do While InStr(Str, " ") > 0
    Str = Replace(Str, " ", "")
Loop
Str = Trim(Str)
DelAllSpace = Str
End Function

Public Function AddVirus(xVirusName, xMD5Code) As Boolean
On Error Resume Next
Dim Xdb As Database
Dim XWS As Workspace
Dim XData As Recordset
Dim PwdString2 As String
    Set XWS = DBEngine.Workspaces(0)
    PwdString2 = "htgtalcmdltnsc"
    Set Xdb = DBEngine.OpenDatabase(AppPath & "Data.PAV", False, False, ";PWD=" & PwdString2)


If xMD5Code <> "D41D8CD98F00B204E9800998ECF8427E" Then
    Set XData = Xdb.OpenRecordset("SELECT * FROM " & "Data" & " WHERE " & "MD5Code" & "='" & xMD5Code & "'")

    If XData.RecordCount > 0 Then
        AddVirus = False
    Else
        AddVirus = True
        XData.AddNew
        XData.Fields("VirusName") = xVirusName
        XData.Fields("MD5Code") = xMD5Code
        XData.Update
    End If
Else
    AddVirus = False
End If
Set Xdb = Nothing
Set XWS = Nothing
Set XData = Nothing

End Function
Public Function DeleteVirus(xMD5Code) As Boolean
Dim Ydb As Database
Dim YWS As Workspace
Dim YData As Recordset
Dim PwdString3 As String

    Set YWS = DBEngine.Workspaces(0)
    PwdString3 = "htgtalcmdltnsc"
    Set Ydb = DBEngine.OpenDatabase(AppPath & "Data.PAV", False, False, ";PWD=" & PwdString3)


If xMD5Code <> "D41D8CD98F00B204E9800998ECF8427E" Then
    Set YData = Ydb.OpenRecordset("SELECT * FROM " & "Data" & " WHERE " & "MD5Code" & "='" & xMD5Code & "'")
    
    If YData.RecordCount > 0 Then
        DeleteVirus = True
        YData.Delete
    Else
        DeleteVirus = False
    End If
Else
    DeleteVirus = False
End If
Set Ydb = Nothing
Set YWS = Nothing
Set YData = Nothing

End Function




Public Function RTPCheckVirus(xFile) As String

'On Error GoTo ThoatRaViError
If FileLen(xFile) > 1000000 Then GoTo ThoatRaViError
Dim MD5Cod
Dim InputData As String
Dim sData As String
Dim Str1
Dim str2
Dim sFile As String
sFile = xFile
'Check Virus by MD5 code
'MD5Cod = HashFile(sFile, MD5)
MD5Cod = GetMD5(sFile)
If UCase(MD5Cod) <> "D41D8CD98F00B204E9800998ECF8427E" Then
    Set RData = Rdb.OpenRecordset("SELECT * FROM " & "Data" & " WHERE " & "MD5Code" & "='" & UCase(MD5Cod) & "'")
    If RData.RecordCount > 0 Then
       If IsNull(RData.Fields("VirusName")) = False Then RTPCheckVirus = RData.Fields("VirusName")
       GoTo eNdKeTtHuC
    Else
        RTPCheckVirus = "No"
    End If
Else
    RTPCheckVirus = "No"
End If
'If MsgBox(MD5Cod & " - " & CheckVirus, vbYesNo) = vbYes Then End

'Check Virus By String
On Error GoTo ThoatRaViError
If FileExists(sFile) = True Then
    Open sFile For Binary As #1
        sData = Space(LOF(1))
        Get #1, , sData
    Close #1
    Set RData = Rdb.OpenRecordset("STRI", dbOpenTable)
    RData.MoveFirst
    Dim ix
    ix = 1
BaTdAuTiMkIeM:
DoEvents
    Dim strSTR As String
    strSTR = RData.Fields("STRING")
    'List1.AddItem strSTR
    If InStr(1, sData, strSTR) <> 0 Then
        RTPCheckVirus = RData.Fields("VirusName")
    Else
        RData.MoveNext
        ix = ix + 1
        If ix > RData.RecordCount Then
            RTPCheckVirus = "No"
        Else
            GoTo BaTdAuTiMkIeM
        End If
    End If
End If
eNdKeTtHuC:
Close #1

Exit Function
ThoatRaViError:
RTPCheckVirus = "No"
End Function




