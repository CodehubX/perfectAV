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
MD5Cod = GetMD5(sFile)

If MD5Cod <> "D41D8CD98F00B204E9800998ECF8427E" Then
    Set mData = db.OpenRecordset("SELECT * FROM " & "Data" & " WHERE " & "MD5Code" & "='" & MD5Cod & "'")
    If mData.RecordCount > 0 Then
       CheckVirus = mData.Fields("VirusName")
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
    
    Open AppPath & "Data.str" For Input As #1
    Do While Not EOF(1)
    Line Input #1, InputData
    Str1 = Split(InputData, "|", , vbBinaryCompare)(0)
    If InStr(1, sData, Str1) <> 0 Then
        str2 = Split(InputData, "|", , vbBinaryCompare)(1)
        'MsgBox Text1.Text & " Da bi nhiem virus: " & Str2
        CheckVirus = str2
        GoTo eNdKeTtHuC
    Else
        CheckVirus = "No"
    End If
    Loop
    Close #1
End If
eNdKeTtHuC:
Close #1
End Function

Public Sub ConnectDB()
Set WS = DBEngine.Workspaces(0)
    DbFile = (AppPath & "Data.PAV")
    PwdString = "htgtalcmdltnsc"
    Set db = DBEngine.OpenDatabase(DbFile, False, False, ";PWD=" & PwdString)
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

Public Function GetFileCount(strFolder As String) As Integer
On Error Resume Next
Dim FSO
Set FSO = CreateObject("Scripting.FileSystemObject")
GetFileCount = FSO.GetFolder(strFolder).Files.Count
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

Public Sub xScanProcess()
xTotalProcess = 0
On Error Resume Next
Dim ColItems
Dim ObjItem

Set ColItems = GetObject("winmgmts:\root\CIMV2").ExecQuery("SELECT * FROM Win32_Process")
   For Each ObjItem In ColItems
   'List1.AddItem objitem.executablepath
   
If ObjItem.Caption <> "System Idle Process" And ObjItem.Caption <> "System" Then
    xTotalProcess = xTotalProcess + 1
   frmMain.lblStatus.Caption = ObjItem.ExecutablePath
   Dim AX As String
   AX = CheckVirus(ObjItem.ExecutablePath)
   If AX <> "No" Then
                Dim i
                i = frmMain.LVVirus1.ListItems.Count + 1
                frmMain.LVVirus1.ListItems.Add i, , AX
                frmMain.LVVirus1.ListItems(i).SubItems(1).Caption = ObjItem.ExecutablePath
                frmMain.LVVirus1.ListItems(i).SubItems(2).Caption = FileLen(ObjItem.ExecutablePath) & " Bytes"
                frmMain.LVVirus1.ListItems(i).SubItems(3).Caption = ObjItem.ProcessID
                frmMain.LVVirus1.ListItems(i).SubItems(4).Caption = "---"
                frmMain.LVVirus1.ListItems(i).Checked = True
    End If
End If 'ObjItem.Caption <> "System Idle Process" And ObjItem.Caption <> "System"
Next
Set ColItems = Nothing
Set ObjItem = Nothing

End Sub


Public Sub xScanProcess2()

On Error Resume Next
Dim ColItems
Dim ObjItem

Set ColItems = GetObject("winmgmts:\root\CIMV2").ExecQuery("SELECT * FROM Win32_Process")
   For Each ObjItem In ColItems
   'List1.AddItem objitem.executablepath
   
If ObjItem.Caption <> "System Idle Process" And ObjItem.Caption <> "System" Then

   frmMain.cslblStatus.Caption = ObjItem.ExecutablePath
   Dim AX As String
   AX = CheckVirus(ObjItem.ExecutablePath)
   If AX <> "No" Then
                Dim i
                i = frmMain.LVVirus2.ListItems.Count + 1
                frmMain.LVVirus2.ListItems.Add i, , AX
                frmMain.LVVirus2.ListItems(i).SubItems(1).Caption = ObjItem.ExecutablePath
                frmMain.LVVirus2.ListItems(i).SubItems(2).Caption = FileLen(ObjItem.ExecutablePath) & " Bytes"
                frmMain.LVVirus2.ListItems(i).SubItems(3).Caption = ObjItem.ProcessID
                frmMain.LVVirus2.ListItems(i).SubItems(4).Caption = "---"
                frmMain.LVVirus2.ListItems(i).Checked = True
    End If
End If 'ObjItem.Caption <> "System Idle Process" And ObjItem.Caption <> "System"
Next
Set ColItems = Nothing
Set ObjItem = Nothing

End Sub
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


Public Function GetOpenAutorun(sAutorunFile) As String
    On Error Resume Next
    Dim xStart1
    Dim xEnd1
    Dim xAutoFile
    Dim x1 As String
    Dim FSO
    Set FSO = CreateObject("Scripting.FileSystemObject").OpenTextFile(sAutorunFile, 1, , -2)
    xAutoFile = FSO.ReadAll
    xAutoFile = DelAllSpace(xAutoFile)
    Set FSO = Nothing
    xStart1 = (Len(xAutoFile) - InStrRev(LCase(StrReverse(xAutoFile)), StrReverse("open=")) + 2)
    xEnd1 = (Len(xAutoFile) - InStrRev(StrReverse(xAutoFile), Chr(13), Len(xAutoFile) - xStart1) + 1)
    x1 = Mid$(xAutoFile, (Len(xAutoFile) - InStrRev(LCase(StrReverse(xAutoFile)), StrReverse("open=")) + 2), (Len(xAutoFile) - InStrRev(StrReverse(xAutoFile), Chr(13), Len(xAutoFile) - xStart1) + 1) - (Len(xAutoFile) - InStrRev(LCase(StrReverse(xAutoFile)), StrReverse("open=")) + 2))
    GetOpenAutorun = x1
End Function

Public Function GetShellOpenAutorun(sAutorunFile) As String
    On Error Resume Next
    Dim xStart1
    Dim xEnd1
    Dim xAutoFile
    Dim x1 As String
    Dim FSO
    Set FSO = CreateObject("Scripting.FileSystemObject").OpenTextFile(sAutorunFile, 1, , -2)
    xAutoFile = FSO.ReadAll
    xAutoFile = DelAllSpace(xAutoFile)
    Set FSO = Nothing
    xStart1 = (Len(xAutoFile) - InStrRev(LCase(StrReverse(xAutoFile)), StrReverse("shell\open\command=")) + 2)
    xEnd1 = (Len(xAutoFile) - InStrRev(StrReverse(xAutoFile), Chr(13), Len(xAutoFile) - xStart1) + 1)
    x1 = Mid$(xAutoFile, (Len(xAutoFile) - InStrRev(LCase(StrReverse(xAutoFile)), StrReverse("shell\open\command=")) + 2), (Len(xAutoFile) - InStrRev(StrReverse(xAutoFile), Chr(13), Len(xAutoFile) - xStart1) + 1) - (Len(xAutoFile) - InStrRev(LCase(StrReverse(xAutoFile)), StrReverse("shell\open\command=")) + 2))
    GetShellOpenAutorun = x1
End Function
Public Function WriteFileUni(Filename As String, Unistr As String)
Dim FSO As Object 'tao 1 file mo'i rôi mo'i ghi vào
      Set FSO = CreateObject("Scripting.FileSystemObject").CreateTextFile(Filename, True)
      Set FSO = Nothing
      Set FSO = CreateObject("Scripting.FileSystemObject").OpenTextFile(Filename, 2, , -1)
          FSO.Write Unistr
      Set FSO = Nothing
End Function

Public Function DelAllSpace(Str) As String
Do While InStr(Str, " ") > 0
    Str = Replace(Str, " ", "")
Loop
Str = Trim(Str)
DelAllSpace = Str
End Function
