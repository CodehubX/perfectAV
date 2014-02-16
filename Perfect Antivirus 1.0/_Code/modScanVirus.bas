Attribute VB_Name = "modScanVirus"

Dim db    As Database

Dim rs    As Recordset

Dim WS    As Workspace

Dim mData As Recordset

Public xTotalProcess

Public Declare Function DeleteFile _
               Lib "kernel32" _
               Alias "DeleteFileA" (ByVal lpFileName As String) As Long

Public Function CheckVirus(xFile) As String

        '<EhHeader>
        On Error GoTo CheckVirus_Err

        '</EhHeader>
100     DoEvents

        Dim MD5Cod

        Dim InputData As String

        Dim sData     As String

        Dim Str1

        Dim str2

        Dim sFile As String

102     If IsNull(xFile) = False Then sFile = xFile

        'Check Virus by MD5 code
        'MD5Cod = HashFile(sFile, MD5)
104     DoEvents
106     MD5Cod = GetMD5(sFile)

108     If UCase(MD5Cod) <> "D41D8CD98F00B204E9800998ECF8427E" Then
110         Set mData = db.OpenRecordset("SELECT * FROM " & "Data" & " WHERE " & "MD5Code" & "='" & UCase(MD5Cod) & "'")

112         If mData.RecordCount > 0 Then
114             If IsNull(mData.Fields("VirusName")) = False Then CheckVirus = mData.Fields("VirusName")
116             GoTo eNdKeTtHuC
            Else
118             CheckVirus = "No"
            End If

        Else
120         CheckVirus = "No"
        End If

        'If MsgBox(MD5Cod & " - " & CheckVirus, vbYesNo) = vbYes Then End

        'Check Virus By String
122     If FileExists(sFile) = True Then
124         Open sFile For Binary As #1
126         sData = Space(LOF(1))
128         Get #1, , sData
130         Close #1
132         Set mData = db.OpenRecordset("STRI", dbOpenTable)
134         mData.MoveFirst

            Dim ix

136         ix = 1
BaTdAuTiMkIeM:
    
            Dim strSTR As String

138         strSTR = mData.Fields("STRING")

            'List1.AddItem strSTR
140         If InStr(1, sData, strSTR) <> 0 Then
142             CheckVirus = mData.Fields("VirusName")
            Else
144             mData.MoveNext
146             ix = ix + 1

148             If ix > mData.RecordCount Then
150                 CheckVirus = "No"
                Else
152                 GoTo BaTdAuTiMkIeM
                End If
            End If

        End If

eNdKeTtHuC:

154     Close #1

        '<EhFooter>
        Exit Function

CheckVirus_Err:
        WriteErr Time & "-" & Date & " - " & Err.Description & " - " & "in PerfectAntivirus2009.modScanVirus.CheckVirus " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Function

Public Sub ConnectDB()

        '<EhHeader>
        On Error GoTo ConnectDB_Err

        '</EhHeader>
        On Error GoTo Connect_ERr

100     Set WS = DBEngine.Workspaces(0)
102     DbFile = (AppPath & "Data.PAV")
104     PwdString = "htgtalcmdltnsc"
106     Set db = DBEngine.OpenDatabase(DbFile, False, False, ";PWD=" & PwdString)

        Exit Sub

Connect_ERr:
108     UniMsgBox "Kho6ng the63 ke61t no61i d9e61n CSDL!", vbOKOnly, "Error!"

110     End

        '<EhFooter>
        Exit Sub

ConnectDB_Err:
        WriteErr Time & "-" & Date & " - " & Err.Description & " - " & "in PerfectAntivirus2009.modScanVirus.ConnectDB " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Public Function GetFileName(ByVal sPath As String) As String

        '<EhHeader>
        On Error GoTo GetFileName_Err

        '</EhHeader>

100     GetFileName = Mid(sPath, InStrRev(sPath, "\") + 1)

        '<EhFooter>
        Exit Function

GetFileName_Err:
        WriteErr Time & "-" & Date & " - " & Err.Description & " - " & "in PerfectAntivirus2009.modScanVirus.GetFileName " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Function

Public Function GetFolderPath(ByVal sPath As String) As String

        '<EhHeader>
        On Error GoTo GetFolderPath_Err

        '</EhHeader>

100     GetFolderPath = Left(sPath, InStrRev(sPath, "\") - 1)

        '<EhFooter>
        Exit Function

GetFolderPath_Err:
        WriteErr Time & "-" & Date & " - " & Err.Description & " - " & "in PerfectAntivirus2009.modScanVirus.GetFolderPath " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Function

Public Function GetFolderCha(ByVal sPath As String) As String

        '<EhHeader>
        On Error GoTo GetFolderCha_Err

        '</EhHeader>
        On Error Resume Next

100     GetFolderCha = Mid(sPath, (InStrRev(sPath, "\", InStrRev(sPath, "\") - 1)) + 1, ((InStrRev(sPath, "\") - 1) - InStrRev(sPath, "\", InStrRev(sPath, "\") - 1)))

        '<EhFooter>
        Exit Function

GetFolderCha_Err:
        WriteErr Time & "-" & Date & " - " & Err.Description & " - " & "in PerfectAntivirus2009.modScanVirus.GetFolderCha " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Function

Public Function GetFileCount(strFolder As String) As Integer

        '<EhHeader>
        On Error GoTo GetFileCount_Err

        '</EhHeader>
        On Error Resume Next

        Dim FSO

100     Set FSO = CreateObject("Scripting.FileSystemObject")
102     GetFileCount = FSO.GetFolder(strFolder).Files.Count

        '<EhFooter>
        Exit Function

GetFileCount_Err:
        WriteErr Time & "-" & Date & " - " & Err.Description & " - " & "in PerfectAntivirus2009.modScanVirus.GetFileCount " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Function

Public Function AppPath() As String

        '<EhHeader>
        On Error GoTo AppPath_Err

        '</EhHeader>
        Dim X As String

100     X = App.Path

102     If Right(X, 1) <> "\" Then X = X & "\"
104     AppPath = X

        '<EhFooter>
        Exit Function

AppPath_Err:
        WriteErr Time & "-" & Date & " - " & Err.Description & " - " & "in PerfectAntivirus2009.modScanVirus.AppPath " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Function

Public Function FileExists(sFile) As Boolean

        '<EhHeader>
        On Error GoTo FileExists_Err

        '</EhHeader>
        On Error Resume Next

100     FileExists = ((GetAttr(sFile) And vbDirectory) = 0)

        '<EhFooter>
        Exit Function

FileExists_Err:
        WriteErr Time & "-" & Date & " - " & Err.Description & " - " & "in PerfectAntivirus2009.modScanVirus.FileExists " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Function

Public Sub xScanProcess()

        '<EhHeader>
        On Error GoTo xScanProcess_Err

        '</EhHeader>

100     xTotalProcess = 0

        On Error Resume Next

        Dim ColItems

        Dim ObjItem

102     Set ColItems = GetObject("winmgmts:\root\CIMV2").ExecQuery("SELECT * FROM Win32_Process")

104     For Each ObjItem In ColItems

            'List1.AddItem objitem.executablepath
   
106         If ObjItem.Caption <> "System Idle Process" And ObjItem.Caption <> "System" Then
108             xTotalProcess = xTotalProcess + 1
110             frmMain.lblStatus.Caption = ObjItem.ExecutablePath

                Dim AX As String

112             AX = CheckVirus(ObjItem.ExecutablePath)

114             If AX <> "No" Then

                    Dim I

116                 I = frmMain.LVVirus1.ListItems.Count + 1
118                 frmMain.LVVirus1.ListItems.Add I, , AX
120                 frmMain.LVVirus1.ListItems(I).SubItems(1).Caption = ObjItem.ExecutablePath
122                 frmMain.LVVirus1.ListItems(I).SubItems(2).Caption = FileLen(ObjItem.ExecutablePath) & " Bytes"
124                 frmMain.LVVirus1.ListItems(I).SubItems(3).Caption = ObjItem.ProcessID
126                 frmMain.LVVirus1.ListItems(I).SubItems(4).Caption = "---"
128                 frmMain.LVVirus1.ListItems(I).Checked = True
                End If
            End If 'ObjItem.Caption <> "System Idle Process" And ObjItem.Caption <> "System"

        Next

130     Set ColItems = Nothing
132     Set ObjItem = Nothing

        '<EhFooter>
        Exit Sub

xScanProcess_Err:
        WriteErr Time & "-" & Date & " - " & Err.Description & " - " & "in PerfectAntivirus2009.modScanVirus.xScanProcess " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Public Sub xScanProcess2()

        '<EhHeader>
        On Error GoTo xScanProcess2_Err

        '</EhHeader>

        On Error Resume Next

        Dim ColItems

        Dim ObjItem

100     Set ColItems = GetObject("winmgmts:\root\CIMV2").ExecQuery("SELECT * FROM Win32_Process")

102     For Each ObjItem In ColItems

            'List1.AddItem objitem.executablepath
   
104         If ObjItem.Caption <> "System Idle Process" And ObjItem.Caption <> "System" Then

106             frmMain.cslblStatus.Caption = ObjItem.ExecutablePath

                Dim AX As String

108             AX = CheckVirus(ObjItem.ExecutablePath)

110             If AX <> "No" Then

                    Dim I

112                 I = frmMain.LVVirus2.ListItems.Count + 1
114                 frmMain.LVVirus2.ListItems.Add I, , AX
116                 frmMain.LVVirus2.ListItems(I).SubItems(1).Caption = ObjItem.ExecutablePath
118                 frmMain.LVVirus2.ListItems(I).SubItems(2).Caption = FileLen(ObjItem.ExecutablePath) & " Bytes"
120                 frmMain.LVVirus2.ListItems(I).SubItems(3).Caption = ObjItem.ProcessID
122                 frmMain.LVVirus2.ListItems(I).SubItems(4).Caption = "---"
124                 frmMain.LVVirus2.ListItems(I).Checked = True
                End If
            End If 'ObjItem.Caption <> "System Idle Process" And ObjItem.Caption <> "System"

        Next

126     Set ColItems = Nothing
128     Set ObjItem = Nothing

        '<EhFooter>
        Exit Sub

xScanProcess2_Err:
        WriteErr Time & "-" & Date & " - " & Err.Description & " - " & "in PerfectAntivirus2009.modScanVirus.xScanProcess2 " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Public Sub DelAllLV(LV As UniListView)

        '<EhHeader>
        On Error GoTo DelAllLV_Err

        '</EhHeader>
        Dim k

100     For k = 1 To LV.ListItems.Count

102         If k > LV.ListItems.Count Then Exit Sub
104         LV.ListItems.Remove k
106         k = k - 1
108     Next k

        '<EhFooter>
        Exit Sub

DelAllLV_Err:
        WriteErr Time & "-" & Date & " - " & Err.Description & " - " & "in PerfectAntivirus2009.modScanVirus.DelAllLV " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Public Sub DelAllChecked(LV As UniListView)

        '<EhHeader>
        On Error GoTo DelAllChecked_Err

        '</EhHeader>
        Dim k

100     For k = 1 To LV.ListItems.Count

102         If k > LV.ListItems.Count Then Exit Sub
104         If LV.ListItems(k).Checked = True Then
106             LV.ListItems.Remove k
108             k = k - 1
            End If

110     Next k

        '<EhFooter>
        Exit Sub

DelAllChecked_Err:
        WriteErr Time & "-" & Date & " - " & Err.Description & " - " & "in PerfectAntivirus2009.modScanVirus.DelAllChecked " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Public Function GetOpenAutorun(sAutorunFile) As String

        '<EhHeader>
        On Error GoTo GetOpenAutorun_Err

        '</EhHeader>
        On Error Resume Next

        Dim xStart1

        Dim xEnd1

        Dim xAutoFile

        Dim x1 As String

        Dim FSO

100     Set FSO = CreateObject("Scripting.FileSystemObject").OpenTextFile(sAutorunFile, 1, , -2)
102     xAutoFile = FSO.ReadAll
104     xAutoFile = DelAllSpace(xAutoFile)
106     Set FSO = Nothing
108     xStart1 = (Len(xAutoFile) - InStrRev(LCase(StrReverse(xAutoFile)), StrReverse("open=")) + 2)
110     xEnd1 = (Len(xAutoFile) - InStrRev(StrReverse(xAutoFile), Chr(13), Len(xAutoFile) - xStart1) + 1)
112     x1 = Mid$(xAutoFile, (Len(xAutoFile) - InStrRev(LCase(StrReverse(xAutoFile)), StrReverse("open=")) + 2), (Len(xAutoFile) - InStrRev(StrReverse(xAutoFile), Chr(13), Len(xAutoFile) - xStart1) + 1) - (Len(xAutoFile) - InStrRev(LCase(StrReverse(xAutoFile)), StrReverse("open=")) + 2))
114     GetOpenAutorun = x1

        '<EhFooter>
        Exit Function

GetOpenAutorun_Err:
        WriteErr Time & "-" & Date & " - " & Err.Description & " - " & "in PerfectAntivirus2009.modScanVirus.GetOpenAutorun " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Function

Public Function GetShellOpenAutorun(sAutorunFile) As String

        '<EhHeader>
        On Error GoTo GetShellOpenAutorun_Err

        '</EhHeader>
        On Error Resume Next

        Dim xStart1

        Dim xEnd1

        Dim xAutoFile

        Dim x1 As String

        Dim FSO

100     Set FSO = CreateObject("Scripting.FileSystemObject").OpenTextFile(sAutorunFile, 1, , -2)
102     xAutoFile = FSO.ReadAll
104     xAutoFile = DelAllSpace(xAutoFile)
106     Set FSO = Nothing
108     xStart1 = (Len(xAutoFile) - InStrRev(LCase(StrReverse(xAutoFile)), StrReverse("shell\open\command=")) + 2)
110     xEnd1 = (Len(xAutoFile) - InStrRev(StrReverse(xAutoFile), Chr(13), Len(xAutoFile) - xStart1) + 1)
112     x1 = Mid$(xAutoFile, (Len(xAutoFile) - InStrRev(LCase(StrReverse(xAutoFile)), StrReverse("shell\open\command=")) + 2), (Len(xAutoFile) - InStrRev(StrReverse(xAutoFile), Chr(13), Len(xAutoFile) - xStart1) + 1) - (Len(xAutoFile) - InStrRev(LCase(StrReverse(xAutoFile)), StrReverse("shell\open\command=")) + 2))
114     GetShellOpenAutorun = x1

        '<EhFooter>
        Exit Function

GetShellOpenAutorun_Err:
        WriteErr Time & "-" & Date & " - " & Err.Description & " - " & "in PerfectAntivirus2009.modScanVirus.GetShellOpenAutorun " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Function

Public Function WriteFileUni(FileName As String, Unistr As String)

        '<EhHeader>
        On Error GoTo WriteFileUni_Err

        '</EhHeader>
        Dim FSO As Object 'tao 1 file mo'i rôi mo'i ghi vào

100     Set FSO = CreateObject("Scripting.FileSystemObject").CreateTextFile(FileName, True)
102     Set FSO = Nothing
104     Set FSO = CreateObject("Scripting.FileSystemObject").OpenTextFile(FileName, 2, , -1)
106     FSO.Write Unistr
108     Set FSO = Nothing

        '<EhFooter>
        Exit Function

WriteFileUni_Err:
        WriteErr Time & "-" & Date & " - " & Err.Description & " - " & "in PerfectAntivirus2009.modScanVirus.WriteFileUni " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Function

Public Function DelAllSpace(Str) As String

        '<EhHeader>
        On Error GoTo DelAllSpace_Err

        '</EhHeader>
100     Do While InStr(Str, " ") > 0
102         Str = Replace(Str, " ", "")
        Loop

104     Str = Trim(Str)
106     DelAllSpace = Str

        '<EhFooter>
        Exit Function

DelAllSpace_Err:
        WriteErr Time & "-" & Date & " - " & Err.Description & " - " & "in PerfectAntivirus2009.modScanVirus.DelAllSpace " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Function
