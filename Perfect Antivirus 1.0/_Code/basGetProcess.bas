Attribute VB_Name = "basOther"


Public Sub GetProcess(lstView As UniListView)

'On Error Resume Next
frmMain.SoLuong = 0
    lstView.BackColor = vbWhite
    lstView.ListItems.Clear
    'lstView.SmallIcons = Nothing
    frmMain.lstPro.Clear
'---------Liet ke process-------
  Dim theloop As Long
  Dim proc As PROCESSENTRY32
  Dim snap As Long
  Dim exename As String
  Dim ID As Long
   snap = CreateToolhelpSnapshot(TH32CS_SNAPall, 0)
   proc.dwSize = Len(proc)
   theloop = ProcessFirst(snap, proc)
   While theloop <> 0

      ID = proc.th32ProcessID
      theloop = ProcessNext(snap, proc)
      If ProcessPathByPID(proc.th32ProcessID) <> "SYSTEM" Then
      'MsgBox ProcessPathByPID(proc.th32ProcessID)

                  'Set lsv = lstView.ListItems.Add()
                  'lsv.Text = proc.szExeFile
                  'lsv.SubItems(1) = ProcessPathByPID(proc.th32ProcessID)
                  'lsv.SubItems(2) = proc.th32ProcessID
                  Dim i
                  i = lstView.ListItems.Count + 1
                  lstView.ListItems.Add i, , proc.szExeFile
                  lstView.ListItems(i).SubItems(1).Caption = ProcessPathByPID(proc.th32ProcessID)
                  lstView.ListItems(i).SubItems(2).Caption = proc.th32ProcessID
                  'lstView.ListItems(i).SubItems(3).Caption = FileLen(ProcessPathByPID(proc.th32ProcessID))
                  
        End If
   Wend
   CloseHandle snap
       EnumWindows AddressOf EnumWindowsProc, ByVal 0&

End Sub

