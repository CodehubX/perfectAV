Attribute VB_Name = "basShowHid"
Public Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Long, lpdwProcessId As Long) As Long
Public Declare Function EnumWindows Lib "user32" (ByVal lpEnumFunc As Long, ByVal lParam As Long) As Boolean

Public Function EnumWindowsProc(ByVal hWnd As Long, ByVal lParam As Long) As Boolean
    Dim ID As Long
    Dim CoChua As Boolean
    Dim i As Integer
    CoChua = False
    GetWindowThreadProcessId hWnd, ID
    
    With frmMain.lstPro
        For i = 0 To .ListCount - 1
            If ID = Val(.List(i)) Then CoChua = True
        Next
        If CoChua = False Then .AddItem ID
    End With
If CoChua = False Then
    If CheckID(ID) <> ID Then
        Dim tmp As String
        tmp = ProcessPathByPID(ID)
    
        CoChua = False
        '/////////////here
        'With frmPro

        '    Set lsv = .LV.ListItems.Add()
        '    lsv.Text = GetFileName(tmp)
        '    lsv.SubItems(1) = tmp
        '    lsv.SubItems(2) = ID
        '    lsv.ForeColor = vbRed
            
        'End With
        frmMain.SoLuong = frmMain.SoLuong + 1
        Dim u
        u = frmMain.LVPro.ListItems.Count + 1
        frmMain.LVPro.ListItems.Add u, , GetFileName(tmp)
        frmMain.LVPro.ListItems(u).SubItems(1).Caption = tmp
        frmMain.LVPro.ListItems(u).SubItems(2).Caption = ID
        frmMain.LVPro.ListItems(u).ForeColor = vbRed
'        frmmain.LV1.ListItems(u).Font.Bold = True
        frmMain.LVPro.Refresh
        
    End If
End If
    EnumWindowsProc = True
End Function
