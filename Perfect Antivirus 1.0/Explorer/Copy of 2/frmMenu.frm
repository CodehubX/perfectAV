VERSION 5.00
Object = "{D1ECD258-D329-4388-AB83-DEC261A66B86}#1.0#0"; "UniControls_v2.0.ocx"
Begin VB.Form frmMenu 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin UniControls.UniMenu UniMenu1 
      Left            =   2760
      Top             =   720
      _ExtentX        =   741
      _ExtentY        =   767
   End
   Begin VB.Menu ttf 
      Caption         =   "Thao Ta1c File"
      Begin VB.Menu open 
         Caption         =   "Mo73 File (Open)"
      End
      Begin VB.Menu Del 
         Caption         =   "Xo1a File (Delete)"
      End
      Begin VB.Menu shide 
         Caption         =   "Xo1a Thuo65c Ti1nh A63n"
      End
      Begin VB.Menu Move 
         Caption         =   "Di Chuye63n (Cut)"
      End
      Begin VB.Menu copy 
         Caption         =   "Sao Che1p (Copy)"
      End
      Begin VB.Menu paste 
         Caption         =   "Da1n (Paste)"
      End
      Begin VB.Menu openwithnotepad 
         Caption         =   "Mo73 Ba82ng Notepad"
      End
      Begin VB.Menu properties 
         Caption         =   "Xem Thuo65c Ti1nh (Properties)"
      End
   End
   Begin VB.Menu ttfo 
      Caption         =   "Thao Ta1c Folder"
      Begin VB.Menu createFoldre 
         Caption         =   "Ta5o Thu7 Mu5c"
      End
      Begin VB.Menu Goto 
         Caption         =   "D9i D9e61n Thu7 Mu5c Na2y"
      End
   End
End
Attribute VB_Name = "frmMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public sClipBoard
Private Sub copy_Click()
sClipBoard = ""
sClipBoard = frmMain.txtPath.Text & "/0"
UniMsgBox "D9a4 lu7u ta65p tin va2o bo65 nho71 d9e63 sao che1p!", vbOKOnly, "Copy"
frmMain.lstFolder.Refresh
End Sub

Private Sub createFoldre_Click()
Dim sPathFolder
sPathFolder = frmMain.FolderDir.List(frmMain.FolderDir.ListIndex) & "\"
Dim sFolderName
sFolderName = UniInputbox("NhËp vµo tªn Folder cÇn t¹o", "T¹o Folder")
If sFolderName <> "" Then
MkDir sPathFolder & sFolderName
UniMsgBox ChrW(272) & ChrW(227) & " t" & ChrW(7841) & "o Folder xong.", vbOKOnly, "!", frmMain.hwnd
Dim sPathSave
sPathSave = ""
sPathSave = frmMain.FolderDir.List(frmMain.FolderDir.ListIndex) & "\"
frmMain.FolderDir.Path = ""
frmMain.FolderDir.Path = sPathSave

End If
End Sub


Private Sub del_Click()
On Error GoTo del_Err
If UniMsgBox("Ba5n co1 muo61ng xo1a File " & GetFileName(frmMain.txtPath.Text) & " kho6ng?", vbYesNo, "Xo1a") = vbYes Then
Dim sPath
sPath = frmMain.txtPath.Text
    If FileExists(sPath) = True Then
        DeleteFile sPath
        If FileExists(sPath) = False Then UniMsgBox "D9a4 xo1a xong!" Else UniMsgBox "Kho6ng xo1a d9u7o75c!"
    Else
        UniMsgBox "Ta65p tin kho6ng to62n ta5i!", vbOKOnly, "Delete"
    End If
End If
Exit Sub
del_Err:
UniMsgBox "Xa3y ra lo64i trong qua1 tri2nh thu75c hie65n." & vbCrLf & " Tho6ng tin lo64i: " & Err.Number & " - " & Err.Description

End Sub



Private Sub Form_Load()
UniMenu1.InitUnicodeMenu frmMenu.hwnd
End Sub


Private Sub Goto_Click()
Dim sPathFolder
sPathFolder = frmMain.FolderDir.List(frmMain.FolderDir.ListIndex) & "\"
Shell "explorer " & sPathFolder, vbNormalFocus
End Sub

Private Sub Move_Click()
sClipBoard = ""
sClipBoard = frmMain.txtPath.Text & "/1"
UniMsgBox "D9a4 lu7u va2o bo65 nho71 d9e63 di chuye63n"
End Sub

Private Sub open_Click()
If FileExists(frmMain.txtPath.Text) = True Then
    If UniMsgBox("Ba5n muo61ng mo73 File " & GetFileName(frmMain.txtPath.Text) & " kho6ng?", vbYesNo) = vbYes Then
        ShellExecute Me.hwnd, vbNullString, frmMain.txtPath.Text, vbNullString, "", 1
    End If
Else
    UniMsgBox "Ta65p tin kho6ng to62n ta5i!", vbOKOnly, "Open"
End If
End Sub

Private Sub openwithnotepad_Click()
On Error GoTo Notepad_Err
Dim sPath
sPath = frmMain.txtPath.Text
If FileExists(sPath) = True Then
    Shell "notepad " & sPath, vbNormalFocus
Else
    UniMsgBox "Ta65p tin kho6ng to62n ta5i!", vbOKOnly, "Open With Notepad"
End If
Exit Sub
Notepad_Err:
UniMsgBox "Xa3y ra lo64i trong qua1 tri2nh thu75c hie65n." & vbCrLf & " Tho6ng tin lo64i: " & Err.Number & " - " & Err.Description

End Sub

Private Sub paste_Click()
On Error GoTo Paste_Err




If sClipBoard <> "" Then 'clip

Dim S1
Dim S2
Dim Gx
    S1 = Left(sClipBoard, Len(sClipBoard) - 2)
    S2 = Right(sClipBoard, 2)
    
    If FolderExists(frmMain.txtPath.Text) = True Then 'folder files
        Gx = frmMain.txtPath.Text & "\" & GetFileName(S1)
    ElseIf FileExists(frmMain.txtPath.Text) = True Then
        Gx = GetFolderPath(frmMain.txtPath.Text) & "\" & GetFileName(S1)
    ElseIf FolderExists(frmMain.txtPath.Text) = False And FileExists(frmMain.txtPath.Text) = False Then
        UniMsgBox "No7i d9e61n kho6ng co1 tha65t!", vbOKOnly
        Exit Sub
    End If 'folder files
    
    If FileExists(S1) = True Then 'exists
    Dim Mn
    Dim M1
    Dim M2
    M1 = GetFolderPath(Gx)
    M2 = GetFileName(Gx)
    Mn = 1

    If S1 = Gx Then Gx = M1 & "\Copy of [" & Mn & "] " & M2
    
    While FileExists(Gx) = True
        MsgBox Mn
        Mn = Mn + 1
        MsgBox Mn
        Gx = M1 & "\Copy of [" & Mn & "] " & M2

        If S2 = "/0" Then ' /0 = copy
            FileCopy S1, Gx
            UniMsgBox "D9a4 Sao Che1p Xong!", vbOKOnly, "Paste"
        Else
            FileCopy S1, Gx
            SetAttr Gx, vbNormal
            DeleteFile S1
            UniMsgBox "D9a4 Di Chuye63n Xong!", vbOKOnly, "Paste"
            sClipBoard = ""
        End If '/0 = copy
    Wend
    
    Else
        UniMsgBox "Ta65p tin ba5n vu72a 'Sao Che1p' hoa85c 'Di Chuye63n' hie65n gio72 kho6ng to62n ta5i!", vbOKOnly, "Paste"
    End If
Else
    UniMsgBox "Ba5n va64n chu7a 'Sao Che1p' hoa85c 'Di Chuye63n' ta65p tin na2o", vbOKOnly, "Paste"
End If

frmMain.lstFolder.Refresh
frmMain.lstFolder.Path = GetFolderPath(frmMain.txtPath.Text)

Exit Sub
Paste_Err:
UniMsgBox "Xa3y ra lo64i trong qua1 tri2nh thu75c hie65n." & vbCrLf & " Tho6ng tin lo64i: " & Err.Number & " - " & Err.Description
End Sub

Private Sub properties_Click()
Dim ssPath As String
ssPath = frmMain.txtPath.Text
ShowProperties ssPath

End Sub



Private Sub shide_Click()
On Error GoTo shide
Dim sPath
sPath = frmMain.txtPath.Text
If FileExists(sPath) = True Then
    SetAttr sPath, vbNormal
    UniMsgBox "Ta65p tin d9a4 d9u7o75c hie65n.", vbOKOnly, "OK"
Else
    UniMsgBox "Ta65p tin kho6ng to62n ta5i!", vbOKOnly
End If

Exit Sub
shide_Err:
UniMsgBox "Xa3y ra lo64i trong qua1 tri2nh thu75c hie65n." & vbCrLf & " Tho6ng tin lo64i: " & Err.Number & " - " & Err.Description

End Sub
