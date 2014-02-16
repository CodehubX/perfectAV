VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Remove Perfect Antivirus 2009"
   ClientHeight    =   1080
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6240
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1080
   ScaleWidth      =   6240
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdNo 
      Caption         =   "Không"
      Height          =   360
      Left            =   3240
      TabIndex        =   2
      Top             =   600
      Width           =   990
   End
   Begin VB.ListBox List1 
      Height          =   2595
      Left            =   120
      TabIndex        =   3
      Top             =   1080
      Width           =   6015
   End
   Begin VB.CommandButton cmdYes 
      Caption         =   "Có"
      Height          =   360
      Left            =   1920
      TabIndex        =   0
      Top             =   600
      Width           =   990
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Ban co muon go bo chuong trinh Perfect Antivirus 2009 ra khoi may tinh khong?"
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   6060
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function DeleteFile Lib "kernel32" Alias "DeleteFileA" (ByVal lpFileName As String) As Long

Private Sub cmdNo_Click()
Unload Me
End Sub

Private Sub cmdYes_Click()
If FileExists(AppPath & "Filelist.usl") = False Then
    MsgBox "File not found:" & vbCrLf & "Filelist.usl"
    Exit Sub
End If
    List1.AddItem "Now Start Remove.............."
    List1.AddItem "............................."

    
    Dim Str As String
    Dim InputData As String
    Open AppPath & "FileList.usl" For Input As #1
    Do While Not EOF(1)
        Line Input #1, InputData
        If FileExists(InputData) = True Then
            'xoa file
BaTdAuXoAfIlE:
            If zXoaFile(InputData) = False Then
                MsgBox "File is running:" & vbCrLf & InputData & vbCrLf & "Please close this file before remove PAV 2009.", vbOKOnly, "File is running!"
                GoTo BaTdAuXoAfIlE
            Else
                List1.AddItem "Deleted: " & InputData
            End If
        End If
    Loop
    Close #1
    DeleteKey HKEY_CLASSES_ROOT, "Folder\shell\[PAV 2009] Quét Virus\command"
    DeleteKey HKEY_CLASSES_ROOT, "Folder\shell\[PAV 2009] Quét Virus"
    SaveString HKEY_CLASSES_ROOT, "exefile\shell\open\command", "", ChrW(34) & "%1" & ChrW(34) & " %*"
    SaveString HKEY_CLASSES_ROOT, "batfile\shell\open\command", "", ChrW(34) & "%1" & ChrW(34) & " %*"
    SaveString HKEY_CLASSES_ROOT, "cmdfile\shell\open\command", "", ChrW(34) & "%1" & ChrW(34) & " %*"
    SaveString HKEY_CLASSES_ROOT, "comfile\shell\open\command", "", ChrW(34) & "%1" & ChrW(34) & " %*"
            
    List1.AddItem "Delete in Registry..........."
    List1.AddItem "............................."
    List1.AddItem "............................."
    List1.AddItem "Complete....................."
    List1.AddItem "Thank for use this program..."
    
    Me.Height = 4300
    Me.cmdNo.Caption = "Thoát"
    Me.cmdNo.Left = 2520
    Me.cmdNo.Top = 600
    Me.cmdYes.Visible = False
    Me.Label1.Caption = "Da go bo xong! Cam on ban da su dung chuong trinh!"
End Sub


Public Function AppPath()
AppPath = App.Path
If Right$(AppPath, 1) <> "\" Then AppPath = AppPath & "\"
End Function
Public Function FileExists(sFile) As Boolean
    On Error Resume Next
    FileExists = ((GetAttr(sFile) And vbDirectory) = 0)
End Function
Public Function zXoaFile(xFile) As Boolean
SetAttr xFile, vbNormal
DeleteFile xFile
zXoaFile = Not FileExists(xFile)
End Function

