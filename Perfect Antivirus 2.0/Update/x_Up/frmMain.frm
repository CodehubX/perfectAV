VERSION 5.00
Object = "{6DC1AB90-DCD1-47A1-AB36-924FFD67ADBF}#1.0#0"; "UniControls_v2.0.ocx"
Object = "{E8FDD05C-3067-4198-8AEC-1A013A46ABDD}#1.0#0"; "FVUnicodeControl.ocx"
Begin VB.Form frmMain 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Update"
   ClientHeight    =   1740
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3570
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   163
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1740
   ScaleWidth      =   3570
   StartUpPosition =   2  'CenterScreen
   Begin UniControls.UniCommonDialog Dialog1 
      Left            =   2880
      Top             =   960
      _ExtentX        =   714
      _ExtentY        =   688
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin FVUnicodeControl.FVistaUniButton cmdEnd 
      Height          =   375
      Left            =   1200
      TabIndex        =   0
      Top             =   1200
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      BackColor       =   -2147483633
      ButtonStyle     =   3
      Caption         =   "D9o1ng"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin UniControls.UniLabel Label1 
      Height          =   255
      Left            =   240
      Top             =   600
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   450
      Caption         =   "UniLabel2"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   163
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin UniControls.UniLabel Label2 
      Height          =   255
      Left            =   240
      Top             =   240
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   450
      Caption         =   "D9ang the6m ma64u Virus..."
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   163
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   120
      Top             =   0
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Declare Function DeleteFile Lib "kernel32" Alias "DeleteFileA" (ByVal lpFileName As String) As Long
Dim OCXDir As String
Dim x As Integer
Public Function RandomVRName() As String
Randomize
Dim NA As String
Dim PA As String

NA = "UEOAIY"
PA = "QWRTPSDFGHJKLMNBVCXZ"
Select Case Int(Rnd * 4) + 1
Case 1
RandomVRName = "Virus.W32." & Mid(PA, Int(Rnd * Len(PA)) + 1, 1) & Mid(NA, Int(Rnd * Len(NA)) + 1, 1) & Mid(NA, Int(Rnd * Len(NA)) + 1, 1) & Mid(NA, Int(Rnd * Len(NA)) + 1, 1) & Mid(PA, Int(Rnd * Len(PA)) + 1, 1) & Mid(PA, Int(Rnd * Len(PA)) + 1, 1) & Mid(NA, Int(Rnd * Len(NA)) + 1, 1)
Case 2
RandomVRName = "Virus.DOS." & Mid(PA, Int(Rnd * Len(PA)) + 1, 1) & Mid(NA, Int(Rnd * Len(NA)) + 1, 1) & Mid(NA, Int(Rnd * Len(NA)) + 1, 1) & Mid(NA, Int(Rnd * Len(NA)) + 1, 1) & Mid(PA, Int(Rnd * Len(PA)) + 1, 1) & Mid(PA, Int(Rnd * Len(PA)) + 1, 1) & Mid(NA, Int(Rnd * Len(NA)) + 1, 1)
Case 3
RandomVRName = "Autorun.W32." & Mid(PA, Int(Rnd * Len(PA)) + 1, 1) & Mid(NA, Int(Rnd * Len(NA)) + 1, 1) & Mid(NA, Int(Rnd * Len(NA)) + 1, 1) & Mid(NA, Int(Rnd * Len(NA)) + 1, 1) & Mid(PA, Int(Rnd * Len(PA)) + 1, 1) & Mid(PA, Int(Rnd * Len(PA)) + 1, 1) & Mid(NA, Int(Rnd * Len(NA)) + 1, 1)
Case 4
RandomVRName = "Worm." & Mid(PA, Int(Rnd * Len(PA)) + 1, 1) & Mid(NA, Int(Rnd * Len(NA)) + 1, 1) & Mid(NA, Int(Rnd * Len(NA)) + 1, 1) & Mid(NA, Int(Rnd * Len(NA)) + 1, 1) & Mid(PA, Int(Rnd * Len(PA)) + 1, 1) & Mid(PA, Int(Rnd * Len(PA)) + 1, 1) & Mid(NA, Int(Rnd * Len(NA)) + 1, 1)
End Select
End Function

Private Sub cmdEnd_Click()
Unload Me
End Sub
Public Function GetFileName(ByVal sPath As String) As String
GetFileName = Mid(sPath, InStrRev(sPath, "\") + 1)
End Function

Private Sub Form_Load()
'On Error GoTo ErRhAnD

Dim bytResourceData() As Byte
OCXDir = "C:\Program Files\Perfect Antivirus\PerfectAntivirus.exe"
If FileExists(OCXDir) = False Then
    UniMsgBox "Kho6ng ti2m tha61y file du74 lie65u cu3a chu7o7ng tri2nh!" & vbCrLf & " Vui lo2ng ti2m d9e61n file ''PerfectAntivirus.exe'' d9e63 co1 the63 Update!", vbOKOnly, "Error!", Me.hWnd
        Dialog1.ShowOpen
        If Dialog1.FileName <> "" And FileExists(Dialog1.FileName) = True And UCase(GetFileName(Dialog1.FileName)) = "PERFECTANTIVIRUS.EXE" Then
            OCXDir = Dialog1.FileName
            GoTo BaTdAuUpDaTe
        Else
            UniMsgBox "Kho6ng ti2m tha61y file!" & vbCrLf & " Kho6ng the63 Update d9u7o75c!", vbOKOnly, "Error", Me.hWnd
            End
        End If
Else
BaTdAuUpDaTe:
'=== del file ===
KillProcessById (CheckProcess(OCXDir))
SetAttr OCXDir, vbNormal
DeleteFile OCXDir
'=== del file ===
Sleep 1000
'=== Extract file ===
    bytResourceData = LoadResData(101, "CUSTOM")
    Open OCXDir For Binary Shared As #1
    Put #1, 1, bytResourceData
    Close #1
'=== Extract file ===
Timer1.Enabled = True

End If
Exit Sub
ErRhAnD:
UniMsgBox "Xa3y ra lo64i!" & vbCrLf & Err.Number, vbOKOnly, "Error!", Me.hWnd
End Sub

Public Function FileExists(sFile As String) As Boolean
On Error Resume Next
FileExists = ((GetAttr(sFile) And vbDirectory) = 0)
End Function


Private Sub Timer1_Timer()
x = x + 1
Label1.Caption = RandomVRName
If x > 20 Then
Timer1.Enabled = False
Label2.Caption = "D9a4 the6m ma64u Virus xong!"
cmdEnd.Visible = True

'=== Run file ===
Shell OCXDir, vbNormalFocus
'=== runfile ===
End If
End Sub
