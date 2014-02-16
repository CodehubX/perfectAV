VERSION 5.00
Object = "{D1ECD258-D329-4388-AB83-DEC261A66B86}#1.0#0"; "UniControls_v2.0.ocx"
Begin VB.Form frmMain 
   BackColor       =   &H80000009&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PAV 2009 - Auto Update"
   ClientHeight    =   4230
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6600
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
   ScaleHeight     =   4230
   ScaleWidth      =   6600
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin UniControls.UniCommonDialog Dialog1 
      Left            =   6120
      Top             =   480
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
   Begin UniControls.UniButton cmdExit 
      Height          =   375
      Left            =   2400
      TabIndex        =   1
      Top             =   3720
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      Icon            =   "frmMain.frx":628A
      Style           =   2
      Caption         =   "D9o1ng"
      IconAlign       =   3
      iNonThemeStyle  =   2
      MaskColor       =   16711935
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      RoundedBordersByTheme=   0   'False
      ShowFocusRectangle=   0   'False
   End
   Begin UniControls.UniListBox lstVR 
      Height          =   2055
      Left            =   240
      TabIndex        =   0
      Top             =   1440
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   3625
      IconMaskColor   =   16711935
      Picture         =   "frmMain.frx":62A6
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      RowHeight       =   19
      GridLines       =   -1  'True
   End
   Begin UniControls.UniLabel UniLabel3 
      Height          =   255
      Left            =   240
      Top             =   1080
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   450
      Caption         =   "Danh Sa1ch Ca1c Virus Mo71i:"
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
   Begin UniControls.UniLabel UniLabel2 
      Height          =   255
      Left            =   360
      Top             =   600
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   450
      Alignment       =   1
      Caption         =   "D9a4 Ca65p Nha65t The6m Virus Mo71i"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   255
   End
   Begin UniControls.UniLabel UniLabel1 
      Height          =   375
      Left            =   360
      Top             =   120
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   661
      Alignment       =   1
      Caption         =   "Tu75 D9o65ng Ca65p Nha65t"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16711680
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function BringWindowToTop Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function DeleteFile Lib "kernel32" Alias "DeleteFileA" (ByVal lpFileName As String) As Long

Dim xUpdateOn As String
Dim xThongTin As String

Dim db As Database
Dim rs As Recordset
Dim WS As Workspace
Dim mData As Recordset
Dim nData As Recordset
Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub Form_Load()
If App.PrevInstance = True Then End
xThongTin = ""
xUpdateOn = "X"
Dim Comd
Comd = Command()
xThongTin = GetUrlSource("http://trung.trung12345.googlepages.com/PAV2009Version.txt")

If Comd = "off" Then
    'Update Offline
    UniMsgBox " Ha4y cho5n File Update Offline d9e63 chu7o7ng tri2nh ca65p nha65t." & vbCrLf & " Ne61u ba5n kho6ng co1 File na2y, co1 the63 ta3i ve62 tu72 d9i5a chi3: http://qts.come.vn"
    Dim OffFile As String
    Dialog1.FileName = ""
    Dialog1.Filter = "*.upd|*.upd|"
    Dialog1.ShowOpen
    OffFile = Dialog1.FileName
    If OffFile <> "" Then
        xUpdateOn = "OFF"
        xThongTin = ReadFile(OffFile)
    End If
ElseIf Comd = "on" Then
    'Update Online
    xUpdateOn = "ON"
    xThongTin = GetUrlSource("http://trung.trung12345.googlepages.com/PAV2009Version.txt")
End If

App.TaskVisible = False
'Lay thong tin tu` Server
'Test Offline

'Ghi du lieu vao xThongTin
'////////////////////////////////////
'xThongTin = ReadFile("C:\PAVVer.txt")
'////////////////////////////////////

'Ghi xThongTin vao File
DeleteFile AppPath & "temp.upd"
WriteFile AppPath & "temp.upd", xThongTin

'Kiem tra phien ban hien tai
Dim xNowVersion As Integer
xNowVersion = ReadIniFile(AppPath & "Version.txt", "Version", "Ver", 0)

'Kiem Tra Phien Ban Moi Tai Ve
Dim xVersion As Integer
xVersion = ReadIniFile(AppPath & "temp.upd", "Version", "Ver", 0)

'Kiem tra phien ban co khac nhau hay ko

If xNowVersion < xVersion Then ' Newversion
    'Neu be hon thi la co phien ban moi

    'hien Form
    Me.Show
    BringWindowToTop Me.hwnd
    App.TaskVisible = True
    
    'Them vao CSDL
    Dim xVrNameMD As String
    Dim xVrMD5 As String
    Dim xVrNameStr As String
    Dim xVrString As String
    Dim xTotal
    
    xTotal = ReadIniFile(AppPath & "temp.upd", "Total", "To", 1)
    

        Dim i
        For i = 1 To xTotal
            xVrNameMD = ReadIniFile(AppPath & "temp.upd", i, "NAMEMD", "")
            xVrMD5 = ReadIniFile(AppPath & "temp.upd", i, "MD5", "")
            xVrNameStr = ReadIniFile(AppPath & "temp.upd", i, "NAMESTR", "")
            xVrString = ReadIniFile(AppPath & "temp.upd", i, "STRING", "")
            
            'MsgBox xVrNameMD & " - " & xVrMD5 & vbCrLf & xVrNameStr & " - " & xVrString

                'Add to DB
                ConnectDB 'connect to database
            If xVrNameMD <> "" And xVrMD5 <> "" Then
                Set mData = db.OpenRecordset("SELECT * FROM " & "Data" & " WHERE " & "MD5Code" & "='" & xVrMD5 & "'")
                If mData.RecordCount > 0 Then
                    'Da co' roi`
                Else
                lstVR.AddItem xVrNameMD
                mData.AddNew
                mData.Fields("VirusName") = xVrNameMD
                mData.Fields("MD5Code") = xVrMD5
                mData.Update
                End If
            End If

            If xVrNameStr <> "" And xVrString <> "" Then
                Set nData = db.OpenRecordset("SELECT * FROM " & "STRI" & " WHERE " & "String" & "='" & xVrString & "'")
                If nData.RecordCount > 0 Then
                    'Da co' roi`
                Else
                lstVR.AddItem xVrNameStr
                nData.AddNew
                nData.Fields("VirusName") = xVrNameStr
                nData.Fields("String") = xVrString
                nData.Update
                End If
            End If
            
        Next i
        'DeleteFile AppPath & "Data.str"
        'WriteFile AppPath & "Data.str", OldStr
        
    'Ghi phien ban moi vao
    WriteIniFile AppPath & "Version.txt", "Version", "Ver", xVersion
    
    DeleteFile AppPath & "temp.upd"
Else
    'Neu ko co phien ban moi thi thoat'
    If xUpdateOn = "X" Then
        DeleteFile AppPath & "temp.upd"
        End
    Else
        UniMsgBox "Kho6ng co1 ma64u Virus na2o mo71i ho7n!", vbOKOnly + vbInformation, "Update"
        DeleteFile AppPath & "temp.upd"
        End
    End If
End If
End Sub
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
