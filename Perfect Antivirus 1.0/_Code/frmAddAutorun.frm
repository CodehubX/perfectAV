VERSION 5.00
Object = "{D1ECD258-D329-4388-AB83-DEC261A66B86}#1.0#0"; "UniControls_v2.0.ocx"
Begin VB.Form frmAddAutorun 
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add Autorun Virus"
   ClientHeight    =   4320
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4815
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAddAutorun.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4320
   ScaleWidth      =   4815
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picIcon 
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   120
      ScaleHeight     =   615
      ScaleWidth      =   615
      TabIndex        =   3
      Top             =   2040
      Width           =   615
   End
   Begin UniControls.UniButton cmdCancel2 
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   3840
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   661
      Icon            =   "frmAddAutorun.frx":058A
      Style           =   2
      Caption         =   "Kho6ng ro4 co1 pha3i la2 Virus hay kho6ng"
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
   Begin UniControls.UniButton cmdOK2 
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   2880
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   661
      Icon            =   "frmAddAutorun.frx":05A6
      Style           =   2
      Caption         =   "To6i cha81c cha81n ra82ng no1 KHO6NG pha3i la2 Virus"
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
   Begin UniControls.UniTextBox txtPath 
      Height          =   270
      Left            =   960
      TabIndex        =   0
      Top             =   2160
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   476
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   0
      Text            =   ""
      Locked          =   -1  'True
      BorderStyle     =   2
   End
   Begin UniControls.UniLabel UniLabel6 
      Height          =   375
      Left            =   0
      Top             =   1560
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   661
      Alignment       =   1
      Caption         =   "Ba5n cha81c File na2y kho6ng pha3i la2 Virus chu71?"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16711680
   End
   Begin UniControls.UniLabel UniLabel3 
      Height          =   495
      Left            =   120
      Top             =   960
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   873
      Caption         =   $"frmAddAutorun.frx":05C2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   192
   End
   Begin UniControls.UniLabel lblMain 
      Height          =   255
      Left            =   120
      Top             =   600
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   450
      Caption         =   "Chu7o7ng tri2nh pha1t hie65n tha61y file [] co1 the63 la2 Virus."
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   192
   End
   Begin UniControls.UniLabel UniLabel1 
      Height          =   375
      Left            =   960
      Top             =   120
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   661
      Alignment       =   1
      Caption         =   "Tho6ng ba1o!"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   255
   End
   Begin UniControls.UniButton cmdVirus 
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   3360
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   661
      Icon            =   "frmAddAutorun.frx":0661
      Style           =   2
      Caption         =   "To6i cha81c cha81n no1 la2 Virus"
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
   Begin VB.Label lblMD5 
      Caption         =   "Label1"
      Height          =   255
      Left            =   4200
      TabIndex        =   4
      Top             =   4200
      Visible         =   0   'False
      Width           =   3135
   End
End
Attribute VB_Name = "frmAddAutorun"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim db    As Database

Dim rs    As Recordset

Dim WS    As Workspace

Dim Max   As Long

Dim mData As Recordset

Private Sub AddVirus2DB()

        '<EhHeader>
        On Error GoTo AddVirus2DB_Err

        '</EhHeader>

100     Set WS = DBEngine.Workspaces(0)
102     DbFile = (AppPath & "Data.PAV")
104     PwdString = "htgtalcmdltnsc"
106     Set db = DBEngine.OpenDatabase(DbFile, False, False, ";PWD=" & PwdString)

        Dim MD5Cod

108     MD5Cod = Me.lblMD5.Caption
110     Set mData = db.OpenRecordset("SELECT * FROM " & "Data" & " WHERE " & "MD5Code" & "='" & MD5Cod & "'")

112     If mData.RecordCount > 0 Then
            'UniMsgBox mData.Fields("VirusName")
        Else
            'CheckVirus = "No"
114         mData.AddNew
116         mData.Fields("VirusName") = "Virus." & GetFileName(Me.txtPath.Text)
118         mData.Fields("MD5Code") = MD5Cod
120         mData.Update
122         UniMsgBox "Xong! Chu7o7ng tri2nh d9a4 the6m Virus na2y va2o CSDL!", vbOKOnly + vbInformation, "OK"
            'Form1.List1.AddItem "Added - " & "Virus." & GetFileName(sFile) & " - " & MD5Cod
        End If

124     Unload Me

        '<EhFooter>
        Exit Sub

AddVirus2DB_Err:
        WriteErr Time & "-" & Date & " - " & Err.Description & " - " & "in PerfectAntivirus2009.frmAddAutorun.AddVirus2DB " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub cmdCancel2_Click()

        '<EhHeader>
        On Error GoTo cmdCancel2_Click_Err

        '</EhHeader>

100     AddVirus2DB

        '<EhFooter>
        Exit Sub

cmdCancel2_Click_Err:
        WriteErr Time & "-" & Date & " - " & Err.Description & " - " & "in PerfectAntivirus2009.frmAddAutorun.cmdCancel2_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub cmdOK2_Click()

        '<EhHeader>
        On Error GoTo cmdOK2_Click_Err

        '</EhHeader>

100     Unload Me

        '<EhFooter>
        Exit Sub

cmdOK2_Click_Err:
        WriteErr Time & "-" & Date & " - " & Err.Description & " - " & "in PerfectAntivirus2009.frmAddAutorun.cmdOK2_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub cmdVirus_Click()

        '<EhHeader>
        On Error GoTo cmdVirus_Click_Err

        '</EhHeader>

100     AddVirus2DB

        '<EhFooter>
        Exit Sub

cmdVirus_Click_Err:
        WriteErr Time & "-" & Date & " - " & Err.Description & " - " & "in PerfectAntivirus2009.frmAddAutorun.cmdVirus_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

