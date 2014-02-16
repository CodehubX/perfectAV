VERSION 5.00
Begin VB.UserControl WM11ToolBar 
   Alignable       =   -1  'True
   AutoRedraw      =   -1  'True
   ClientHeight    =   5550
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7845
   EditAtDesignTime=   -1  'True
   ScaleHeight     =   5550
   ScaleWidth      =   7845
   Begin VB.PictureBox picDropDownBox 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Left            =   5640
      ScaleHeight     =   345
      ScaleWidth      =   2025
      TabIndex        =   22
      Top             =   1200
      Visible         =   0   'False
      Width           =   2025
   End
   Begin VB.Timer tmrMoveOff 
      Left            =   7080
      Top             =   2520
   End
   Begin VB.PictureBox picArrowRight 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   525
      Left            =   480
      ScaleHeight     =   525
      ScaleWidth      =   450
      TabIndex        =   21
      Top             =   1320
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.PictureBox picArrowLeft 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   525
      Left            =   0
      ScaleHeight     =   525
      ScaleWidth      =   450
      TabIndex        =   20
      Top             =   1320
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.PictureBox picOption 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   150
      Index           =   0
      Left            =   0
      ScaleHeight     =   150
      ScaleWidth      =   240
      TabIndex        =   19
      Top             =   600
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox picOptionRight 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   600
      Index           =   0
      Left            =   3080
      Picture         =   "WM11ToolBar.ctx":0000
      ScaleHeight     =   600
      ScaleWidth      =   75
      TabIndex        =   18
      Top             =   1920
      Width           =   75
   End
   Begin VB.PictureBox picBackOption 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   450
      Index           =   1
      Left            =   2240
      Picture         =   "WM11ToolBar.ctx":02C2
      ScaleHeight     =   450
      ScaleWidth      =   840
      TabIndex        =   17
      Top             =   2640
      Width           =   840
   End
   Begin VB.PictureBox picBackOption 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   600
      Index           =   0
      Left            =   2240
      Picture         =   "WM11ToolBar.ctx":16B4
      ScaleHeight     =   600
      ScaleWidth      =   840
      TabIndex        =   16
      Top             =   1920
      Width           =   840
   End
   Begin VB.PictureBox picOptionRight 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   450
      Index           =   1
      Left            =   3060
      Picture         =   "WM11ToolBar.ctx":3136
      ScaleHeight     =   450
      ScaleWidth      =   75
      TabIndex        =   15
      Top             =   2640
      Width           =   75
   End
   Begin VB.PictureBox picOptionLeft 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   450
      Index           =   1
      Left            =   2160
      Picture         =   "WM11ToolBar.ctx":3358
      ScaleHeight     =   450
      ScaleWidth      =   75
      TabIndex        =   14
      Top             =   2640
      Width           =   75
   End
   Begin VB.PictureBox picOptionLeft 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   600
      Index           =   0
      Left            =   2160
      Picture         =   "WM11ToolBar.ctx":357A
      ScaleHeight     =   600
      ScaleWidth      =   75
      TabIndex        =   13
      Top             =   1920
      Width           =   75
   End
   Begin VB.PictureBox picBackButton 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1035
      Index           =   3
      Left            =   1140
      Picture         =   "WM11ToolBar.ctx":383C
      ScaleHeight     =   1035
      ScaleWidth      =   840
      TabIndex        =   12
      Top             =   3360
      Width           =   840
   End
   Begin VB.PictureBox picRight 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1035
      Index           =   3
      Left            =   1980
      Picture         =   "WM11ToolBar.ctx":65C6
      ScaleHeight     =   1035
      ScaleWidth      =   75
      TabIndex        =   11
      Top             =   3360
      Width           =   75
   End
   Begin VB.PictureBox picLeft 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1035
      Index           =   3
      Left            =   1080
      Picture         =   "WM11ToolBar.ctx":6A58
      ScaleHeight     =   1035
      ScaleWidth      =   75
      TabIndex        =   10
      Top             =   3360
      Width           =   75
   End
   Begin VB.PictureBox picBackButton 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1380
      Index           =   2
      Left            =   1140
      Picture         =   "WM11ToolBar.ctx":6EEA
      ScaleHeight     =   1380
      ScaleWidth      =   840
      TabIndex        =   9
      Top             =   1920
      Width           =   840
   End
   Begin VB.PictureBox picRight 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1380
      Index           =   2
      Left            =   1980
      Picture         =   "WM11ToolBar.ctx":AB8C
      ScaleHeight     =   1380
      ScaleWidth      =   75
      TabIndex        =   8
      Top             =   1920
      Width           =   75
   End
   Begin VB.PictureBox picLeft 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1380
      Index           =   2
      Left            =   1080
      Picture         =   "WM11ToolBar.ctx":B18E
      ScaleHeight     =   1380
      ScaleWidth      =   75
      TabIndex        =   7
      Top             =   1920
      Width           =   75
   End
   Begin VB.PictureBox picBackButton 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1485
      Index           =   1
      Left            =   60
      Picture         =   "WM11ToolBar.ctx":B790
      ScaleHeight     =   1485
      ScaleWidth      =   840
      TabIndex        =   6
      Top             =   3480
      Width           =   840
   End
   Begin VB.PictureBox picBackButton 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1485
      Index           =   0
      Left            =   60
      Picture         =   "WM11ToolBar.ctx":F8CA
      ScaleHeight     =   1485
      ScaleWidth      =   840
      TabIndex        =   5
      Top             =   1920
      Width           =   840
   End
   Begin VB.PictureBox picRight 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1485
      Index           =   1
      Left            =   900
      Picture         =   "WM11ToolBar.ctx":13A04
      ScaleHeight     =   1485
      ScaleWidth      =   75
      TabIndex        =   4
      Top             =   3480
      Width           =   75
   End
   Begin VB.PictureBox picRight 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1485
      Index           =   0
      Left            =   900
      Picture         =   "WM11ToolBar.ctx":14076
      ScaleHeight     =   1485
      ScaleWidth      =   75
      TabIndex        =   3
      Top             =   1920
      Width           =   75
   End
   Begin VB.PictureBox picLeft 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1485
      Index           =   1
      Left            =   0
      Picture         =   "WM11ToolBar.ctx":146E8
      ScaleHeight     =   1485
      ScaleWidth      =   75
      TabIndex        =   2
      Top             =   3480
      Width           =   75
   End
   Begin VB.PictureBox picLeft 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1485
      Index           =   0
      Left            =   0
      Picture         =   "WM11ToolBar.ctx":14D5A
      ScaleHeight     =   1485
      ScaleWidth      =   75
      TabIndex        =   1
      Top             =   1920
      Width           =   75
   End
   Begin VB.PictureBox picButton 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Index           =   0
      Left            =   60
      ScaleHeight     =   495
      ScaleWidth      =   1200
      TabIndex        =   0
      Top             =   60
      Width           =   1200
   End
   Begin VB.Image imgRight 
      Height          =   555
      Left            =   7680
      Picture         =   "WM11ToolBar.ctx":153CC
      Top             =   0
      Width           =   75
   End
   Begin VB.Image imgLeft 
      Height          =   555
      Left            =   0
      Picture         =   "WM11ToolBar.ctx":1565E
      Top             =   0
      Width           =   75
   End
   Begin VB.Image imgDropDownBoxArrow 
      Height          =   1035
      Left            =   5880
      Picture         =   "WM11ToolBar.ctx":158F0
      Top             =   2040
      Width           =   195
   End
   Begin VB.Image imgDropDownBoxBG 
      Height          =   1035
      Left            =   3840
      Picture         =   "WM11ToolBar.ctx":163FA
      Stretch         =   -1  'True
      Top             =   2040
      Width           =   2025
   End
   Begin VB.Image imgDropDownBoxLeft 
      Height          =   1035
      Left            =   3600
      Picture         =   "WM11ToolBar.ctx":169A0
      Top             =   2040
      Width           =   210
   End
   Begin VB.Image imgArrowRight 
      Height          =   2100
      Left            =   2640
      Picture         =   "WM11ToolBar.ctx":175BE
      Top             =   3120
      Width           =   450
   End
   Begin VB.Image imgArrowLeft 
      Height          =   2100
      Left            =   2160
      Picture         =   "WM11ToolBar.ctx":1A850
      Top             =   3120
      Width           =   450
   End
   Begin VB.Image imgSeparator 
      Height          =   555
      Index           =   0
      Left            =   0
      Picture         =   "WM11ToolBar.ctx":1DAE2
      Top             =   720
      Width           =   45
   End
   Begin VB.Image imgBackGround 
      Height          =   555
      Left            =   60
      Picture         =   "WM11ToolBar.ctx":1DCE0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1575
   End
End
Attribute VB_Name = "WM11ToolBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Please do not remove these lines
'Control Windows Media Player 11 Toolbar
'Author: vie87vn
'From: www.caulacbovb.com

Option Explicit

'API Function
Private Declare Function AppendMenu Lib "user32" Alias "AppendMenuW" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As Long) As Long
Private Declare Function CreateMenu Lib "user32" () As Long
Private Declare Function CreatePopupMenu Lib "user32" () As Long
Private Declare Function DestroyMenu Lib "user32" (ByVal hMenu As Long) As Long
Private Declare Function FindWindow Lib "user32.dll" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function FindWindowEx Lib "user32.dll" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Private Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Private Declare Function GetMenuItemInfo Lib "user32" Alias "GetMenuItemInfoA" (ByVal hMenu As Long, ByVal un As Long, ByVal b As Boolean, lpmii As MENUITEMINFO) As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Private Declare Function GetWindowRect Lib "user32.dll" (ByVal hwnd As Long, ByRef lpRect As RECT) As Long
Private Declare Function InsertMenuItem Lib "user32" Alias "InsertMenuItemW" (ByVal hMenu As Long, ByVal un As Long, ByVal bool As Boolean, ByRef lpcMenuItemInfo As MENUITEMINFO) As Long
Private Declare Function ModifyMenu Lib "user32" Alias "ModifyMenuW" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpString As Long) As Long
Private Declare Function SetMenuItemInfo Lib "user32" Alias "SetMenuItemInfoW" (ByVal hMenu As Long, ByVal uItem As Long, ByVal fByPosition As Long, lpmii As MENUITEMINFO) As Long
Private Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Private Declare Function SetTextAlign Lib "gdi32.dll" (ByVal hdc As Long, ByVal wFlags As Long) As Long
Private Declare Function TextOut Lib "gdi32.dll" Alias "TextOutW" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal lpString As Long, ByVal nCount As Long) As Long
Private Declare Function TrackPopupMenu Lib "user32" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal X As Long, ByVal Y As Long, ByVal nReserved As Long, ByVal hwnd As Long, ByVal lprc As Any) As Long
Private Declare Function WindowFromPoint Lib "user32.dll" (ByVal xPoint As Long, ByVal yPoint As Long) As Long

'Control Events
Public Event Click()
Public Event DblClick()
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

'Button Events
Public Event ButtonChange(ButtonIndex As Integer)
Public Event ButtonClick(ButtonIndex As Integer)
Public Event ButtonDblClick(ButtonIndex As Integer)
Public Event ButtonMouseDown(ButtonIndex As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event ButtonMouseMove(ButtonIndex As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event ButtonMouseUp(ButtonIndex As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

'Option Button Events
Public Event OptionClick(ButtonIndex As Integer, MenuID As Long)
Public Event OptionDblClick(ButtonIndex As Integer)
Public Event OptionMouseDown(ButtonIndex As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event OptionMouseMove(ButtonIndex As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event OptionMouseUp(ButtonIndex As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

'Left Arrow Button Events
Public Event LArrowClick()
Public Event LArrowDblClick()
Public Event LArrowMouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event LArrowMouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event LArrowMouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

'Left Arrow Button Events
Public Event RArrowClick()
Public Event RArrowDblClick()
Public Event RArrowMouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event RArrowMouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event RArrowMouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

'Drop Down Box Events
Public Event DDBClick()
Public Event DDBDblClick()
Public Event DDBMouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event DDBMouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event DDBMouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

'Const
Const TA_BASELINE = 24
Const TA_BOTTOM = 8
Const TA_CENTER = 6
Const TA_LEFT = 0
Const TA_NOUPDATECP = 0
Const TA_RIGHT = 2
Const TA_TOP = 0
Const TA_UPDATECP = 1
Const TA_MASK = (TA_BASELINE + TA_CENTER + TA_UPDATECP)
Const TPM_RETURNCMD = &H100&
Const MF_SEPARATOR = &H800&
Const MF_STRING = &H0&
Const MF_POPUP = &H10&
Const MF_CHECKED = 8
Const MF_UNCHECKED = &H0&
Const MF_GRAYED As Long = &H1&
Const MIIM_TYPE = &H10
Const MIIM_SUBMENU = &H4

'Type
Private Type POINTAPI
    X As Long
    Y As Long
End Type

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Type MENUITEMINFO
    cbSize As Long
    fMask As Long
    fType As Long
    fState As Long
    wID As Long
    hSubMenu As Long
    hbmpChecked As Long
    hbmpUnchecked As Long
    dwItemData As Long
    dwTypeData As Long
    cch As Long
End Type

'For Control
Private bUnicode As Boolean
Private iButtonCount As Integer
Private bArrowButton As Boolean
Private bDropDownBox As Boolean
Private sDropDownBoxCaption As String
Private iDropDownBoxWidth As Integer
Private iIndexButtonSelected As Integer
Private iSeparatorCount As Integer
Private bLeftArrowEnabled As Boolean
Private bRightArrowEnabled As Boolean
Private lFirstButtonLeft As Long

'Default Value
Const DEF_ENABLED As Boolean = True
Const DEF_UNICODE As Boolean = True
Const DEF_BUTTONCOUNT As Integer = 0
Const DEF_ARROWBUTTON As Boolean = False
Const DEF_DROPDOWNBOX As Boolean = False
Const DEF_DROPDOWNBOXCAPTION As String = "Drop Down Box"
Const DEF_DROPDOWNBOXWIDTH As Integer = 2000
Const DEF_INDEXBUTTONSELECTED As Integer = -1
Const DEF_SEPARATORCOUNT As Integer = 0
Const DEF_LEFTARROWENABLED As Boolean = False
Const DEF_RIGHTARROWENABLED As Boolean = False
Const DEF_FIRSTBUTTONLEFT As Long = 0

'For each Button
Private sButtonCaption() As String
Private bButtonEnable() As Boolean
Private bButtonOption() As Boolean
Private iSeparator() As Integer
Private lButtonMenuID() As Long
Private lButtonMenuHandle() As Long

'For each Menu
    'Parent Menu
Private iParentMenuCount As Integer
Private lParentMenuID() As Long
Private lParentMenuHandle() As Long
    'Sub Menu
Private iMenuCount As Integer
Private lMenuID() As Long
Private lMyParentMenuHandle() As Long
Private sMenuCaption() As String
Private iMenuIndex() As Integer

Private Sub imgBackGround_Click()
    RaiseEvent Click
End Sub

Private Sub imgBackGround_DragDrop(Source As Control, X As Single, Y As Single)
    RaiseEvent DblClick
End Sub

Private Sub imgBackGround_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub imgBackGround_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
    ResetAllCurrentState
End Sub

Private Sub ResetAllCurrentState()
    Dim i As Integer, ss_Tmp  As String
    For i = 0 To iButtonCount - 1
        If i = iIndexButtonSelected Then
            DrawButton picButton(i), 0, 1
            DrawOption picOption(i), 0, 1
        Else
            DrawButton picButton(i), 0, 0
            DrawOption picOption(i), 0, 0
        End If
        If bUnicode Then ss_Tmp = ToUni(sButtonCaption(i)) Else ss_Tmp = sButtonCaption(i)
        SetTextAlign picButton(i).hdc, TA_TOP Or TA_CENTER Or TA_NOUPDATECP
        TextOut picButton(i).hdc, picButton(i).Width / (2 * Screen.TwipsPerPixelX), 8, StrPtr(ss_Tmp), Len(ss_Tmp)
    Next i
    With picArrowLeft
        If bLeftArrowEnabled Then
            .PaintPicture imgArrowLeft.Picture, 0, 0, .Width, .Height, 0, .Height, .Width, .Height
        Else
            .PaintPicture imgArrowLeft.Picture, 0, 0, .Width, .Height, 0, 0, .Width, .Height
        End If
    End With
    With picArrowRight
        If bRightArrowEnabled Then
            .PaintPicture imgArrowRight.Picture, 0, 0, .Width, .Height, 0, .Height, .Width, .Height
        Else
            .PaintPicture imgArrowRight.Picture, 0, 0, .Width, .Height, 0, 0, .Width, .Height
        End If
    End With
    If bDropDownBox Then
        Dim sString  As String
        DrawDropDownBox iDropDownBoxWidth, 0
        SetTextAlign picDropDownBox.hdc, TA_LEFT Or TA_TOP Or TA_NOUPDATECP
        sString = IIf(bUnicode, ToUni(sDropDownBoxCaption), sDropDownBoxCaption)
        TextOut picDropDownBox.hdc, 10, 5, StrPtr(sString), Len(sString)
        If iButtonCount > 0 Then
            If picDropDownBox.Left < (picButton(iButtonCount - 1).Left + picButton(iButtonCount - 1).Width) Then
                picDropDownBox.Visible = False
            Else
                picDropDownBox.Visible = True
            End If
        End If
    End If
End Sub

Private Sub imgBackGround_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Private Sub picArrowLeft_Click()
    If bLeftArrowEnabled Then
        RaiseEvent LArrowClick
        ResetAllCurrentState
    End If
End Sub

Private Sub picArrowLeft_DblClick()
    If bLeftArrowEnabled Then
        RaiseEvent LArrowDblClick
        ResetAllCurrentState
    End If
End Sub

Private Sub picArrowLeft_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    With picArrowLeft
        If bLeftArrowEnabled Then
            RaiseEvent LArrowMouseDown(Button, Shift, X + picArrowLeft.Left, Y + picArrowLeft.Top)
            .PaintPicture imgArrowLeft.Picture, 0, 0, .Width, .Height, 0, 3 * .Height, .Width, .Height
        Else
            .PaintPicture imgArrowLeft.Picture, 0, 0, .Width, .Height, 0, 0, .Width, .Height
        End If
    End With
End Sub

Private Sub picArrowLeft_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    With picArrowLeft
        If bLeftArrowEnabled Then
            RaiseEvent LArrowMouseMove(Button, Shift, X + picArrowLeft.Left, Y + picArrowLeft.Top)
            .PaintPicture imgArrowLeft.Picture, 0, 0, .Width, .Height, 0, 2 * .Height, .Width, .Height
        Else
            .PaintPicture imgArrowLeft.Picture, 0, 0, .Width, .Height, 0, 0, .Width, .Height
        End If
    End With
    With picArrowRight
        If bRightArrowEnabled Then
            .PaintPicture imgArrowRight.Picture, 0, 0, .Width, .Height, 0, .Height, .Width, .Height
        Else
            .PaintPicture imgArrowRight.Picture, 0, 0, .Width, .Height, 0, 0, .Width, .Height
        End If
    End With
End Sub

Private Sub picArrowLeft_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent LArrowMouseUp(Button, Shift, X + picArrowLeft.Left, Y + picArrowLeft.Top)
End Sub

Private Sub picArrowRight_Click()
    If bRightArrowEnabled Then
        RaiseEvent RArrowClick
        ResetAllCurrentState
    End If
End Sub

Private Sub picArrowRight_DblClick()
    If bRightArrowEnabled Then
        RaiseEvent RArrowDblClick
        ResetAllCurrentState
    End If
End Sub

Private Sub picArrowRight_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    With picArrowRight
        If bRightArrowEnabled Then
            .PaintPicture imgArrowRight.Picture, 0, 0, .Width, .Height, 0, 3 * .Height, .Width, .Height
            RaiseEvent RArrowMouseDown(Button, Shift, X + picArrowLeft.Left, Y + picArrowLeft.Top)
        Else
            .PaintPicture imgArrowRight.Picture, 0, 0, .Width, .Height, 0, 0, .Width, .Height
        End If
    End With
End Sub

Private Sub picArrowRight_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    With picArrowRight
        If bRightArrowEnabled Then
            RaiseEvent RArrowMouseMove(Button, Shift, X + picArrowLeft.Left, Y + picArrowLeft.Top)
            .PaintPicture imgArrowRight.Picture, 0, 0, .Width, .Height, 0, 2 * .Height, .Width, .Height
        Else
            .PaintPicture imgArrowRight.Picture, 0, 0, .Width, .Height, 0, 0, .Width, .Height
        End If
    End With
    With picArrowLeft
        If bLeftArrowEnabled Then
            .PaintPicture imgArrowLeft.Picture, 0, 0, .Width, .Height, 0, .Height, .Width, .Height
        Else
            .PaintPicture imgArrowLeft.Picture, 0, 0, .Width, .Height, 0, 0, .Width, .Height
        End If
    End With
End Sub

Private Sub picArrowRight_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent RArrowMouseUp(Button, Shift, X + picArrowLeft.Left, Y + picArrowLeft.Top)
End Sub

Private Sub picButton_Click(Index As Integer)
    If bButtonEnable(Index) Then
        iIndexButtonSelected = Index
        RaiseEvent ButtonClick(Index)
        ResetAllCurrentState
    End If
End Sub

Private Sub picButton_DblClick(Index As Integer)
    If bButtonEnable(Index) Then
        iIndexButtonSelected = Index
        RaiseEvent ButtonDblClick(Index)
        ResetAllCurrentState
    End If
End Sub

Private Sub picButton_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ss_Tmp As String
    If bButtonEnable(Index) Then
        RaiseEvent ButtonMouseDown(Index, Button, Shift, X + picButton(Index).Left, Y + picButton(Index).Top)
        DrawButton picButton(Index), 2, IIf(iIndexButtonSelected <> Index, 0, 1)
        If bButtonOption(Index) Then DrawOption picOption(Index), IIf(iIndexButtonSelected <> Index, 1, 0), IIf(iIndexButtonSelected <> Index, 0, 1)
    Else
        DrawButton picButton(Index), 0, 0
        If bButtonOption(Index) Then DrawOption picOption(Index), 0, 0
    End If
    If bUnicode Then ss_Tmp = ToUni(sButtonCaption(Index)) Else ss_Tmp = sButtonCaption(Index)
    SetTextAlign picButton(Index).hdc, TA_TOP Or TA_CENTER Or TA_NOUPDATECP
    TextOut picButton(Index).hdc, picButton(Index).Width / (2 * Screen.TwipsPerPixelX), 8, StrPtr(ss_Tmp), Len(ss_Tmp)
End Sub

Private Sub picButton_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ss_Tmp As String
    ResetAllCurrentState
    If bButtonEnable(Index) Then
        RaiseEvent ButtonMouseMove(Index, Button, Shift, X + picButton(Index).Left, Y + picButton(Index).Top)
        If Button = 0 Then
            DrawButton picButton(Index), 1, IIf(iIndexButtonSelected <> Index, 0, 1)
            If bButtonOption(Index) Then DrawOption picOption(Index), IIf(iIndexButtonSelected <> Index, 1, 0), IIf(iIndexButtonSelected <> Index, 0, 1)
        Else
            DrawButton picButton(Index), 2, IIf(iIndexButtonSelected <> Index, 0, 1)
            If bButtonOption(Index) Then DrawOption picOption(Index), IIf(iIndexButtonSelected <> Index, 1, 0), IIf(iIndexButtonSelected <> Index, 0, 1)
        End If
    Else
        DrawButton picButton(Index), 0, 0
        If bButtonOption(Index) Then DrawOption picOption(Index), 0, 0
    End If
    If bUnicode Then ss_Tmp = ToUni(sButtonCaption(Index)) Else ss_Tmp = sButtonCaption(Index)
    SetTextAlign picButton(Index).hdc, TA_TOP Or TA_CENTER Or TA_NOUPDATECP
    TextOut picButton(Index).hdc, picButton(Index).Width / (2 * Screen.TwipsPerPixelX), 8, StrPtr(ss_Tmp), Len(ss_Tmp)
End Sub

Private Sub picButton_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If bButtonEnable(Index) Then RaiseEvent ButtonMouseUp(Index, Button, Shift, X + picButton(Index).Left, Y + picButton(Index).Top)
End Sub

Private Sub picDropDownBox_Click()
    RaiseEvent DDBClick
    ResetAllCurrentState
End Sub

Private Sub picDropDownBox_DblClick()
    RaiseEvent DDBDblClick
    ResetAllCurrentState
End Sub

Private Sub picDropDownBox_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim sString As String
    RaiseEvent DDBMouseDown(Button, Shift, X + picDropDownBox.Left, Y + picDropDownBox.Top)
    DrawDropDownBox iDropDownBoxWidth, 2
    SetTextAlign picDropDownBox.hdc, TA_LEFT Or TA_TOP Or TA_NOUPDATECP
    sString = IIf(bUnicode, ToUni(sDropDownBoxCaption), sDropDownBoxCaption)
    TextOut picDropDownBox.hdc, 10, 5, StrPtr(sString), Len(sString)
End Sub

Private Sub picDropDownBox_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim sString As String
    RaiseEvent DDBMouseMove(Button, Shift, X + picDropDownBox.Left, Y + picDropDownBox.Top)
    If Button = 0 Then
        DrawDropDownBox iDropDownBoxWidth, 1
    Else
        DrawDropDownBox iDropDownBoxWidth, 2
    End If
    SetTextAlign picDropDownBox.hdc, TA_LEFT Or TA_TOP Or TA_NOUPDATECP
    sString = IIf(bUnicode, ToUni(sDropDownBoxCaption), sDropDownBoxCaption)
    TextOut picDropDownBox.hdc, 10, 5, StrPtr(sString), Len(sString)
End Sub

Private Function DrawDropDownBox(Optional iWidth As Integer = 2000, Optional iMouse As Integer = 0)
    Dim sString As String
    Dim iBGWidth As Single
    If iWidth < 1500 Then iWidth = 2000
    iBGWidth = iWidth - 405
    With picDropDownBox
        .Top = 90
        .Left = UserControl.Width - iWidth - 120
        .Width = iWidth
        .PaintPicture imgDropDownBoxLeft.Picture, 0, 0, 210, 345, 0, iMouse * 345, 210, 345
        .PaintPicture imgDropDownBoxBG.Picture, 210, 0, iBGWidth, 345, 0, iMouse * 345, 20, 345
        .PaintPicture imgDropDownBoxArrow.Picture, iBGWidth + 210, 0, 195, 345, 0, iMouse * 345, 195, 345
    End With
End Function

Private Sub picDropDownBox_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent DDBMouseUp(Button, Shift, X + picDropDownBox.Left, Y + picDropDownBox.Top)
End Sub

Private Sub picOption_Click(Index As Integer)
    If bButtonEnable(Index) Then
        Dim lReturnMenuID As Long
        Dim R As RECT
        If lButtonMenuHandle(Index) > 0 Then
            GetWindowRect picOption(Index).hwnd, R
            lReturnMenuID = TrackPopupMenu(lButtonMenuHandle(Index), TPM_RETURNCMD, R.Left, R.Top + (picOption(Index).Height / Screen.TwipsPerPixelY), 0, UserControl.hwnd, 0&)
        End If
        RaiseEvent OptionClick(Index, lReturnMenuID)
        ResetAllCurrentState
    End If
End Sub

Private Sub picOption_DblClick(Index As Integer)
    If bButtonEnable(Index) Then
        RaiseEvent OptionDblClick(Index)
        ResetAllCurrentState
    End If
End Sub

Private Sub picOption_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ss_Tmp As String
    If bButtonEnable(Index) Then
        RaiseEvent OptionMouseDown(Index, Button, Shift, X + picOption(Index).Left, Y + picOption(Index).Top)
        If iIndexButtonSelected = Index Then
            DrawButton picButton(Index), 0, 1
            DrawOption picOption(Index), 2, 1
        Else
            DrawButton picButton(Index), 1, 2
            DrawOption picOption(Index), 3, 2
        End If
    Else
        DrawButton picButton(Index), 0, 0
        DrawOption picOption(Index), 0, 0
    End If
    If bUnicode Then ss_Tmp = ToUni(sButtonCaption(Index)) Else ss_Tmp = sButtonCaption(Index)
    SetTextAlign picButton(Index).hdc, TA_TOP Or TA_CENTER Or TA_NOUPDATECP
    TextOut picButton(Index).hdc, picButton(Index).Width / (2 * Screen.TwipsPerPixelX), 8, StrPtr(ss_Tmp), Len(ss_Tmp)
End Sub

Private Sub picOption_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim i As Integer, ss_Tmp As String
    ResetAllCurrentState
    If bButtonEnable(Index) Then
        RaiseEvent OptionMouseMove(Index, Button, Shift, X + picOption(Index).Left, Y + picOption(Index).Top)
        If Button = 0 Then
            If iIndexButtonSelected = Index Then
                DrawButton picButton(Index), 0, 1
                DrawOption picOption(Index), 1, 1
            Else
                DrawButton picButton(Index), 1, 2
                DrawOption picOption(Index), 2, 2
            End If
        Else
            If iIndexButtonSelected = Index Then
                DrawButton picButton(Index), 0, 1
                DrawOption picOption(Index), 2, 1
            Else
                DrawButton picButton(Index), 1, 2
                DrawOption picOption(Index), 3, 2
            End If
        End If
    Else
        DrawButton picButton(Index), 0, 0
        DrawOption picOption(Index), 0, 0
    End If
    If bUnicode Then ss_Tmp = ToUni(sButtonCaption(Index)) Else ss_Tmp = sButtonCaption(Index)
    SetTextAlign picButton(Index).hdc, TA_TOP Or TA_CENTER Or TA_NOUPDATECP
    TextOut picButton(Index).hdc, picButton(Index).Width / (2 * Screen.TwipsPerPixelX), 8, StrPtr(ss_Tmp), Len(ss_Tmp)
End Sub

Private Sub picOption_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If bButtonEnable(Index) Then RaiseEvent OptionMouseUp(Index, Button, Shift, X + picOption(Index).Left, Y + picOption(Index).Top)
End Sub

Private Sub tmrMoveOff_Timer()
    Dim lpPoint As POINTAPI, i As Integer
    Dim bIsUserControl As Boolean, bIsButton As Boolean, bIsOption As Boolean, bIsArrow As Boolean, bIsDDB As Boolean
    Dim bIsBG As Boolean
    Dim hMyWnd As Long
    GetCursorPos lpPoint
    hMyWnd = WindowFromPoint(lpPoint.X, lpPoint.Y)
    bIsUserControl = IIf(hMyWnd = UserControl.hwnd, True, False)
    bIsButton = IIf(hMyWnd = picButton(0).hwnd, True, False)
    For i = 1 To iButtonCount - 1
        bIsButton = bIsButton Or IIf(hMyWnd = picButton(i).hwnd, True, False)
    Next i
    bIsOption = IIf(hMyWnd = picOption(0).hwnd, True, False)
    For i = 1 To iButtonCount - 1
        bIsOption = bIsOption Or IIf(hMyWnd = picOption(i).hwnd, True, False)
    Next i
    bIsArrow = IIf(hMyWnd = picArrowLeft.hwnd, True, False)
    bIsArrow = bIsArrow Or IIf(hMyWnd = picArrowRight.hwnd, True, False)
    bIsDDB = IIf(hMyWnd = picDropDownBox.hwnd, True, False)
    If Not (bIsUserControl Or bIsButton Or bIsOption Or bIsArrow Or bIsDDB) Then
        ResetAllCurrentState
    End If
End Sub

Private Sub UserControl_Initialize()
    ReDim sButtonCaption(1) As String
    ReDim bButtonEnable(1) As Boolean
    ReDim bButtonOption(1) As Boolean
    ReDim iSeparator(1) As Integer
    iIndexButtonSelected = -1
    iButtonCount = 0
    iSeparatorCount = 0
    iSeparator(0) = 0
    bButtonEnable(0) = True
    bButtonOption(0) = False
    sButtonCaption(0) = "Button 0"
    bUnicode = True
    bArrowButton = False
    lFirstButtonLeft = 0
    picOption(0).Left = picButton(0).Left + lFirstButtonLeft
    picOption(0).Width = picButton(0).Width
    picOption(0).Top = 405
    picOption(0).Height = 150
    picOption(0).Visible = False
    iMenuCount = 0
    iParentMenuCount = 0
    With picArrowLeft
        .Left = 90
        .Top = 30
        .PaintPicture imgArrowLeft.Picture, 0, 0, .Width, .Height, 0, 0, .Width, .Height
    End With
    With picArrowRight
        .Left = 540
        .Top = 30
        .PaintPicture imgArrowRight.Picture, 0, 0, .Width, .Height, 0, 0, .Width, .Height
    End With
    bLeftArrowEnabled = False
    bRightArrowEnabled = False
    picButton(0).Visible = False
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    With PropBag
        UserControl.Enabled = .ReadProperty("Enabled", DEF_ENABLED)
        bUnicode = .ReadProperty("AutoUnicode", DEF_UNICODE)
        bArrowButton = .ReadProperty("ArrowButton", DEF_ARROWBUTTON)
        bLeftArrowEnabled = .ReadProperty("LeftArrowEnabled", DEF_LEFTARROWENABLED)
        bRightArrowEnabled = .ReadProperty("RightArrowEnabled", DEF_RIGHTARROWENABLED)
        iIndexButtonSelected = .ReadProperty("ActiveButton", DEF_INDEXBUTTONSELECTED)
        bDropDownBox = .ReadProperty("DropDownBox", DEF_DROPDOWNBOX)
        sDropDownBoxCaption = .ReadProperty("DDBCaption", DEF_DROPDOWNBOXCAPTION)
        iDropDownBoxWidth = .ReadProperty("DDBWidth", DEF_DROPDOWNBOXWIDTH)
        lFirstButtonLeft = .ReadProperty("FirstButtonLeft", DEF_FIRSTBUTTONLEFT)
        DrawDropDownBox iDropDownBoxWidth
        picArrowLeft.Visible = bArrowButton
        picArrowRight.Visible = bArrowButton
        picButton(0).Left = IIf(bArrowButton, 1080, 60) + lFirstButtonLeft
        picDropDownBox.Visible = bDropDownBox
        If bDropDownBox Then DrawDropDownBox iDropDownBoxWidth
        picOption(0).Left = picButton(0).Left + lFirstButtonLeft
        ResetAllCurrentState
    End With
End Sub

Private Sub UserControl_Resize()
    With UserControl
        .Height = 555
        imgRight.Left = .Width - 60
        imgBackGround.Width = imgRight.Left - 60
        picDropDownBox.Left = .Width - picDropDownBox.Width - 120
    End With
End Sub

Public Function AddButton(Optional sBtnCaption As String = "", Optional bBtnOption As Boolean = False, Optional bBtnEnable As Boolean = True, Optional iBtnWidth As Integer, Optional MenuID As Long = -1)
Attribute AddButton.VB_Description = "Add new a button on your Windows Media Player 11 Tool Bar"
    Dim sf_strTmp As String, i%
    tmrMoveOff.Interval = 100
    ReDim Preserve sButtonCaption(iButtonCount)
    ReDim Preserve bButtonOption(iButtonCount)
    ReDim Preserve bButtonEnable(iButtonCount)
    ReDim Preserve iSeparator(iButtonCount)
    ReDim Preserve lButtonMenuID(iButtonCount)
    ReDim Preserve lButtonMenuHandle(iButtonCount)
    
    If Len(sBtnCaption) = 0 Then sBtnCaption = "Button " & iButtonCount
    
    sButtonCaption(iButtonCount) = sBtnCaption
    bButtonOption(iButtonCount) = bBtnOption
    bButtonEnable(iButtonCount) = bBtnEnable
    
    If bUnicode Then sf_strTmp = ToUni(sBtnCaption) Else sf_strTmp = sBtnCaption
    If iButtonCount > 0 Then Load picButton(iButtonCount)
    With picButton(iButtonCount)
        .Top = 60
        If iButtonCount > 0 Then .Left = picButton(iButtonCount - 1).Left + picButton(iButtonCount - 1).Width + IIf(iSeparator(iButtonCount - 1) > 0, 105, 0)
        .Height = IIf(bBtnOption, 345, 495)
        .Width = IIf(iBtnWidth = 0, IIf(TextWidth(sf_strTmp) < 1200, 1200, TextWidth(sf_strTmp) + 240 * 2), iBtnWidth)
        .Visible = True
    End With
    DrawButton picButton(iButtonCount), 0, 0 + IIf(bBtnOption, 2, 0)
    If iButtonCount > 0 Then Load picOption(iButtonCount)
    With picOption(iButtonCount)
        .Top = 405
        .Left = picButton(iButtonCount).Left
        .Width = picButton(iButtonCount).Width
        If bBtnOption Then
            .Visible = True
            DrawOption picOption(iButtonCount), 0, 0
        Else
            .Visible = False
        End If
    End With
    If bBtnEnable Then
        picButton(iButtonCount).ForeColor = vbWhite
    Else
        picButton(iButtonCount).ForeColor = vbGrayText
    End If
    SetTextAlign picButton(iButtonCount).hdc, TA_TOP Or TA_CENTER Or TA_NOUPDATECP
    TextOut picButton(iButtonCount).hdc, picButton(iButtonCount).Width / (2 * Screen.TwipsPerPixelX), 8, StrPtr(sf_strTmp), Len(sf_strTmp)
    
    Dim lPMH&
    lButtonMenuID(iButtonCount) = MenuID
    lButtonMenuHandle(iButtonCount) = 0
    
    For i = 0 To iParentMenuCount - 1
        If lParentMenuID(i) = MenuID Then
            lPMH = lParentMenuHandle(i)
            lButtonMenuHandle(iButtonCount) = lPMH
            Exit For
        End If
    Next i
    
    iButtonCount = iButtonCount + 1
End Function

Public Function ModifyButton(iBtnIndex As Integer, Optional sBtnCaption As String, Optional bBtnOption As Boolean = False, Optional bBtnEnable As Boolean = False, Optional iBtnWidth As Integer) As Long
    If iBtnIndex < 0 Or iBtnIndex > iButtonCount - 1 Then ModifyButton = 3: Exit Function
    Dim sf_strTmp As String, i As Integer
    sButtonCaption(iBtnIndex) = sBtnCaption
    bButtonOption(iBtnIndex) = bBtnOption
    bButtonEnable(iBtnIndex) = bBtnEnable
    If Len(sBtnCaption) = 0 Then sBtnCaption = "Button " & iBtnIndex
    If bUnicode Then sf_strTmp = ToUni(sBtnCaption) Else sf_strTmp = sBtnCaption
    picButton(iBtnIndex).Width = IIf(iBtnWidth = 0, IIf(TextWidth(sf_strTmp) < 1200, 1200, TextWidth(sf_strTmp) + 240 * 2), iBtnWidth)
    picOption(iBtnIndex).Visible = bBtnOption
    picButton(iBtnIndex).Height = IIf(bBtnOption, 345, 495)
    picOption(iBtnIndex).Width = picButton(iBtnIndex).Width
    If bBtnEnable Then
        picButton(iBtnIndex).ForeColor = vbWhite
    Else
        picButton(iBtnIndex).ForeColor = vbGrayText
    End If
    SetTextAlign picButton(iBtnIndex).hdc, TA_TOP Or TA_CENTER Or TA_NOUPDATECP
    TextOut picButton(iBtnIndex).hdc, picButton(iBtnIndex).Width / (2 * Screen.TwipsPerPixelX), 8, StrPtr(sf_strTmp), Len(sf_strTmp)
    If iIndexButtonSelected = iBtnIndex Then
        If bBtnEnable Then
            iIndexButtonSelected = iBtnIndex
        Else
            iIndexButtonSelected = -1
            For i = iBtnIndex To iButtonCount - 1
                If bButtonEnable(i) Then
                    iIndexButtonSelected = i
                    Exit For
                End If
            Next i
            For i = iBtnIndex - 1 To 0 Step -1
                If bButtonEnable(i) Then
                    iIndexButtonSelected = i
                    Exit For
                End If
            Next i
        End If
    End If
    If iIndexButtonSelected >= 0 Then
        Call picButton_Click(iIndexButtonSelected)
    End If
    ResetAllCurrentState
    ResetPositionButton
End Function

Public Function AddSeparator() As Long
Attribute AddSeparator.VB_Description = "Add a separator after button."
    If iButtonCount > 0 Then
        If iSeparatorCount > 0 Then Load imgSeparator(iSeparatorCount)
        imgSeparator(iSeparatorCount).Top = 0
        imgSeparator(iSeparatorCount).Left = picButton(iButtonCount - 1).Left + picButton(iButtonCount - 1).Width + 30
        imgSeparator(iSeparatorCount).Visible = True
        imgSeparator(iSeparatorCount).ZOrder vbBringToFront
        iSeparatorCount = iSeparatorCount + 1
        iSeparator(iButtonCount - 1) = iButtonCount
    Else
        AddSeparator = 1
    End If
End Function

Public Function hwnd() As Long
    hwnd = UserControl.hwnd
End Function

Public Function hButtonWnd(iBtnIndex As Integer) As Long
    hButtonWnd = picButton(iBtnIndex).hwnd
End Function

Public Function hOptionWnd(iBtnIndex) As Long
    hOptionWnd = picOption(iBtnIndex).hwnd
End Function

Private Sub ResetPositionButton()
    Dim iBtn As Integer, iC As Integer
    picButton(0).Left = IIf(bArrowButton, 1080 + lFirstButtonLeft, lFirstButtonLeft)
    picOption(0).Left = picButton(0).Left
    For iBtn = 1 To iButtonCount - 1
        picButton(iBtn).Left = picButton(iBtn - 1).Left + picButton(iBtn - 1).Width + IIf(iSeparator(iBtn - 1) > 0, 105, 0)
        picOption(iBtn).Left = picButton(iBtn - 1).Left + picButton(iBtn - 1).Width + IIf(iSeparator(iBtn - 1) > 0, 105, 0)
    Next iBtn
    For iBtn = 0 To iButtonCount - 1
        If iSeparator(iBtn) > 0 Then
            imgSeparator(iC).Left = picButton(iSeparator(iBtn) - 1).Left + picButton(iSeparator(iBtn) - 1).Width + 30
            iC = iC + 1
        End If
    Next iBtn
    ResetAllCurrentState
End Sub

Private Sub DrawButton(picB As PictureBox, Optional iButtonWhenMouse As Integer, Optional iStateButtonNow As Integer)
    Dim iH As Integer
    picB.Cls
    If iStateButtonNow < 0 Or iStateButtonNow > 3 Then iStateButtonNow = 0
    If iButtonWhenMouse > 3 Or iButtonWhenMouse < 0 Then iButtonWhenMouse = 0
    If iStateButtonNow > 1 Then iH = 345 Else iH = 495
    'iButtonWhenMouse : Normal = 0 - OverButUnsel = 1 - OverAndSel = 2 - Down = 3
    picB.PaintPicture picLeft(iStateButtonNow).Picture, 0, 0, 75, iH, 0, iButtonWhenMouse * iH, 75, iH
    picB.PaintPicture picBackButton(iStateButtonNow).Picture, 75, 0, picB.Width - 150, iH, 0, iButtonWhenMouse * iH, 840, iH
    picB.PaintPicture picRight(iStateButtonNow).Picture, picB.Width - 75, 0, 75, iH, 0, iButtonWhenMouse * iH, 75, iH
End Sub

Private Sub DrawOption(picB As PictureBox, Optional iButtonWhenMouse As Integer, Optional iStateButtonNow As Integer = 0)
    picB.Cls
    If iStateButtonNow <> 0 And iStateButtonNow <> 1 Then iStateButtonNow = 0
    If iButtonWhenMouse > 3 Or iButtonWhenMouse < 0 Then iButtonWhenMouse = 0
    'iButtonWhenMouse : Normal = 0 ; OverButUnsel = 1 ; OverAndSel = 2 ; Down = 3
    picB.PaintPicture picOptionLeft(iStateButtonNow).Picture, 0, 0, 75, 150, 0, iButtonWhenMouse * 150, 75, 150
    picB.PaintPicture picBackOption(iStateButtonNow).Picture, 75, 0, picB.Width - 150, 150, 0, iButtonWhenMouse * 150, 840, 150
    picB.PaintPicture picOptionRight(iStateButtonNow).Picture, picB.Width - 75, 0, 75, 150, 0, iButtonWhenMouse * 150, 75, 150
    If iButtonWhenMouse > 0 Or iStateButtonNow = 1 Then
        picB.PSet (picB.Width / 2, 45), vbWhite: picB.PSet (picB.Width / 2 - 15, 45), vbWhite: picB.PSet (picB.Width / 2 - 30, 45), vbWhite: picB.PSet (picB.Width / 2 + 15, 45), vbWhite: picB.PSet (picB.Width / 2 + 30, 45), vbWhite
        picB.PSet (picB.Width / 2, 60), vbWhite: picB.PSet (picB.Width / 2 - 15, 60), vbWhite: picB.PSet (picB.Width / 2 + 15, 60), vbWhite
        picB.PSet (picB.Width / 2, 75), vbWhite
    End If
End Sub

Private Sub UserControl_Terminate()
    Dim i%
    If iParentMenuCount > 0 Then
        For i = 0 To iParentMenuCount - 1
            DestroyMenu lParentMenuHandle(i)
        Next i
    End If
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    With PropBag
        Call .WriteProperty("Enabled", UserControl.Enabled, DEF_ENABLED)
        Call .WriteProperty("AutoUnicode", bUnicode, DEF_UNICODE)
        Call .WriteProperty("ArrowButton", bArrowButton, DEF_ARROWBUTTON)
        Call .WriteProperty("LeftArrowEnabled", bLeftArrowEnabled, DEF_LEFTARROWENABLED)
        Call .WriteProperty("RightArrowEnabled", bRightArrowEnabled, DEF_RIGHTARROWENABLED)
        Call .WriteProperty("ActiveButton", iIndexButtonSelected, DEF_INDEXBUTTONSELECTED)
        Call .WriteProperty("DropDownBox", bDropDownBox, DEF_DROPDOWNBOX)
        Call .WriteProperty("DDBCaption", sDropDownBoxCaption, DEF_DROPDOWNBOXCAPTION)
        Call .WriteProperty("DDBWidth", iDropDownBoxWidth, DEF_DROPDOWNBOXWIDTH)
        Call .WriteProperty("FirstButtonLeft", lFirstButtonLeft, DEF_FIRSTBUTTONLEFT)
    End With
End Sub

Public Property Get LeftArrowEnabled() As Boolean
Attribute LeftArrowEnabled.VB_Description = "Returns/sets a value that determines whether a left arrow can respond to user-generated events."
Attribute LeftArrowEnabled.VB_ProcData.VB_Invoke_Property = ";Buttons"
    LeftArrowEnabled = bLeftArrowEnabled
End Property

Public Property Let LeftArrowEnabled(ByVal NewEnabled As Boolean)
    bLeftArrowEnabled = NewEnabled
   With picArrowLeft
        .PaintPicture imgArrowLeft.Picture, 0, 0, .Width, .Height, 0, IIf(bLeftArrowEnabled, .Height, 0), .Width, .Height
    End With
    PropertyChanged "LeftArrowEnabled"
End Property

Public Property Get RightArrowEnabled() As Boolean
Attribute RightArrowEnabled.VB_Description = "Returns/sets a value that determines whether a right arrow can respond to user-generated events."
Attribute RightArrowEnabled.VB_ProcData.VB_Invoke_Property = ";Buttons"
    RightArrowEnabled = bRightArrowEnabled
End Property

Public Property Let RightArrowEnabled(ByVal NewEnabled As Boolean)
    bRightArrowEnabled = NewEnabled
    With picArrowRight
        .PaintPicture imgArrowRight.Picture, 0, 0, .Width, .Height, 0, IIf(bRightArrowEnabled, .Height, 0), .Width, .Height
    End With
    PropertyChanged "RightArrowEnabled"
End Property

Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal NewEnabled As Boolean)
    UserControl.Enabled = NewEnabled
    PropertyChanged "Enabled"
End Property

Public Property Get AutoUnicode() As Boolean
Attribute AutoUnicode.VB_Description = "Set buttons caption to Vietnamese Unicode or not."
Attribute AutoUnicode.VB_ProcData.VB_Invoke_Property = ";Misc"
    AutoUnicode = bUnicode
End Property

Public Property Let AutoUnicode(ByVal NewAutoUnicode As Boolean)
    bUnicode = NewAutoUnicode
    PropertyChanged "AutoUnicode"
End Property

Public Property Get ArrowButton() As Boolean
Attribute ArrowButton.VB_Description = "Shows/hides left arrow and right arrow beforce all buttons."
Attribute ArrowButton.VB_ProcData.VB_Invoke_Property = ";Buttons"
    ArrowButton = bArrowButton
End Property

Public Property Let ArrowButton(ByVal NewArrowButton As Boolean)
    bArrowButton = NewArrowButton
    picArrowLeft.Visible = bArrowButton
    picArrowRight.Visible = bArrowButton
    picButton(0).Left = IIf(bArrowButton, 1080, 60) + lFirstButtonLeft
    picOption(0).Left = picButton(0).Left + lFirstButtonLeft
    ResetPositionButton
    PropertyChanged "ArrowButton"
End Property

Public Property Get DropDownBox() As Boolean
    DropDownBox = bDropDownBox
End Property

Public Property Let DropDownBox(ByVal NewDropDownBox As Boolean)
    bDropDownBox = NewDropDownBox
    picDropDownBox.Visible = bDropDownBox
    If bDropDownBox Then DrawDropDownBox iDropDownBoxWidth
    PropertyChanged "DropDownBox"
    ResetAllCurrentState
End Property

Public Property Get DDBX() As Integer
    'DDBX = picDropDownBox.Left
End Property

Public Property Get DDBY() As Integer
    'DDBY = picDropDownBox.Top
End Property

Public Property Get DDBHeight() As Integer
    'DDBY = picDropDownBox.Height
End Property

Public Property Get DDBCaption() As String
    DDBCaption = sDropDownBoxCaption
End Property

Public Property Let DDBCaption(ByVal NewCaption As String)
    sDropDownBoxCaption = NewCaption
    DrawDropDownBox iDropDownBoxWidth
    PropertyChanged "DDBCaption"
End Property

Public Property Get DDBWidth() As String
    DDBWidth = iDropDownBoxWidth
End Property

Public Property Let DDBWidth(ByVal NewWidth As String)
    iDropDownBoxWidth = NewWidth
    PropertyChanged "DDBWidth"
End Property

Public Property Get ButtonCount() As Long
Attribute ButtonCount.VB_Description = "Return button quantity."
Attribute ButtonCount.VB_ProcData.VB_Invoke_Property = ";Buttons"
    ButtonCount = iButtonCount
End Property

Public Property Get SeparatorCount() As Long
Attribute SeparatorCount.VB_Description = "Return separator quantity."
    SeparatorCount = iSeparatorCount
End Property

Public Property Get ActiveButton() As Integer
Attribute ActiveButton.VB_Description = "Returns/sets a button as a selected button."
Attribute ActiveButton.VB_ProcData.VB_Invoke_Property = ";Buttons"
    ActiveButton = iIndexButtonSelected
End Property

Public Property Let ActiveButton(iNewIndex As Integer)
    If iNewIndex > iButtonCount - 1 Then iNewIndex = 0
    If iNewIndex < 0 Then iNewIndex = -1
    iIndexButtonSelected = iNewIndex
    If iNewIndex >= 0 Then
        If bButtonEnable(iNewIndex) Then Call picButton_Click(iNewIndex) Else iIndexButtonSelected = -1
    End If
    ResetAllCurrentState
    PropertyChanged "ActiveButton"
End Property

Public Property Get FirstButtonLeft() As Long
    FirstButtonLeft = lFirstButtonLeft
End Property

Public Property Let FirstButtonLeft(ByVal NewFirstButtonLeft As Long)
    lFirstButtonLeft = NewFirstButtonLeft
    ResetPositionButton
    PropertyChanged "FirstButtonLeft"
End Property

Public Property Get ButtonCaption(iBtnIndex As Integer) As String
    ButtonCaption = sButtonCaption(iBtnIndex)
End Property

Public Property Get ButtonEnabled(iBtnIndex As Integer) As Boolean
    ButtonEnabled = bButtonEnable(iBtnIndex)
End Property

Public Property Get ButtonOption(iBtnIndex As Integer) As Boolean
    ButtonOption = bButtonOption(iBtnIndex)
End Property

Public Property Get ButtonX(iBtnIndex As Integer) As Single
    ButtonX = picButton(iBtnIndex).Left
End Property

Public Property Get ButtonY(iBtnIndex As Integer) As Single
    ButtonY = picButton(iBtnIndex).Top
End Property

Public Property Get ButtonWidth(iBtnIndex As Integer) As Single
    ButtonWidth = picButton(iBtnIndex).Width
End Property

Public Property Get ButtonHeight(iBtnIndex As Integer) As Single
    ButtonHeight = picButton(iBtnIndex).Height
End Property

Public Property Get ButtonOptionHeight(iBtnIndex As Integer) As Single
    ButtonOptionHeight = picOption(iBtnIndex).Height
End Property

Public Function RemoveButton(ButtonIndex As Integer)
    Dim i As Integer
    For i = ButtonIndex To iButtonCount - 2
        picButton(i).Left = picButton(i + 1).Left
        picOption(i).Left = picButton(i).Left
        sButtonCaption(i) = sButtonCaption(i + 1)
        bButtonEnable(i) = bButtonEnable(i + 1)
        bButtonOption(i) = bButtonOption(i + 1)
    Next i
    If iButtonCount > 1 Then Unload picButton(iButtonCount - 1) Else picButton(0).Visible = False
    If iButtonCount > 1 Then Unload picOption(iButtonCount - 1) Else picOption(0).Visible = False
    iButtonCount = IIf(iButtonCount > 0, iButtonCount - 1, 0)
    ResetPositionButton
End Function

Public Function RemoveButtonFromCaption(Caption As String)
    Dim i As Integer
    For i = 0 To iButtonCount - 1
        If bUnicode Then
            If sButtonCaption(i) = ToUni(Caption) Then
                RemoveButton i
                Exit Function
            End If
        Else
            If sButtonCaption(i) = Caption Then
                RemoveButton i
                Exit Function
            End If
        End If
    Next i
End Function

Public Function RemoveSeparator(SeparatorIndex As Integer)
    Dim i As Integer, iC As Integer, iPos As Integer, a As String, b As String
    For i = SeparatorIndex To iSeparatorCount - 1
        If i = 0 Then imgSeparator(0).Visible = False Else Unload imgSeparator(i)
        iSeparatorCount = IIf(iSeparatorCount > 0, iSeparatorCount - 1, 0)
    Next i
    For i = 0 To iButtonCount - 1
        If iSeparator(i) > 0 Then iC = iC + 1
        If iC = SeparatorIndex + 1 Then iSeparator(i) = 0: iPos = i: Exit For
    Next i
    For i = iPos + 1 To iButtonCount - 1
        If iSeparator(i) > 0 Then
            If iSeparatorCount > 0 Then Load imgSeparator(iSeparatorCount)
            imgSeparator(iSeparatorCount).Top = 0
            imgSeparator(iSeparatorCount).Left = picButton(i - 1).Left + picButton(i - 1).Width + 30
            imgSeparator(iSeparatorCount).Visible = True
            imgSeparator(iSeparatorCount).ZOrder vbBringToFront
            iSeparatorCount = iSeparatorCount + 1
        End If
    Next i
    ResetPositionButton
End Function

Public Function RemoveAllSeparators() As Long
    Dim i As Integer
    For i = 0 To iSeparatorCount - 1
        If iSeparatorCount > 0 Then RemoveSeparator 0 Else Exit For
    Next i
End Function

Public Function RemoveAllButtons() As Long
    Dim i As Integer
    For i = 0 To iButtonCount - 1
        If iButtonCount > 0 Then RemoveButton 0 Else Exit For
    Next i
    RemoveAllSeparators
End Function

Public Function AddMenu(ParentMenuID As Long) As Long
    
    ReDim Preserve lParentMenuHandle(iParentMenuCount) As Long
    ReDim Preserve lParentMenuID(iParentMenuCount) As Long
    
    lParentMenuHandle(iParentMenuCount) = CreatePopupMenu()
    lParentMenuID(iParentMenuCount) = ParentMenuID
    
    AddMenu = lParentMenuHandle(iParentMenuCount)
    
    iParentMenuCount = iParentMenuCount + 1
End Function

Public Function AddSubMenu(ParentMenuID As Long, SubMenuID As Long, MenuCaption As String) As Long
    If iParentMenuCount = 0 Then Exit Function
    Dim i%, lPMH&

    For i = 0 To iParentMenuCount - 1
        If lParentMenuID(i) = ParentMenuID Then
            lPMH = lParentMenuHandle(i)
            Exit For
        End If
    Next

    ReDim Preserve lMenuID(iMenuCount) As Long
    ReDim Preserve sMenuCaption(iMenuCount) As String
    ReDim Preserve iMenuIndex(iMenuCount) As Integer
    ReDim Preserve lMyParentMenuHandle(iMenuCount) As Long
    
    lMenuID(iMenuCount) = SubMenuID
    sMenuCaption(iMenuCount) = MenuCaption
    iMenuIndex(iMenuCount) = iMenuCount
    lMyParentMenuHandle(iMenuCount) = lPMH
    
    AddSubMenu = AppendMenu(lPMH, IIf(MenuCaption = "-", MF_SEPARATOR, MF_STRING), lMenuID(iMenuCount), StrPtr(sMenuCaption(iMenuCount)))
    iMenuCount = iMenuCount + 1
End Function

Public Function CheckedMenu(SubMenuID As Long, Checked As Boolean) As Long
    Dim i%
    For i = 0 To iMenuCount - 1
        If lMenuID(i) = SubMenuID Then
            ModifyMenu lMyParentMenuHandle(i), SubMenuID, IIf(Checked, MF_CHECKED, MF_UNCHECKED), SubMenuID, StrPtr(sMenuCaption(i))
            CheckedMenu = 1
            Exit For
        End If
    Next i
End Function

Public Function EnabledMenu(SubMenuID As Long, Enabled As Boolean) As Long
    Dim i%
    For i = 0 To iMenuCount - 1
        If lMenuID(i) = SubMenuID Then
            ModifyMenu lMyParentMenuHandle(i), SubMenuID, IIf(Enabled, MF_STRING, MF_GRAYED), SubMenuID, StrPtr(sMenuCaption(i))
            EnabledMenu = 1
            Exit For
        End If
    Next i
End Function

Public Function ToUni(str$) As String
    Dim ANSI$, UNI$, i&, sTem$, sUni$, arrUNI() As String
    ANSI = "a1|a2|a3|a4|a5|a6|a8|a61a62a63a64a65a81a82a83a84a85A1|A2|A3|A4|A5|A6|A8|A61A62A63A64A65A81A82A83A84A85e1|e2|e3|e4|e5|e6|e61e62e63e64e65E1|E2|E3|E4|E5|E6|E61E62E63E64E65i1|i2|i3|i4|i5|I1|I2|I3|I4|I5|o1|o2|o3|o4|o5|o6|o7|o61o62o63o64o65o71o72o73o74o75O1|O2|O3|O4|O5|O6|O7|O61O62O63O64O65O71O72O73O74O75u1|u2|u3|u4|u5|u7|u71u72u73u74u75U1|U2|U3|U4|U5|U7|U71U72U73U74U75y1|y2|y3|y4|y5|Y1|Y2|Y3|Y4|Y5|d9|D9|"
    UNI = "E1,E0,1EA3,E3,1EA1,E2,103,1EA5,1EA7,1EA9,1EAB,1EAD,1EAF,1EB1,1EB3,1EB5,1EB7,C1,C0,1EA2,C3,1EA0,C2,102,1EA4,1EA6,1EA8,1EAA,1EAC,1EAE,1EB0,1EB2,1EB4,1EB6,E9,E8,1EBB,1EBD,1EB9,EA,1EBF,1EC1,1EC3,1EC5,1EC7,C9,C8,1EBA,1EBC,1EB8,CA,1EBE,1EC0,1EC2,1EC4,1EC6,ED,EC,1EC9,129,1ECB,CD,CC,1EC8,128,1ECA,F3,F2,1ECF,F5,1ECD,F4,1A1,1ED1,1ED3,1ED5,1ED7,1ED9,1EDB,1EDD,1EDF,1EE1,1EE3,D3,D2,1ECE,D5,1ECC,D4,1A0,1ED0,1ED2,1ED4,1ED6,1ED8,1EDA,1EDC,1EDE,1EE0,1EE2,FA,F9,1EE7,169,1EE5,1B0,1EE9,1EEB,1EED,1EEF,1EF1,DA,D9,1EE6,168,1EE4,1AF,1EE8,1EEA,1EEC,1EEE,1EF0,FD,1EF3,1EF7,1EF9,1EF5,DD,1EF2,1EF6,1EF8,1EF4,111,110"
    arrUNI = Split(UNI, ",")

    For i = 1 To Len(str)
        If IsNumeric(Mid(str, i + 1, 1)) = False Then
            sUni = sUni & Mid(str, i, 1)
        Else
            sTem = IIf(IsNumeric(Mid(str, i + 2, 1)), Mid(str, i, 3), Mid(str, i, 2))
            i = i + IIf(IsNumeric(Mid(str, i + 2, 1)), 2, 1)
            If InStr(ANSI, sTem) > 0 Then sTem = ChrW("&h" & arrUNI(InStr(ANSI, sTem) \ 3))
            sUni = sUni & sTem
        End If
    Next
    ToUni = sUni
End Function

Private Function ToANSI(str$) As String
    Dim ANSI$, UNI$, i&, sTem$, arrANSI() As String
    Dim lPos&
    ANSI = "a1,a2,a3,a4,a5,a6,a8,a61,a62,a63,a64,a65,a81,a82,a83,a84,a85,A1,A2,A3,A4,A5,A6,A8,A61,A62,A63,A64,A65,A81,A82,A83,A84,A85,e1,e2,e3,e4,e5,e6,e61,e62,e63,e64,e65,E1,E2,E3,E4,E5,E6,E61,E62,E63,E64,E65,i1,i2,i3,i4,i5,I1,I2,I3,I4,I5,o1,o2,o3,o4,o5,o6,o7,o61,o62,o63,o64,o65,o71,o72,o73,o74,o75,O1,O2,O3,O4,O5,O6,O7,O61,O62,O63,O64,O65,O71,O72,O73,O74,O75,u1,u2,u3,u4,u5,u7,u71,u72,u73,u74,u75,U1,U2,U3,U4,U5,U7,U71,U72,U73,U74,U75,y1,y2,y3,y4,y5,Y1,Y2,Y3,Y4,Y5,d9,D9"
    UNI = "E1|||E0|||1EA3|E3|||1EA1|E2|||103||1EA5|1EA7|1EA9|1EAB|1EAD|1EAF|1EB1|1EB3|1EB5|1EB7|C1|||C0|||1EA2|C3|||1EA0|C2|||102||1EA4|1EA6|1EA8|1EAA|1EAC|1EAE|1EB0|1EB2|1EB4|1EB6|E9|||E8|||1EBB|1EBD|1EB9|EA|||1EBF|1EC1|1EC3|1EC5|1EC7|C9|||C8|||1EBA|1EBC|1EB8|CA|||1EBE|1EC0|1EC2|1EC4|1EC6|ED|||EC|||1EC9|129||1ECB|CD|||CC|||1EC8|128||1ECA|F3|||F2|||1ECF|F5|||1ECD|F4|||1A1||1ED1|1ED3|1ED5|1ED7|1ED9|1EDB|1EDD|1EDF|1EE1|1EE3|D3|||D2|||1ECE|D5|||1ECC|D4|||1A0||1ED0|1ED2|1ED4|1ED6|1ED8|1EDA|1EDC|1EDE|1EE0|1EE2|FA|||F9|||1EE7|169||1EE5|1B0||1EE9|1EEB|1EED|1EEF|1EF1|DA|||D9|||1EE6|168||1EE4|1AF||1EE8|1EEA|1EEC|1EEE|1EF0|FD|||1EF3|1EF7|1EF9|1EF5|DD|||1EF2|1EF6|1EF8|1EF4|111||110||"
    arrANSI = Split(ANSI, ",")
    
    For i = 1 To Len(str)
        If InStr(1, UNI, Left(Hex(AscW(Mid(str, i, 1))) & "||||", 5)) > 0 Then
            sTem = sTem & arrANSI(InStr(1, UNI, Left(Hex(AscW(Mid(str, i, 1))) & "||||", 5)) \ 5)
        Else
            sTem = sTem & Mid(str, i, 1)
        End If
    Next
    ToANSI = sTem
End Function
