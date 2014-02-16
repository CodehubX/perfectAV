Attribute VB_Name = "modDialogOpen"
Option Explicit
Private Const WH_CBT = 5
Private Const WM_SETFONT = &H30
 
 
Public Declare Function CreateFont Lib "gdi32" Alias "CreateFontA" (ByVal H As Long, ByVal w As Long, ByVal E As Long, ByVal O As Long, ByVal w As Long, ByVal i As Long, ByVal U As Long, ByVal s As Long, ByVal C As Long, ByVal OP As Long, ByVal CP As Long, ByVal Q As Long, ByVal PAF As Long, ByVal f As String) As Long
Private Declare Function GetCurrentThreadId Lib "kernel32" () As Long
Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal ParenthWnd As Long, ByVal ChildhWnd As Long, ByVal ClassName As String, ByVal Caption As String) As Long
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
' Tim cua so
Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Declare Function SetWindowTextW Lib "user32" (ByVal hwnd As Long, ByVal lpString As Long) As Long
Public Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long
Private Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hMod As Long, ByVal dwThreadId As Long) As Long
 
 
 
Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias _
         "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
 
Private Type OPENFILENAME
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type
Private hHook As Long
Private Tit As String
Private fi As String
Private i As Long
Private di As String
Private Function OpenFile() As String
    Dim hOpenFile As OPENFILENAME
    Dim retval As Long
    With hOpenFile
        .lStructSize = Len(hOpenFile)
        .hwndOwner = frmMain.hwnd
        .hInstance = App.hInstance
        .lpstrFilter = fi
        .nFilterIndex = 1
        .lpstrFile = String(257, 0)
        .nMaxFile = Len(hOpenFile.lpstrFile) - 1
        .lpstrFileTitle = hOpenFile.lpstrFile
        .nMaxFileTitle = hOpenFile.nMaxFile
        .lpstrInitialDir = di
        .lpstrTitle = "CCCCC"
        .flags = &H2 Or &H8
        i = 0
    End With
    hHook = SetWindowsHookEx(WH_CBT, AddressOf OpenHookProc, App.hInstance, GetCurrentThreadId())
    retval = GetOpenFileName(hOpenFile)
    UnhookWindowsHookEx hHook
    If retval = 0 Then
        OpenFile = ""
    Else
        OpenFile = Trim(hOpenFile.lpstrFile)
    End If
End Function
Private Function OpenHookProc(ByVal ncode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
On Error Resume Next
Dim hButton As Long, hOp As Long, hFont As Long
 
 
If ncode <> 6 Then Exit Function
i = i + 1
If i > 5 Then UnhookWindowsHookEx hHook
hFont = CreateFont(13, 0, 0, 0, 500, 0, 0, 0, 0, 0, 0, 0, 0, "Tahoma")
 
hOp = FindWindow("#32770", "CCCCC")
SendMessage hOp, WM_SETFONT, hFont, 0
SetWindowTextW hOp, StrPtr(Tit)
 
hButton = FindWindowEx(hOp, 0&, "Button", "&Open")
SendMessage hButton, WM_SETFONT, hFont, 0
SetWindowTextW hButton, StrPtr(ToUnicode("&Mo73"))
 
hButton = FindWindowEx(hOp, 0&, "Button", "Cancel")
SendMessage hButton, WM_SETFONT, hFont, 0
SetWindowTextW hButton, StrPtr(ToUnicode("&Bo3 qua"))
 
hButton = FindWindowEx(hOp, 0&, "Button", "Open as &read-only")
SendMessage hButton, WM_SETFONT, hFont, 0
SetWindowTextW hButton, StrPtr(ToUnicode("Mo73 chi3 d9o5c"))
 
hButton = FindWindowEx(hOp, 0&, "Static", "Files of &type:")
SendMessage hButton, WM_SETFONT, hFont, 0
SetWindowTextW hButton, StrPtr(ToUnicode("&Kie63u ta65p tin :"))
 
hButton = FindWindowEx(hOp, 0&, "Static", "File &name:")
SendMessage hButton, WM_SETFONT, hFont, 0
SetWindowTextW hButton, StrPtr(ToUnicode("&Te6n ta65p tin :"))
   
hButton = FindWindowEx(hOp, 0&, "Static", "Look &in:")
SendMessage hButton, WM_SETFONT, hFont, 0
SetWindowTextW hButton, StrPtr(ToUnicode("Ti2m trong :"))
 
End Function
Public Function MoFile(FormOner As Form, Optional Title As String, Optional Filte As String, Optional Dir As String) As String
If Title <> "" Then
    Tit = ToUnicode(Title)
Else
    Tit = ToUnicode("Ha4y cho5n ta65p tin")
End If
If Filte = "" Then
    fi = Replace("All file (*.*)|*.*", "|", Chr(0)) & Chr(0)
Else
    fi = Replace(Filte, "|", Chr(0)) & Chr(0)
End If
di = Dir
MoFile = OpenFile
End Function


