Attribute VB_Name = "modFolderBrowser"
Private Type BrowseInfo
    hWndOwner As Long
    pIDLRoot As Long
    pszDisplayName As Long
    lpszTitle As Long
    ulFlags As Long
    lpfnCallback As Long
    lParam As Long
    iImage As Long
End Type
Private Const WM_SETFONT = &H30
Declare Function CreateFont Lib "gdi32" Alias "CreateFontA" (ByVal h As Long, ByVal w As Long, ByVal E As Long, ByVal O As Long, ByVal w As Long, ByVal i As Long, ByVal U As Long, ByVal s As Long, ByVal C As Long, ByVal OP As Long, ByVal CP As Long, ByVal Q As Long, ByVal PAF As Long, ByVal F As String) As Long
' Tao font Unicode cho tieng Viet
Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal ParenthWnd As Long, ByVal ChildhWnd As Long, ByVal ClassName As String, ByVal Caption As String) As Long
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
' Tim cua so
Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Declare Function SetWindowTextW Lib "user32" (ByVal hwnd As Long, ByVal lpString As Long) As Long
' Gui thong diep doi  font va doi van ban
Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal hMem As Long)
Private Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
Private Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseInfo) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long
' Cac ham API dung cho hop thoai Browse
Private DialogTit$
Private DialogTxt$
Private Han As Long
Private Function Address(ByVal Add As Long) As Long
Address = Add
End Function
Private Function Browse() As String
    Dim iNull As Integer, lpIDList As Long, lResult As Long
    Dim sPath As String, udtBI As BrowseInfo
    
    With udtBI
        .hWndOwner = Han
        .lpfnCallback = Address(AddressOf BrowseCallbackProc)
        .lpszTitle = lstrcat("CCC", "")
        .ulFlags = 1
    End With

    lpIDList = SHBrowseForFolder(udtBI)
    
    
    If lpIDList Then
        sPath = String$(256, 0)
        SHGetPathFromIDList lpIDList, sPath
        CoTaskMemFree lpIDList
        iNull = InStr(sPath, vbNullChar)
        If iNull Then
            sPath = Left$(sPath, iNull - 1)
        End If
    End If
    If sPath <> "" Then
        If Right(sPath, 1) <> "\" Then sPath = sPath & "\"
    End If
    Browse = sPath
End Function
Public Function BrowseCallbackProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal lParam As Long, ByVal lpdata As Long) As Long
On Error Resume Next
        hFont = CreateFont(13, 0, 0, 0, 300, 0, 0, 0, 0, 0, 0, 0, 0, "Tahoma")
       
        hButton = FindWindowEx(hwnd, 0&, "Button", "OK")
        SendMessage hButton, WM_SETFONT, hFont, 0
        SetWindowTextW hButton, StrPtr("Ch" & ChrW(7885) & "n")
       
        hButton = FindWindowEx(hwnd, 0&, "Button", "Cancel")
        SendMessage hButton, WM_SETFONT, hFont, 0
        SetWindowTextW hButton, StrPtr("B" & ChrW(7887) & " qua")
       
        hButton = FindWindowEx(hwnd, 0&, "static", "CCC")
        SendMessage hwnd, WM_SETFONT, hFont, 0
        SetWindowTextW hButton, StrPtr(DialogTxt)
       
        SendMessage hwnd, WM_SETFONT, hFont, 0
        SetWindowTextW hwnd, StrPtr(DialogTit)
End Function

Public Function ChonThuMuc(YouForm As Form) As String
DialogTit = "Select folder to scan"
DialogTxt = "Ch" & ChrW(7885) & "n th" & ChrW(432) & " m" & ChrW(7909) & "c c" & ChrW(7847) & "n qu" & ChrW(233) & "t, sau " & ChrW(273) & ChrW(243) & " nh" & ChrW(7845) & "n n" & ChrW(250) & "t " & ChrW(34) & "Ch" & ChrW(7885) & "n" & ChrW(34)
Han = YouForm.hwnd
ChonThuMuc = Browse()
End Function

