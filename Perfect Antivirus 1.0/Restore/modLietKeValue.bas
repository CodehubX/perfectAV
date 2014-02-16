Attribute VB_Name = "modLietKeValue"
Option Explicit
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, lpcbValueName As Long, ByVal lpReserved As Long, lpType As Long, lpData As Byte, lpcbData As Long) As Long
Private Const HKEY_LOCAL_MACHINE = &H80000002
Private Const HKEY_CURRENT_USER = &H80000001
Private Const KEY_ALL_ACCESS = &HF003F
Private Const REG_SZ = 1
Private Const REG_BINARY = 3                     ' Free form binary
Private Const REG_DWORD = 4                      ' 32-bit number
Private Const REG_EXPAND_SZ = 2                  ' Unicode nul terminated string
Private Const REG_MULTI_SZ = 7                   ' Multiple Unicode strings
Dim RetVal As Long
Dim hKey As Long
Dim NameKey As String
Dim lpType As Long
Dim LenName As Long
Dim Data(0 To 255) As Byte
Dim DataLen As Long
Dim DataString As String
Dim index As Long
Dim i As Long
Dim KetQua As String
Public xTotalStartUp
Public Function GetKeyValue(FullKeyName)
xTotalStartUp = 0
Dim Key1, Key2, i, Ua
Ua = 10
For i = 1 To Len(FullKeyName)
    If Mid(FullKeyName, i, 1) = "\" Then
        Ua = Ua + 10
        If Ua = 20 Then
            Key1 = Left(FullKeyName, i - 1)
            Key2 = Right(FullKeyName, Len(FullKeyName) - i)
        End If
    End If
Next i
'frmMain.Cls
If Key1 = "HKEY_LOCAL_MACHINE" Then
RetVal = RegOpenKeyEx(HKEY_LOCAL_MACHINE, Key2, 0, KEY_ALL_ACCESS, hKey)
ElseIf Key1 = "HKEY_CURRENT_USER" Then
RetVal = RegOpenKeyEx(HKEY_CURRENT_USER, Key2, 0, KEY_ALL_ACCESS, hKey)
End If

index = 0
Do While RetVal = 0
    NameKey = Space(255)
    DataString = Space(255)
    LenName = 255
    DataLen = 255
    RetVal = RegEnumValue(hKey, index, NameKey, LenName, ByVal 0, lpType, Data(0), DataLen)
    If RetVal = 0 Then
        NameKey = Left(NameKey, LenName) 'R�t b? kho?n tr?ng th?a
        DataString = ""
' X? l� th�ng tin theo ki?u c?a n� v� ??a v�o bi?n DataString
        Select Case lpType
             Case REG_SZ
                For i = 0 To DataLen - 1
                    DataString = DataString & Chr(Data(i)) ' N?i c�c ch? c�i th�nh chu?i
                Next
             Case REG_BINARY
                For i = 0 To DataLen - 1
                    Dim temp As String
                    temp = Hex(Data(i))
                    If Len(temp) < 2 Then temp = String(2 - Len(temp), "0") & temp
                    DataString = DataString & temp & " "
 ' N?i c�c c?p s? nh? ph�n l?i v?i nhau
                Next
            Case REG_DWORD
                For i = DataLen - 1 To 0 Step -1
                    DataString = DataString & Hex(Data(i)) 'N?i c�c s� hexa v?i nhau
                Next
            Case REG_MULTI_SZ
                For i = 0 To DataLen - 1
                    DataString = DataString & Chr(Data(i))
    'N?i c�c k� t? bao g?m k� t? vbNullChar (?? c�ch d�ng) th�nh m?t chu?i, b?n c� th? s? d?ng m?t m?ng g?m nhi?u string thay v� l� m?t
                Next
            Case REG_EXPAND_SZ
                For i = 0 To DataLen - 2
                    DataString = DataString & Chr(Data(i))
    'N?i c�c k� t? l?i v?i nhau, b? k� t? NULL cu?i c�ng
                Next
            Case Else
                DataString = " Khong xac dinh duoc !"
        ' Tr�n ?�y l� 5 ki?u c� tr�n WinXP
        End Select
    End If
    If Left(Left(NameKey, LenName), 1) <> " " Then
    '///////////////////
    'Form1.List1.AddItem DataString
    frmMain.lblStatus.Caption = DataString
    
    Dim AX As String
    xTotalStartUp = xTotalStartUp + 1
   AX = CheckVirus(DataString)
   If AX <> "No" Then
                Dim ia
                ia = frmMain.LVVirus1.ListItems.Count + 1
                frmMain.LVVirus1.ListItems.Add ia, , AX
                frmMain.LVVirus1.ListItems(ia).SubItems(1).Caption = DataString
                frmMain.LVVirus1.ListItems(ia).SubItems(2).Caption = FileLen(DataString) & " Bytes"
                frmMain.LVVirus1.ListItems(ia).SubItems(3).Caption = "---"
                frmMain.LVVirus1.ListItems(ia).SubItems(4).Caption = Key1 & "-" & Key2 & ":" & Left(NameKey, LenName)
                
                frmMain.LVVirus1.ListItems(ia).Checked = True
    End If
    
    
    '///////////////
    End If
    index = index + 1
    'frmMain.Print Left(NameKey, LenName) & "=" & DataString
Loop
RetVal = RegCloseKey(hKey)
End Function

Public Function GetKeyValue2(FullKeyName)
Dim Key1, Key2, i, Ua
Ua = 10
For i = 1 To Len(FullKeyName)
    If Mid(FullKeyName, i, 1) = "\" Then
        Ua = Ua + 10
        If Ua = 20 Then
            Key1 = Left(FullKeyName, i - 1)
            Key2 = Right(FullKeyName, Len(FullKeyName) - i)
        End If
    End If
Next i
'frmMain.Cls
If Key1 = "HKEY_LOCAL_MACHINE" Then
RetVal = RegOpenKeyEx(HKEY_LOCAL_MACHINE, Key2, 0, KEY_ALL_ACCESS, hKey)
ElseIf Key1 = "HKEY_CURRENT_USER" Then
RetVal = RegOpenKeyEx(HKEY_CURRENT_USER, Key2, 0, KEY_ALL_ACCESS, hKey)
End If

index = 0
Do While RetVal = 0
    NameKey = Space(255)
    DataString = Space(255)
    LenName = 255
    DataLen = 255
    RetVal = RegEnumValue(hKey, index, NameKey, LenName, ByVal 0, lpType, Data(0), DataLen)
    If RetVal = 0 Then
        NameKey = Left(NameKey, LenName) 'R�t b? kho?n tr?ng th?a
        DataString = ""
' X? l� th�ng tin theo ki?u c?a n� v� ??a v�o bi?n DataString
        Select Case lpType
             Case REG_SZ
                For i = 0 To DataLen - 1
                    DataString = DataString & Chr(Data(i)) ' N?i c�c ch? c�i th�nh chu?i
                Next
             Case REG_BINARY
                For i = 0 To DataLen - 1
                    Dim temp As String
                    temp = Hex(Data(i))
                    If Len(temp) < 2 Then temp = String(2 - Len(temp), "0") & temp
                    DataString = DataString & temp & " "
 ' N?i c�c c?p s? nh? ph�n l?i v?i nhau
                Next
            Case REG_DWORD
                For i = DataLen - 1 To 0 Step -1
                    DataString = DataString & Hex(Data(i)) 'N?i c�c s� hexa v?i nhau
                Next
            Case REG_MULTI_SZ
                For i = 0 To DataLen - 1
                    DataString = DataString & Chr(Data(i))
    'N?i c�c k� t? bao g?m k� t? vbNullChar (?? c�ch d�ng) th�nh m?t chu?i, b?n c� th? s? d?ng m?t m?ng g?m nhi?u string thay v� l� m?t
                Next
            Case REG_EXPAND_SZ
                For i = 0 To DataLen - 2
                    DataString = DataString & Chr(Data(i))
    'N?i c�c k� t? l?i v?i nhau, b? k� t? NULL cu?i c�ng
                Next
            Case Else
                DataString = " Khong xac dinh duoc !"
        ' Tr�n ?�y l� 5 ki?u c� tr�n WinXP
        End Select
    End If
    If Left(Left(NameKey, LenName), 1) <> " " Then
    '///////////////////
    'Form1.List1.AddItem DataString
    frmMain.cslblStatus.Caption = DataString
    
    Dim AX As String

   AX = CheckVirus(DataString)
   If AX <> "No" Then
                Dim ia
                ia = frmMain.LVVirus2.ListItems.Count + 1
                frmMain.LVVirus2.ListItems.Add ia, , AX
                frmMain.LVVirus2.ListItems(ia).SubItems(1).Caption = DataString
                frmMain.LVVirus2.ListItems(ia).SubItems(2).Caption = FileLen(DataString) & " Bytes"
                frmMain.LVVirus2.ListItems(ia).SubItems(3).Caption = "---"
                frmMain.LVVirus2.ListItems(ia).SubItems(4).Caption = Key1 & "-" & Key2 & ":" & Left(NameKey, LenName)
                
                frmMain.LVVirus2.ListItems(ia).Checked = True
    End If
    
    
    '///////////////
    End If
    index = index + 1
    'frmMain.Print Left(NameKey, LenName) & "=" & DataString
Loop
RetVal = RegCloseKey(hKey)
End Function


