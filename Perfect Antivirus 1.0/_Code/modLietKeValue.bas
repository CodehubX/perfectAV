Attribute VB_Name = "modLietKeValue"
Option Explicit

Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long

Private Declare Function RegOpenKeyEx _
                Lib "advapi32.dll" _
                Alias "RegOpenKeyExA" (ByVal hKey As Long, _
                                       ByVal lpSubKey As String, _
                                       ByVal ulOptions As Long, _
                                       ByVal samDesired As Long, _
                                       phkResult As Long) As Long

Private Declare Function RegEnumValue _
                Lib "advapi32.dll" _
                Alias "RegEnumValueA" (ByVal hKey As Long, _
                                       ByVal dwIndex As Long, _
                                       ByVal lpValueName As String, _
                                       lpcbValueName As Long, _
                                       ByVal lpReserved As Long, _
                                       lpType As Long, _
                                       lpData As Byte, _
                                       lpcbData As Long) As Long

Private Const HKEY_LOCAL_MACHINE = &H80000002

Private Const HKEY_CURRENT_USER = &H80000001

Private Const KEY_ALL_ACCESS = &HF003F

Private Const REG_SZ = 1

Private Const REG_BINARY = 3                     ' Free form binary

Private Const REG_DWORD = 4                      ' 32-bit number

Private Const REG_EXPAND_SZ = 2                  ' Unicode nul terminated string

Private Const REG_MULTI_SZ = 7                   ' Multiple Unicode strings

Dim RetVal         As Long

Dim hKey           As Long

Dim NameKey        As String

Dim lpType         As Long

Dim LenName        As Long

Dim Data(0 To 255) As Byte

Dim DataLen        As Long

Dim DataString     As String

Dim index          As Long

Dim I              As Long

Dim KetQua         As String

Public xTotalStartUp

Public Function GetKeyValue(FullKeyName)

        '<EhHeader>
        On Error GoTo GetKeyValue_Err

        '</EhHeader>

100     xTotalStartUp = 0

        Dim Key1, Key2, I, Ua

102     Ua = 10

104     For I = 1 To Len(FullKeyName)

106         If Mid(FullKeyName, I, 1) = "\" Then
108             Ua = Ua + 10

110             If Ua = 20 Then
112                 Key1 = Left(FullKeyName, I - 1)
114                 Key2 = Right(FullKeyName, Len(FullKeyName) - I)
                End If
            End If

116     Next I

        'frmMain.Cls
118     If Key1 = "HKEY_LOCAL_MACHINE" Then
120         RetVal = RegOpenKeyEx(HKEY_LOCAL_MACHINE, Key2, 0, KEY_ALL_ACCESS, hKey)
122     ElseIf Key1 = "HKEY_CURRENT_USER" Then
124         RetVal = RegOpenKeyEx(HKEY_CURRENT_USER, Key2, 0, KEY_ALL_ACCESS, hKey)
        End If

126     index = 0

128     Do While RetVal = 0
130         NameKey = Space(255)
132         DataString = Space(255)
134         LenName = 255
136         DataLen = 255
138         RetVal = RegEnumValue(hKey, index, NameKey, LenName, ByVal 0, lpType, Data(0), DataLen)

140         If RetVal = 0 Then
142             NameKey = Left(NameKey, LenName) 'Rút b? kho?n tr?ng th?a
144             DataString = ""

                ' X? lý thông tin theo ki?u c?a nó và ??a vào bi?n DataString
146             Select Case lpType

                    Case REG_SZ

148                     For I = 0 To DataLen - 1
150                         DataString = DataString & Chr(Data(I)) ' N?i các ch? cái thành chu?i
                        Next

152                 Case REG_BINARY

154                     For I = 0 To DataLen - 1

                            Dim temp As String

156                         temp = Hex(Data(I))

158                         If Len(temp) < 2 Then temp = String(2 - Len(temp), "0") & temp
160                         DataString = DataString & temp & " "
                            ' N?i các c?p s? nh? phân l?i v?i nhau
                        Next

162                 Case REG_DWORD

164                     For I = DataLen - 1 To 0 Step -1
166                         DataString = DataString & Hex(Data(I)) 'N?i các sô hexa v?i nhau
                        Next

168                 Case REG_MULTI_SZ

170                     For I = 0 To DataLen - 1
172                         DataString = DataString & Chr(Data(I))
                            'N?i các ký t? bao g?m ký t? vbNullChar (?? cách dòng) thành m?t chu?i, b?n có th? s? d?ng m?t m?ng g?m nhi?u string thay vì là m?t
                        Next

174                 Case REG_EXPAND_SZ

176                     For I = 0 To DataLen - 2
178                         DataString = DataString & Chr(Data(I))
                            'N?i các ký t? l?i v?i nhau, b? ký t? NULL cu?i cùng
                        Next

180                 Case Else
182                     DataString = " Khong xac dinh duoc !"
                        ' Trên ?ây là 5 ki?u có trên WinXP
                End Select

            End If

184         If Left(Left(NameKey, LenName), 1) <> " " Then
                '///////////////////
                'Form1.List1.AddItem DataString
186             frmMain.lblStatus.Caption = DataString
    
                Dim AX As String

188             xTotalStartUp = xTotalStartUp + 1
190             AX = CheckVirus(DataString)

192             If AX <> "No" Then

                    Dim ia

194                 ia = frmMain.LVVirus1.ListItems.Count + 1
196                 frmMain.LVVirus1.ListItems.Add ia, , AX
198                 frmMain.LVVirus1.ListItems(ia).SubItems(1).Caption = DataString
200                 frmMain.LVVirus1.ListItems(ia).SubItems(2).Caption = FileLen(DataString) & " Bytes"
202                 frmMain.LVVirus1.ListItems(ia).SubItems(3).Caption = CheckProcess(DataString)
204                 frmMain.LVVirus1.ListItems(ia).SubItems(4).Caption = Key1 & "-" & Key2 & ":" & Left(NameKey, LenName)
                
206                 frmMain.LVVirus1.ListItems(ia).Checked = True
                End If
    
                '///////////////
            End If

208         index = index + 1
            'frmMain.Print Left(NameKey, LenName) & "=" & DataString
        Loop

210     RetVal = RegCloseKey(hKey)

        '<EhFooter>
        Exit Function

GetKeyValue_Err:
        WriteErr Time & "-" & Date & " - " & Err.Description & " - " & "in PerfectAntivirus2009.modLietKeValue.GetKeyValue " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Function

Public Function GetKeyValue2(FullKeyName)

        '<EhHeader>
        On Error GoTo GetKeyValue2_Err

        '</EhHeader>
        Dim Key1, Key2, I, Ua

100     Ua = 10

102     For I = 1 To Len(FullKeyName)

104         If Mid(FullKeyName, I, 1) = "\" Then
106             Ua = Ua + 10

108             If Ua = 20 Then
110                 Key1 = Left(FullKeyName, I - 1)
112                 Key2 = Right(FullKeyName, Len(FullKeyName) - I)
                End If
            End If

114     Next I

        'frmMain.Cls
116     If Key1 = "HKEY_LOCAL_MACHINE" Then
118         RetVal = RegOpenKeyEx(HKEY_LOCAL_MACHINE, Key2, 0, KEY_ALL_ACCESS, hKey)
120     ElseIf Key1 = "HKEY_CURRENT_USER" Then
122         RetVal = RegOpenKeyEx(HKEY_CURRENT_USER, Key2, 0, KEY_ALL_ACCESS, hKey)
        End If

124     index = 0

126     Do While RetVal = 0
128         NameKey = Space(255)
130         DataString = Space(255)
132         LenName = 255
134         DataLen = 255
136         RetVal = RegEnumValue(hKey, index, NameKey, LenName, ByVal 0, lpType, Data(0), DataLen)

138         If RetVal = 0 Then
140             NameKey = Left(NameKey, LenName) 'Rút b? kho?n tr?ng th?a
142             DataString = ""

                ' X? lý thông tin theo ki?u c?a nó và ??a vào bi?n DataString
144             Select Case lpType

                    Case REG_SZ

146                     For I = 0 To DataLen - 1
148                         DataString = DataString & Chr(Data(I)) ' N?i các ch? cái thành chu?i
                        Next

150                 Case REG_BINARY

152                     For I = 0 To DataLen - 1

                            Dim temp As String

154                         temp = Hex(Data(I))

156                         If Len(temp) < 2 Then temp = String(2 - Len(temp), "0") & temp
158                         DataString = DataString & temp & " "
                            ' N?i các c?p s? nh? phân l?i v?i nhau
                        Next

160                 Case REG_DWORD

162                     For I = DataLen - 1 To 0 Step -1
164                         DataString = DataString & Hex(Data(I)) 'N?i các sô hexa v?i nhau
                        Next

166                 Case REG_MULTI_SZ

168                     For I = 0 To DataLen - 1
170                         DataString = DataString & Chr(Data(I))
                            'N?i các ký t? bao g?m ký t? vbNullChar (?? cách dòng) thành m?t chu?i, b?n có th? s? d?ng m?t m?ng g?m nhi?u string thay vì là m?t
                        Next

172                 Case REG_EXPAND_SZ

174                     For I = 0 To DataLen - 2
176                         DataString = DataString & Chr(Data(I))
                            'N?i các ký t? l?i v?i nhau, b? ký t? NULL cu?i cùng
                        Next

178                 Case Else
180                     DataString = " Khong xac dinh duoc !"
                        ' Trên ?ây là 5 ki?u có trên WinXP
                End Select

            End If

182         If Left(Left(NameKey, LenName), 1) <> " " Then
                '///////////////////
                'Form1.List1.AddItem DataString
184             frmMain.cslblStatus.Caption = DataString
    
                Dim AX As String

186             AX = CheckVirus(DataString)

188             If AX <> "No" Then

                    Dim ia

190                 ia = frmMain.LVVirus2.ListItems.Count + 1
192                 frmMain.LVVirus2.ListItems.Add ia, , AX
194                 frmMain.LVVirus2.ListItems(ia).SubItems(1).Caption = DataString
196                 frmMain.LVVirus2.ListItems(ia).SubItems(2).Caption = FileLen(DataString) & " Bytes"
198                 frmMain.LVVirus2.ListItems(ia).SubItems(3).Caption = CheckProcess(DataString)
200                 frmMain.LVVirus2.ListItems(ia).SubItems(4).Caption = Key1 & "-" & Key2 & ":" & Left(NameKey, LenName)
                
202                 frmMain.LVVirus2.ListItems(ia).Checked = True
                End If
    
                '///////////////
            End If

204         index = index + 1
            'frmMain.Print Left(NameKey, LenName) & "=" & DataString
        Loop

206     RetVal = RegCloseKey(hKey)

        '<EhFooter>
        Exit Function

GetKeyValue2_Err:
        WriteErr Time & "-" & Date & " - " & Err.Description & " - " & "in PerfectAntivirus2009.modLietKeValue.GetKeyValue2 " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Function

