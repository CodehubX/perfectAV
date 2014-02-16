Attribute VB_Name = "Module3"
Private Declare Function InternetGetConnectedStateEx Lib "wininet.dll" (ByRef lpdwFlags As Long, ByVal lpszConnectionName As String, ByVal dwNameLen As Integer, ByVal dwReserved As Long) As Long
Dim sConnType As String * 255

Public Function CheckInternet() As Boolean
On Error GoTo KhOnGkEtNoI
    Dim Ret As Long
    Ret = InternetGetConnectedStateEx(Ret, sConnType, 254, 0)
    If Ret = 1 Then
        'MsgBox "Ban da ket noi Internet " & sConnType, vbInformation
        CheckInternet = True
    Else
        'MsgBox "Ban chua kt noi internet", vbInformation
        CheckInternet = False
    End If
Exit Function
KhOnGkEtNoI:
CheckInternet = False
End Function


