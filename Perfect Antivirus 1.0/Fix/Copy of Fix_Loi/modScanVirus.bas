Attribute VB_Name = "modScanVirus"
Dim db As Database
Dim rs As Recordset
Dim WS As Workspace
Dim mData As Recordset
Dim nData As Recordset

Public Sub FixError()

On Error Resume Next


Dim MD5Cod
MD5Cod = "ED0EF0A136DEC83DF69F04118870003E"
Set mData = db.OpenRecordset("SELECT * FROM " & "Data" & " WHERE " & "MD5Code" & "='" & MD5Cod & "'")
mData.Delete
MD5Cod = "1BED46E005F56926898798B8A04D188D"
Set mData = db.OpenRecordset("SELECT * FROM " & "Data" & " WHERE " & "MD5Code" & "='" & MD5Cod & "'")
mData.Delete
MD5Cod = "5302EEE7E82AA83CC6E0490D60F330DB"
Set mData = db.OpenRecordset("SELECT * FROM " & "Data" & " WHERE " & "MD5Code" & "='" & MD5Cod & "'")
mData.Delete
MD5Cod = "3C7B5613517DA2712DB47348C40E9B33"
Set mData = db.OpenRecordset("SELECT * FROM " & "Data" & " WHERE " & "MD5Code" & "='" & MD5Cod & "'")
mData.Delete
MD5Cod = "0E776ED5F7CC9F94299E70461B7B8185"
Set mData = db.OpenRecordset("SELECT * FROM " & "Data" & " WHERE " & "MD5Code" & "='" & MD5Cod & "'")
mData.Delete

Call CreateField2(AppPath & "DATA.PAV")

End Sub

Public Sub ConnectDB()
On Error GoTo CoLoIxAyRa
Set WS = DBEngine.Workspaces(0)
    DbFile = (AppPath & "DATA.PAV")
    PwdString = "htgtalcmdltnsc"
    Set db = DBEngine.OpenDatabase(DbFile, False, False, ";PWD=" & PwdString)
    
Exit Sub
CoLoIxAyRa:
MsgBox "Khong tim thay CSDL!" & vbCrLf & "Hay dat file nay vao thu muc cai dat cua PAV 2009!" & vbCrLf & "Tat PAV 2009 truoc khi chay file Fix loi nay!", vbOKOnly + vbCritical, "Error"
End
End Sub

Public Function AppPath()
AppPath = App.Path
If Right(AppPath, 1) <> "\" Then AppPath = AppPath & "\"
End Function

Public Sub CreateField2(AccessPath$)
On Error GoTo GaPlOiRoItHoAt

Set objConnection = CreateObject("ADODB.Connection")
    objConnection.Open "Provider = Microsoft.Jet.OLEDB.4.0;Data Source=" & AccessPath & ";Jet OLEDB:Database Password=" & "htgtalcmdltnsc" & ";"
    objConnection.Execute "CREATE TABLE STRI([ID] COUNTER,[String] MEMO, VirusName MEMO)"
    objConnection.Close
    
Exit Sub
GaPlOiRoItHoAt:
MsgBox "Chuong trinh da duoc Update phien ban 1.2 roi!" & vbCrLf & "Ban khong can phai Update them lan nua!", vbOKOnly + vbInformation, "OK!"
End
End Sub
