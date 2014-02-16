VERSION 5.00
Begin VB.UserControl Downloader 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
End
Attribute VB_Name = "Downloader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' Code by Le Cao Minh Thanh
' Thanh vien caulacbovb
' http://www.caulacbovb.com
' Email: thanh3em@gmail.com
' OCX ho tro download Multi va dem duong luong da download ve,
' su dung cac su kien san co cua Usercontrol, tham khao Code tu nhieu nguon tong hop nen
' Chu y: Cac ban phai ghi ro nguon goc Source neu co su dung lai

Dim m_Path As New Collection
Dim m_Name As New Collection
Event Progress(ByteDownloaded As Long, FileSize As Long, URL As String)
Event Complete(URL As String)

Public Sub DownloadFile(ByVal URL As String, ByVal Path As String)
Dim i As Integer
For i = 1 To m_Name.Count
    If m_Name(i) = URL Then
        Debug.Print "URL nay da ton tai trong tien trinh Download"
        Exit Sub
    End If
Next i
m_Name.Add URL, URL
m_Path.Add Path, URL
UserControl.AsyncRead URL, vbAsyncTypeFile, URL, vbAsyncReadForceUpdate
End Sub

Public Sub StopDownload(ByVal URL As String)
Dim i As Integer
For i = 1 To m_Name.Count
    If m_Name(i) = URL Then
        GoTo TiepTuc
    End If
Next i
Debug.Print "Khong ton tai URL nay trong tien trinh Download"
Exit Sub
TiepTuc:
UserControl.CancelAsyncRead URL
End Sub

Private Sub UserControl_AsyncReadComplete(AsyncProp As AsyncProperty)
On Error Resume Next
Name AsyncProp.Value As m_Path.Item(AsyncProp.PropertyName)
m_Path.Remove AsyncProp.PropertyName
m_Name.Remove AsyncProp.PropertyName
RaiseEvent Complete(AsyncProp.PropertyName)
End Sub

Private Sub UserControl_AsyncReadProgress(AsyncProp As AsyncProperty)
RaiseEvent Progress(AsyncProp.BytesRead, AsyncProp.BytesMax, AsyncProp.PropertyName)
End Sub
