Attribute VB_Name = "modString"
Option Explicit

Public Declare Function WideCharToMultiByte Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long, ByRef lpMultiByteStr As Any, ByVal cchMultiByte As Long, ByVal lpDefaultChar As String, ByVal lpUsedDefaultChar As Long) As Long
Public Declare Function MultiByteToWideChar Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByRef lpMultiByteStr As Any, ByVal cchMultiByte As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)
Public Const CP_UTF8 = 65001

Public Function UTF82Unicode(ByVal sUTF8 As String) As String

    Dim UTF8Size      As Long
    Dim BufferSize    As Long
    Dim BufferUNI    As String
    Dim LenUNI        As Long
    Dim bUTF8()      As Byte

    If LenB(sUTF8) = 0 Then Exit Function

    bUTF8 = StrConv(sUTF8, vbFromUnicode)
    UTF8Size = UBound(bUTF8) + 1

    BufferSize = UTF8Size * 2
    BufferUNI = String$(BufferSize, vbNullChar)

    LenUNI = MultiByteToWideChar(CP_UTF8, 0, bUTF8(0), UTF8Size, StrPtr(BufferUNI), BufferSize)

    If LenUNI Then
        UTF82Unicode = Left$(BufferUNI, LenUNI)
    End If

End Function


Public Function Unicode2UTF8(ByVal strUnicode As String) As String

    Dim LenUNI    As Long
    Dim BufferSize As Long
    Dim LenUTF8    As Long
    Dim bUTF8()    As Byte

    LenUNI = Len(strUnicode)

    If LenUNI = 0 Then Exit Function

    BufferSize = LenUNI * 3 + 1
    ReDim bUTF8(BufferSize - 1)

    LenUTF8 = WideCharToMultiByte(CP_UTF8, 0, StrPtr(strUnicode), LenUNI, bUTF8(0), BufferSize, vbNullString, 0)

    If LenUTF8 Then
        ReDim Preserve bUTF8(LenUTF8 - 1)
        Unicode2UTF8 = StrConv(bUTF8, vbUnicode)
    End If

End Function

