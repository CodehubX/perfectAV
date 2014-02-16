Attribute VB_Name = "modString"
Option Explicit

Public Declare Function WideCharToMultiByte _
               Lib "kernel32" (ByVal CodePage As Long, _
                               ByVal dwFlags As Long, _
                               ByVal lpWideCharStr As Long, _
                               ByVal cchWideChar As Long, _
                               ByRef lpMultiByteStr As Any, _
                               ByVal cchMultiByte As Long, _
                               ByVal lpDefaultChar As String, _
                               ByVal lpUsedDefaultChar As Long) As Long

Public Declare Function MultiByteToWideChar _
               Lib "kernel32" (ByVal CodePage As Long, _
                               ByVal dwFlags As Long, _
                               ByRef lpMultiByteStr As Any, _
                               ByVal cchMultiByte As Long, _
                               ByVal lpWideCharStr As Long, _
                               ByVal cchWideChar As Long) As Long

Public Declare Sub CopyMemory _
               Lib "kernel32" _
               Alias "RtlMoveMemory" (lpvDest As Any, _
                                      lpvSource As Any, _
                                      ByVal cbCopy As Long)

Public Const CP_UTF8 = 65001

Public Function UTF82Unicode(ByVal sUTF8 As String) As String

        '<EhHeader>
        On Error GoTo UTF82Unicode_Err

        '</EhHeader>

        Dim UTF8Size   As Long

        Dim BufferSize As Long

        Dim BufferUNI  As String

        Dim LenUNI     As Long

        Dim bUTF8()    As Byte

100     If LenB(sUTF8) = 0 Then Exit Function

102     bUTF8 = StrConv(sUTF8, vbFromUnicode)
104     UTF8Size = UBound(bUTF8) + 1

106     BufferSize = UTF8Size * 2
108     BufferUNI = String$(BufferSize, vbNullChar)

110     LenUNI = MultiByteToWideChar(CP_UTF8, 0, bUTF8(0), UTF8Size, StrPtr(BufferUNI), BufferSize)

112     If LenUNI Then
114         UTF82Unicode = Left$(BufferUNI, LenUNI)
        End If

        '<EhFooter>
        Exit Function

UTF82Unicode_Err:
        WriteErr Time & "-" & Date & " - " & Err.Description & " - " & "in PerfectAntivirus2009.modString.UTF82Unicode " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Function

Public Function Unicode2UTF8(ByVal strUnicode As String) As String

        '<EhHeader>
        On Error GoTo Unicode2UTF8_Err

        '</EhHeader>

        Dim LenUNI     As Long

        Dim BufferSize As Long

        Dim LenUTF8    As Long

        Dim bUTF8()    As Byte

100     LenUNI = Len(strUnicode)

102     If LenUNI = 0 Then Exit Function

104     BufferSize = LenUNI * 3 + 1
106     ReDim bUTF8(BufferSize - 1)

108     LenUTF8 = WideCharToMultiByte(CP_UTF8, 0, StrPtr(strUnicode), LenUNI, bUTF8(0), BufferSize, vbNullString, 0)

110     If LenUTF8 Then
112         ReDim Preserve bUTF8(LenUTF8 - 1)
114         Unicode2UTF8 = StrConv(bUTF8, vbUnicode)
        End If

        '<EhFooter>
        Exit Function

Unicode2UTF8_Err:
        WriteErr Time & "-" & Date & " - " & Err.Description & " - " & "in PerfectAntivirus2009.modString.Unicode2UTF8 " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Function

