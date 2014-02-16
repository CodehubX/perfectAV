Attribute VB_Name = "basMD5"
'This module is used to gather the contents of a file quickly and to grab the MD5 of a file quickly by using API functions. Use this
'code in any projects you wish, no need to give credit. Please vote though.

'marcin@malwarebytes.org if you have any questions.

Option Explicit

Public Const OPEN_EXISTING = 3

Public Const GENERIC_READ = &H80000000

Public Const FILE_SHARE_READ = &H1

Public Const FILE_SHARE_WRITE = &H2

Public Const BUFFER_SIZE            As Long = 255

'MD5 Hashing
Public Const MS_ENHANCED_PROV       As String = "Microsoft Enhanced Cryptographic Provider v1.0"

Public Const MS_BASE_PROV           As String = "Microsoft Base Cryptographic Provider v1.0"

Public Const PROV_RSA_FULL          As Long = 1

Public Const ALG_CLASS_DATA_ENCRYPT As Long = 24576

Public Const ALG_TYPE_STREAM        As Long = 2048

Public Const ALG_TYPE_ANY           As Long = 0

Public Const ALG_SID_RC4            As Long = 1

Public Const ALG_SID_MD5            As Long = 3

Public Const CALG_RC4               As Long = ALG_CLASS_DATA_ENCRYPT Or ALG_TYPE_STREAM Or ALG_SID_RC4

Public Const CRYPT_VERIFYCONTEXT = &HF0000000

Public Const CRYPT_NEWKEYSET   As Long = 8

Public Const ENCRYPT_ALGORITHM As Long = CALG_RC4

Public Const ENCRYPT_NUMBERKEY As String = "16006833"

Public Const ALG_CLASS_HASH    As Long = 32768

Public Const HP_HASHVAL        As Long = 2

Public Const HP_HASHSIZE       As Long = 4

'Faster hashing
Public Const HASH_TYPE = ALG_TYPE_ANY Or ALG_CLASS_HASH Or ALG_SID_MD5

Public Declare Function GetTickCount Lib "kernel32.dll" () As Long

Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Public Declare Function ReadFile _
               Lib "kernel32" (ByVal hFile As Long, _
                               lpBuffer As Any, _
                               ByVal nNumberOfBytesToRead As Long, _
                               lpNumberOfBytesRead As Long, _
                               lpOverlapped As Any) As Long

Public Declare Function CreateFile _
               Lib "kernel32" _
               Alias "CreateFileA" (ByVal lpFileName As String, _
                                    ByVal dwDesiredAccess As Long, _
                                    ByVal dwShareMode As Long, _
                                    lpSecurityAttributes As Any, _
                                    ByVal dwCreationDisposition As Long, _
                                    ByVal dwFlagsAndAttributes As Long, _
                                    ByVal hTemplateFile As Long) As Long

Public Declare Function GetFileSize _
               Lib "kernel32" (ByVal hFile As Long, _
                               lpFileSizeHigh As Long) As Long

'MD5 Hashing
Public Declare Function CryptAcquireContext _
               Lib "advapi32.dll" _
               Alias "CryptAcquireContextA" (ByRef phProv As Long, _
                                             ByVal pszContainer As String, _
                                             ByVal pszProvider As String, _
                                             ByVal dwProvType As Long, _
                                             ByVal dwFlags As Long) As Long

Public Declare Function CryptCreateHash _
               Lib "advapi32.dll" (ByVal hProv As Long, _
                                   ByVal Algid As Long, _
                                   ByVal hKey As Long, _
                                   ByVal dwFlags As Long, _
                                   ByRef phHash As Long) As Long

Public Declare Function CryptHashData _
               Lib "advapi32.dll" (ByVal hHash As Long, _
                                   ByVal pbData As String, _
                                   ByVal dwDataLen As Long, _
                                   ByVal dwFlags As Long) As Long

Public Declare Function CryptDeriveKey _
               Lib "advapi32.dll" (ByVal hProv As Long, _
                                   ByVal Algid As Long, _
                                   ByVal hBaseData As Long, _
                                   ByVal dwFlags As Long, _
                                   ByRef phKey As Long) As Long

Public Declare Function CryptEncrypt _
               Lib "advapi32.dll" (ByVal hKey As Long, _
                                   ByVal hHash As Long, _
                                   ByVal Final As Long, _
                                   ByVal dwFlags As Long, _
                                   ByVal pbData As String, _
                                   ByRef pdwDataLen As Long, _
                                   ByVal dwBufLen As Long) As Long

Public Declare Function CryptDecrypt _
               Lib "advapi32.dll" (ByVal hKey As Long, _
                                   ByVal hHash As Long, _
                                   ByVal Final As Long, _
                                   ByVal dwFlags As Long, _
                                   ByVal pbData As String, _
                                   ByRef pdwDataLen As Long) As Long

Public Declare Function CryptGetHashParam _
               Lib "advapi32.dll" (ByVal pCryptHash As Long, _
                                   ByVal dwParam As Long, _
                                   ByRef pbData As Any, _
                                   ByRef pcbData As Long, _
                                   ByVal dwFlags As Long) As Long

Public Declare Function CryptDestroyKey Lib "advapi32.dll" (ByVal hKey As Long) As Long

Public Declare Function CryptDestroyHash Lib "advapi32.dll" (ByVal hHash As Long) As Long

Public Declare Function CryptReleaseContext _
               Lib "advapi32.dll" (ByVal hProv As Long, _
                                   ByVal dwFlags As Long) As Long

Public Function InputFile$(ByRef sFile$)

        '<EhHeader>
        On Error GoTo InputFile_Err

        '</EhHeader>
        Dim hFile&, uBuffer() As Byte, lFileSize&, lBytesRead&
    
        'Get a handle to the file
100     hFile = CreateFile(sFile, GENERIC_READ, FILE_SHARE_READ Or FILE_SHARE_WRITE, ByVal 0, OPEN_EXISTING, 0, 0)
        
        'Can't open file
102     If hFile = -1 Then Exit Function
104     lFileSize = GetFileSize(hFile, 0)
    
106     If lFileSize < 1 Then
108         CloseHandle hFile

            Exit Function

        End If
    
        'Prepare the buffer
110     ReDim uBuffer(lFileSize - 1)
    
        'Read the file
112     If ReadFile(hFile, uBuffer(0), lFileSize, lBytesRead, ByVal 0) <> 0 Then
114         If lBytesRead <> lFileSize Then
116             ReDim Preserve uBuffer(lBytesRead)
            End If
    
118         InputFile = StrConv(uBuffer, vbUnicode)
        End If
        
        'Close the handle to the file
120     CloseHandle hFile

        '<EhFooter>
        Exit Function

InputFile_Err:
        WriteErr Time & "-" & Date & " - " & Err.Description & " - " & "in PerfectAntivirus2009.basMD5.InputFile " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Function

Public Function GetMD5$(ByRef sFileName$)

        'Get the MD5
        '<EhHeader>
        On Error GoTo GetMD5_Err

        '</EhHeader>

100     GetMD5 = MD5String(InputFile(sFileName))

        '<EhFooter>
        Exit Function

GetMD5_Err:
        WriteErr Time & "-" & Date & " - " & Err.Description & " - " & "in PerfectAntivirus2009.basMD5.GetMD5 " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Function

Public Function MD5String(ByRef sString$)

        '<EhHeader>
        On Error GoTo MD5String_Err

        '</EhHeader>
        Dim uMD5() As Byte, lMD5Len&, I&, sMD5$, hCrypt&, hHash&
    
        'Prepare the byte array
100     ReDim uMD5(BUFFER_SIZE)

102     DoEvents

        'Acquire the MD5 hash generator
104     If CryptAcquireContext(hCrypt, vbNullString, MS_ENHANCED_PROV, PROV_RSA_FULL, CRYPT_VERIFYCONTEXT) <> 0 Then
106         If CryptCreateHash(hCrypt, HASH_TYPE, 0, 0, hHash) <> 0 Then
108             If CryptHashData(hHash, sString, Len(sString), 0) <> 0 Then
110                 If CryptGetHashParam(hHash, HP_HASHSIZE, uMD5(0), UBound(uMD5) + 1, 0) <> 0 Then
112                     lMD5Len = uMD5(0)
                    
114                     If CryptGetHashParam(hHash, HP_HASHVAL, uMD5(0), UBound(uMD5) + 1, 0) <> 0 Then

116                         For I = 0 To lMD5Len - 1
118                             sMD5 = sMD5 & (Right$("0" & Hex$(uMD5(I)), 2))
120                         Next I
                            
122                         MD5String = sMD5
                        End If
                    End If
                End If
            End If
        End If
        
        'Destroy the MD5 hash generator
124     CryptDestroyHash hHash
126     CryptReleaseContext hCrypt, 0

        '<EhFooter>
        Exit Function

MD5String_Err:
        WriteErr Time & "-" & Date & " - " & Err.Description & " - " & "in PerfectAntivirus2009.basMD5.MD5String " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Function

