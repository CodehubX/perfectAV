Attribute VB_Name = "modUSBDetect"

Private Declare Function SetWindowLong _
                Lib "User32.dll" _
                Alias "SetWindowLongA" (ByVal hWnd As Long, _
                                        ByVal nIndex As Long, _
                                        ByVal dwNewLong As Long) As Long

'The CallWindowProc function passes message information to the specified window procedure
Private Declare Function CallWindowProc _
                Lib "User32.dll" _
                Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, _
                                         ByVal hWnd As Long, _
                                         ByVal Msg As Long, _
                                         ByVal wParam As Long, _
                                         ByVal lParam As Long) As Long

'This function converts a globally unique identifier (GUID) into a string of printable characters.
Private Declare Function StringFromGUID2 _
                Lib "OLE32.dll" (ByRef rGUID As Any, _
                                 ByVal lpSz As String, _
                                 ByVal cchMax As Long) As Long

'Convert API Calls from 16-bit to 32-bit
Private Declare Function lstrcpyA _
                Lib "kernel32.dll" (ByVal lpString1 As String, _
                                    ByVal lpString2 As Long) As Long

Private Declare Function lstrlenA Lib "kernel32.dll" (ByVal lpString As Long) As Long

'The GetDriveType function determines whether a disk drive is a removable, fixed, CD-ROM, RAM disk, or network drive
Private Declare Function GetDriveType _
                Lib "kernel32.dll" _
                Alias "GetDriveTypeA" (ByVal nDrive As String) As Long

'The RtlMoveMemory routine moves memory either forward or backward,
'aligned or unaligned, in 4-byte blocks, followed by any remaining bytes
Private Declare Sub RtlMoveMemory _
                Lib "kernel32.dll" (ByRef Destination As Any, _
                                    ByRef Source As Any, _
                                    ByVal Length As Long)

'The GetDWORD method retrieves a DWORD property
Private Declare Sub GetDWORD _
                Lib "MSVBVM60.dll" _
                Alias "GetMem4" (ByRef inSrc As Any, _
                                 ByRef inDst As Long)

' GetWORD method retrieves a WORD property
Private Declare Sub GetWord _
                Lib "MSVBVM60.dll" _
                Alias "GetMem2" (ByRef inSrc As Any, _
                                 ByRef inDst As Integer)

'The DEV_BROADCAST_HDR structure is a standard header for information related to a device event reported
'through the WM_DEVICECHANGE message
Private Type DEV_BROADCAST_HDR

    dbch_size As Long
    dbch_devicetype As Long
    dbch_reserved As Long

End Type

'Used with the GUIDString Function ( GUID = Globally Unique Identifier )
Public Type Guid

    D1 As Long
    D2 As Integer
    D3 As Integer
    D4(7) As Byte

End Type

Dim OldProc                              As Long

'Window handle
Dim WHnd                                 As Long

'use the GWL_WNDPROC constant to tell the SetWindowLong function that you
'want to change the address of the target window's WindowProc function
Private Const GWL_WNDPROC                As Long = (-4)

'The WM_DEVICECHANGE device message notifies an application of a change to the hardware
'configuration of a device or the computer
Private Const WM_DEVICECHANGE            As Long = &H219

'The system broadcasts the DBT_DEVNODES_CHANGED device event when a device has been added to or removed from the system
Private Const DBT_DEVNODES_CHANGED       As Long = &H7

'The system broadcasts the DBT_DEVICEARRIVAL device event when a device or piece of media has been inserted and becomes available
Private Const DBT_DEVICEARRIVAL          As Long = &H8000&

'The system broadcasts the DBT_DEVICEREMOVECOMPLETE device event when a device or piece of media has been physically removed
Private Const DBT_DEVICEREMOVECOMPLETE   As Long = &H8004&

'The application must check the event to ensure that the type of device arriving is a volume
Private Const DBT_DEVTYP_VOLUME          As Long = &H2 ' Logical volume

Private Const DBT_DEVTYP_DEVICEINTERFACE As Long = &H5 ' Device interface class

Private Const DBTF_MEDIA                 As Long = &H1 ' Media comings and goings

Private Const DBTF_NET                   As Long = &H2 ' Network volume

'No Root Directory
Private Const DRIVE_NO_ROOT_DIR          As Long = 1

'Removeable drive
Private Const DRIVE_REMOVABLE            As Long = 2

'fixed drive ( not removeable )
Private Const DRIVE_FIXED                As Long = 3

'remote drive ( network )
Private Const DRIVE_REMOTE               As Long = 4

'CD rom
Private Const DRIVE_CDROM                As Long = 5

'RAM disk ( USB stick )
Private Const DRIVE_RAMDISK              As Long = 6

Public Sub SubClass(ByVal iWnd As Long)

        '<EhHeader>
        On Error GoTo SubClass_Err

        '</EhHeader>
100     If (WHnd) Then Call UnSubClass

102     OldProc = SetWindowLong(iWnd, GWL_WNDPROC, AddressOf WndProc)
104     WHnd = iWnd

        '<EhFooter>
        Exit Sub

SubClass_Err:
        WriteErr Time & "-" & Date & " - " & Err.Description & " - " & "in PerfectAntivirus2009.modUSBDetect.SubClass " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Public Sub UnSubClass()

        '<EhHeader>
        On Error GoTo UnSubClass_Err

        '</EhHeader>
100     If (WHnd = 0) Then Exit Sub
102     Call SetWindowLong(WHnd, GWL_WNDPROC, OldProc)

104     WHnd = 0
106     OldProc = 0

        '<EhFooter>
        Exit Sub

UnSubClass_Err:
        WriteErr Time & "-" & Date & " - " & Err.Description & " - " & "in PerfectAntivirus2009.modUSBDetect.UnSubClass " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Function WndProc(ByVal hWnd As Long, _
                         ByVal uMsg As Long, _
                         ByVal wParam As Long, _
                         ByVal lParam As Long) As Long

        '<EhHeader>
        On Error GoTo WndProc_Err

        '</EhHeader>
        Dim DevBroadcastHead As DEV_BROADCAST_HDR

        Dim UMask            As Long, Flags As Integer

100     If (uMsg = WM_DEVICECHANGE) Then

102         Select Case wParam

                Case DBT_DEVICEARRIVAL, DBT_DEVICEREMOVECOMPLETE
104                 Call RtlMoveMemory(DevBroadcastHead, ByVal lParam, Len(DevBroadcastHead))

106                 If (DevBroadcastHead.dbch_devicetype = DBT_DEVTYP_VOLUME) Then
108                     Call GetDWORD(ByVal (lParam + Len(DevBroadcastHead)), UMask)
110                     Call GetWord(ByVal (lParam + Len(DevBroadcastHead) + 4), Flags)

                        '/////////// Here is Detected /////////////
                        'MsgBox "Drive(s): " & UMaskString(UMask) & " " & IIf(wParam = DBT_DEVICEARRIVAL, "Inserted", "Ejected")
112                     If wParam = DBT_DEVICEARRIVAL Then
114                         USB_IN UMaskString(UMask)
                        Else
                            'USB_OUT UMaskString(UMask)
116                         frmMessenger.zShowMessenger "D9a4 ru1t USB", "D9a4 nga81t ke61t no71i vo71i [" & UMaskString(UMask) & ":\]", 7000, xTrang

                        End If
                    
                    End If

            End Select

        End If

118     WndProc = CallWindowProc(OldProc, hWnd, uMsg, wParam, lParam)

        '<EhFooter>
        Exit Function

WndProc_Err:
        WriteErr Time & "-" & Date & " - " & Err.Description & " - " & "in PerfectAntivirus2009.modUSBDetect.WndProc " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Function

Public Function CopyString(ByVal lPtr As Long) As String

        '<EhHeader>
        On Error GoTo CopyString_Err

        '</EhHeader>
        Dim BufferLen As Long

100     BufferLen = lstrlenA(lPtr)

102     If (BufferLen > 0) Then
104         CopyString = Space$(BufferLen)
106         Call lstrcpyA(CopyString, lPtr)
        End If

        '<EhFooter>
        Exit Function

CopyString_Err:
        WriteErr Time & "-" & Date & " - " & Err.Description & " - " & "in PerfectAntivirus2009.modUSBDetect.CopyString " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Function

Private Function DriveTypeString(ByVal lDriveType As Long) As String

        '<EhHeader>
        On Error GoTo DriveTypeString_Err

        '</EhHeader>
100     Select Case lDriveType

            Case DRIVE_NO_ROOT_DIR: DriveTypeString = "No root directory"

102         Case DRIVE_REMOVABLE:  DriveTypeString = "Removable"

104         Case DRIVE_FIXED:      DriveTypeString = "Fixed"

106         Case DRIVE_REMOTE:      DriveTypeString = "Remote"

108         Case DRIVE_CDROM:      DriveTypeString = "CD-ROM"

110         Case DRIVE_RAMDISK:    DriveTypeString = "RAM disk"

112         Case Else:              DriveTypeString = "[ Unknown ]"
        End Select

        '<EhFooter>
        Exit Function

DriveTypeString_Err:
        WriteErr Time & "-" & Date & " - " & Err.Description & " - " & "in PerfectAntivirus2009.modUSBDetect.DriveTypeString " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Function

Private Function UMaskString(ByVal iUnitMask As Long) As String

        '<EhHeader>
        On Error GoTo UMaskString_Err

        '</EhHeader>
        Dim Bits As Long

100     For Bits = 0 To 30

102         If (iUnitMask And (2 ^ Bits)) Then UMaskString = UMaskString & Chr$(Asc("A") + Bits)
104     Next Bits

        '<EhFooter>
        Exit Function

UMaskString_Err:
        WriteErr Time & "-" & Date & " - " & Err.Description & " - " & "in PerfectAntivirus2009.modUSBDetect.UMaskString " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Function

Private Function GUIDString(ByRef inGUID As Guid) As String

        '<EhHeader>
        On Error GoTo GUIDString_Err

        '</EhHeader>
        Dim RetBuffer   As String, GUILen As Long

        Const BufferLen As Long = 80

100     RetBuffer = Space$(BufferLen)
102     GUILen = StringFromGUID2(inGUID, RetBuffer, BufferLen)

104     If (GUILen) Then GUIDString = StrConv(Left$(RetBuffer, (GUILen - 1) * 2), vbFromUnicode)

        '<EhFooter>
        Exit Function

GUIDString_Err:
        WriteErr Time & "-" & Date & " - " & Err.Description & " - " & "in PerfectAntivirus2009.modUSBDetect.GUIDString " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Function

Public Sub USB_IN(xDriverLetter)

        '<EhHeader>
        On Error GoTo USB_IN_Err

        '</EhHeader>

100     zfrmScanUSB.zScanUSB xDriverLetter & ":\"

        '<EhFooter>
        Exit Sub

USB_IN_Err:
        WriteErr Time & "-" & Date & " - " & Err.Description & " - " & "in PerfectAntivirus2009.modUSBDetect.USB_IN " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

