Attribute VB_Name = "modUSBDetect"
Private Declare Function SetWindowLong Lib "User32.dll" Alias "SetWindowLongA" ( _
    ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

'The CallWindowProc function passes message information to the specified window procedure
Private Declare Function CallWindowProc Lib "User32.dll" Alias "CallWindowProcA" ( _
    ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, _
    ByVal wParam As Long, ByVal lParam As Long) As Long

'This function converts a globally unique identifier (GUID) into a string of printable characters.
Private Declare Function StringFromGUID2 Lib "OLE32.dll" ( _
    ByRef rGUID As Any, ByVal lpSz As String, ByVal cchMax As Long) As Long

'Convert API Calls from 16-bit to 32-bit
Private Declare Function lstrcpyA Lib "kernel32.dll" (ByVal lpString1 As String, ByVal lpString2 As Long) As Long
Private Declare Function lstrlenA Lib "kernel32.dll" (ByVal lpString As Long) As Long

'The GetDriveType function determines whether a disk drive is a removable, fixed, CD-ROM, RAM disk, or network drive
Private Declare Function GetDriveType Lib "kernel32.dll" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long

'The RtlMoveMemory routine moves memory either forward or backward,
'aligned or unaligned, in 4-byte blocks, followed by any remaining bytes
Private Declare Sub RtlMoveMemory Lib "kernel32.dll" ( _
    ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)

'The GetDWORD method retrieves a DWORD property
Private Declare Sub GetDWORD Lib "MSVBVM60.dll" Alias "GetMem4" (ByRef inSrc As Any, ByRef inDst As Long)

' GetWORD method retrieves a WORD property
Private Declare Sub GetWord Lib "MSVBVM60.dll" Alias "GetMem2" (ByRef inSrc As Any, ByRef inDst As Integer)

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

Dim OldProc As Long
'Window handle
Dim WHnd As Long

'use the GWL_WNDPROC constant to tell the SetWindowLong function that you
'want to change the address of the target window's WindowProc function
Private Const GWL_WNDPROC As Long = (-4)

'The WM_DEVICECHANGE device message notifies an application of a change to the hardware
'configuration of a device or the computer
Private Const WM_DEVICECHANGE As Long = &H219

'The system broadcasts the DBT_DEVNODES_CHANGED device event when a device has been added to or removed from the system
Private Const DBT_DEVNODES_CHANGED As Long = &H7

'The system broadcasts the DBT_DEVICEARRIVAL device event when a device or piece of media has been inserted and becomes available
Private Const DBT_DEVICEARRIVAL As Long = &H8000&

'The system broadcasts the DBT_DEVICEREMOVECOMPLETE device event when a device or piece of media has been physically removed
Private Const DBT_DEVICEREMOVECOMPLETE As Long = &H8004&

'The application must check the event to ensure that the type of device arriving is a volume
Private Const DBT_DEVTYP_VOLUME As Long = &H2 ' Logical volume
Private Const DBT_DEVTYP_DEVICEINTERFACE As Long = &H5 ' Device interface class

Private Const DBTF_MEDIA As Long = &H1 ' Media comings and goings
Private Const DBTF_NET As Long = &H2 ' Network volume

'No Root Directory
Private Const DRIVE_NO_ROOT_DIR As Long = 1

'Removeable drive
Private Const DRIVE_REMOVABLE As Long = 2

'fixed drive ( not removeable )
Private Const DRIVE_FIXED As Long = 3

'remote drive ( network )
Private Const DRIVE_REMOTE As Long = 4

'CD rom
Private Const DRIVE_CDROM As Long = 5

'RAM disk ( USB stick )
Private Const DRIVE_RAMDISK As Long = 6

Public Sub SubClass(ByVal iWnd As Long)
    If (WHnd) Then Call UnSubClass

    OldProc = SetWindowLong(iWnd, GWL_WNDPROC, AddressOf WndProc)
    WHnd = iWnd
End Sub

Public Sub UnSubClass()
    If (WHnd = 0) Then Exit Sub
    Call SetWindowLong(WHnd, GWL_WNDPROC, OldProc)

    WHnd = 0
    OldProc = 0
End Sub

Private Function WndProc(ByVal hWnd As Long, ByVal uMsg As Long, _
    ByVal wParam As Long, ByVal lParam As Long) As Long
    Dim DevBroadcastHead As DEV_BROADCAST_HDR
    Dim UMask As Long, Flags As Integer

    If (uMsg = WM_DEVICECHANGE) Then
        Select Case wParam
            Case DBT_DEVICEARRIVAL, DBT_DEVICEREMOVECOMPLETE
                Call RtlMoveMemory(DevBroadcastHead, ByVal lParam, Len(DevBroadcastHead))

                If (DevBroadcastHead.dbch_devicetype = DBT_DEVTYP_VOLUME) Then
                    Call GetDWORD(ByVal (lParam + Len(DevBroadcastHead)), UMask)
                    Call GetWord(ByVal (lParam + Len(DevBroadcastHead) + 4), Flags)
                    '/////////// Here is Detected /////////////
                    'MsgBox "Drive(s): " & UMaskString(UMask) & " " & IIf(wParam = DBT_DEVICEARRIVAL, "Inserted", "Ejected")
                    If wParam = DBT_DEVICEARRIVAL Then
                        USB_IN UMaskString(UMask)
                    Else
                        'USB_OUT UMaskString(UMask)
                        frmMessenger.zShowMessenger "D9a4 ru1t USB", "D9a4 nga81t ke61t no71i vo71i [" & UMaskString(UMask) & ":\]", 7000, xTrang

                    End If
                    
                End If

        End Select
    End If

    WndProc = CallWindowProc(OldProc, hWnd, uMsg, wParam, lParam)
End Function

Public Function CopyString(ByVal lPtr As Long) As String
    Dim BufferLen As Long

    BufferLen = lstrlenA(lPtr)

    If (BufferLen > 0) Then
        CopyString = Space$(BufferLen)
        Call lstrcpyA(CopyString, lPtr)
    End If
End Function

Private Function DriveTypeString(ByVal lDriveType As Long) As String
    Select Case lDriveType
        Case DRIVE_NO_ROOT_DIR: DriveTypeString = "No root directory"
        Case DRIVE_REMOVABLE:  DriveTypeString = "Removable"
        Case DRIVE_FIXED:      DriveTypeString = "Fixed"
        Case DRIVE_REMOTE:      DriveTypeString = "Remote"
        Case DRIVE_CDROM:      DriveTypeString = "CD-ROM"
        Case DRIVE_RAMDISK:    DriveTypeString = "RAM disk"
        Case Else:              DriveTypeString = "[ Unknown ]"
    End Select
End Function

Private Function UMaskString(ByVal iUnitMask As Long) As String
    Dim Bits As Long

    For Bits = 0 To 30
        If (iUnitMask And (2 ^ Bits)) Then _
            UMaskString = UMaskString & Chr$(Asc("A") + Bits)
    Next Bits
End Function


Private Function GUIDString(ByRef inGUID As Guid) As String
    Dim RetBuffer As String, GUILen As Long

    Const BufferLen As Long = 80

    RetBuffer = Space$(BufferLen)
    GUILen = StringFromGUID2(inGUID, RetBuffer, BufferLen)

    If (GUILen) Then GUIDString = StrConv(Left$(RetBuffer, (GUILen - 1) * 2), vbFromUnicode)
End Function

Public Sub USB_IN(xDriverLetter)
zfrmScanUSB.zScanUSB xDriverLetter & ":\"
End Sub

