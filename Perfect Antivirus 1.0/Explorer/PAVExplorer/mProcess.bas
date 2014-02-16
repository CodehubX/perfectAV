Attribute VB_Name = "mProcess"
Public Declare Function GetFileAttributes Lib "kernel32" Alias "GetFileAttributesA" (ByVal lpFileName As String) As Long

Function GetAttribute(ByVal sFilePath As String) As String
    Select Case GetFileAttributes(sFilePath)
        Case 1: GetAttribute = "Read Only"
        Case 2: GetAttribute = "Hidden"
        Case 3: GetAttribute = "Read Only + Hidden"
        Case 4: GetAttribute = "System"
        Case 5: GetAttribute = "Read Only + System"
        Case 6: GetAttribute = "Hidden + System"
        Case 7: GetAttribute = "Read Only + Hidden + System"
        '-------------------------------------------------'
        Case 32: GetAttribute = "Archive"
        Case 33: GetAttribute = "Read Only + Archive"
        Case 34: GetAttribute = "Hidden + Archive"
        Case 35: GetAttribute = "Read Only + Hidden + Archive"
        Case 36: GetAttribute = "System + Archive"
        Case 37: GetAttribute = "Read Only + System + Archive"
        Case 38: GetAttribute = "HSA"
        Case 39: GetAttribute = "Read Only + Hidden + System + Archive"
        '-------------------------------------------------'
        Case 128: GetAttribute = "Normal"
        '-------------------------------------------------'
        Case Else: GetAttribute = "N/A"
    End Select
End Function
