Attribute VB_Name = "modSound"

Option Explicit

Private Const SND_ASYNC As Long = &H1    '  play asynchronously

Private Declare Function sndPlaySound _
                Lib "winmm.dll" _
                Alias "sndPlaySoundA" (ByVal lpszSoundName As String, _
                                       ByVal uFlags As Long) As Long

Public Sub PLaySound(F As String)

        '<EhHeader>
        On Error GoTo PLaySound_Err

        '</EhHeader>

100     sndPlaySound F, SND_ASYNC

        '<EhFooter>
        Exit Sub

PLaySound_Err:
        WriteErr Time & "-" & Date & " - " & Err.Description & " - " & "in PerfectAntivirus2009.modSound.PLaySound " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

