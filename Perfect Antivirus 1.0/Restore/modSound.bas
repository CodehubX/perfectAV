Attribute VB_Name = "modSound"

Option Explicit
Private Const SND_ASYNC   As Long = &H1    '  play asynchronously
Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, _
                                                                             ByVal uFlags As Long) As Long

Public Sub PLaySound(F As String)
    sndPlaySound F, SND_ASYNC
End Sub


