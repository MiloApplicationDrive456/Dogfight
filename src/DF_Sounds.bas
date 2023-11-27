Attribute VB_Name = "DF_Sounds"
Option Explicit
Option Base 1

Sub PlaySound(Clip As String)
    If SoundOn Then
        sndPlaySound64 "C:\Users\white\Documents\Programming\DFsound\" & Clip & ".wav", &H1
    End If
End Sub

Sub StopSound()
    sndPlaySound64 vbNullString, &H1
End Sub
