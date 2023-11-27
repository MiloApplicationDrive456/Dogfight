Attribute VB_Name = "DF_Sounds"
Option Explicit
Option Base 1

Sub PlaySound(Clip As String)
    If SoundOn Then
        sndPlaySound64 ThisWorkbook.path & "\res\" & Clip & ".wav", &H1
    End If
End Sub

Sub StopSound()
    sndPlaySound64 vbNullString, &H1
End Sub
