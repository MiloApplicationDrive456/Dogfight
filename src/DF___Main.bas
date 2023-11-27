Attribute VB_Name = "DF___Main"
'Module: DF_Main
'25 November 2023
Option Explicit

'***************************
'
'  Global Arrays Declaired:
'
'***************************

Public StopGame As Boolean
'Board(rows, columns)
Public Board(1 To 14, 1 To 12) As Variant   'Array Map of Board
Public AAGun(1 To 14, 1 To 12) As Variant   'Array Map of AAGuns
'
'Move Matricies()
'column content:
'1) Dist.
'2) Move"(row,col)"
'3) Attack"N,E,S,W"
'4) Code = Dist + Move + Attack
'5->7)  Roll: 1/2, Roll: 2/4 Roll: 3/6

'Public MoveMatrixEven() As Variant          'Array of Even Moves to Attack
'Public MoveMatrixOdd() As Variant           'Array of Odd Moves to Attack
'

'Plane/Foe Matrix
Public Plane(1 To 2, 1 To 3) As Variant    '(Friend Name, Row, Col)
Public Foe(1 To 2, 1 To 3) As Variant      '(Foe Name, Row, Col)

'Plane Matricies(Name, row, col)
Public SQA(1 To 6, 1 To 3) As Variant       'Array of Allied Planes (Name, Row, Col)
Public JAG(1 To 6, 1 To 3) As Variant       'Array of Axis Planes (Name, Row, Col)

Public DieOptN(1 To 2, 1 To 8) As Integer       'Nominal Option Matrix
Public DieOptS(1 To 2, 1 To 8) As Integer      'Switch Option Matrix
Public DieMovN(1 To 2, 1 To 8) As String       'Nominal Move Matrix
Public DieMovS(1 To 2, 1 To 8) As String       'Switch Move Matrix

'Card Arrays
Public Deck_SQ1(), Deck_SQ2(), Deck_JA1(), Deck_JA2() As Variant 'Unused Deck
Public FltDeck_SQ1(), FltDeck_SQ2(), FltDeck_JA1(), FltDeck_JA2() As Variant 'InFlight Cards
Public DisDeck_SQ1(), DisDeck_SQ2(), DisDeck_JA1(), DisDeck_JA2() As Variant 'Dis Cards
Public DeckCards() As Variant   'Array of Unused Deck Arrays
Public FltCards() As Variant    'Array of FltDeck Arrays
Public DisCards() As Variant    'Array of DisDeck Arrays

Public Inflt(1 To 4)    'Index of infight plane e.g. JA1[X]_N begining with 1

Public SoundOn As Boolean

Public AA_SQ(1 To 8, 1 To 2)   'Squadron AA Gun Row and and Col
Public AA_JA(1 To 8, 1 To 2)   'Jagdstaffel AA Gun Row and and Col

'
'Declare to use Sleep function
Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal Milliseconds As LongPtr)
'Declare to play Wav sound files
Declare PtrSafe Function sndPlaySound64 Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long

''Declare Screen Size
'Declare PtrSafe Function GetSystemMetrics32 Lib "User32" Alias "GetSystemMetrics" (ByVal nIndex As Long) As Long


'Public Sub UseClass()
'    Dim plane(6) As New clsPlanes
'    plane(1).Name = "JAG"
'    plane(1).Col = 1
'    plane(1).Row = 1
'    plane(1).Dir = "0"
'    MsgBox "erimjh"
'End Sub


Sub UpdatePlanePos(PlaneName As String, PRow As Integer, PCol As Integer)
'Input:     Plane name and current positon
'Output:    Erases old position from Board
'           Writes plane name to new position
'           Updates Plane Array coords
'
    Dim i As Integer
    Dim OldProw As Integer
    Dim OldPcol As Integer
    
    Call FindPlaneOnBoard(PlaneName, OldProw, OldPcol)
    Board(OldProw, OldPcol) = ""
    
    Board(PRow, PCol) = PlaneName
    
    'Update Plane Array
    If left(PlaneName, 2) = "SQ" Then
        For i = 1 To 6
            If PlaneName = SQA(i, 1) Then
                SQA(i, 2) = PRow
                SQA(i, 3) = PCol
                Exit For
            End If
        Next i
    Else
        If left(PlaneName, 2) = "JA" Then
            For i = 1 To 6
                If PlaneName = JAG(i, 1) Then
                    JAG(i, 2) = PRow
                    JAG(i, 3) = PCol
                    Exit For
                End If
            Next i
        End If
    End If

End Sub

Sub UpdatePlaneAng(PlaneName As String)
'Input:     full plane name
'Output:    Writes plane name to Board
'           Updates Plane Array direcion only
'           row and column are unchanged
'
    Dim i As Integer
    Dim PRow As Integer
    Dim PCol As Integer
    
    Call FindPlaneOnBoard(PlaneName, PRow, PCol)
    Board(PRow, PCol) = PlaneName
    
    'Update Plane Array
    If left(PlaneName, 2) = "SQ" Then
        For i = 1 To 6
            If left(PlaneName, 4) = left(SQA(i, 1), 4) Then
                SQA(i, 1) = PlaneName
                Exit For
            End If
        Next i
    Else
        If left(PlaneName, 2) = "JA" Then
            For i = 1 To 6
                If left(PlaneName, 4) = left(JAG(i, 1), 4) Then
                    JAG(i, 1) = PlaneName
                    Exit For
                End If
            Next i
        End If
    End If

End Sub
Sub ReadBoard()
'Input:     Array Board
'Output:    Fill SaveBoard Range
'***************** Couldn't get to work without loop??!!
Dim i As Integer
Dim j As Integer

For i = 1 To 14
    For j = 1 To 12
        Board(i, j) = Range("Board_Save").Cells(i, j).Value
    Next j
Next i
'PrintBoard
End Sub

Sub ActionDecision(nPlane As Integer, nFoe As Integer, PlaneAct() As Integer)
'Input:     Die Option Matricies (N/S)
'           Move Option Maticies(N/S)
'Output:    Plane Action Matrix (Using Dice 1/2, Attacking Foe 1/2, from Side 1 to 8)
Dim i As Integer
Dim j As Integer
Dim MaxBonus(2, 2) As Integer 'Index of Max & Max
Dim P1D1 As Integer
Dim P2D2 As Integer
Dim P1D2 As Integer
Dim P2D1 As Integer
Dim Dist1 As Integer
Dim Dist2 As Integer
Dim Dist2Foe(3) As Integer
Dim IsSwitch As Boolean

'Find if Nominal or Swithc Die give max bonus
Nominal_Switch:
Erase MaxBonus()
Erase PlaneAct()
P1D1 = 0
P2D2 = 0
P1D2 = 0
P2D1 = 0
For i = 1 To 8
    P1D1 = P1D1 + DieOptN(1, i)
    P2D2 = P2D2 + DieOptN(2, i)
    P1D2 = P1D2 + DieOptS(1, i)
    P2D1 = P2D1 + DieOptS(2, i)
Next i

If P1D1 + P2D2 >= P1D2 + P2D1 Then   'Nominal
    PlaneAct(1, 1) = 1
    PlaneAct(2, 1) = 2
    IsSwitch = False
Else    'Switch
    PlaneAct(1, 1) = 2
    PlaneAct(2, 1) = 1
    IsSwitch = True
End If

' If PlaneAct(1, 1) = 1 Then
' MovSeq = ValidMove(DieMovN(i, PlaneAct(i, 3)), Plane(), i)
                
'Find where Max is  Bonus & Build PlaneAct()
For j = 1 To nPlane      'Plane
    For i = 1 To 8  'Foe1 & Foe2 attack preference
        If PlaneAct(1, 1) = 1 Then  'NOMINAL DICE
            If DieOptN(j, i) > MaxBonus(j, 2) Then
                MaxBonus(j, 1) = i
                MaxBonus(j, 2) = DieOptN(j, i)
            End If
        Else                        'SWITCH DICE
            If DieOptS(j, i) > MaxBonus(j, 2) Then
                MaxBonus(j, 1) = i
                MaxBonus(j, 2) = DieOptS(j, i)
            End If
        End If
    Next i
Next j

'
'  Code to optimize attacks when both attacks are the same
'  Look for optimized max and zero out conflict and recalcualte max bonus
For j = 1 To nPlane
    'Decide attack/approach Foe
    If MaxBonus(j, 1) <> 0 Then 'Plane j has a target
        PlaneAct(j, 2) = IIf(MaxBonus(j, 1) < 5, 1, 2)
    Else    'Approach Nearest
        Dist1 = Distance(CStr(Plane(j, 1)), CStr(Foe(1, 1)), Dist2Foe())
        Dist2 = Distance(CStr(Plane(j, 1)), CStr(Foe(2, 1)), Dist2Foe())
        PlaneAct(j, 2) = IIf(Dist1 <= Dist2, 1, 2)
    End If
    'Decide side of attack, if 0 then Approch not attack
    PlaneAct(j, 3) = MaxBonus(j, 1)
Next j

'Debug.Print plane(1, 1) & " " & plane(2, 1)
'Debug.Print MaxBonus(1, 1), MaxBonus(1, 2)
'Debug.Print MaxBonus(2, 1), MaxBonus(2, 2)
'Debug.Print PlaneAct(1, 1), PlaneAct(1, 2), PlaneAct(1, 3)
'Debug.Print PlaneAct(2, 1), PlaneAct(2, 2), PlaneAct(2, 3)

'Remove second duplicate goal and recalcualte
'But allow both to stage same plane (<>0)
If PlaneAct(1, 3) <> 0 And PlaneAct(1, 3) = PlaneAct(2, 3) Then
    If IsSwitch Then
        DieOptS(2, PlaneAct(1, 3)) = 0
        DieMovS(2, PlaneAct(1, 3)) = ""
        GoTo Nominal_Switch
    Else
        DieOptN(2, PlaneAct(1, 3)) = 0
        DieMovN(2, PlaneAct(1, 3)) = ""
        GoTo Nominal_Switch
    End If
End If

'Check if plane2 has achievable goal and choose best alternate


End Sub
'Sub OptimizeMoves(OptiMove() As Integer)
''Input:     Die Option Matricies (N/S)
''           Move Option Maticies(N/S)
''Output:    Plane Action Matrix (Using Dice 1/2, from Side 1 to 8)
'Dim i As Integer
'Dim j As Integer
'Dim MaxBonus(2, 2) As Integer 'Index of Max & Max
'Dim P1D1 As Integer
'Dim P2D2 As Integer
'Dim P1D2 As Integer
'Dim P2D1 As Integer
'Dim Dist1 As Integer
'Dim Dist2 As Integer
'Dim Dist2Foe(3) As Integer
'Dim IsSwitch As Boolean
'
''Find if Nominal or Switch Die give max bonus
'Nominal_Switch:
'Erase MaxBonus()
'P1D1 = 0
'P2D2 = 0
'P1D2 = 0
'P2D1 = 0
'For i = 1 To 8
'    P1D1 = P1D1 + DieOptN(1, i)
'    P2D2 = P2D2 + DieOptN(2, i)
'    P1D2 = P1D2 + DieOptS(1, i)
'    P2D1 = P2D1 + DieOptS(2, i)
'Next i
'
'End Sub

Sub GameOver(side As String)
    Call PlaySound(side)
    With ActiveSheet.Shapes(side)
        .Visible = msoTrue
        .ZOrder msoBringToFront
    End With
End Sub
