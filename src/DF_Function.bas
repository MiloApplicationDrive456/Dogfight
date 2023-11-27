Attribute VB_Name = "DF_Function"
'Module: DF_Functions
'12 March 2022
Option Explicit
'(rows, columns)

Function FindPlaneName(Plane) As String
'Input:     The left 4 letters of a plane name
'Output:    The full plane name as it is on shape
Dim shp As Shape

For Each shp In ActiveSheet.Shapes
    If left(shp.Name, 4) = left(Plane, 4) Then
        FindPlaneName = shp.Name
        Exit For
    End If
Next shp

End Function
    
Sub IncrGridSpace(grid_dx As Single, grid_dy As Single, Increment As Integer)
'Input:     Cell Height and Width of sheet using C2:C3
'Output:    Fractions of Cell Height and Width

    grid_dx = (Cells(3, 3).left - Cells(3, 2).left) / Increment
    grid_dy = (Cells(3, 3).top - Cells(2, 3).top) / Increment
'    Debug.Print grid_dx, grid_dy
End Sub

Function FindPlane1(Plane As Variant)
'Same as FindPlaneOnBoard but with array
'Input:  plane name
'Output:  plane position on Board
Dim i As Integer
Dim j As Integer
Dim k As Integer

'Error Check missing plane from board
If Plane(1, 1) = "" Or Plane(2, 1) = "" Then
    MsgBox "Error in Function FindPlane1!" & vbCr & _
    "plane1: " & Plane(1, 1) & vbCr & _
    "plane2: " & Plane(2, 1), vbCritical
    End
End If

'Scan Board for plane
For k = 1 To 2 'planes
    For i = 2 To 13 'rows
        For j = 2 To 12 'columns
            If Board(i, j) <> "" And Board(i, j) <> "X" Then
                If left(Board(i, j), 4) = left(Plane(k, 1), 4) Then
                    Plane(k, 2) = i
                    Plane(k, 3) = j
                    'exit i,j loop
                    j = 12
                    i = 13
                End If
            End If
        Next j
    Next i
Next k

End Function



Function FindPlaneOnBoard(Plane As String, PRow As Integer, PCol As Integer)
'Input:  plane name
'Output:  plane position on Board
Dim i As Integer
Dim j As Integer

If Plane = "" Then
    MsgBox "Error in Function FindPlaneOnBoard, plan = <null>!", vbCritical
    End
End If
    
For i = 2 To 13 'rows
    For j = 2 To 12 'columns
        If left(Board(i, j), 4) = left(Plane, 4) Then
            PRow = i
            PCol = j
            Exit Function
        End If
    Next j
Next i

End Function

Function Distance(Plane1 As String, Plane2 As String, Dist2Plane() As Integer) As Integer
'Input: two plane names
'Output: Array(Dist, rows, cols) between plane1 & 2
Dim i As Integer
Dim j As Integer
Dim Planes(1 To 2) As String
Dim Coord(1 To 2, 1 To 2) As Integer

Planes(1) = Plane1
Planes(2) = Plane2

For j = 1 To 2
    If left(Planes(j), 2) = "SQ" Then
        For i = 1 To 6
            If left(Planes(j), 4) = left(SQA(i, 1), 4) Then
                Coord(j, 1) = SQA(i, 2)
                Coord(j, 2) = SQA(i, 3)
            End If
        Next i
    Else
        If left(Planes(j), 2) = "JA" Then
            For i = 1 To 6
                If left(Planes(j), 4) = left(JAG(i, 1), 4) Then
                    Coord(j, 1) = JAG(i, 2)
                    Coord(j, 2) = JAG(i, 3)
                End If
            Next i
        End If
    End If
Next j
Dist2Plane(2) = Coord(2, 1) - Coord(1, 1)
Dist2Plane(3) = Coord(2, 2) - Coord(1, 2)
Dist2Plane(1) = Abs(Dist2Plane(2)) + Abs(Dist2Plane(3))
Distance = Dist2Plane(1)
End Function

'Function FindFoe(plane() As Variant, Foe() As Variant)
''Input:  first 4 char. of plane name
''Output:  plane distance to Foe
'Dim i As Integer
'Dim j As Integer
'Dim Prow As Integer
'Dim Pcol As Integer
'Dim k As Integer
'Dim temp As Variant
'
'Foe(1, 1) = IIf(Left(plane, 1) = "J", "S", "J")
'
'k = 1
'For i = 2 To 13 'rows
'    For j = 2 To 12 'columns
'        If Not Board(i, j) = "" Then
'            If Left(Board(i, j), 4) = plane Then
'                Prow = i
'                Pcol = j
'            Else
'                If Left(Board(i, j), 1) = Left(Foe(1, 1), 1) And Not Right(Board(i, j), 1) = "0" Then
'                    Foe(k, 1) = Board(i, j)
'                    Foe(k, 3) = i
'                    Foe(k, 4) = j
'                    k = k + 1
'                End If
'            End If
'        End If
'    Next j
'Next i
'
''Calculate distances
'Foe(1, 3) = Foe(1, 3) - Prow
'Foe(1, 4) = Foe(1, 4) - Pcol
'Foe(1, 2) = Abs(Foe(1, 3)) + Abs(Foe(1, 4))
'
'Foe(2, 3) = Foe(2, 3) - Prow
'Foe(2, 4) = Foe(2, 4) - Pcol
'Foe(2, 2) = Abs(Foe(2, 3)) + Abs(Foe(2, 4))
'
''check if Foe2 is closer then switch
'If Foe(2, 2) < Foe(1, 2) Then
'    For i = 1 To 4
'        temp = Foe(1, i)
'        Foe(1, i) = Foe(2, i)
'        Foe(2, i) = temp
'    Next i
'End If
'
'End Function
'Function UpdatePlaneBoardPos(plane As String, Prow As Integer, Pcol As Integer)
''Input:     Plane name and current positon
''Output:    Erases old position from Board
''           Writes plane name to new position
''
'    Dim OldProw As Integer
'    Dim OldPcol As Integer
'
'    Call FindPlaneOnBoard(plane, OldProw, OldPcol)
'    Board(OldProw, OldPcol) = ""
'    Board(Prow, Pcol) = plane
''    PrintBoard
'
'End Function



Function PauseUpdate(MS As Integer)
'Put Statement at top of Main:  Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal Milliseconds As LongPtr)
'This doesn't work:    Application.Wait Now + TimeValue("0:00:01") / 2
    Sleep MS   'Miliseconds
    ThisWorkbook.Sheets("Board").Range("A20").Select
    DoEvents
End Function

Function PrintBoard()
'Input:     <None>
'Output:    Debug.Print Board
Dim i As Integer
Dim j As Integer
Dim RowString As String

    For i = LBound(Board, 1) To UBound(Board, 1)
        RowString = ""
        For j = LBound(Board, 2) To UBound(Board, 2)
            Select Case Board(i, j)
            Case "": RowString = RowString + "    "
            Case "X": RowString = RowString + "XXXX"
            Case Else: RowString = RowString + Board(i, j) + " "
            End Select
            If left(Board(i, j), 1) = "G" Then RowString = RowString + " "
            If j <> UBound(Board, 2) Then RowString = RowString + ","
       Next j
'       Debug.Print RowString
    Next i
'    Debug.Print
End Function

Function SaveBoard()
'Input:     Array Board
'Output:    Fill SaveBoard Range
Range("Board_Save") = Board
End Function

Function LoadComboBoxPlanes(Planes() As String)
Dim i As Integer
Dim j As Integer
Dim found As Integer

found = 0
For i = 2 To 13
    For j = 2 To 11
        If left(Board(i, j), 1) = "J" Or left(Board(i, j), 1) = "S" Then
            found = found + 1
            ReDim Preserve Planes(1 To found)
            Planes(found) = Board(i, j)
        End If
    Next j
Next i

End Function

Function MMLookup(MMCode As String, Die As Integer) As String
'Input:  Move Matrix Code & Die role
'Output:  Move
Dim MMcol As Integer
    'PMove = MMLookup(MMCode, Die(k))
    'MoveMatrixEven 2, 4, 6
    'MoveMatrixOdd 1, 3, 5
    
If Die Mod 2 Then   'Odd dice -->  goto Odd Move Matrix
    MMcol = (Die + 1) / 2 + 1
    MMLookup = Application.WorksheetFunction.VLookup(MMCode, Range("MoveMatrixOdd"), MMcol, False)
Else                'Even dice -->  goto Even Move Matrix
    MMcol = Die / 2 + 1
    MMLookup = Application.WorksheetFunction.VLookup(MMCode, Range("MoveMatrixEven"), MMcol, False)
End If

End Function

Function ValidMove(DieMov As String, planeNum As Integer) As String
'Input:  All Move Codes (single or CSV list)
'Output:  Single Valid Move Code

Dim i As Integer
Dim j As Integer
Dim k As Integer
Dim step As String
Dim DX As Integer
Dim DY As Integer
Dim PRow As Integer
Dim PCol As Integer
Dim Codes() As String
Dim Plane2 As String
Dim GunFound As Boolean
           
Plane2 = Choose(planeNum, Plane(2, 1), Plane(1, 1))

    Codes() = Split(DieMov, ",")
    For i = 0 To UBound(Codes)
        GunFound = False
        PRow = Plane(planeNum, 2)
        PCol = Plane(planeNum, 3)
        For j = 1 To Len(Codes(i))
            step = Mid(Codes(i), j, 1)
            'Calculate Step
            Select Case step
                Case "N": DX = 0: DY = -1
                Case "E": DX = 1: DY = 0
                Case "S": DX = 0: DY = 1
                Case "W": DX = -1: DY = 0
            End Select
            
            'Check Ocupancy and other plane
            If Board(PRow + DY, PCol + DX) <> "" Then
                Exit For
            End If
            
            'Check for Enemy AA Gun
            If (PRow + DY) > 10 And left(Plane(1, 1), 2) = "JA" Then
                For k = 1 To 8
                    If (PRow + DY) = AA_SQ(k, 1) And (PCol + DX) = AA_SQ(k, 2) Then
                        GunFound = True
                        Exit For
                    End If
                Next k
            End If
            If (PRow + DY) < 5 And left(Plane(1, 1), 2) = "SQ" Then
                For k = 1 To 8
                    If (PRow + DY) = AA_JA(k, 1) And (PCol + DX) = AA_JA(k, 2) Then
                        GunFound = True
                        Exit For
                    End If
                Next k
            End If
            If GunFound = True Then Exit For
            
            'Update Position
            PRow = PRow + DY
            PCol = PCol + DX
            'Valid Code found?
            If j = Len(Codes(i)) Then
                ValidMove = Codes(i)
                Exit Function
            End If
        Next j
    Next i
End Function
Function ValidMoves(DieMov As String, planeNum As Integer) As String
'Same as Function ValidMove but returns CSV list
'Input:  All Move Codes
'Output: set of all Valid Moves as comma separated list

Dim i As Integer
Dim j As Integer
Dim k As Integer
Dim step As String
Dim DX As Integer
Dim DY As Integer
Dim PRow As Integer
Dim PCol As Integer
Dim Codes() As String
Dim Plane2 As String
Dim GunFound As Boolean
            
Plane2 = Choose(planeNum, Plane(2, 1), Plane(1, 1))

    Codes() = Split(DieMov, ",")
    For i = 0 To UBound(Codes)
        GunFound = False
        PRow = Plane(planeNum, 2)
        PCol = Plane(planeNum, 3)
        For j = 1 To Len(Codes(i))
            step = Mid(Codes(i), j, 1)
            'Calculate Step
            Select Case step
                Case "N": DX = 0: DY = -1
                Case "E": DX = 1: DY = 0
                Case "S": DX = 0: DY = 1
                Case "W": DX = -1: DY = 0
            End Select
            
            'Check Ocupancy and other plane
            If Board(PRow + DY, PCol + DX) <> "" And Board(PRow + DY, PCol + DX) <> Plane2 Then
                Exit For
            End If
            
            'Check for Enemy AA Gun
            If (PRow + DY) > 10 And left(Plane(1, 1), 2) = "JA" Then
                For k = 1 To 8
                    If (PRow + DY) = AA_SQ(k, 1) And (PCol + DX) = AA_SQ(k, 2) Then
                        GunFound = True
                        Exit For
                    End If
                Next k
            End If
            If (PRow + DY) < 5 And left(Plane(1, 1), 2) = "SQ" Then
                For k = 1 To 8
                    If (PRow + DY) = AA_JA(k, 1) And (PCol + DX) = AA_JA(k, 2) Then
                        GunFound = True
                        Exit For
                    End If
                Next k
            End If
            If GunFound = True Then Exit For

            'Update Position
            PRow = PRow + DY
            PCol = PCol + DX
            'Valid Code found?
            If j = Len(Codes(i)) Then
                ValidMoves = ValidMoves & Codes(i) & ","
            End If
        Next j
    Next i
    'Remove last *,"
    If ValidMoves <> "" Then ValidMoves = left(ValidMoves, Len(ValidMoves) - 1)
End Function
Function StageMove(Dist2Foe() As Integer, DiceRoll As Integer, Plane() As Variant, planeNum As Integer) As String
'Input:  Dist2Foe(Dist,row,col), DieRoll, Plane(name,row,col), Plane#
'Output:  Single Valid Move Code
'
Dim i As Integer
Dim j As Integer
Dim k As Integer
Dim m As Integer
Dim n As Integer
Dim p As Integer
Dim q As Integer
Dim r As Integer
Dim step(2 To 6) As Integer

Dim Trans As String
Dim PosX As Integer
Dim PosY As Integer
Dim TargetX As Integer
Dim TargetY As Integer

Dim explore As String
Dim TempMove As String
Dim MovePair As String
Dim DblBck As Boolean
Dim NewDist2Foe As Integer
Dim OldDist2Foe As Integer
'
' move to stage s:
'|   |   | s |   |   |
'|   | s |   | s |   |
'| s |   | X |   | s |
'|   | s |   | s |   |
'|   |   | s |   |   |
' Initialize!!!!

'OldDist2Foe = IIf(Abs(Dist2Foe(1)) = 1, 2, Abs(Dist2Foe(1)))
OldDist2Foe = 999
TargetY = Plane(planeNum, 2) + Dist2Foe(2)
TargetX = Plane(planeNum, 3) + Dist2Foe(3)
Trans = "NESW"

For i = 1 To 5
    If DiceRoll > i Then step(i + 1) = 1
Next i

'Build Possible paths from 4 compas points
For q = 1 To 1 + 3 * step(6)
    For p = 1 To 1 + 3 * step(5)
        For i = 1 To 1 + 3 * step(4)
            For j = 1 To 1 + 3 * step(3)
                For k = 1 To 1 + 3 * step(2)
                    For m = 1 To 4

                        explore = left(CStr(m) + CStr(k) + CStr(j) + CStr(i) + CStr(p) + CStr(q), DiceRoll)
                        DblBck = False
                        
                        'Find doublebacks
                        For r = 2 To DiceRoll
                            MovePair = Mid(explore, r - 1, 2)
                            If (MovePair = "13" Or MovePair = "24" Or MovePair = "31" Or MovePair = "42") Then
                                DblBck = True
                            End If
                        Next r
                        
                        'Find Stage position
                        If Not DblBck Then
                            For n = 1 To 4     'translate 1234 code into compass moves --> NESW
                                explore = Replace(explore, n, Mid(Trans, n, 1))
                            Next n
                            
                            TempMove = ValidMove(explore, planeNum)
                            If TempMove <> "" Then
                                PosX = Plane(planeNum, 3)
                                PosY = Plane(planeNum, 2)
                            
                                For r = 1 To DiceRoll   'Find New Best Temp Position
                                    Select Case Mid(TempMove, r, 1)
                                        Case "N": 'North
                                            PosY = PosY - 1
                                        Case "E": 'East
                                            PosX = PosX + 1
                                        Case "S": 'South
                                            PosY = PosY + 1
                                        Case "W": 'West
                                            PosX = PosX - 1
                                    End Select
                                Next r
                                NewDist2Foe = Abs(PosX - TargetX) + Abs(PosY - TargetY)
                                If NewDist2Foe <= OldDist2Foe Then
                                    StageMove = TempMove
                                    OldDist2Foe = NewDist2Foe
                                End If
                            End If
                        End If
                        
                    Next m
                Next k
            Next j
        Next i
    Next p
Next q
If StageMove = "" Then
    MsgBox "StageMove failed to find valid move for " & Plane(planeNum, 1) & " with Dice Roll of " & DiceRoll, vbCritical
    Stop
End If
End Function

Function FindFoeDir(Foe As String, FoeDir As Integer) As String
'Same as FindFoeDir but single value
'Input:  Name of plane
'Output:  Translate vector NESW --> FRBL
Dim i As Integer
Dim Point As String
Dim FoeDirStr As String

Point = Right(Foe, 1)   'Direction Plane is pointed but coord system with North Down

Select Case Point
    Case "N":
        FoeDirStr = "SWEN"
    Case "E":
        FoeDirStr = "WNSE"
    Case "S":
        FoeDirStr = "NEWS"
    Case "W":
        FoeDirStr = "ESNW"
End Select

FindFoeDir = Mid(FoeDirStr, FoeDir, 1)

End Function
Function FindFltCard(Deck As Integer, Card As String) As Integer
'Public FltDeck_SQ1(), FltDeck_SQ2(), FltDeck_JA1(), FltDeck_JA2() As Variant 'InFlight Cards
'Public FltCards() As Variant    'Array of FltDeck Arrays
Dim FltDeck As String
Dim i As Integer, j As Integer, MaxBurst As Integer, MinBurst As Integer

    FindFltCard = 0
    
    
    Select Case Card
        Case "B":
            For i = 1 To UBound(FltCards(Deck))
                If left(FltCards(Deck)(i), 1) = "B" Then
                    FindFltCard = i
                    MaxBurst = Mid(FltCards(Deck)(i), 2, 1)
                    For j = i To UBound(FltCards(Deck))
                        If left(FltCards(Deck)(j), 1) = "B" Then
                            If Mid(FltCards(Deck)(j), 2, 1) > MaxBurst Then
                                MaxBurst = Mid(FltCards(Deck)(j), 2, 1)
                                FindFltCard = j
                            End If
                        End If
                    Next j
                    Exit For
                End If
            Next i
            
        Case "S":
            For i = 1 To UBound(FltCards(Deck))
                If left(FltCards(Deck)(i), 1) = "B" Then
                    FindFltCard = i
                    MinBurst = Mid(FltCards(Deck)(i), 2, 1)
                    For j = i + 1 To UBound(FltCards(Deck))
                        If left(FltCards(Deck)(j), 1) = "B" Then
                            If Mid(FltCards(Deck)(j), 2, 1) < MinBurst Then
                                MinBurst = Mid(FltCards(Deck)(j), 2, 1)
                                FindFltCard = j
                            End If
                        End If
                    Next j
                    Exit For
                End If
            Next i
            
        Case "L":
            For i = 1 To UBound(FltCards(Deck))
                If FltCards(Deck)(i) = "L" Then
                    FindFltCard = i
                End If
            Next i
            
        Case "R":
            For i = 1 To UBound(FltCards(Deck))
                If FltCards(Deck)(i) = "R" Then
                    FindFltCard = i
                End If
            Next i
            
    End Select

End Function

