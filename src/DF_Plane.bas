Attribute VB_Name = "DF_Plane"
'module: DF_Plane
'5 March 2023
Option Explicit

Sub SelectPlansForRound(PTurnOE As Integer, nPlane As Integer, nFoe As Integer) 'PTurn: 0=even, 1=odd
Dim i As Integer
Dim Squad(1 To 4) As String
Dim SquadAlive(1 To 4) As String

    Squad(1) = "JA1"
    Squad(2) = "JA2"
    Squad(3) = "SQ1"
    Squad(4) = "SQ2"
    nPlane = 0
    Plane(1, 1) = 0
    Plane(2, 1) = 0
    Foe(1, 1) = 0
    Foe(2, 1) = 0

    For i = 1 To 4
        SquadAlive(i) = 0
        If Inflt(i) < 4 Then SquadAlive(i) = 1
    Next i
    
    Select Case PTurnOE
    
        Case 1:  'Odd Germans FRIEND / Allies FOE
            Select Case SquadAlive(1) * 2 + SquadAlive(2) 'convert to binary then integer
                Case 3: 'Both JA Alive
                    Plane(1, 1) = Squad(1) & Inflt(1)
                    Plane(2, 1) = Squad(2) & Inflt(2)
                    nPlane = 2
                Case 2: 'Only JA1 Alive
                    Plane(1, 1) = Squad(1) & Inflt(1)
                    nPlane = 1
                Case 1: 'Only JA2 Alive >> Move to Plane Pos. 1
                    Plane(1, 1) = Squad(2) & Inflt(2)
                    nPlane = 1
            End Select
            
            Select Case SquadAlive(3) * 2 + SquadAlive(4) 'convert to binary then integer
                Case 3: 'Both SQ Alive
                    Foe(1, 1) = Squad(3) & Inflt(3)
                    Foe(2, 1) = Squad(4) & Inflt(4)
                    nFoe = 2
                Case 2: 'Only SQ1 Alive
                    Foe(1, 1) = Squad(3) & Inflt(3)
                    nFoe = 1
                Case 1: 'Only SQ2 Alive >> Move it to Foe Pos. 1
                    Foe(1, 1) = Squad(4) & Inflt(4)
                    nFoe = 1
            End Select

        Case 0:  'even Allies FRIEND / Germans FOE
            Select Case SquadAlive(1) * 2 + SquadAlive(2) 'convert to binary then integer
                Case 3:
                    Foe(1, 1) = Squad(1) & Inflt(1)
                    Foe(2, 1) = Squad(2) & Inflt(2)
                    nFoe = 2
                Case 2:
                    Foe(1, 1) = Squad(1) & Inflt(1)
                    nFoe = 1
                Case 1:
                    Foe(1, 1) = Squad(2) & Inflt(2)
                    nFoe = 1
            End Select
            Select Case SquadAlive(3) * 2 + SquadAlive(4) 'convert to binary then integer
                Case 3:
                    Plane(1, 1) = Squad(3) & Inflt(3)
                    Plane(2, 1) = Squad(4) & Inflt(4)
                    nPlane = 2
                Case 2:
                    Plane(1, 1) = Squad(3) & Inflt(3)
                    nPlane = 1
                Case 1:
                    Plane(1, 1) = Squad(4) & Inflt(4)
                    nPlane = 1
            End Select
    End Select

End Sub

Sub BuildFriendFoeMatrix(PTurnOE As Integer, nPlane As Integer, nFoe As Integer)
'Input:     2x3 Plane Array with only 2x plane ident loaded
'Output:    completed 2x3 Plane Array with full names, rows & columns

'Public plane(1 To 2, 1 To 3) As Variant    '(Friend Name, Row, Col)
'Public Foe(1 To 2, 1 To 3) As Variant      '(Foe Name, Row, Col)

Dim i As Integer
Dim j As Integer

    Select Case PTurnOE
    
        Case 1:  'Odd   JAG is Friend
            For j = 1 To nPlane
                For i = 1 To 6
                    If left(Plane(j, 1), 4) = left(JAG(i, 1), 4) Then
                        Plane(j, 1) = JAG(i, 1)
                        Plane(j, 2) = JAG(i, 2)
                        Plane(j, 3) = JAG(i, 3)
                    End If
                Next i
            Next j
            For j = 1 To nFoe
                For i = 1 To 6
                    If left(Foe(j, 1), 4) = left(SQA(i, 1), 4) Then
                        Foe(j, 1) = SQA(i, 1)
                        Foe(j, 2) = SQA(i, 2)
                        Foe(j, 3) = SQA(i, 3)
                    End If
                Next i
            Next j
        
        Case 0:  'Even SQA is Friend
            For j = 1 To nPlane
                For i = 1 To 6
                    If left(Plane(j, 1), 4) = left(SQA(i, 1), 4) Then
                        Plane(j, 1) = SQA(i, 1)
                        Plane(j, 2) = SQA(i, 2)
                        Plane(j, 3) = SQA(i, 3)
                    End If
                Next i
            Next j
            For j = 1 To nFoe
                For i = 1 To 6
                    If left(Foe(j, 1), 4) = left(JAG(i, 1), 4) Then
                        Foe(j, 1) = JAG(i, 1)
                        Foe(j, 2) = JAG(i, 2)
                        Foe(j, 3) = JAG(i, 3)
                    End If
                Next i
            Next j
    End Select
'Debug.Print plane(1, 1), plane(1, 2), plane(1, 3)
'Debug.Print plane(2, 1), plane(2, 2), plane(2, 3)
End Sub
Sub PlaneOrientations(nFoe As Integer, FoeDir() As String, AttackDir() As String)
'Get Position(Foe: Name,R,C), Foe Side Orientations(FoeDir) and Attack Orentations(AttackDir) Matricies
'    Direction Matricies
'               Attack Orent.       Foe Side Orent.
'Side   F  -->  S   W   N   E       N   E   S   W
'       R  -->  W   N   E   S       E   S   W   N
'       L  -->  E   S   W   N       W   N   E   S
'       B  -->  N   E   S   W       S   W   N   E
'               |   |   |   |       |   |   |   |
'plane Orient:  N   E   S   W       N   E   S   W
'
Dim i As Integer
'    Erase FoeDir            'Not Necessary?
'    Erase AttackDir         'Not Necessary?
    For i = 1 To nFoe
        Call FindAttackDir(CStr(Foe(i, 1)), AttackDir(), i)
        Call FindFoeDirs(CStr(Foe(i, 1)), FoeDir(), i)
    Next i
End Sub
Sub Find2Planes(nPlane As Integer)
'Input:     2x3 Plane Array with only 2x plane ident loaded
'Output:    completed 2x3 Plane Array with full names, rows & columns
'           Number of planes found

Dim i As Integer
Dim j As Integer

For j = 1 To nPlane
    If left(Plane(j, 1), 2) = "SQ" Then
        For i = 1 To 6
            If left(Plane(j, 1), 4) = left(SQA(i, 1), 4) Then
                Plane(j, 1) = SQA(i, 1)
                Plane(j, 2) = SQA(i, 2)
                Plane(j, 3) = SQA(i, 3)
            End If
        Next i
    Else
        If left(Plane(j, 1), 2) = "JA" Then
            For i = 1 To 6
                If left(Plane(j, 1), 4) = left(JAG(i, 1), 4) Then
                    Plane(j, 1) = JAG(i, 1)
                    Plane(j, 2) = JAG(i, 2)
                    Plane(j, 3) = JAG(i, 3)
                End If
            Next i
        End If
    End If
Next j
'Debug.Print plane(1, 1), plane(1, 2), plane(1, 3)
'Debug.Print plane(2, 1), plane(2, 2), plane(2, 3)
End Sub

Sub FindAttackDir(Foe As String, AttackDir() As String, NumP As Integer)
'Input:  Name of plane
'Output:  Find the direction of attack for each face based on plane orientation
Dim i As Integer
Dim Point As String
Dim AttackDirStr As String

Point = Right(Foe, 1)   'Direction Plane is pointed
Select Case Point
    Case "N":
        AttackDirStr = "SWEN"
    Case "E":
        AttackDirStr = "WNSE"
    Case "S":
        AttackDirStr = "NEWS"
    Case "W":
        AttackDirStr = "ESNW"
    Case "0"    'Hack:  For Planes that are on the ground
        AttackDirStr = "SWEN"
End Select

For i = 1 To 4
    AttackDir(NumP, i) = Mid(AttackDirStr, i, 1)
Next i
'Debug.Print AttackDir(1), AttackDir(2), AttackDir(3), AttackDir(4)
End Sub

Sub FindFoeDirs(Foe As String, FoeDir() As String, NumP As Integer)
'Input:  Name of plane
'Output:  Translate vector NESW --> FRBL
Dim i As Integer
Dim Point As String
Dim FoeDirStr As String

Point = Right(Foe, 1)   'Direction Plane is pointed but coord system with North Down

Select Case Point
    Case "N":
        FoeDirStr = "NEWS"
    Case "E":
        FoeDirStr = "ESNW"
    Case "S":
        FoeDirStr = "SWEN"
    Case "W":
        FoeDirStr = "WNSE"
End Select

For i = 1 To 4
    FoeDir(NumP, i) = Mid(FoeDirStr, i, 1)
Next i

End Sub

Function FindAllPlanes() As Integer
'Same as FindPlaneOnBoard but with array
'Input:  Nothing
'Output:  Fill SQA() & JAG() Arrays & return # planes found

Dim i As Integer
Dim j As Integer
Dim incr As Integer

FindAllPlanes = 0
'Scan Board for plane

For i = 2 To 13 'rows
    For j = 2 To 12 'columns
        If left(Board(i, j), 2) = "SQ" Then
            incr = Mid(Board(i, j), 4, 1) + IIf(Mid(Board(i, j), 3, 1) = "2", 3, 0)
            SQA(incr, 1) = Board(i, j)
            SQA(incr, 2) = i
            SQA(incr, 3) = j
            FindAllPlanes = FindAllPlanes + 1
        Else
            If left(Board(i, j), 2) = "JA" Then
                incr = Mid(Board(i, j), 4, 1) + IIf(Mid(Board(i, j), 3, 1) = "2", 3, 0)
                JAG(incr, 1) = Board(i, j)
                JAG(incr, 2) = i
                JAG(incr, 3) = j
                FindAllPlanes = FindAllPlanes + 1
            End If
        End If
    Next j
Next i
'Dump Plane Location Array
'Debug.Print "Found: " & FindAllPlanes & " Planes"
'For i = 1 To 6
'    Debug.Print SQA(i, 1), SQA(i, 2), SQA(i, 3)
'Next i
'For i = 1 To 6
'    Debug.Print JAG(i, 1), JAG(i, 2), JAG(i, 3)
'Next i
End Function
