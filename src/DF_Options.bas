Attribute VB_Name = "DF_Options"
Option Explicit

Sub MoveOptionsNS(nPlane As Integer, nFoe As Integer, Die() As Integer, Dist2Foe() As Integer, AttackDir() As String, FoeDir() As String)
'Input:
'Output:    DieOpt & DiMov Matricies

Dim i As Integer, j As Integer, k As Integer, m As Integer, n As Integer
Dim MMCode As String
Dim PMove As String
Dim AddOpt As Integer       'integer offset for DieOpt/Mov columns 5-8
Dim FoeR As Integer         'Foe Row
Dim FoeC As Integer         'Foe Column
Dim AttackBonus As Integer  'Attack Optimization

Erase DieOptN
Erase DieOptS
Erase DieMovN
Erase DieMovS

For i = 1 To nPlane  'Friend Plane
    For j = 1 To nFoe 'Foe Plane
        For k = 1 To 2  'Dice
            'Check Condition for Attack with each dice
            Call Distance(CStr(Plane(i, 1)), CStr(Foe(j, 1)), Dist2Foe())
            If (Dist2Foe(1) - Die(k)) Mod 2 And (Die(k) - Dist2Foe(1)) >= -1 Then
            
                'Search for moves for Plane(i) to Foe(j) with Dice(k) per Distance()
                For n = 1 To 4  'Loop through FRLB
                    AttackBonus = Choose(n, 1, 2, 2, 3)
'                    AttackBonus = Choose(n, 3, 2, 2, 1)
                    'Build Move Matrix Code
                    MMCode = Dist2Foe(1) & "(" & Dist2Foe(2) & "," & Dist2Foe(3) & ")" & AttackDir(j, n)
                    PMove = MMLookup(MMCode, Die(k))
                    
                    If PMove <> "" Then
                        AddOpt = IIf(j = 2, 4, 0)
                        If ((i + k) Mod 2) = 0 Then 'Plane+Dice = even = Not Switched!
                            DieOptN(i, n + AddOpt) = 1 * AttackBonus
                            DieMovN(i, n + AddOpt) = PMove
                        Else    'Switched!
                            DieOptS(i, n + AddOpt) = 1 * AttackBonus
                            DieMovS(i, n + AddOpt) = PMove
                        End If
                    End If
                    
                Next n
            End If
        Next k
    Next j
    
    'Zero out board occupied squared for possible attaks from plane i
    'Could Optimize code to zero out only if non-zero entry?
    For j = 1 To nFoe  'Foe Plane
        FoeR = Foe(j, 2)
        FoeC = Foe(j, 3)
'            Call FindFoeDirs(CStr(Foe(j, 1)), FoeDir())
        For n = 1 To 4  'Loop through FRLB
        m = n + (j - 1) * 4 'Column in DieOpt Matricies
            Select Case FoeDir(j, n)
                Case "N": 'DX = 0: DY = -1
                    If Board(FoeR - 1, FoeC) <> "" And Board(FoeR - 1, FoeC) <> Plane(i, 1) Then
                        For k = 1 To 2
                            DieOptN(k, m) = 0
                            DieMovN(k, m) = ""
                            DieOptS(k, m) = 0
                            DieMovS(k, m) = ""
                        Next k
                    End If
                Case "E": 'DX = 1: DY = 0
                    If Board(FoeR, FoeC + 1) <> "" And Board(FoeR, FoeC + 1) <> Plane(i, 1) Then
                        For k = 1 To 2
                            DieOptN(k, m) = 0
                            DieMovN(k, m) = ""
                            DieOptS(k, m) = 0
                            DieMovS(k, m) = ""
                        Next k
                    End If
                Case "S": 'DX = 0: DY = 1
                    If Board(FoeR + 1, FoeC) <> "" And Board(FoeR + 1, FoeC) <> Plane(i, 1) Then
                        For k = 1 To 2
                            DieOptN(k, m) = 0
                            DieMovN(k, m) = ""
                            DieOptS(k, m) = 0
                            DieMovS(k, m) = ""
                        Next k
                    End If
                Case "W": 'DX = -1: DY = 0
                    If Board(FoeR, FoeC - 1) <> "" And Board(FoeR, FoeC - 1) <> Plane(i, 1) Then
                        For k = 1 To 2
                            DieOptN(k, m) = 0
                            DieMovN(k, m) = ""
                            DieOptS(k, m) = 0
                            DieMovS(k, m) = ""
                        Next k
                    End If
            End Select
        Next n
    Next j
    
    'Zero out Routes over edges, enemy guns and enemy planes
    'Keep routs over partner plane
    For j = 1 To nFoe  'Foe Plane
        For m = 1 To 8  'Loop through Foe1&2 Moves
            If DieOptN(i, m) <> 0 Then
                DieMovN(i, m) = ValidMoves(DieMovN(i, m), i)
                If DieMovN(i, m) = "" Then DieOptN(i, m) = 0
            End If
            If DieOptS(i, m) <> 0 Then
                 DieMovS(i, m) = ValidMoves(DieMovS(i, m), i)
                If DieMovS(i, m) = "" Then DieOptS(i, m) = 0
            End If
        Next m
    Next j
    
Next i
'Dump Option Matricies
'Debug.Print DieOptN(1, 1), DieOptN(1, 2), DieOptN(1, 3), DieOptN(1, 4), DieOptN(1, 5), DieOptN(1, 6), DieOptN(1, 7), DieOptN(1, 8)
'Debug.Print DieOptN(2, 1), DieOptN(2, 2), DieOptN(2, 3), DieOptN(2, 4), DieOptN(2, 5), DieOptN(2, 6), DieOptN(2, 7), DieOptN(2, 8)
'Debug.Print DieOptS(1, 1), DieOptS(1, 2), DieOptS(1, 3), DieOptS(1, 4), DieOptS(1, 5), DieOptS(1, 6), DieOptS(1, 7), DieOptS(1, 8)
'Debug.Print DieOptS(2, 1), DieOptS(2, 2), DieOptS(2, 3), DieOptS(2, 4), DieOptS(2, 5), DieOptS(2, 6), DieOptS(2, 7), DieOptS(2, 8)

End Sub
