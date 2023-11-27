Attribute VB_Name = "DF_MoveMatrix_Gen"
'Module: DF_MoveMatrix_Gen
'5 Mar 2022
Option Explicit

Sub Build_Move_Matrix()
Dim EvenArray As Variant
Dim OddArray As Variant
Dim i As Long
Dim j As Integer
Dim k As Integer
Dim m As Integer
Dim p As Integer
Dim q As Integer
Dim n As Long

Dim rMatrix As Integer
Dim moveRC() As String
Dim PrintCol As Integer

Dim DiceRoll As Integer
Dim PosC As Integer
Dim PosR As Integer
Dim newPosC As Integer
Dim newPosR As Integer
Dim TargetC As Integer
Dim TargetR As Integer
Dim direct As Integer
Dim path As String
Dim Path_(1 To 4) As String
Dim step As String
Dim step2 As Integer
Dim step3 As Integer
Dim step4 As Integer
Dim step5 As Integer
Dim step6 As Integer
Dim explore() As String
Dim nSolutions As Integer
Dim nSolutions_(1 To 4) As Integer
Dim Solution_(1 To 4) As String
Dim MovePair As String

EvenArray = Range("Move_Even").Value
OddArray = Range("Move_Odd").Value

For rMatrix = 1 To UBound(EvenArray, 1) Step 4
'For rMatrix = 1 To UBound(OddArray, 1) Step 4

    'split coord string
    '(R,C) = (deltaY,deltaX)
    
    moveRC = Split(Mid(EvenArray(rMatrix, 2), 2, Len(EvenArray(rMatrix, 2)) - 2), ",")
'    moveRC = Split(Mid(OddArray(rMatrix, 2), 2, Len(OddArray(rMatrix, 2)) - 2), ",")
    TargetR = CInt(moveRC(0))   'Row
    TargetC = CInt(moveRC(1))   'Col
    
'    For DiceRoll = 2 To 6 Step 2
    For DiceRoll = 1 To 5 Step 2
        ' Initialize!!!!
        PrintCol = Choose(DiceRoll, 10, 4, 11, 5, 12, 6)
        PosC = 0
        PosR = 0
        n = 0
        step2 = 0
        step3 = 0
        step4 = 0
        step5 = 0
        step6 = 0
        nSolutions = 0
        For i = 1 To 4
            Solution_(i) = ""
        Next i
        
        If DiceRoll > 1 Then step2 = 1
        If DiceRoll > 2 Then step3 = 1
        If DiceRoll > 3 Then step4 = 1
        If DiceRoll > 4 Then step5 = 1
        If DiceRoll > 5 Then step6 = 1
        
        Erase explore()
        
        'Build Possible paths from 4 compas points
        For q = 1 To 1 + 3 * step6
            For p = 1 To 1 + 3 * step5
                For i = 1 To 1 + 3 * step4
                    For j = 1 To 1 + 3 * step3
                        For k = 1 To 1 + 3 * step2
                            For m = 1 To 4
                                n = n + 1
                                ReDim Preserve explore(1 To n)
                                explore(n) = CStr(m) + CStr(k) + CStr(j) + CStr(i) + CStr(p) + CStr(q)
                            Next m
                        Next k
                    Next j
                Next i
            Next p
        Next q
        'Debug.Print explore(n)
        
        For i = 1 To n
            path = ""
            PosC = 0
            PosR = 0
        
            For j = 1 To DiceRoll
                direct = Mid(explore(i), j, 1)
                'check for switchback pairs
                If j > 1 Then
                    MovePair = Mid(explore(i), j - 1, 2)
                    If (MovePair = "13" Or MovePair = "24" Or MovePair = "31" Or MovePair = "42") Then
                        Exit For
                    End If
                End If
                    
                Select Case direct
                    Case 1: 'North move
                        newPosC = PosC
                        newPosR = PosR - 1
                        step = "N"
                    Case 2: 'East move
                        newPosC = PosC + 1
                        newPosR = PosR
                        step = "E"
                    Case 3: 'South move
                        newPosC = PosC
                        newPosR = PosR + 1
                        step = "S"
                    Case 4: 'West move
                        newPosC = PosC - 1
                        newPosR = PosR
                        step = "W"
                End Select
                    
                    If (newPosC = TargetC And newPosR = TargetR) Then
                        Exit For
                    Else
                        'Accept New Position
                        PosC = newPosC
                        PosR = newPosR
                        path = path + step
                    End If
        
                    If Len(path) = DiceRoll And Abs(PosC - TargetC) + Abs(PosR - TargetR) = 1 Then
                        nSolutions = nSolutions + 1
                        'Find facing orientation of attacking plane
                        
                        'Attack to North from South
                        If PosR - TargetR = 1 Then
                            nSolutions_(1) = nSolutions_(1) + 1
                            Solution_(1) = Solution_(1) + path + ","
                        End If
        
                        'Attack to East from West
                        If PosC - TargetC = -1 Then
                            nSolutions_(2) = nSolutions_(2) + 1
                            Solution_(2) = Solution_(2) + path + ","
                        End If
                        
                        'Attack to South from North
                        If PosR - TargetR = -1 Then
                            nSolutions_(3) = nSolutions_(3) + 1
                            Solution_(3) = Solution_(3) + path + ","
                        End If
                        
                        'Attack to West from East
                        If PosC - TargetC = 1 Then
                            nSolutions_(4) = nSolutions_(4) + 1
                            Solution_(4) = Solution_(4) + path + ","
                        End If
                    End If
                Next j
        Next i

        'remove last ","
        On Error Resume Next
        For i = 1 To 4
            Solution_(i) = left(Solution_(i), Len(Solution_(i)) - 1)
        Next i
        On Error GoTo 0
        
        'N, E, S, W solutions
        For i = 1 To 4
            ActiveWorkbook.Sheets("MoveMatrix").Cells(rMatrix + i, PrintCol).Value = Solution_(i)
        Next i
        
    Next DiceRoll
Next rMatrix
'Debug.Print "Topo R/C:  "; TargetY - 1; "/"; TargetX - 1; " DiceRoll: "; DiceRoll
'Debug.Print "Solutions Found: "; nSolutions
'Debug.Print "N "; nSolutions_(1)
'Debug.Print Solution_(1)
'Debug.Print "E "; nSolutions_(2)
'Debug.Print Solution_(2)
'Debug.Print "S "; nSolutions_(3)
'Debug.Print Solution_(3)
'Debug.Print "W "; nSolutions_(4)
'Debug.Print Solution_(4)
'End Sub
End Sub


