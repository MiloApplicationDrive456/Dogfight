Attribute VB_Name = "DF__Simulation"
'module: DF_Simulaiton
'12 March 2022
Option Explicit


Sub Simulation()
'***************************
'
'       Simulation
'
'***************************
'http://www.cpearson.com/excel/PlaySound.aspx
Dim i As Integer, j As Integer, k As Integer, n As Integer, m As Integer
Dim nPlane As Integer, nFoe As Integer
Dim PTurn As Integer    'Turn Odd JAG / even = SQA
Dim PTurnOE As Integer  '0 = even / 1 = odd
Dim Round As Integer
Dim Die(1 To 2) As Integer
Dim ReThinkDie As Integer
Dim AttackDir(1 To 2, 1 To 4) As String    'Translates foes FLRB into NSEW with row col  attack
Dim FoeDir(1 To 2, 1 To 4) As String 'Translates foes FLRB into NSEW attack direction
Dim PlaneAct(1 To 2, 1 To 3) As Integer '1/2 = Dice1/2 ; Foe= 1/2; side 1 to 4
Dim MovSeq As String
Dim AttackTo As String  'Direction of plane attacking, final rotation
Dim FoeDir2 As Integer    'Necessary to correct Dir 5-8 back to 1-4
Dim Dist2Foe(1 To 3) As Integer     '(Distance, Rows, Cols)
Dim EndGame As Boolean
Dim ShootDown As Boolean

EndGame = False
Call InitAAGuns

For i = 1 To 4
'    Inflt(i) = 1
    Inflt(i) = 3
Next i

For PTurn = 1 To 40

    'Switch teams for Friend to Foe on Odd & Even rounds
    PTurnOE = PTurn Mod 2
    
    'Build Friend/Foe Plane Matricies & Count nPlane & nFoe
    Call SelectPlansForRound(PTurnOE, nPlane, nFoe)
    
    'Roll SQA(Blue) or JAG(Red) die
    Call Roll_Dice(Die(1), Die(2), Choose(PTurnOE + 1, "SQA", "JAG"))
    
    'Build Plane & Foe Matricies From JAG(2,6) & SQA(2,6) matricies
    Call BuildFriendFoeMatrix(PTurnOE, nPlane, nFoe)
    
    'Get Position(Foe: Name,R,C), Foe Side Orientations(FoeDir) and Attack Orentations(AttackDir) Matricies
    Call PlaneOrientations(nFoe, FoeDir(), AttackDir())
    
    'Buld Options Matrix
    Call MoveOptionsNS(nPlane, nFoe, Die(), Dist2Foe(), AttackDir(), FoeDir())
    
    'Make Attack Decissions PlaneAct(2,3): Dice [1/2], Attacking Foe [1/2], from Side [1 to 8])
    Call ActionDecision(nPlane, nFoe, PlaneAct())
    
    '  Combat Melees for Attacking Planes
    '____________________________________________________________________
    ShootDown = False
    For i = 1 To nPlane  'Friend Plane
        
        'Rethink Required if i = 2 (Second Plane) and first Plane shot down its Foe:
        If ShootDown Then
            'Hack:  Set die equal to avoid ActionDecision re-write for second plane only
            ReThinkDie = Choose(PlaneAct(1, 1), 2, 1)
            Die(PlaneAct(1, 1)) = Die(ReThinkDie)
            'ReThink:
            Call SelectPlansForRound(PTurnOE, nPlane, nFoe)
            Call BuildFriendFoeMatrix(PTurnOE, nPlane, nFoe)
            Call PlaneOrientations(nFoe, FoeDir(), AttackDir())
            Call MoveOptionsNS(nPlane, nFoe, Die(), Dist2Foe(), AttackDir(), FoeDir())
            Call ActionDecision(nPlane, nFoe, PlaneAct())
            PlaneAct(2, 1) = ReThinkDie 'Insure pointer to correct die
        End If
        
        If PlaneAct(i, 3) <> 0 Then 'Attack move possible, Move to Attack
        
            If PlaneAct(1, 1) = 1 Then  'Nominal Dice
                MovSeq = ValidMove(DieMovN(i, PlaneAct(i, 3)), i)
                
                'Create Alternate Move
                '********  Needs optimzed Step through per MaxBonus!
                If MovSeq = "" Then
                    For n = 8 To 1 Step -1  'Loop through FRLB backwards
                        If Not DieMovN(i, n) = "" Then
                            If Not n = PlaneAct(i, 3) Then
                                MovSeq = ValidMove(DieMovN(i, n), i)
                                If MovSeq <> "" Then
                                    PlaneAct(i, 2) = Fix(n / 5) + 1
                                    PlaneAct(i, 3) = n
                                    Exit For
                                End If
                            End If
                        End If
                    Next n
                    PlaneAct(i, 2) = Fix(n / 5) + 1
                End If
                
                If MovSeq = "" Then PlaneAct(i, 3) = 0
            Else    'Switch Dice
                MovSeq = ValidMove(DieMovS(i, PlaneAct(i, 3)), i)
                
                'Create Alternate Move
                '********  Needs optimzed Step through per MaxBonus!
                If MovSeq = "" Then
                    For n = 8 To 1 Step -1  'Loop through FRLB backwards
                        If Not DieMovS(i, n) = "" Then
                            If Not n = PlaneAct(i, 3) Then
                                MovSeq = ValidMove(DieMovS(i, n), i)
                                If MovSeq <> "" Then
                                    PlaneAct(i, 2) = Fix(n / 5) + 1
                                    PlaneAct(i, 3) = n
                                    Exit For
                                End If
                            End If
                        End If
                    Next n
                End If
                
                If MovSeq = "" Then PlaneAct(i, 3) = 0
            End If
        End If
        
        'Move to Stage Position if no attack possible
        If PlaneAct(i, 3) = 0 Then
            Call Distance(CStr(Plane(i, 1)), CStr(Foe(PlaneAct(i, 2), 1)), Dist2Foe())
            MovSeq = StageMove(Dist2Foe(), Die(PlaneAct(i, 1)), Plane(), i)
        End If
        
        
        Call PointDice(PlaneAct(i, 1))
        Call PlaySound("BiPlane")
        
        'Move Plane
        For j = 1 To Len(MovSeq)
            Call PlaneRotMove(CStr(Plane(i, 1)), Mid(MovSeq, j, 1))
        Next j

        'Must be re-created after move
        Call Find2Planes(nPlane)
        
        FoeDir2 = IIf(PlaneAct(i, 3) < 5, PlaneAct(i, 3), PlaneAct(i, 3) - 4)
        
        If PlaneAct(i, 3) <> 0 Then 'Attack move possible, Turn to Attack
            AttackTo = FindFoeDir(CStr(Foe(PlaneAct(i, 2), 1)), FoeDir2)
            
            If Plane(i, 1) = "" Then
                MsgBox "NO PLANE BEFORE TURN"
                Stop
            End If
            
            'Turn Plane to Attack
            Call TurnPlane(CStr(Plane(i, 1)), AttackTo)
            Call StopSound
            
            '********************************************************
            Call DogFight(CStr(Plane(i, 1)), CStr(Foe(PlaneAct(i, 2), 1)), PlaneAct(i, 3), AttackTo, ShootDown, EndGame)
            If EndGame Then Exit Sub
            '********************************************************
            
            Call Place_All_Cards
            
        End If
        Call HidePointer
        If StopGame Then Exit Sub
    Next i
    
Next PTurn

'Call PauseUpdate(500)

End Sub
