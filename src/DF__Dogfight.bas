Attribute VB_Name = "DF__Dogfight"
Option Explicit

Sub DogFight(Plane As String, Foe As String, Pattack As Integer, AttackTo As String, ShootDown As Boolean, EndGame As Boolean)
'Input:     Attack Array
'Output:    play cards CJA1/2 CSQ1/2
Dim i As Integer
Dim shp As Shape
Dim DeckP As Integer, DeckF As Integer
Dim pickP As Integer, pickF As Integer
Dim FindDeck() As Variant
Dim CTop As Single, CLeft As Single     'Card Top & Left position
Dim DTop As Single                      'Delta Card Top for lower placement
Dim MidField(2) As Single
Dim grid_dx As Single                   'Cell width in
Dim grid_dy As Single
Dim Cheight As Single
Dim Cwidth As Single
'
Plane = FindPlaneName(Plane)
Foe = FindPlaneName(Foe)

Call IncrGridSpace(grid_dx, grid_dy, 1)
Cheight = 58.5  'print ActiveSheet.Shapes("CJA1_C2").height
Cwidth = 38.5   'print ActiveSheet.Shapes("CJA1_C2").width

'Find Board Center
MidField(1) = grid_dx * 5 + Range("B1").left
MidField(2) = grid_dy * 6 + Range("A2").top

'Determine "High/Low" card play placement
If ActiveSheet.Shapes(Plane).top < MidField(2) Then DTop = grid_dy * 4.5
CTop = MidField(2) - grid_dy * 3 + DTop

FindDeck = Array("empty", "JA1", "JA2", "SQ1", "SQ2")
For i = 1 To 4
    If left(Plane, 3) = FindDeck(i) Then DeckP = i
    If left(Foe, 3) = FindDeck(i) Then DeckF = i
Next i
        
    Select Case Pattack 'B = Bigest Burst, S = Smallest Burst, R = Roll, L = Loop
        Case 1, 5:  'Front play largest Burst
            pickP = FindFltCard(DeckP, "B")
            pickF = FindFltCard(DeckF, "B")
            
        Case 2, 3, 6, 7:    'Side play smallest Burst + Roll
            pickP = FindFltCard(DeckP, "S")
            pickF = FindFltCard(DeckF, "R")

        Case 4, 8:          'Rear play smallest Burst + Loop
            pickP = FindFltCard(DeckP, "S")
            pickF = FindFltCard(DeckF, "L")
    End Select

    If Not pickP = 0 Then
    
        'PLAY FRIEND Burst CARD::
        Set shp = ActiveSheet.Shapes("C" & left(Plane, 3) & "_C" & pickP)
        CLeft = MidField(1) - grid_dx * 0.5 - Cwidth
        Call Move_Card_to_Play(shp, CTop, CLeft)
        Call BlazeGuns(Plane, AttackTo, Mid(FltCards(DeckP)(pickP), 2, 1))
        
        'FOE RESPONSE:: Shoot, Loops or Roll
        If Not pickF = 0 Then
        
            'PLAY FOE CARD::
            Set shp = ActiveSheet.Shapes("C" & left(Foe, 3) & "_C" & pickF)
            CLeft = MidField(1) + grid_dx * 0.5
            Call Move_Card_to_Play(shp, CTop, CLeft)
            
            If left(FltCards(DeckF)(pickF), 1) = "B" Then
            
                'BURST::
                AttackTo = Choose(InStr(1, "NSEW", AttackTo), "S", "N", "W", "E")   'FLip direction for defender
                Call BlazeGuns(Foe, AttackTo, Mid(FltCards(DeckF)(pickF), 2, 1))
                
                If Mid(FltCards(DeckP)(pickP), 2, 1) > Mid(FltCards(DeckF)(pickF), 2, 1) Then
                
                    'FOE SHOT DOWN::
                    Call CrashPlane(Foe, EndGame)
                    ShootDown = True
                    If EndGame Then Exit Sub
                    
                    'DISCARD:: Friend
                    Call Discard(DeckP, pickP)
                    
                Else
                    
                    If Mid(FltCards(DeckF)(pickF), 2, 1) > Mid(FltCards(DeckP)(pickP), 2, 1) Then
                    
                        'FRIEND SHOT DOWN::
                        Call CrashPlane(Plane, EndGame)
                        ShootDown = True
                        If EndGame Then Exit Sub
                        
                        'DISCARD:: Foe
                        Call Discard(DeckF, pickF)
                        
                    Else
                    
                        If Mid(FltCards(DeckF)(pickF), 2, 1) = Mid(FltCards(DeckP)(pickP), 2, 1) Then
                    
                            'DISCARD:: Foe & Firend
                            Call Discard(DeckF, pickF)
                            Call Discard(DeckP, pickP)
                            
                        End If
                        
                    End If
                    
                End If
                
            Else
            
                'LOOP/ROLL::
                If FltCards(DeckF)(pickF) = "L" Then
                    Call LoopPlane(Foe)
                Else
                    Call SpinPlane(Foe)
                End If
                
                'DISCARD:: Foe & Firend
                Call Discard(DeckF, pickF)
                Call Discard(DeckP, pickP)
                
            End If
            
        Else
        
            'FOE HAS NO CARDS!
            'FOE SHOT DOWN::
            Call CrashPlane(Foe, EndGame)
            ShootDown = True
            If EndGame Then Exit Sub
            
            'DISCARD:: Friend
            Call Discard(DeckP, pickP)
            
        End If
        
        
        PauseUpdate (100)
    End If
End Sub

