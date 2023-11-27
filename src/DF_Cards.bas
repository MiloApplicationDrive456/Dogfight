Attribute VB_Name = "DF_Cards"
'Module: DF_Cards
'25 Feb 2023
Option Explicit
Sub Create_Start_Cards()
    ''Card Arrays
    'Public Deck_SQ1(), Deck_SQ2(), Deck_JA1(), Deck_JA1() As Variant 'Unused Deck
    'Public FltDeck_SQ1(), FltDeck_SQ2(), FltDeck_JA1(), FltDeck_JA1() As Variant 'InFlight Cards
    'Public DisDeck_SQ1(), DisDeck_SQ2(), DisDeck_JA1(), DisDeck_JA1() As Variant 'InFlight Cards
    'Public DeckCards() As Variant
    'Public FltCards() As Variant
    Dim Temp As String
    Dim TmpArry() As Variant
    Dim TmpArry1() As Variant

    Dim i As Integer, j As Integer, k As Integer
    
    'Set Up 4 decks of 16 cards
    TmpArry = Array("empty", "R", "R", "R", "L", "L", "B1", "B1", "B1", "B2", "B2", "B3", "B3", "B4", "B4", "B5", "B5")
    Deck_SQ1 = TmpArry
    Deck_SQ2 = TmpArry
    Deck_JA1 = TmpArry
    Deck_JA2 = TmpArry
    TmpArry = Array("empty")
    DisDeck_SQ1 = TmpArry
    DisDeck_SQ2 = TmpArry
    DisDeck_JA1 = TmpArry
    DisDeck_JA2 = TmpArry
    Erase TmpArry
    
    'Set Up Array of Arrays
    DeckCards = Array("empty", Deck_SQ1, Deck_SQ2, Deck_JA1, Deck_JA2)
    FltCards = Array("empty", FltDeck_SQ1, FltDeck_SQ2, FltDeck_JA1, FltDeck_JA2)
    DisCards = Array("empty", DisDeck_SQ1, DisDeck_SQ2, DisDeck_JA1, DisDeck_JA1)
    
    ' Shuffle Decks using the Fisher-Yates algorithm
    For i = 1 To 4
        For j = 16 To 2 Step -1
            k = Int((j - 1 + 1) * Rnd + 1)
            Temp = DeckCards(i)(j)
            DeckCards(i)(j) = DeckCards(i)(k)
            DeckCards(i)(k) = Temp
        Next j
    Next i
    
    ' ::Print:: the shuffled deck cards
'    For i = 1 To 4
'        For j = 1 To UBound(DeckCards(i))
'            Debug.Print DeckCards(i)(j) & " ";
'        Next j
'        Debug.Print
'    Next i
'    Debug.Print
    
    '  Draw Flight Cards from Decks
    For i = 1 To 4
        For j = 1 To 4
            ReDim Preserve TmpArry(j)                         'Temp FltCards
            TmpArry(j) = DeckCards(i)(UBound(DeckCards(i)))     'Get Last card of each squadrons deck cards (considerd "top" card)
            TmpArry1 = DeckCards(i)                             'Temp DeckCards
            ReDim Preserve TmpArry1(UBound(DeckCards(i)) - 1)   'Reduce Temp DeckCards by 1
            DeckCards(i) = TmpArry1                             'Save DeckCards now 1 card less
        Next j
        FltCards(i) = TmpArry
        Erase TmpArry
        Erase TmpArry1
    Next i
    ' ::Print:: the remaining deck cards
'    For i = 1 To 4
'        For j = 1 To UBound(DeckCards(i))
'            Debug.Print DeckCards(i)(j) & " ";
'        Next j
'        Debug.Print
'    Next i
'    Debug.Print
    ' ::Print:: the flight cards
'    For i = 1 To 4
'        For j = 1 To UBound(FltCards(i))
'            Debug.Print FltCards(i)(j) & " ";
'        Next j
'        Debug.Print
'    Next i
'    Debug.Print
End Sub

Sub Place_All_Cards()
    'Set Up Squadren Cards
    Dim i As Integer, j As Integer
    Dim shp As Shape
    Dim BoardSheet As Worksheet, ShapeSheet As Worksheet
    Dim Cards As String
    Dim DX, DY As Integer
    
    Call Clear_All_Cards
    Set ShapeSheet = Sheets("Shapes")
    Set BoardSheet = Sheets("Board")
    Application.ScreenUpdating = False
    
    For i = 1 To 4
        Cards = Choose(i, "CJA1", "CJA2", "CSQ1", "CSQ2")
        DX = Choose(i, 0, 4, 0, 4)
        DY = Choose(i, 0, 0, 6, 6)
        
        'Show Flight Cards
        For j = 1 To UBound(FltCards(i))
            Set shp = ShapeSheet.Shapes(Cards & FltCards(i)(j))
            Call CopyShape(ShapeSheet, shp, 15)
            Call PasteShape(BoardSheet, 15)
            With Selection
                .Name = Cards & "_C" & j
                .top = Cells(2 + DY, 13).top + 1 + 2 * DY + (j - 1) * 16
                .left = 5 + Cells(2, 13 + DX).left + (j - 1) * 8
                .Height = 0.64 * 82
                .Width = 0.64 * 54
            End With
        Next j
        
        'Show Remaining Deck
        For j = 1 To UBound(DeckCards(i))
            Set shp = ShapeSheet.Shapes(Cards)
            
            Call CopyShape(ShapeSheet, shp, 15)
            Call PasteShape(BoardSheet, 15)
            
            With Selection
                .Name = Cards & "_Back"
                .top = Cells(2 + DY, 13).top + 1 + 2 * DY
                .left = 5 + Cells(2, 13 + DX + 2).left + 30 + (j - 1) * 2
                .Height = 0.64 * 82
                .Width = 0.64 * 54
            End With
        Next j
        'Show Discard
            For j = 1 To UBound(DisCards(i))
                Set shp = ShapeSheet.Shapes(Cards & DisCards(i)(j))
                
                Call CopyShape(ShapeSheet, shp, 15)
                Call PasteShape(BoardSheet, 15)
                
                With Selection
                    .Name = Cards & "_Dis"
                    .top = Cells(2 + 4 + DY, 13).top + 1 + 2 * DY
                    .left = 5 + Cells(2, 13 + DX).left + (j - 1) * 10
                    .Height = 0.64 * 82
                    .Width = 0.64 * 54
                End With
            Next j
    Next i
    
    ThisWorkbook.Sheets("Board").Range("A20").Select
    Application.ScreenUpdating = True
End Sub

Sub Discard(Deck As Integer, pick As Integer)
'Public DeckCards() As Variant   'Array of Unused Deck Arrays
'Public FltCards() As Variant    'Array of FltDeck Arrays
'Public DisCards() As Variant    'Array of DisDeck Arrays
    Dim i As Integer
    Dim TmpFlt As Variant
    Dim TmpDis As Variant

    TmpFlt = FltCards(Deck)
    TmpDis = DisCards(Deck)
    
    ReDim Preserve TmpDis(UBound(TmpDis) + 1)
    TmpDis(UBound(TmpDis)) = FltCards(Deck)(pick)
    DisCards(Deck) = TmpDis
    For i = pick To UBound(TmpFlt) - 1
        TmpFlt(i) = TmpFlt(i + 1)
    Next i
    ReDim Preserve TmpFlt(UBound(TmpFlt) - 1)
    FltCards(Deck) = TmpFlt
End Sub
Sub CrashPlane(CrashIt As String, EndGame As Boolean)
    Dim i As Integer, PRow As Integer, PCol As Integer
    Dim shpItem As String
    PlaySound ("Crash")
    Call Explosion(CrashIt)
    Call FindPlaneOnBoard(CrashIt, PRow, PCol)
    Board(PRow, PCol) = ""
    Select Case left(CrashIt, 3)
        Case "JA1":
            Inflt(1) = Inflt(1) + 1
            If Inflt(1) < 4 Then Call NewCards(1)
        Case "JA2":
            Inflt(2) = Inflt(2) + 1
            If Inflt(2) < 4 Then Call NewCards(2)
        Case "SQ1":
            Inflt(3) = Inflt(3) + 1
            If Inflt(3) < 4 Then Call NewCards(3)
        Case "SQ2":
            Inflt(4) = Inflt(4) + 1
            If Inflt(4) < 4 Then Call NewCards(4)
    End Select
    ActiveSheet.Shapes(CrashIt).Visible = msoFalse
    
    'Check for dead Squadron
    For i = 1 To 4
        shpItem = Choose(i, "Board_JAG11", "Board_JAG10", "Board_SQD94", "Board_SQD95")
        If Inflt(i) = 4 Then
            Range(shpItem).Interior.Color = RGB(64, 64, 64)
        End If
    Next i
    
    'Check for End Game
    If Inflt(1) * Inflt(2) = 16 Then
        Call GameOver("AlliesWin")
        EndGame = True
    End If
    If Inflt(3) * Inflt(4) = 16 Then
        Call GameOver("GermanyWins")
        EndGame = True
    End If
    
End Sub
Sub SpinPlane(Plane As String)
Dim shp As Shape
Dim Angle As Double
Dim MoveIncr As Integer, i As Integer
    
    Set shp = ActiveSheet.Shapes(Plane)
    MoveIncr = 40
    Angle = 360 / MoveIncr
    PlaySound ("Roll")
    For i = 1 To MoveIncr
        shp.rotation = shp.rotation + Angle
        PauseUpdate (5)
    Next i
End Sub

Sub LoopPlane(Plane As String)
Dim shp As Shape
Dim Angle As Double
Dim MoveIncr As Integer, i As Integer
    
    Set shp = ActiveSheet.Shapes(Plane)
    MoveIncr = 40
    Angle = 360 / MoveIncr
    PlaySound ("Loop")
    For i = 1 To MoveIncr
        shp.rotation = shp.rotation + Angle
        PauseUpdate (5)
    Next i
End Sub

Sub Move_Card_to_Play(shp As Shape, Tnew As Single, Lnew As Single)
Dim NudgeT As Double, NudgeL As Double, Angle As Double
Dim MoveIncr As Integer, i As Integer
    
    MoveIncr = 20
    NudgeT = (Tnew - shp.top) / MoveIncr
    NudgeL = (Lnew - shp.left) / MoveIncr
    Angle = 360 / MoveIncr
    PlaySound ("flipcard")
    For i = 1 To MoveIncr
        shp.top = shp.top + NudgeT
        shp.left = shp.left + NudgeL
        shp.rotation = shp.rotation + Angle
        PauseUpdate (5)
    Next i
    
End Sub

Sub Clear_All_Cards()
Dim shp As Shape
Dim brdsheet As Worksheet
    
    Set brdsheet = ThisWorkbook.Worksheets("Board")
    
    For Each shp In brdsheet.Shapes
        If left(shp.Name, 3) = "CSQ" Or left(shp.Name, 3) = "CJA" Then
            shp.Delete
        End If
    Next shp

End Sub

Sub NewCards(Deck As Integer)
'Public DeckCards() As Variant   'Array of Unused Deck Arrays
'Public FltCards() As Variant    'Array of FltDeck Arrays
'Public DisCards() As Variant    'Array of DisDeck Arrays
    Dim i As Integer
    Dim TmpDeck As Variant
    Dim TmpFlt As Variant
    Dim TmpDis As Variant

    TmpDeck = DeckCards(Deck)
    TmpDis = DisCards(Deck)
    TmpFlt = FltCards(Deck)
    
    'Add all FltCards to Discards
    For i = 1 To UBound(TmpFlt)
        ReDim Preserve TmpDis(UBound(TmpDis) + 1)
        TmpDis(UBound(TmpDis)) = TmpFlt(i)
    Next i
    DisCards(Deck) = TmpDis
    
'    Erase TmpFlt
    ReDim TmpFlt(0 To 4)
    TmpFlt(0) = "empty"

    
    For i = 1 To 4
        TmpFlt(i) = TmpDeck(UBound(TmpDeck))
        ReDim Preserve TmpDeck(UBound(TmpDeck) - 1)
    Next i
    FltCards(Deck) = TmpFlt
    DeckCards(Deck) = TmpDeck

End Sub


