Attribute VB_Name = "DF_Setup"
'Module: DF_Board_Setup
Option Explicit

Function LoadBoard()
'(rows, columns)
Dim i As Integer
Dim j As Integer
Dim rowsting As String
Dim Guns() As Variant

Guns = Array("Hit", "Hit", "Miss", "Miss")

'
'Definitions:
' Board edge:       X
' Gun Hit/Mis:      GH, GM
' Jagdstaffel 11:   J
' Jagdstaffel 10:   K
' Squadron 94:      S
' Squadron 95:      T
' Plane#:           1,2,3
' Plane flying:     P
' Plane Landed:     L
' Direction:        N,S,E,W
' Examples:         PJ1N, PS2E, LT3

Erase Board
    'Load Edges
    For i = 1 To 14
        Board(i, 1) = "X"
        Board(i, 12) = "X"
    Next i
    For i = 2 To 11
        Board(1, i) = "X"
        Board(14, i) = "X"
    Next i
    Board(2, 2) = "X"
    Board(2, 11) = "X"
    Board(13, 2) = "X"
    Board(13, 11) = "X"
    
    'Load Planes
    Board(3, 3) = "JA11_0"
    Board(2, 3) = "JA13_0"
    Board(3, 2) = "JA12_0"
    
    Board(2, 10) = "JA23_0"
    Board(3, 10) = "JA21_0"
    Board(3, 11) = "JA22_0"
    
    Board(12, 3) = "SQ11_0"
    Board(12, 2) = "SQ12_0"
    Board(13, 3) = "SQ13_0"
    
    Board(12, 10) = "SQ21_0"
    Board(13, 10) = "SQ23_0"
    Board(12, 11) = "SQ22_0"

    'Anti-Aircraft Guns
    Guns = Resample(Guns)
    AAGun(4, 2) = Guns(1)
    AAGun(4, 3) = Guns(2)
    AAGun(3, 4) = Guns(3)
    AAGun(2, 4) = Guns(0)
    
    Guns = Resample(Guns)
    AAGun(2, 9) = Guns(1)
    AAGun(3, 9) = Guns(2)
    AAGun(4, 10) = Guns(3)
    AAGun(4, 11) = Guns(0)
    
    Guns = Resample(Guns)
    AAGun(13, 9) = Guns(1)
    AAGun(12, 9) = Guns(2)
    AAGun(11, 10) = Guns(3)
    AAGun(11, 11) = Guns(0)
    
    Guns = Resample(Guns)
    AAGun(11, 2) = Guns(1)
    AAGun(11, 3) = Guns(2)
    AAGun(12, 4) = Guns(3)
    AAGun(13, 4) = Guns(0)
    
'Print Board
    For i = LBound(Board, 1) To UBound(Board, 1)
        rowsting = ""
        For j = LBound(Board, 2) To UBound(Board, 2)
            Select Case Board(i, j)
            Case "": rowsting = rowsting + "    "
            Case "X": rowsting = rowsting + "XXXX"
            Case Else: rowsting = rowsting + Board(i, j) + " "
            End Select
            If left(Board(i, j), 1) = "G" Then rowsting = rowsting + " "
            If j <> UBound(Board, 2) Then rowsting = rowsting + ","
       Next j
'       Debug.Print rowsting
    Next i
'    Debug.Print
End Function

Function Resample(data_vector() As Variant) As Variant()
    Dim shuffled_vector() As Variant
    shuffled_vector = data_vector
    Dim i As Long
    For i = UBound(shuffled_vector) To LBound(shuffled_vector) Step -1
        Dim t As Variant
        t = shuffled_vector(i)
        Dim j As Long
        j = Application.RandBetween(LBound(shuffled_vector), UBound(shuffled_vector))
        shuffled_vector(i) = shuffled_vector(j)
        shuffled_vector(j) = t
    Next i
    Resample = shuffled_vector
End Function

Sub ResetGame()
'Input: Nothing
'Output: Places all markers at start positions with correct names and sizes
Dim i As Integer
Dim shp As Shape
Dim shpItem As String
'name, row, column, orientation
Dim GA(1 To 16, 1 To 3) As Variant  'Gun Array
Dim PosLeft As Single
Dim PosTop As Single
Dim PosRot As Single
Dim gdx As Single
Dim gdy As Single
Dim PInc As Integer 'retuns 0 for i = 1-3 & 1 for i = 4-6
Dim GInc As Integer 'retuns 0 for i = 1-4 & 1 for i = 5-8
    
'Clear Dice
On Error Resume Next
    ActiveSheet.Shapes("Dice1").Delete
    ActiveSheet.Shapes("Dice2").Delete
On Error GoTo 0
        
Call IncrGridSpace(gdx, gdy, 1)
Call LoadBoard

'Set Hidden Shape Size and Positons
For i = 1 To 5
    Set shp = ActiveSheet.Shapes(Choose(i, "GermanyWins", "AlliesWin", "Explosion", "GunBlaze", "DicePointer"))
    
    With shp
        .LockAspectRatio = msoFalse
        .Visible = msoFalse
        .rotation = 0
        
        Select Case i
        
        Case 1, 2   'GermanyWins, AlliesWin
            .Height = 143.2     '2in
            .Width = 286.5      '4in
            .left = 261.5       'center
            .top = 163          'center
        Case 3      'Explosion
            .Height = 37
            .Width = 37
            .left = 55          'Start only
            .top = 125          'Start only
        Case 4      'GunBlaze
            .Height = 22
            .Width = 10
            .left = 55          'Start only
            .top = 175          'Start only
        Case 5      'DicePointer
            .Height = 13.6       '0.19in
            .Width = 27.2       '0.38in
            .left = 55          'Start only
            .top = 225          'Start only
            .rotation = 45
        End Select

        .LockAspectRatio = msoTrue
    End With
Next i

For Each shp In ActiveSheet.Shapes

    For i = 1 To 8  '1 to 6 for planes, 1 to 8 for Guns
        '____________Planes____________________________________
        If i < 7 Then
            PInc = Fix(i / 4)   'retuns 0 for i = 1-3 & 1 for i = 4-6
            JAG(i, 1) = "JA" & 1 + PInc & i - 3 * PInc & "_0"
            JAG(i, 2) = Choose(i, 3, 3, 2, 3, 3, 2)
            JAG(i, 3) = Choose(i, 3, 2, 3, 10, 11, 10)
            SQA(i, 1) = "SQ" & 1 + PInc & i - 3 * PInc & "_0"
            SQA(i, 2) = Choose(i, 12, 12, 13, 12, 12, 13)
            SQA(i, 3) = Choose(i, 3, 2, 3, 10, 11, 10)
            'place,name & rotate
            If left(shp.Name, 4) = left(JAG(i, 1), 4) Then
                shp.Name = JAG(i, 1)
                PosLeft = Cells(JAG(i, 2), JAG(i, 3)).left
                PosTop = Cells(JAG(i, 2), JAG(i, 3)).top
                PosRot = 135 + PInc * 90
                Call RePosSize(shp, PosLeft, PosTop, PosRot, gdy, gdx)
            Else
                If left(shp.Name, 4) = left(SQA(i, 1), 4) Then
                    shp.Name = SQA(i, 1)
                    PosLeft = Cells(SQA(i, 2), SQA(i, 3)).left
                    PosTop = Cells(SQA(i, 2), SQA(i, 3)).top
                    PosRot = 45 - PInc * 90
                    Call RePosSize(shp, PosLeft, PosTop, PosRot, gdy, gdx)
                End If
            End If

        End If
        '____________Guns____________________________________
        GInc = Fix(i / 5)
        GA(i, 1) = "JA" & 1 + GInc & "G_" & i - 4 * GInc
        GA(i, 2) = Choose(i, 4, 4, 3, 2, 2, 3, 4, 4)
        GA(i, 3) = Choose(i, 2, 3, 4, 4, 9, 9, 10, 11)
        GA(i + 8, 1) = "SQ" & 1 + GInc & "G_" & i - 4 * GInc
        GA(i + 8, 2) = Choose(i, 11, 11, 12, 13, 13, 12, 11, 11)
        GA(i + 8, 3) = Choose(i, 2, 3, 4, 4, 9, 9, 10, 11)
        'place & rotate
        If shp.Name = GA(i, 1) Then
            PosLeft = Cells(GA(i, 2), GA(i, 3)).left
            PosTop = Cells(GA(i, 2), GA(i, 3)).top
            PosRot = 0
            Call RePosSize(shp, PosLeft, PosTop, PosRot, gdy, gdx)
        Else
            If shp.Name = GA(i + 8, 1) Then
                PosLeft = Cells(GA(i + 8, 2), GA(i + 8, 3)).left
                PosTop = Cells(GA(i + 8, 2), GA(i + 8, 3)).top
                PosRot = 0
                Call RePosSize(shp, PosLeft, PosTop, PosRot, gdy, gdx)
            End If
        End If
    Next i
Next shp

Call SaveBoard
Call HidePointer
Call HideGunBlaze
Call Clear_All_Cards
Call Place_All_Cards

'Reset Card Area Backgrounds
Range("Board_JAG11").Interior.Color = RGB(240, 220, 220)
Range("Board_Jag10").Interior.Color = RGB(255, 255, 200)
Range("Board_SQD94").Interior.Color = RGB(180, 220, 230)
Range("Board_SQD95").Interior.Color = RGB(200, 255, 255)

''Reset Card Area Bkgrd colors
'With ActiveSheet.Shapes("Card Area").GroupItems
'    .Item("Jag11").Fill.ForeColor.RGB = RGB(240, 220, 220)
'    .Item("Jag10").Fill.ForeColor.RGB = RGB(255, 255, 200)
'    .Item("94th").Fill.ForeColor.RGB = RGB(180, 220, 230)
'    .Item("95th").Fill.ForeColor.RGB = RGB(200, 255, 255)
'End With

ActiveSheet.Shapes("DicePointer").Visible = False

End Sub

Sub RePosSize(shp As Shape, PosLeft As Single, PosTop As Single, PosRot As Single, ht As Single, wi As Single)

    With shp
        .left = PosLeft
        .top = PosTop
        .rotation = PosRot
        .LockAspectRatio = msoFalse
        .Height = ht
        .Width = wi
        .LockAspectRatio = msoTrue
        .Visible = msoTrue
    End With
    
End Sub
