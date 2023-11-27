Attribute VB_Name = "DF_Move"
'Module: Move
Option Explicit

Sub PlaneRotMove(Plane As String, move As String)
'Inupt      plane: First 4 letters of Plane = Plane Ident
'           move: String of compass direcions
'Output:    Truns in direcion of Move & Moves Plane
'           Saves Board position Matrix
'           Saves updated Plane Matix

    Dim i As Integer
    Dim DX As Integer       'Delta X
    Dim DY As Integer       'Delta Y
    Dim gdx As Single       'incremental grid Delta X
    Dim gdy As Single       'incremental grid Delta Y
    Dim planeAng As Single  'Plane Angle
    Dim PRow As Integer     'Plane row
    Dim PCol As Integer     'Plane Column
    Dim Prow_start As Integer   'Plane row initial
    Dim Pcol_start As Integer   'Plane Column initial
    Dim step As String
    Dim shp As Shape
    
'   Find plane from existing shapes on board
    Call FindPlaneOnBoard(Plane, Prow_start, Pcol_start)
    PRow = Prow_start
    PCol = Pcol_start
    
    For i = 1 To Len(move)
        Plane = FindPlaneName(Plane)    'Find Full plane name from ident
        step = Mid(move, i, 1)
        'rotate & move
        Call TurnPlane(Plane, step)
        Call MovePlane(Plane, step, PRow, PCol)
    Next i
    
    Call SaveBoard
    Call FindAllPlanes
End Sub

Sub MovePlane(Plane As String, step As String, PRow As Integer, PCol As Integer)
'Inupt      plane: Complete 6char. plane name
'           Step: Compass direction N,S,E,W
'Output:    Prow, Pcol: New row and column
'           Moves shape to new position

    Dim i As Integer
    Dim DX As Integer   'Delta X
    Dim DY As Integer   'Delta Y
    Dim gdx As Single   'incremental grid Delta X
    Dim gdy As Single   'incremental grid Delta Y
    Dim incr As Integer '#steps in move (more = slower)


    incr = 15
    Call IncrGridSpace(gdx, gdy, incr)  'Move increment = GridSpace/Incr
    
    'Calculate Step
    Select Case step
        Case "N": DX = 0: DY = -1
        Case "E": DX = 1: DY = 0
        Case "S": DX = 0: DY = 1
        Case "W": DX = -1: DY = 0
    End Select
    
    'Check Ocupancy
        If Board(PRow + DY, PCol + DX) <> "" Then
            MsgBox "Illegal Move: " & step & vbCr & "Program Stopped!", vbCritical
            Stop
        End If
    'Update Position
    PRow = PRow + DY
    PCol = PCol + DX
    
    'Half Step
    For i = 1 To incr - 1
        ActiveSheet.Shapes(Plane).left = ActiveSheet.Shapes(Plane).left + DX * gdx
        ActiveSheet.Shapes(Plane).top = ActiveSheet.Shapes(Plane).top + DY * gdy
        Call PauseUpdate(5)
    Next i
    
    'Full Step
    ActiveSheet.Shapes(Plane).left = Cells(PRow, PCol).left
    ActiveSheet.Shapes(Plane).top = Cells(PRow, PCol).top
            
    'Update Board
    Call UpdatePlanePos(Plane, PRow, PCol)
    
    Call PauseUpdate(5)
End Sub

Sub TurnPlane(Plane As String, step As String)
'Inupt:     First 4 letters of Plane = Plane Ident
'Output:    Turns Plane shape on board per Step

    Dim Angle As Single
    Dim planeAng As Single
    Dim PlaneTurn As Single
    Dim NewAng As Single
    Dim AngStep As Integer
    Dim i As Single
    Dim plane_new_name As String
    Dim TurnRate As Integer
    
    TurnRate = 5
    
    'Calculate Turn
    Select Case step
        Case "N": Angle = 0
        Case "E": Angle = 90
        Case "S": Angle = 180
        Case "W": Angle = 270
    End Select
    planeAng = ActiveSheet.Shapes(Plane).rotation
    PlaneTurn = Angle - planeAng
    
    'Max trun 180°
    If PlaneTurn > 180 Then
        PlaneTurn = PlaneTurn - 360
    End If
    If PlaneTurn < -180 Then
        PlaneTurn = PlaneTurn + 360
    End If
    
    'Rotate
    If Plane = "" Then
        MsgBox "NO PLANE BEFORE TURN" & vbCr & _
        "step = " & step
        Stop
    End If
    AngStep = Abs(Round(PlaneTurn / TurnRate))
    For i = 1 To AngStep
        ActiveSheet.Shapes(Plane).rotation = planeAng + PlaneTurn / AngStep
        planeAng = planeAng + PlaneTurn / AngStep
        Call PauseUpdate(5)
    Next i
    
    'Rename & Update Board
    plane_new_name = left(Plane, Len(Plane) - 1) & step
    ActiveSheet.Shapes(Plane).Name = plane_new_name
    Plane = plane_new_name
    Call UpdatePlaneAng(plane_new_name)
    Call SaveBoard
End Sub

Sub BlazeGuns(Plane As String, AttackTo As String, Burst As Integer)
Dim i As Integer
Dim DLeft As Integer
Dim DTop As Integer
Dim Angle As Integer
Dim grid_dx As Single
Dim grid_dy As Single

Plane = FindPlaneName(Plane)
Call IncrGridSpace(grid_dx, grid_dy, 1)

Select Case AttackTo    'Positions adjusted by hand
    Case "N":
        Angle = 0
        DTop = -0.57 * grid_dy
        DLeft = 0.37 * grid_dx
    Case "E":
        Angle = 90
        DTop = 0.15 * grid_dy
        DLeft = 1# * grid_dx
    Case "S":
        Angle = 180
        DTop = 0.86 * grid_dy
        DLeft = 0.41 * grid_dx
    Case "W":
        Angle = 270
        DTop = 0.16 * grid_dy
        DLeft = -0.29 * grid_dx
End Select
    
    With ActiveSheet.Shapes("GunBlaze")
        .Visible = False
        .left = ActiveSheet.Shapes(Plane).left + DLeft
        .top = ActiveSheet.Shapes(Plane).top + DTop
        .rotation = Angle
        Call PlaySound("MachineGun")
        For i = 1 To 3 * Burst
            .Visible = True
            Call PauseUpdate(30)
            .Visible = False
            Call PauseUpdate(30)
        Next i
        Call StopSound
    End With
End Sub
Sub Explosion(Plane As String)
Dim i As Integer
Dim DLeft As Integer
Dim DTop As Integer
Dim Angle As Integer
Dim grid_dx As Single
Dim grid_dy As Single

Plane = FindPlaneName(Plane)
Call IncrGridSpace(grid_dx, grid_dy, 1)

    With ActiveSheet.Shapes("Explosion")
        .Visible = False
        .left = ActiveSheet.Shapes(Plane).left
        .top = ActiveSheet.Shapes(Plane).top
        For i = 1 To 6
            .Visible = True
            Call PauseUpdate(90)
            .Visible = False
            Call PauseUpdate(90)
        Next i
    End With
End Sub

Sub PointDice(Dice As Integer)
Dim i As Integer
Dim DLeft As Integer
Dim DTop As Integer
Dim Angle As Integer
Dim grid_dx As Single
Dim grid_dy As Single

Call IncrGridSpace(grid_dx, grid_dy, 1)
DTop = -0.55 * grid_dy
DLeft = -0.6 * grid_dx

With ActiveSheet.Shapes("DicePointer")
    .Visible = False
    .left = ActiveSheet.Shapes("Dice" & Dice).left + DLeft
    .top = ActiveSheet.Shapes("Dice" & Dice).top + DTop
    .rotation = 45
    .Visible = True
    Call PauseUpdate(500)
End With
    
End Sub

Sub HidePointer()
    ActiveSheet.Shapes("DicePointer").Visible = False
End Sub

Sub HideGunBlaze()
    ActiveSheet.Shapes("GunBlaze").Visible = False
End Sub
