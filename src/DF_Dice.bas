Attribute VB_Name = "DF_Dice"
Option Explicit

Public Sub Roll_Dice(roll_1 As Integer, roll_2 As Integer, Team As String)
    Dim i As Integer
    Dim shp1 As Shape, shp2 As Shape
    Dim DiePos As Single
    Dim DieLeftPosAdd As Single, DieTopPosAdd As Single, SQAOffset As Single
    Dim Board As Worksheet, ShapeSheet As Worksheet
'   Debug.Print ActiveSheet.Shapes("Dice1").height .width
'   18.41488
 
    Randomize
    DiePos = 189.5
    DieLeftPosAdd = 30
    SQAOffset = 68
    DieTopPosAdd = 0
    For i = 1 To 6
        roll_1 = Int(6 * Rnd) + 1
        roll_2 = Int(6 * Rnd) + 1
    
'        'Set for testing
'        roll_1 = 4
'        roll_2 = 6
        
        Set Board = Sheets("Board")
        Set ShapeSheet = Sheets("Shapes")
        Set shp1 = ShapeSheet.Shapes("DRoll" & roll_1)
        Set shp2 = ShapeSheet.Shapes("DRoll" & roll_2)
        
        On Error Resume Next
        ActiveSheet.Shapes("Dice1").Delete
        ActiveSheet.Shapes("Dice2").Delete
        On Error GoTo 0
'        Call PauseUpdate(500)

        Call CopyShape(ShapeSheet, shp1, 60)
'        DoEvents
        Call PasteShape(Board, 60)
        
        Application.ScreenUpdating = False
        DoEvents
        Application.ScreenUpdating = True
        
        Set shp1 = Board.Shapes("DRoll" & roll_1)
        If Team = "SQA" Then
            DieTopPosAdd = SQAOffset
            Board.Shapes.Range(Array("DRoll" & roll_1 & "_bkg")).Select
            With Selection
                .Name = "colored" 'named changed to avoid double color error
                .ShapeRange.Fill.ForeColor.RGB = RGB(0, 0, 255)
            End With
        End If
        shp1.Name = "Dice1"
        shp1.top = DiePos + DieTopPosAdd
        shp1.left = DiePos
'        Call PauseUpdate(500)

        Call CopyShape(ShapeSheet, shp2, 60)
        Call PasteShape(Board, 60)
        Application.ScreenUpdating = False
        DoEvents
        Application.ScreenUpdating = True
        
        Set shp2 = Board.Shapes("DRoll" & roll_2)
        If Team = "SQA" Then
            Board.Shapes.Range(Array("DRoll" & roll_2 & "_bkg")).Select
            With Selection.ShapeRange.Fill
                .ForeColor.RGB = RGB(0, 0, 255)
            End With
        End If
        shp2.Name = "Dice2"
        shp2.top = DiePos + DieTopPosAdd
        shp2.left = DiePos + DieLeftPosAdd
        
        Board.Range("A1").Select
        Application.CutCopyMode = False 'Remove Focus on shape
        PauseUpdate (5)
    Next i
    Call PlaySound("Dice")
End Sub

Public Sub CopyShape(ShpSheet As Worksheet, shp As Shape, nTry As Integer)
Dim n As Integer
                
    On Error Resume Next

    For n = 1 To nTry
        Err.Clear
        ShpSheet.Shapes(shp.Name).Copy
        If Err.Number = 0 Then Exit For
    Next n
    
    If Err.Number <> 0 Then GoTo errorHandler
    On Error GoTo 0
    Exit Sub
    
errorHandler:
    MsgBox "Something went wrong with Dice shape copy" & vbCrLf & _
    "#loops: " & n, vbCritical
    
End Sub

Public Sub PasteShape(BoardSheet As Worksheet, nTry As Integer)
Dim n As Integer

    On Error Resume Next

    For n = 1 To nTry
        Err.Clear
        BoardSheet.Paste
        If Err.Number = 0 Then Exit For
    Next n
    If Err.Number <> 0 Then GoTo errorHandler
    
    On Error GoTo 0
    Exit Sub
    
errorHandler:
    MsgBox "Something went wrong with Dice shape paste" & vbCrLf & _
    "#loops: " & n, vbCritical
    
End Sub


