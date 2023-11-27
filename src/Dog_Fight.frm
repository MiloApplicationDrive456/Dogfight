VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Dog_Fight 
   Caption         =   "DogFight"
   ClientHeight    =   4980
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3765
   OleObjectBlob   =   "Dog_Fight.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Dog_Fight"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Userform Dog_Fight
'27 Feb 2022
Option Explicit

Private Sub Button_Close_Click()
    Unload Me
End Sub

Private Sub Button_EraseCards_Click()

    Call Clear_All_Cards

End Sub

Private Sub Button_GetCards_Click()
    Call Button_EraseCards_Click
    Call Create_Start_Cards
    Call Place_All_Cards
End Sub

Private Sub Button_RestGame_Click()
    Call ResetGame
End Sub

Private Sub Button_RunSim_Click()

StopGame = False
Call Simulation

End Sub

Private Sub Button_StopGame_Click()
    Call StopSound
    StopGame = True
End Sub

Private Sub Button_TestMove_Click()
Dim Planes() As String

    Call PlaneRotMove(Me.ComboBox_Planes, Me.TextBox_move.Value)
    Call LoadComboBoxPlanes(Planes)
    Me.ComboBox_Planes.List = Planes
End Sub

Private Sub Check_Sound_Click()
    SoundOn = IIf(Me.Check_Sound, True, False)
End Sub

Private Sub ComboBox_Planes_Change()
Dim chrt As ChartObject
Dim Strpath As String
Dim w As Single
Dim h As Single

Sheets("Board").Shapes(Me.ComboBox_Planes).Select
Selection.Copy

'Copy image to UserForm
w = Sheets("Board").Shapes(Me.ComboBox_Planes).Width
h = Sheets("Board").Shapes(Me.ComboBox_Planes).Height
Set chrt = ActiveSheet.ChartObjects.Add(0, 0, w, h)
With chrt
    .Select
    ActiveChart.Paste
    .Border.LineStyle = 0 'no border around chart (and picture)
'    .ShapeRange.Fill.Visible = msoFalse
'    .ShapeRange.Line.Visible = msoFalse
    Strpath = ThisWorkbook.path & "\Temp.jpg"
    .Chart.Export Strpath
    .Delete
End With

Me.Image_plane.Picture = LoadPicture(Strpath)

End Sub

Private Sub Option_Big_Click()
    ActiveWindow.Zoom = 155
End Sub

Private Sub Option_Small_Click()
    ActiveWindow.Zoom = 125
End Sub

Private Sub TextBox_move_Change()
    Me.TextBox_move.Text = UCase(Me.TextBox_move.Text)
End Sub

Private Sub userform_initialize()
'***************************
'        Initialize
'***************************
Dim i As Integer
Dim Planes() As String
    
    'load move matrix
'    MoveMatrixEven = Range("MoveMatrixEven").Value
'    MoveMatrixOdd = Range("MoveMatrixOdd").Value
    Sheets("Board").ScrollArea = "A1:A2"
    Me.Image_plane.BorderStyle = fmBorderStyleNone
    
    StopGame = False
    Call ReadBoard
    Call FindAllPlanes
    Call LoadComboBoxPlanes(Planes)
    
    Call Button_EraseCards_Click
    
    Call Create_Start_Cards
    Call Place_All_Cards
    
    Me.ComboBox_Planes.List = Planes
    
    'Sheet setup
'    Application.ExecuteExcel4Macro "SHOW.TOOLBAR(""Ribbon"",False)"
    Application.DisplayFormulaBar = False
'    ThisWorkbook.Sheets("Board").View.DisplayGridlines = False

'    Call GetMonitorInfo
'    ActiveWindow.Zoom = 155
End Sub

Sub GetMonitorInfo()

    'WMI query
    Dim objWmiInterface As Object
    Dim objWmiQuery As Object
    Dim objWmiQueryItem As Object
    Dim strWQL As String
    'outputs
    Dim strDeviceId As String
    Dim strScreenName As String
    Dim varScreenHeight As Variant
    Dim varScreenWidth As Variant

    'run query
    strWQL = "Select * From Win32_DesktopMonitor"
    Set objWmiInterface = GetObject("winmgmts:root/CIMV2")
    Set objWmiQuery = objWmiInterface.ExecQuery(strWQL)
    'iterate output
    For Each objWmiQueryItem In objWmiQuery
        strDeviceId = objWmiQueryItem.DeviceId
        strScreenName = objWmiQueryItem.Name
        varScreenHeight = objWmiQueryItem.ScreenHeight
        varScreenWidth = objWmiQueryItem.ScreenWidth
        Debug.Print strDeviceId
        Debug.Print strScreenName
        Debug.Print varScreenHeight
        Debug.Print varScreenWidth
        Debug.Print ""
    Next

End Sub
