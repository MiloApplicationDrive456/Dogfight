Attribute VB_Name = "Module1"
Sub Macro1()
Attribute Macro1.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro1 Macro
'

'
    ActiveSheet.Shapes.Range(Array("Jag11")).Select
    With Selection.ShapeRange.Fill
        .Visible = msoTrue
        .ForeColor.ObjectThemeColor = msoThemeColorBackground1
        .ForeColor.TintAndShade = 0
        .ForeColor.Brightness = -0.150000006
        .Transparency = 0
        .Solid
    End With
    With Selection.ShapeRange.Fill
        .Visible = msoTrue
        .ForeColor.ObjectThemeColor = msoThemeColorText1
        .ForeColor.TintAndShade = 0
        .ForeColor.Brightness = 0.5
        .Transparency = 0
        .Solid
    End With
    ActiveWindow.SmallScroll Down:=33
    Application.DisplayFormulaBar = False
    ActiveSheet.Shapes.Range(Array("94th")).Select
    ActiveSheet.Shapes.Range(Array("95th")).Select
    ActiveWorkbook.Save
    ActiveSheet.Shapes.Range(Array("Jag11")).Select
    Application.DisplayFormulaBar = True
    ActiveSheet.Shapes.Range(Array("95th")).Select
End Sub
Sub Macro2()
Attribute Macro2.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro2 Macro
'

'
    ActiveSheet.Shapes.Range(Array("95th")).Select
    With Selection.ShapeRange.Fill
        .Visible = msoTrue
        .ForeColor.ObjectThemeColor = msoThemeColorText1
        .ForeColor.TintAndShade = 0
        .ForeColor.Brightness = 0.25
        .Transparency = 0
        .Solid
    End With
End Sub
Sub Macro3()
Attribute Macro3.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro3 Macro
'

'
    With Selection.ShapeRange.Fill
        .Visible = msoTrue
        .ForeColor.RGB = RGB(16, 30, 26)
        .Transparency = 0
        .Solid
    End With
End Sub
