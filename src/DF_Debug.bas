Attribute VB_Name = "DF_Debug"
Function SaveWatchList()
    Dim FilePath As String
    Dim VBComp As VBIDE.VBComponent
    Dim exp As Variant
    Dim Line As String
    
    FilePath = "C:\Temp\WatchList.txt" 'replace with your desired file path
    
    Open FilePath For Output As #1
    For Each VBComp In Application.VBE.VBProjects(1).VBComponents
        If VBComp.Type = vbext_ct_StdModule Or VBComp.Type = vbext_ct_Document Then
            If VBComp.Name <> "ThisWorkbook" Then 'skip ThisWorkbook object
                For Each win In VBComp.Windows
                    If win.Caption Like "Watch*" Then 'only process Watch window
                        For Each exp In win.WatchExpressions
                            Line = VBComp.Name & " - " & exp.Expression & " - " & exp.Value
                            Print #1, Line
                        Next exp
                    End If
                Next win
            End If
        End If
    Next VBComp
    Close #1
    
    MsgBox "Watch list saved to " & FilePath
End Function

