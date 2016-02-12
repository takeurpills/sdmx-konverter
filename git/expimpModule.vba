Attribute VB_Name = "expimpModule"
Sub SaveCodeModules()

'This code Exports all VBA modules
Dim i As Integer, name As String

With ThisWorkbook.VBProject
    For i = .VBComponents.Count To 1 Step -1
        If .VBComponents(i).Type <> 100 Then
            If .VBComponents(i).CodeModule.CountOfLines > 0 Then
                name = .VBComponents(i).CodeModule.name
                .VBComponents(i).Export Application.ThisWorkbook.Path & _
                                            "\git\" & name & ".vba"
            End If
        End If
    Next i
End With

End Sub

Sub ImportCodeModules()
Dim i As Integer
Dim ModuleName As String

With ThisWorkbook.VBProject
    For i = .VBComponents.Count To 1 Step -1
        ModuleName = .VBComponents(i).CodeModule.name
        If ModuleName <> "expimpModule" Then
            If .VBComponents(i).Type <> 100 Then
                .VBComponents.Remove .VBComponents(ModuleName)
                .VBComponents.Import Application.ThisWorkbook.Path & _
                                         "\git\" & ModuleName & ".vba"
            End If
        End If
    Next i
End With

End Sub
