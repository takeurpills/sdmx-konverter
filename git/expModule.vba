Attribute VB_Name = "expModule"
Option Explicit

Sub SaveCodeModules()

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
