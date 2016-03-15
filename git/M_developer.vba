Attribute VB_Name = "M_developer"
Option Explicit

'--------------------------------------------------------
'Procedúra na export zdrojového kódu modulov a formulárov
'--------------------------------------------------------
Sub saveSourcecode()

'Typy VB komponentov vo VBA projektoch
Const VBEXT_CT_ACTIVEXDESIGNER = 11
Const VBEXT_CT_CLASSMODULE = 2
Const VBEXT_CT_DOCUMENT = 100
Const VBEXT_CT_MSFORM = 3
Const VBEXT_CT_STDMODULE = 1

Dim i As Integer
Dim name As String
Dim saveFolder As String

    saveFolder = "\git\"
    
    With ThisWorkbook.VBProject
        For i = .VBComponents.Count To 1 Step -1
            If .VBComponents(i).Type <> VBEXT_CT_DOCUMENT Then
                If .VBComponents(i).CodeModule.CountOfLines > 0 Then
                    name = .VBComponents(i).CodeModule.name
                    .VBComponents(i).Export Application.ThisWorkbook.Path & saveFolder & name & ".vba"
                End If
            End If
        Next i
    End With

End Sub

'-------------------------------------------------
'Procedúra na uloženie kópie súboru pre testovanie
'-------------------------------------------------
Sub saveTestVersion()

Dim wbName As String
Dim saveName As String
Dim folderPath As String

    wbName = "SDMXtester"
    folderPath = "C:\Users\Martin\Desktop\project_tester"
    saveName = folderPath & "\" & wbName & ".xlsm"

    ThisWorkbook.SaveCopyAs fileName:=saveName

End Sub
