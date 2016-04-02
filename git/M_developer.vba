Attribute VB_Name = "M_developer"
Option Explicit

'Typy VB komponentov vo VBA projektoch
Private Const VBEXT_CT_ACTIVEXDESIGNER = 11
Private Const VBEXT_CT_CLASSMODULE = 2
Private Const VBEXT_CT_DOCUMENT = 100
Private Const VBEXT_CT_MSFORM = 3
Private Const VBEXT_CT_STDMODULE = 1

'--------------------------------------------------------
'Procedura na export zdrojoveho kodu modulov a formularov
'--------------------------------------------------------
Sub SaveSourceCode()

Dim i As Integer
Dim moduleName As String
Dim saveFolder As String

    saveFolder = "\git\"
    
    With ThisWorkbook.VBProject
        For i = .VBComponents.Count To 1 Step -1
            If .VBComponents(i).Type <> VBEXT_CT_DOCUMENT Then
                If .VBComponents(i).CodeModule.CountOfLines > 0 Then
                    moduleName = .VBComponents(i).CodeModule.name
                    .VBComponents(i).Export Application.ThisWorkbook.Path & saveFolder & moduleName & ".vba"
                End If
            End If
        Next i
    End With

End Sub

'-------------------------------------------------
'Procedura na ulozenie kopie suboru pre testovanie
'-------------------------------------------------
Sub SaveTestVersion()

Dim wbName As String
Dim saveName As String
Dim folderPath As String

    wbName = "SDMXtester"
    folderPath = "C:\Users\Martin\Desktop\project_tester"
    saveName = folderPath & "\" & wbName & ".xlsm"

    ThisWorkbook.SaveCopyAs fileName:=saveName

End Sub
