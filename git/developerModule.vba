Attribute VB_Name = "developerModule"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                                                                                       '
'  N�zov:  Konvertor ESA2010                                                            '
'  Autor:  Martin T�th - �tatistick� �rad SR                                            '
'                                                                                       '
'  Popis:                                                                               '
'                                                                                       '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Explicit

' Typy VB komponentov vo VBA projektoch
Public Const VBEXT_CT_ACTIVEXDESIGNER = 11
Public Const VBEXT_CT_CLASSMODULE = 2
Public Const VBEXT_CT_DOCUMENT = 100
Public Const VBEXT_CT_MSFORM = 3
Public Const VBEXT_CT_STDMODULE = 1

'---------------------------------------------------------
' Proced�ra na export zdrojov�ho k�du modulov a formul�rov
'---------------------------------------------------------
Sub saveSourcecode()

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
