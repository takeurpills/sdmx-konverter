Attribute VB_Name = "errorModule"
Option Explicit

Const ERR_SOURCEFILE = "openSourceFile"
Const ERR_SECDATA = "SECdataConversion"
Const ERR_REGDATA = "REGdataConversion"
Const ERR_PENSDATA = "PENSdataConversion"
Const ERR_MAINDATA = "MAINdataConversion"

Sub errorHandler(errProcedure As String, Optional miscInfo As String)

Dim errText As String

    Select Case errProcedure
        Case ERR_SOURCEFILE
            errText = "Nastala chyba pri otv�ran� s�boru. Sk�ste znova alebo kontaktujte spr�vcu aplik�cie!"
            unloadForms
        Case ERR_SECDATA, ERR_REGDATA, ERR_PENSDATA, ERR_MAINDATA
            errText = "Zvolen� h�rok - """ & miscInfo & """ nem� spr�vny form�t alebo bol zvolen� nespr�vny typ konverzie!" & vbNewLine & vbNewLine
            errText = errText & "Konverzia h�rku sa nevykon�!"
            incrementConversions (C_FAIL)
        End Select
                       
    MsgBox errText, vbCritical, "Kritick� chyba"
    
End Sub
