Attribute VB_Name = "M_errorHandler"
Option Explicit

Private Const ERR_SOURCEFILE = "openSourceFile"
Private Const ERR_SECDATA = "SECdataConversion"
Private Const ERR_REGDATA = "REGdataConversion"
Private Const ERR_PENSDATA = "PENSdataConversion"
Private Const ERR_MAINDATA = "MAINdataConversion"
Private Const ERR_MAIN = "mainSub"

'------------------------------------------
'Procedura na centralizovany error-handling
'------------------------------------------
Sub errorHandler(errProcedure As String, Optional miscInfo As String)

Dim errText As String

    Select Case errProcedure
        Case ERR_SOURCEFILE
            errText = "Nastala chyba pri otv�ran� s�boru. Sk�ste znova alebo kontaktujte spr�vcu aplik�cie!"
            
            Call UnloadForms
        Case ERR_SECDATA, ERR_REGDATA, ERR_PENSDATA, ERR_MAINDATA
            errText = "Zvolen� h�rok - """ & miscInfo & """ nem� spr�vny form�t pre zvolen� typ konverzie!" & vbNewLine & vbNewLine
            errText = errText & "Konverzia h�rku sa nevykon�!"
            
            PBL_conversionFail = IncrementConversions(PBL_FAIL)
        Case ERR_MAIN
            errText = "Zvolen� h�rok - """ & miscInfo & """ nem� spr�vny form�t pre zvolen� typ konverzie!" & vbNewLine & vbNewLine
            errText = errText & "Konverzia h�rku sa nevykon�!"

            PBL_conversionFail = IncrementConversions(PBL_FAIL)
        End Select
                       
    MsgBox errText, vbCritical, "Kritick� chyba"
    
End Sub
