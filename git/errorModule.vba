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
            errText = "Nastala chyba pri otváraní súboru. Skúste znova alebo kontaktujte správcu aplikácie!"
            unloadForms
        Case ERR_SECDATA, ERR_REGDATA, ERR_PENSDATA, ERR_MAINDATA
            errText = "Zvolený hárok - """ & miscInfo & """ nemá správny formát alebo bol zvolený nesprávny typ konverzie!" & vbNewLine & vbNewLine
            errText = errText & "Konverzia hárku sa nevykoná!"
            incrementConversions (C_FAIL)
        End Select
                       
    MsgBox errText, vbCritical, "Kritická chyba"
    
End Sub
