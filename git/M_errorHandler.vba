Attribute VB_Name = "M_errorHandler"
Option Explicit

Private Const ERR_SOURCEFILE = "openSourceFile"
Private Const ERR_SECDATA = "SECdataConversion"
Private Const ERR_REGDATA = "REGdataConversion"
Private Const ERR_PENSDATA = "PENSdataConversion"
Private Const ERR_MAINDATA = "MAINdataConversion"

'-----------------------------
'Centralizovany error-handling
'-----------------------------
Sub errorHandler(errProcedure As String, Optional miscInfo As String)

Dim errText As String

    Select Case errProcedure
        Case ERR_SOURCEFILE
            errText = "Nastala chyba pri otváraní súboru. Skúste znova alebo kontaktujte správcu aplikácie!"
            
            Call UnloadForms
        Case ERR_SECDATA, ERR_REGDATA, ERR_PENSDATA, ERR_MAINDATA
            errText = "Vyskytla sa neoèakávaná chyba na hárku: """ & miscInfo & """. Prosím kontaktujte správcu aplikácie!"
            
            PBL_conversionFail = IncrementConversions(PBL_FAIL)
        End Select
                       
    MsgBox errText, vbCritical, "Kritická chyba"
    
End Sub
