Attribute VB_Name = "M_errorHandler"
Option Explicit

Private Const ERR_SOURCEFILE = "openSourceFile"
Private Const ERR_SECDATA = "SECdataConversion"
Private Const ERR_T1100DATA = "T1100dataConversion"
Private Const ERR_T9XXDATA = "T9XXdataConversion"
Private Const ERR_T200DATA = "T200dataConversion"
Private Const ERR_REGDATA = "REGdataConversion"
Private Const ERR_PENSDATA = "PENSdataConversion"
Private Const ERR_MAINDATA = "MAINdataConversion"
Private Const ERR_SUDATA = "SUdataConversion"

'-----------------------------
'Centralizovany error-handling
'-----------------------------
Sub errorHandler(errProcedure As String, Optional miscInfo As String)

Dim errText As String

    Select Case errProcedure
        Case ERR_SOURCEFILE
            errText = "Nastala chyba pri otváraní súboru. Skúste znova alebo kontaktujte správcu aplikácie!"
            
            Call UnloadForms
        Case ERR_SECDATA, ERR_REGDATA, ERR_PENSDATA, ERR_MAINDATA, ERR_SUDATA, ERR_T9XXDATA, ERR_T200DATA, ERR_T1100DATA
            errText = "Vyskytla sa neoèakávaná chyba na hárku: """ & miscInfo & """. Prosím kontaktujte správcu aplikácie!"
            
            PBL_conversionFail = IncrementConversions(PBL_FAIL)
        End Select
                       
    MsgBox errText, vbCritical, "Kritická chyba"
    
End Sub
