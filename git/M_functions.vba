Attribute VB_Name = "M_functions"
Option Explicit

'----------------------------------------------
'Funkcia na inkrementovanie pocitadla konverzii
'----------------------------------------------
Function IncrementConversions(successType As String)

    Select Case successType
        Case PBL_OK
            IncrementConversions = PBL_conversionOk + 1
        Case PBL_FAIL
            IncrementConversions = PBL_conversionFail + 1
        End Select
        
End Function
