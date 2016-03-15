Attribute VB_Name = "M_functions"
Option Explicit

Function incrementConversions(successType)

    Select Case successType
        Case PBL_OK
            incrementConversions = PBL_conversionOk + 1
        Case PBL_FAIL
            incrementConversions = PBL_conversionFail + 1
        End Select
        
End Function
