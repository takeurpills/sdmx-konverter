Attribute VB_Name = "functionModule"
Option Explicit

Function incrementConversions(successType)

    Select Case successType
        Case C_OK
            conversionOk = conversionOk + 1
        Case C_FAIL
            conversionFail = conversionFail + 1
        End Select

End Function
