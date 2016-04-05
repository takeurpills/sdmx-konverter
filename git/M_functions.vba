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

'-------------------------------------------------------------------
'Funkcia kontroluje ci boli spravne vyplnene bunky riadiacich hodnot
'-------------------------------------------------------------------
Function cellValueRefTest(startValue As Range, endValue As Range) As Boolean

Dim stringStart As String
Dim stringEnd As String
Dim testStart As Range
Dim testEnd As Range

    stringStart = startValue.Value
    stringEnd = endValue.Value
    
    On Error Resume Next
    
    Set testStart = Range(stringStart)
    Set testEnd = Range(stringEnd)
        
    If Err.Number <> 0 Then
        cellValueRefTest = False
        Err.Clear
    Else
        cellValueRefTest = True
    End If

End Function
