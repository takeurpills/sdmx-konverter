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

'-------------------------------------------------------------
'Funkcia na vymazanie posledneho carriage return z csv vystupu
'-------------------------------------------------------------
Function deleteLastLine(fileName)

Dim myFile As String
Dim objFSO As FileSystemObject
Dim objFile As Object
Dim strfile As String
Dim intLength As Long
Dim strend As String

Const ForReading = 1
Const ForWriting = 2

myFile = fileName & ".csv"
 
Set objFSO = CreateObject("Scripting.FileSystemObject")
 
Set objFile = objFSO.OpenTextFile(myFile, ForReading)
strfile = objFile.ReadAll
objFile.Close
 
intLength = Len(strfile)
strend = Right(strfile, 2)
 
If strend = vbCrLf Then
    strfile = Left(strfile, intLength - 2)
    Set objFile = objFSO.OpenTextFile(myFile, ForWriting)
    objFile.Write strfile
    objFile.Close
End If

End Function

'------------------------------
'Funkcia na traspoziciu 1D pola
'------------------------------
Function transposeArray(myArray As Variant) As Variant
Dim x As Long
Dim Y As Long
Dim Xupper As Long
Dim Yupper As Long
Dim tempArray As Variant
    Xupper = UBound(myArray, 1)
    Yupper = 0
    Y = Yupper
    ReDim tempArray(Xupper, Yupper)
    For x = 0 To Xupper
        tempArray(x, Y) = myArray(x)
    Next x
    transposeArray = tempArray
End Function
