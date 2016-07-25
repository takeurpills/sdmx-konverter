Attribute VB_Name = "M_conversions"
Option Explicit

'---------------------------------------------------------------
'Procedura na vyber vstupneho suboru cez OpenFile dialogove okno
'---------------------------------------------------------------
Sub OpenSourceFile()

Const SUB_NAME = "openSourceFile"

'Spustenie vetvy "errHandler" ak sa vyskytne chyba
    On Error GoTo errHandler

'Otvorenie dialogoveho okna na vyber zdrojoveho suboru
    PBL_fileToOpen = Application.GetOpenFilename(FileFilter:="Excel,*.xls; *.xlsx; *.xlsm", _
                                                 Title:="Otvoriù s˙bor", MultiSelect:=False)
    
    If PBL_fileToOpen <> False Then
    
'Skontroluje ci bezi instancia "xlApp" do ktoreho sa otvara zdrojovy subor
        If PBL_xlApp Is Nothing Then
        Else
            PBL_xlApp.Quit
            Set PBL_xlApp = Nothing
        End If

'Otvorenie zdrojoveho suboru do novej instancie v pozadi
        Set PBL_xlOld = GetObject(PBL_programName).Application
        Set PBL_xlApp = CreateObject("Excel.Application")
        Set PBL_inputWb = PBL_xlApp.Workbooks.Open(fileName:=PBL_fileToOpen, ReadOnly:=True)

'Vypisanie cesty zvoleneho suboru do formularu
        F_main.tbSourceFile.Value = PBL_inputWb.FullName
        F_main.tbSourceFile.SetFocus
    
        On Error GoTo 0
        
'Volanie procedury na vyplnenie listboxov
        Call InputItems
        
    End If
    Exit Sub
    
'Vetva na error-handling
errHandler:
    errorHandler (SUB_NAME)
    
End Sub


'----------------------------------------------------------------
' Procedura na naplnenie listboxov nazvami vstupnych harkov v GUI
'----------------------------------------------------------------
Sub InputItems()

Const SUB_NAME = "inputItems"

Dim n As Integer  'Pocitadlo
Dim w As Integer  'Sirka
Dim m As Integer  'Maximum
    
    F_main.lbLeft.Clear
    F_main.lbRight.Clear
    
    m = 0
    
    For n = 1 To PBL_xlApp.Workbooks(1).Sheets.count
        F_main.lbLeft.AddItem PBL_xlApp.Workbooks(1).Sheets(n).Index
        F_main.lbLeft.List(n - 1, 1) = PBL_xlApp.Workbooks(1).Sheets(n).name
        F_main.labWidth.Caption = PBL_xlApp.Workbooks(1).Sheets(n).name
        w = F_main.labWidth.Width
        If w > m Then
            m = w
        End If
    Next n
    
'Rozdelenie sirky listboxu medzi dva stlpce pre spravne zobrazenie
    F_main.lbLeft.ColumnWidths = 18 & ";" & m + 20
    F_main.lbRight.ColumnWidths = 18 & ";" & m + 20
    
End Sub


'----------------
'Reset formularov
'----------------
Sub UnloadForms()

Const SUB_NAME = "unloadForms"

    F_main.chbLeft.Value = False
    F_main.chbRight.Value = False
    F_main.tbSourceFile.Value = ""
    F_main.lbLeft.Clear
    F_main.lbRight.Clear
    F_main.optSEC.Value = False
    F_main.optT9XX.Value = False
    F_main.optT200.Value = False
    F_main.optREG.Value = False
    F_main.optPENS.Value = False
    F_main.optMAIN.Value = False
    F_main.optSU.Value = False

    Unload F_progress
    
    F_main.Show vbModeless

End Sub


'-------------------------------------------------------------
'Procedura na nacitanie hodnot hlavickovych premennych do pola
'-------------------------------------------------------------
Sub ArrayPush(conversionType As Integer)

Const SUB_NAME = "arrayPush"

Dim i As Integer
Dim r As Integer
Dim s As Integer
Dim x As Integer
Dim paramValue As String
    
    Select Case conversionType
        Case PBL_SEC
        
'Nacitanie 13+12 = 25 hodnot
            ReDim PBL_parameterFix(1 To 25)
            
            s = 2
            For r = 1 To 13
                paramValue = PBL_inputWs.Cells(r, s)
                PBL_parameterFix(r) = paramValue
            Next
            
            s = 4
            For r = 1 To 12
                x = r + 12
                paramValue = PBL_inputWs.Cells(r, s)
                PBL_parameterFix(x) = paramValue
            Next
            
        Case PBL_T9XX
        
'Nacitanie 2*19 = 38 hodnot
            i = 0
            ReDim PBL_parameterFix(1 To 38)
        
            For s = 2 To 4 Step 2
                For r = 1 To 19
                    x = r + 19 * i
                    paramValue = PBL_inputWs.Cells(r, s)
                    PBL_parameterFix(x) = paramValue
                Next
                i = i + 1
            Next
        
        Case PBL_T200
        
'Nacitanie 13+2*12 = 37 hodnot
            i = 1
            ReDim PBL_parameterFix(1 To 37)
            
            s = 2
            For r = 1 To 13
                paramValue = PBL_inputWs.Cells(r, s)
                PBL_parameterFix(r) = paramValue
            Next
            
            For s = 4 To 6 Step 2
                For r = 1 To 12
                    x = r + 13 * i + (1 - i)
                    paramValue = PBL_inputWs.Cells(r, s)
                    PBL_parameterFix(x) = paramValue
                Next
                i = i + 1
            Next
            
        Case PBL_REG
        
'Nacitanie 2*11 = 22 hodnot
            i = 0
            ReDim PBL_parameterFix(1 To 22)
        
            For s = 2 To 4 Step 2
                For r = 1 To 11
                    x = r + 11 * i
                    paramValue = PBL_inputWs.Cells(r, s)
                    PBL_parameterFix(x) = paramValue
                Next
                i = i + 1
            Next
            
        Case PBL_PENS
        
'Nacitanie 12+11 = 23 hodnot
            ReDim PBL_parameterFix(1 To 23)
            
            s = 2
            For r = 1 To 12
                paramValue = PBL_inputWs.Cells(r, s)
                PBL_parameterFix(r) = paramValue
            Next
            
            s = 4
            For r = 1 To 11
                x = r + 12
                paramValue = PBL_inputWs.Cells(r, s)
                PBL_parameterFix(x) = paramValue
            Next
            
        Case PBL_MAIN
        
'Nacitanie 2*11 = 22 hodnot
            i = 0
            ReDim PBL_parameterFix(1 To 22)
        
            For s = 2 To 4 Step 2
                For r = 1 To 11
                    x = r + 11 * i
                    paramValue = PBL_inputWs.Cells(r, s)
                    PBL_parameterFix(x) = paramValue
                Next
                i = i + 1
            Next
            
        Case PBL_SU
        
'Nacitanie 2*12 = 24 hodnot
            i = 0
            ReDim PBL_parameterFix(1 To 24)
        
            For s = 2 To 4 Step 2
                For r = 1 To 12
                    x = r + 12 * i
                    paramValue = PBL_inputWs.Cells(r, s)
                    PBL_parameterFix(x) = paramValue
                Next
                i = i + 1
            Next
    End Select

End Sub


'-----------------------------------------------------------
'Procedura na kopirovanie hodnot z pola do vystupnej tabulky
'-----------------------------------------------------------
Sub ArrayFill(conversionType As Integer)

Const SUB_NAME = "arrayFill"

Dim i As Integer
Dim usedColumns As Variant
Dim copyStart As Integer
Dim copyEnd As Integer
    
'Naplnanie dat od riadok = copyStart po riadok = copyEnd
'usedColumns = do ktorych stlpcov vystupnej tabulky sa maju naplnat data
'copyStart   = spocita od ktoreho riadku sa maju naplnat data
'copyEnd     = spocita kolko riadkov je vyplnenych vo vystupnej tabulke v stlpci OBS_VALUE (tolko riadkov bude naplnenych)
    
    Select Case conversionType
        Case PBL_SEC
            usedColumns = Array(0, 1, 2, 3, 17, 23, 26, 28, 29, 30, 14, 18, 31, 39, 15, 33, 34, 35, 37, 38, 27, 40, 43, 44, 45, 25)
            copyStart = PBL_copyStart
            copyEnd = PBL_outputWs.Cells(Rows.count, "T").End(xlUp).Row
            
            For PBL_rowStep = copyStart To copyEnd
                For i = 1 To 25
                    PBL_outputWs.Cells(PBL_rowStep, usedColumns(i)).Value = PBL_parameterFix(i)
                Next i
            Next PBL_rowStep
            
        Case PBL_T9XX
            usedColumns = Array(0, 1, 2, 3, 4, 5, 6, 7, 8, 10, 11, 12, 13, 14, 15, 16, 17, 21, 22, 25, 23, 26, 24, 28, 29, 30, 32, 33, 34, 35, 36, 37, 38, 27, 40, 43, 44, 45, 39)
            copyStart = PBL_copyStart
            copyEnd = PBL_outputWs.Cells(Rows.count, "T").End(xlUp).Row
            
            For PBL_rowStep = copyStart To copyEnd
                For i = 1 To 38
                    PBL_outputWs.Cells(PBL_rowStep, usedColumns(i)).Value = PBL_parameterFix(i)
                Next i
            Next PBL_rowStep
            
        Case PBL_T200
            usedColumns = Array(0, 1, 2, 3, 4, 10, 11, 12, 13, 14, 15, 16, 17, 39, 18, 19, 23, 24, 25, 26, 27, 28, 29, 30, 31, 32, 36, 33, 34, 35, 37, 38, 40, 41, 42, 43, 44, 45)
            copyStart = PBL_copyStart
            copyEnd = PBL_outputWs.Cells(Rows.count, "T").End(xlUp).Row
            
            For PBL_rowStep = copyStart To copyEnd
                For i = 1 To 37
                    PBL_outputWs.Cells(PBL_rowStep, usedColumns(i)).Value = PBL_parameterFix(i)
                Next i
            Next PBL_rowStep
                        
        Case PBL_REG
            usedColumns = Array(0, 1, 5, 3, 4, 17, 19, 18, 21, 22, 23, 13, 24, 25, 29, 26, 28, 20, 30, 31, 32, 33)
            copyStart = PBL_copyStart
            copyEnd = PBL_outputWs.Cells(Rows.count, "N").End(xlUp).Row
            
            For PBL_rowStep = copyStart To copyEnd
                For i = 1 To 21
                    PBL_outputWs.Cells(PBL_rowStep, usedColumns(i)).Value = PBL_parameterFix(i)
                Next i
            Next PBL_rowStep
            
        Case PBL_PENS
            usedColumns = Array(0, 1, 2, 5, 15, 17, 16, 19, 20, 21, 11, 10, 25, 9, 22, 23, 24, 26, 27, 18, 28, 29, 30, 31)
            copyStart = PBL_copyStart
            copyEnd = PBL_outputWs.Cells(Rows.count, "L").End(xlUp).Row
            
            For PBL_rowStep = copyStart To copyEnd
                For i = 1 To 23
                    PBL_outputWs.Cells(PBL_rowStep, usedColumns(i)).Value = PBL_parameterFix(i)
                Next i
            Next PBL_rowStep
            
        Case PBL_MAIN
            usedColumns = Array(0, 1, 2, 3, 6, 14, 19, 21, 20, 23, 24, 25, 26, 27, 28, 29, 31, 32, 22, 33, 34, 35, 36)
            copyStart = PBL_copyStart
            copyEnd = PBL_outputWs.Cells(Rows.count, "P").End(xlUp).Row
            
            For PBL_rowStep = copyStart To copyEnd
                For i = 1 To 22
                    PBL_outputWs.Cells(PBL_rowStep, usedColumns(i)).Value = PBL_parameterFix(i)
                Next i
            Next PBL_rowStep
            
        Case PBL_SU
            usedColumns = Array(0, 1, 2, 5, 14, 20, 22, 21, 24, 25, 26, 30, 12, 36, 27, 28, 29, 31, 32, 23, 34, 35, 37, 18, 19)
            copyStart = PBL_copyStart
            copyEnd = PBL_outputWs.Cells(Rows.count, "Q").End(xlUp).Row
            
            For PBL_rowStep = copyStart To copyEnd
                For i = 1 To 24
                    PBL_outputWs.Cells(PBL_rowStep, usedColumns(i)).Value = PBL_parameterFix(i)
                Next i
            Next PBL_rowStep
    End Select

End Sub


'---------------------------------------------------------------------
'Procedura na definiciu riadiacich prvkov jednotlivych typov konverzii
'---------------------------------------------------------------------
Sub DefineConversion(conversionType As Integer)

Const SUB_NAME = "defineConversion"

Dim startRange As String
Dim endRange As String
Dim startRangeInstrument As String
Dim endRangeInstrument As String
Dim startRangeBalance As String
Dim endRangeBalance As String
Dim typeFlag As String

    Select Case conversionType
        Case PBL_SEC
            startRangeInstrument = PBL_inputWs.Range("F2").Value
            endRangeInstrument = PBL_inputWs.Range("F3").Value
            
            startRange = startRangeInstrument
            endRange = endRangeInstrument
            
            Call SECdataConversion(startRange, endRange)
            
         Case PBL_T9XX
            startRangeInstrument = PBL_inputWs.Range("K2").Value
            endRangeInstrument = PBL_inputWs.Range("K3").Value
            
            startRange = startRangeInstrument
            endRange = endRangeInstrument
            
            Call T9XXdataConversion(startRange, endRange)
                              
         Case PBL_T200
            startRangeInstrument = PBL_inputWs.Range("H2").Value
            endRangeInstrument = PBL_inputWs.Range("H3").Value
            
            startRange = startRangeInstrument
            endRange = endRangeInstrument
            
            Call T200dataConversion(startRange, endRange)
                                               
        Case PBL_REG
            startRangeInstrument = PBL_inputWs.Range("F2").Value
            endRangeInstrument = PBL_inputWs.Range("F3").Value
            
            startRange = startRangeInstrument
            endRange = endRangeInstrument
            
            Call REGdataConversion(startRange, endRange)
             
        Case PBL_PENS
            startRangeInstrument = PBL_inputWs.Range("F2").Value
            endRangeInstrument = PBL_inputWs.Range("F3").Value
            
            startRange = startRangeInstrument
            endRange = endRangeInstrument
            
            Call PENSdataConversion(startRange, endRange)
             
        Case PBL_MAIN
            startRangeInstrument = PBL_inputWs.Range("F2").Value
            endRangeInstrument = PBL_inputWs.Range("F3").Value
            
            startRange = startRangeInstrument
            endRange = endRangeInstrument
            
            Call MAINdataConversion(startRange, endRange)
            
        Case PBL_SU
            startRangeInstrument = PBL_inputWs.Range("F2").Value
            endRangeInstrument = PBL_inputWs.Range("F3").Value
            
            startRange = startRangeInstrument
            endRange = endRangeInstrument
            
            Call SUdataConversion(startRange, endRange)
    End Select

End Sub


'---------------------------------
'Konverzia dat z tabuliek typu SEC
'---------------------------------
Sub SECdataConversion(startRange As String, endRange As String)

Const SUB_NAME = "SECdataConversion"

Dim i As Integer
Dim j As Integer
Dim firstRow As Integer
Dim lastRow As Integer
Dim leadingRowStart As Integer
Dim leadingRowEnd As Integer
Dim leadingColStart As Integer
Dim leadingColEnd As Integer
Dim leadingValueStart As Range
Dim leadingValueEnd As Range

Dim headerNames As Variant
Dim columnNames As Variant
    
Dim counterpartArea() As String
Dim counterpartSector() As String
Dim refSector() As String
Dim accountingEntry() As String
Dim sto() As String
Dim unitMeasure() As String
Dim unitMult() As String
Dim gfsEcofunc() As String
Dim gfsTaxcat() As String
Dim consolidation() As String
Dim prices() As String
Dim refYearPrice() As String
Dim embargoDate() As String
Dim timePeriod() As String
Dim expenditure() As String
Dim instrAsset() As String
Dim maturity() As String
Dim obsValue() As Variant
Dim obsStatus() As String
Dim confStatus() As String

Dim boolString As String

'Spustenie vetvy "errHandler" ak sa vyskytne chyba
    On Error GoTo errHandler

'Vymedzenie riadiacich prvkov cyklu
    Set leadingValueStart = Range(startRange)
    Set leadingValueEnd = Range(endRange)
    
    leadingRowStart = leadingValueStart.Row
    leadingRowEnd = leadingValueEnd.Row
    
    firstRow = leadingRowStart
    lastRow = leadingRowEnd
    
'Inicializacia riadiacich hodnot - zac. stlpec, kon. stlpec
    leadingColStart = leadingValueStart.Column
    leadingColEnd = leadingValueEnd.Column
    
'Vyplnenie hlavicky vystupnej tabulky
    headerNames = Array("FREQ", "ADJUSTMENT", "REF_AREA", "COUNTERPART_AREA", "REF_SECTOR", "COUNTERPART_SECTOR", "CONSOLIDATION", _
                        "ACCOUNTING_ENTRY", "STO", "INSTR_ASSET", "MATURITY", "EXPENDITURE", "UNIT_MEASURE", "CURRENCY_DENOM", _
                        "VALUATION", "PRICES", "TRANSFORMATION", "CUST_BREAKDOWN", "TIME_PERIOD", "OBS_VALUE", "OBS_STATUS", _
                        "CONF_STATUS", "COMMENT_OBS", "EMBARGO_DATE", "OBS_EDP_WBB", "PRE_BREAK_VALUE", "COMMENT_DSET", _
                        "REF_PERIOD_DETAIL", "TIME_FORMAT", "TIME_PER_COLLECT", "CUST_BREAKDOWN_LB", "REF_YEAR_PRICE", _
                        "DECIMALS", "TABLE_IDENTIFIER", "TITLE", "UNIT_MULT", "LAST_UPDATE", "COMPILING_ORG", "COLL_PERIOD", _
                        "COMMENT_TS", "GFS_ECOFUNC", "GFS_TAXCAT", "DATA_COMP", "CURRENCY", "DISS_ORG")
    columnNames = Array("A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", _
                        "S", "T", "U", "V", "W", "X", "Y", "Z", "AA", "AB", "AC", "AD", "AE", "AF", "AG", "AH", _
                        "AI", "AJ", "AK", "AL", "AM", "AN", "AO", "AP", "AQ", "AR", "AS")
    For i = 0 To 44
        PBL_outputWs.Range(columnNames(i) & "1").Value = headerNames(i)
    Next i

'Vypocet posledneho vyplneneho riadku vo vystupnej tabulke (i), od [i+1] sa zacnu kopirovat nove hodnoty
    i = PBL_outputWs.Cells(Rows.count, "T").End(xlUp).Row
    i = i + 1
    PBL_copyStart = i
    
'Pocitadlo hodnot OBS_VALUE pre spravne dimenzovanie poli
    j = 0
    
'Hlavny cyklus
    For PBL_rowStep = firstRow To lastRow
        
'Kontrola nacitania riadku
        boolString = PBL_inputWs.Cells(PBL_rowStep, leadingColStart - 5).Value
        If boolString = "1" Then
            For PBL_colStep = leadingColStart To leadingColEnd
                
'Kontrola nacitania stlpca
                boolString = PBL_inputWs.Cells(firstRow - 1, PBL_colStep).Value
                If boolString = "1" Then
                        
'Nacitanie dat do pomocnych premennych
                    ReDim Preserve counterpartArea(j)
                    ReDim Preserve counterpartSector(j)
                    ReDim Preserve refSector(j)
                    ReDim Preserve accountingEntry(j)
                    ReDim Preserve sto(j)
                    ReDim Preserve unitMeasure(j)
                    ReDim Preserve unitMult(j)
                    ReDim Preserve gfsEcofunc(j)
                    ReDim Preserve gfsTaxcat(j)
                    ReDim Preserve consolidation(j)
                    ReDim Preserve prices(j)
                    ReDim Preserve refYearPrice(j)
                    ReDim Preserve embargoDate(j)
                    ReDim Preserve timePeriod(j)
                    ReDim Preserve expenditure(j)
                    ReDim Preserve instrAsset(j)
                    ReDim Preserve maturity(j)
                    ReDim Preserve obsValue(j)
                    ReDim Preserve obsStatus(j)
                    ReDim Preserve confStatus(j)
                    
                    counterpartArea(j) = PBL_inputWs.Cells(leadingRowStart - 14, PBL_colStep).Value
                    counterpartSector(j) = PBL_inputWs.Cells(leadingRowStart - 13, PBL_colStep).Value
                    refSector(j) = PBL_inputWs.Cells(leadingRowStart - 12, PBL_colStep).Value
                    accountingEntry(j) = PBL_inputWs.Cells(leadingRowStart - 11, PBL_colStep).Value
                    sto(j) = PBL_inputWs.Cells(leadingRowStart - 10, PBL_colStep).Value
                    unitMeasure(j) = PBL_inputWs.Cells(leadingRowStart - 9, PBL_colStep).Value
                    unitMult(j) = PBL_inputWs.Cells(leadingRowStart - 8, PBL_colStep).Value
                    gfsEcofunc(j) = PBL_inputWs.Cells(leadingRowStart - 7, PBL_colStep).Value
                    gfsTaxcat(j) = PBL_inputWs.Cells(leadingRowStart - 6, PBL_colStep).Value
                    consolidation(j) = PBL_inputWs.Cells(leadingRowStart - 5, PBL_colStep).Value
                    prices(j) = PBL_inputWs.Cells(leadingRowStart - 4, PBL_colStep).Value
                    refYearPrice(j) = PBL_inputWs.Cells(leadingRowStart - 3, PBL_colStep).Value
                    embargoDate(j) = PBL_inputWs.Cells(leadingRowStart - 2, PBL_colStep).Value
                    timePeriod(j) = PBL_inputWs.Cells(PBL_rowStep, leadingColStart - 4).Value
                    expenditure(j) = PBL_inputWs.Cells(PBL_rowStep, leadingColStart - 3).Value
                    instrAsset(j) = PBL_inputWs.Cells(PBL_rowStep, leadingColStart - 2).Value
                    maturity(j) = PBL_inputWs.Cells(PBL_rowStep, leadingColStart - 1).Value
                    obsValue(j) = PBL_inputWs.Cells(PBL_rowStep, PBL_colStep).Value
                    obsStatus(j) = PBL_inputWs.Cells(PBL_rowStep, PBL_colStep + 1).Value
                    confStatus(j) = PBL_inputWs.Cells(PBL_rowStep, PBL_colStep + 2).Value

                    j = j + 1
                End If
            Next PBL_colStep
        End If
    Next PBL_rowStep
    
'Ulozenie hodnot z pomocnych premennych do prislusnych stlpcov vystupneho harku na riadku "i"
    PBL_outputWs.Cells(i, 20).Resize(UBound(obsValue) + 1, 1).Value = PBL_xlNew.Transpose(obsValue)
    PBL_outputWs.Cells(i, 4).Resize(UBound(obsValue) + 1, 1).Value = PBL_xlNew.Transpose(counterpartArea)
    PBL_outputWs.Cells(i, 5).Resize(UBound(obsValue) + 1, 1).Value = PBL_xlNew.Transpose(refSector)
    PBL_outputWs.Cells(i, 6).Resize(UBound(obsValue) + 1, 1).Value = PBL_xlNew.Transpose(counterpartSector)
    PBL_outputWs.Cells(i, 7).Resize(UBound(obsValue) + 1, 1).Value = PBL_xlNew.Transpose(consolidation)
    PBL_outputWs.Cells(i, 8).Resize(UBound(obsValue) + 1, 1).Value = PBL_xlNew.Transpose(accountingEntry)
    PBL_outputWs.Cells(i, 9).Resize(UBound(obsValue) + 1, 1).Value = PBL_xlNew.Transpose(sto)
    PBL_outputWs.Cells(i, 10).Resize(UBound(obsValue) + 1, 1).Value = PBL_xlNew.Transpose(instrAsset)
    PBL_outputWs.Cells(i, 11).Resize(UBound(obsValue) + 1, 1).Value = PBL_xlNew.Transpose(maturity)
    PBL_outputWs.Cells(i, 12).Resize(UBound(obsValue) + 1, 1).Value = PBL_xlNew.Transpose(expenditure)
    PBL_outputWs.Cells(i, 13).Resize(UBound(obsValue) + 1, 1).Value = PBL_xlNew.Transpose(unitMeasure)
    PBL_outputWs.Cells(i, 16).Resize(UBound(obsValue) + 1, 1).Value = PBL_xlNew.Transpose(prices)
    PBL_outputWs.Cells(i, 19).Resize(UBound(obsValue) + 1, 1).Value = PBL_xlNew.Transpose(timePeriod)
    PBL_outputWs.Cells(i, 21).Resize(UBound(obsValue) + 1, 1).Value = PBL_xlNew.Transpose(obsStatus)
    PBL_outputWs.Cells(i, 22).Resize(UBound(obsValue) + 1, 1).Value = PBL_xlNew.Transpose(confStatus)
    PBL_outputWs.Cells(i, 24).Resize(UBound(obsValue) + 1, 1).Value = PBL_xlNew.Transpose(embargoDate)
    PBL_outputWs.Cells(i, 32).Resize(UBound(obsValue) + 1, 1).Value = PBL_xlNew.Transpose(refYearPrice)
    PBL_outputWs.Cells(i, 36).Resize(UBound(obsValue) + 1, 1).Value = PBL_xlNew.Transpose(unitMult)
    PBL_outputWs.Cells(i, 41).Resize(UBound(obsValue) + 1, 1).Value = PBL_xlNew.Transpose(gfsEcofunc)
    PBL_outputWs.Cells(i, 42).Resize(UBound(obsValue) + 1, 1).Value = PBL_xlNew.Transpose(gfsTaxcat)
    
    On Error GoTo 0
    Exit Sub

'Vetva na error-handling
errHandler:

    Call errorHandler(SUB_NAME, PBL_worksheetName)

End Sub


'-------------------------------------
'Konverzia dat z tabuliek typu SEC_9XX
'-------------------------------------
Sub T9XXdataConversion(startRange As String, endRange As String)

Const SUB_NAME = "T9XXdataConversion"

Dim i As Integer
Dim j As Integer
Dim firstRow As Integer
Dim lastRow As Integer
Dim leadingRowStart As Integer
Dim leadingRowEnd As Integer
Dim leadingColStart As Integer
Dim leadingColEnd As Integer
Dim leadingValueStart As Range
Dim leadingValueEnd As Range

Dim headerNames As Variant
Dim columnNames As Variant
    
Dim timePeriod() As String
Dim sto() As String
Dim custBreakdown() As String
Dim custBreakdownLb() As String
Dim gfsEcofunc() As String
Dim gfsTaxcat() As String
Dim obsValue() As Variant

Dim boolString As String

'Spustenie vetvy "errHandler" ak sa vyskytne chyba
    On Error GoTo errHandler

'Vymedzenie riadiacich prvkov cyklu
    Set leadingValueStart = Range(startRange)
    Set leadingValueEnd = Range(endRange)
    
    leadingRowStart = leadingValueStart.Row
    leadingRowEnd = leadingValueEnd.Row
    
    firstRow = leadingRowStart
    lastRow = leadingRowEnd
    
'Inicializacia riadiacich hodnot - zac. stlpec, kon. stlpec
    leadingColStart = leadingValueStart.Column
    leadingColEnd = leadingValueEnd.Column
    
'Vyplnenie hlavicky vystupnej tabulky
    headerNames = Array("FREQ", "ADJUSTMENT", "REF_AREA", "COUNTERPART_AREA", "REF_SECTOR", "COUNTERPART_SECTOR", "CONSOLIDATION", _
                        "ACCOUNTING_ENTRY", "STO", "INSTR_ASSET", "MATURITY", "EXPENDITURE", "UNIT_MEASURE", "CURRENCY_DENOM", _
                        "VALUATION", "PRICES", "TRANSFORMATION", "CUST_BREAKDOWN", "TIME_PERIOD", "OBS_VALUE", "OBS_STATUS", _
                        "CONF_STATUS", "COMMENT_OBS", "EMBARGO_DATE", "OBS_EDP_WBB", "PRE_BREAK_VALUE", "COMMENT_DSET", _
                        "REF_PERIOD_DETAIL", "TIME_FORMAT", "TIME_PER_COLLECT", "CUST_BREAKDOWN_LB", "REF_YEAR_PRICE", _
                        "DECIMALS", "TABLE_IDENTIFIER", "TITLE", "UNIT_MULT", "LAST_UPDATE", "COMPILING_ORG", "COLL_PERIOD", _
                        "COMMENT_TS", "GFS_ECOFUNC", "GFS_TAXCAT", "DATA_COMP", "CURRENCY", "DISS_ORG")
    columnNames = Array("A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", _
                        "S", "T", "U", "V", "W", "X", "Y", "Z", "AA", "AB", "AC", "AD", "AE", "AF", "AG", "AH", _
                        "AI", "AJ", "AK", "AL", "AM", "AN", "AO", "AP", "AQ", "AR", "AS")
    For i = 0 To 44
        PBL_outputWs.Range(columnNames(i) & "1").Value = headerNames(i)
    Next i

'Vypocet posledneho vyplneneho riadku vo vystupnej tabulke (i), od [i+1] sa zacnu kopirovat nove hodnoty
    i = PBL_outputWs.Cells(Rows.count, "T").End(xlUp).Row
    i = i + 1
    PBL_copyStart = i
    
'Pocitadlo hodnot OBS_VALUE pre spravne dimenzovanie poli
    j = 0
    
'Hlavny cyklus
    For PBL_rowStep = firstRow To lastRow
        
'Kontrola nacitania riadku
        boolString = PBL_inputWs.Cells(PBL_rowStep, leadingColStart - 7).Value
        If boolString = "1" Then
            For PBL_colStep = leadingColStart To leadingColEnd
                
'Kontrola nacitania stlpca
                boolString = PBL_inputWs.Cells(firstRow - 1, PBL_colStep).Value
                If boolString = "1" Then
                        
'Nacitanie dat do pomocnych premennych
                    ReDim Preserve timePeriod(j)
                    ReDim Preserve sto(j)
                    ReDim Preserve custBreakdown(j)
                    ReDim Preserve custBreakdownLb(j)
                    ReDim Preserve gfsEcofunc(j)
                    ReDim Preserve gfsTaxcat(j)
                    ReDim Preserve obsValue(j)
                    
                    timePeriod(j) = PBL_inputWs.Cells(leadingRowStart - 2, PBL_colStep).Value
                    sto(j) = PBL_inputWs.Cells(PBL_rowStep, leadingColStart - 6).Value
                    custBreakdown(j) = PBL_inputWs.Cells(PBL_rowStep, leadingColStart - 5).Value
                    custBreakdownLb(j) = PBL_inputWs.Cells(PBL_rowStep, leadingColStart - 3).Value
                    gfsEcofunc(j) = PBL_inputWs.Cells(PBL_rowStep, leadingColStart - 2).Value
                    gfsTaxcat(j) = PBL_inputWs.Cells(PBL_rowStep, leadingColStart - 1).Value
                    obsValue(j) = PBL_inputWs.Cells(PBL_rowStep, PBL_colStep).Value

                    j = j + 1
                End If
            Next PBL_colStep
        End If
    Next PBL_rowStep
    
'Ulozenie hodnot z pomocnych premennych do prislusnych stlpcov vystupneho harku na riadku "i"
    PBL_outputWs.Cells(i, 20).Resize(UBound(obsValue) + 1, 1).Value = PBL_xlNew.Transpose(obsValue)
    PBL_outputWs.Cells(i, 9).Resize(UBound(obsValue) + 1, 1).Value = PBL_xlNew.Transpose(sto)
    PBL_outputWs.Cells(i, 18).Resize(UBound(obsValue) + 1, 1).Value = PBL_xlNew.Transpose(custBreakdown)
    PBL_outputWs.Cells(i, 19).Resize(UBound(obsValue) + 1, 1).Value = PBL_xlNew.Transpose(timePeriod)
    PBL_outputWs.Cells(i, 31).Resize(UBound(obsValue) + 1, 1).Value = PBL_xlNew.Transpose(custBreakdownLb)
    PBL_outputWs.Cells(i, 41).Resize(UBound(obsValue) + 1, 1).Value = PBL_xlNew.Transpose(gfsEcofunc)
    PBL_outputWs.Cells(i, 42).Resize(UBound(obsValue) + 1, 1).Value = PBL_xlNew.Transpose(gfsTaxcat)
    
    On Error GoTo 0
    Exit Sub

'Vetva na error-handling
errHandler:

    Call errorHandler(SUB_NAME, PBL_worksheetName)

End Sub


'-------------------------------------
'Konverzia dat z tabuliek typu SEC_200
'-------------------------------------
Sub T200dataConversion(startRange As String, endRange As String)

Const SUB_NAME = "T200dataConversion"

Dim i As Integer
Dim j As Integer
Dim firstRow As Integer
Dim lastRow As Integer
Dim leadingRowStart As Integer
Dim leadingRowEnd As Integer
Dim leadingColStart As Integer
Dim leadingColEnd As Integer
Dim leadingValueStart As Range
Dim leadingValueEnd As Range

Dim headerNames As Variant
Dim columnNames As Variant
    
Dim refSector() As String
Dim sto() As String
Dim counterpartSector() As String
Dim accountingEntry() As String
Dim consolidation() As String
Dim obsValue() As Variant
Dim obsStatus() As String
Dim confStatus() As String

Dim boolString As String

'Spustenie vetvy "errHandler" ak sa vyskytne chyba
    On Error GoTo errHandler

'Vymedzenie riadiacich prvkov cyklu
    Set leadingValueStart = Range(startRange)
    Set leadingValueEnd = Range(endRange)
    
    leadingRowStart = leadingValueStart.Row
    leadingRowEnd = leadingValueEnd.Row
    
    firstRow = leadingRowStart
    lastRow = leadingRowEnd
    
'Inicializacia riadiacich hodnot - zac. stlpec, kon. stlpec
    leadingColStart = leadingValueStart.Column
    leadingColEnd = leadingValueEnd.Column
    
'Vyplnenie hlavicky vystupnej tabulky
    headerNames = Array("FREQ", "ADJUSTMENT", "REF_AREA", "COUNTERPART_AREA", "REF_SECTOR", "COUNTERPART_SECTOR", "CONSOLIDATION", _
                        "ACCOUNTING_ENTRY", "STO", "INSTR_ASSET", "MATURITY", "EXPENDITURE", "UNIT_MEASURE", "CURRENCY_DENOM", _
                        "VALUATION", "PRICES", "TRANSFORMATION", "CUST_BREAKDOWN", "TIME_PERIOD", "OBS_VALUE", "OBS_STATUS", _
                        "CONF_STATUS", "COMMENT_OBS", "EMBARGO_DATE", "OBS_EDP_WBB", "PRE_BREAK_VALUE", "COMMENT_DSET", _
                        "REF_PERIOD_DETAIL", "TIME_FORMAT", "TIME_PER_COLLECT", "CUST_BREAKDOWN_LB", "REF_YEAR_PRICE", _
                        "DECIMALS", "TABLE_IDENTIFIER", "TITLE", "UNIT_MULT", "LAST_UPDATE", "COMPILING_ORG", "COLL_PERIOD", _
                        "COMMENT_TS", "GFS_ECOFUNC", "GFS_TAXCAT", "DATA_COMP", "CURRENCY", "DISS_ORG")
    columnNames = Array("A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", _
                        "S", "T", "U", "V", "W", "X", "Y", "Z", "AA", "AB", "AC", "AD", "AE", "AF", "AG", "AH", _
                        "AI", "AJ", "AK", "AL", "AM", "AN", "AO", "AP", "AQ", "AR", "AS")
    For i = 0 To 44
        PBL_outputWs.Range(columnNames(i) & "1").Value = headerNames(i)
    Next i

'Vypocet posledneho vyplneneho riadku vo vystupnej tabulke (i), od [i+1] sa zacnu kopirovat nove hodnoty
    i = PBL_outputWs.Cells(Rows.count, "T").End(xlUp).Row
    i = i + 1
    PBL_copyStart = i
    
'Pocitadlo hodnot OBS_VALUE pre spravne dimenzovanie poli
    j = 0
    
'Hlavny cyklus
    For PBL_rowStep = firstRow To lastRow
        
'Kontrola nacitania riadku
        boolString = PBL_inputWs.Cells(PBL_rowStep, leadingColStart - 5).Value
        If boolString = "1" Then
            For PBL_colStep = leadingColStart To leadingColEnd
                
'Kontrola nacitania stlpca
                boolString = PBL_inputWs.Cells(firstRow - 1, PBL_colStep).Value
                If boolString = "1" Then
                        
'Nacitanie dat do pomocnych premennych
                    ReDim Preserve refSector(j)
                    ReDim Preserve sto(j)
                    ReDim Preserve counterpartSector(j)
                    ReDim Preserve accountingEntry(j)
                    ReDim Preserve consolidation(j)
                    ReDim Preserve obsValue(j)
                    ReDim Preserve obsStatus(j)
                    ReDim Preserve confStatus(j)
                    
                    refSector(j) = PBL_inputWs.Cells(leadingRowStart - 2, PBL_colStep).Value
                    sto(j) = PBL_inputWs.Cells(PBL_rowStep, leadingColStart - 4).Value
                    counterpartSector(j) = PBL_inputWs.Cells(PBL_rowStep, leadingColStart - 3).Value
                    accountingEntry(j) = PBL_inputWs.Cells(PBL_rowStep, leadingColStart - 2).Value
                    consolidation(j) = PBL_inputWs.Cells(PBL_rowStep, leadingColStart - 1).Value
                    obsValue(j) = PBL_inputWs.Cells(PBL_rowStep, PBL_colStep).Value
                    obsStatus(j) = PBL_inputWs.Cells(PBL_rowStep, PBL_colStep + 1).Value
                    confStatus(j) = PBL_inputWs.Cells(PBL_rowStep, PBL_colStep + 2).Value

                    j = j + 1
                End If
            Next PBL_colStep
        End If
    Next PBL_rowStep
    
'Ulozenie hodnot z pomocnych premennych do prislusnych stlpcov vystupneho harku na riadku "i"
    PBL_outputWs.Cells(i, 20).Resize(UBound(obsValue) + 1, 1).Value = PBL_xlNew.Transpose(obsValue)
    PBL_outputWs.Cells(i, 5).Resize(UBound(obsValue) + 1, 1).Value = PBL_xlNew.Transpose(refSector)
    PBL_outputWs.Cells(i, 6).Resize(UBound(obsValue) + 1, 1).Value = PBL_xlNew.Transpose(counterpartSector)
    PBL_outputWs.Cells(i, 7).Resize(UBound(obsValue) + 1, 1).Value = PBL_xlNew.Transpose(consolidation)
    PBL_outputWs.Cells(i, 8).Resize(UBound(obsValue) + 1, 1).Value = PBL_xlNew.Transpose(accountingEntry)
    PBL_outputWs.Cells(i, 9).Resize(UBound(obsValue) + 1, 1).Value = PBL_xlNew.Transpose(sto)
    PBL_outputWs.Cells(i, 21).Resize(UBound(obsValue) + 1, 1).Value = PBL_xlNew.Transpose(obsStatus)
    PBL_outputWs.Cells(i, 22).Resize(UBound(obsValue) + 1, 1).Value = PBL_xlNew.Transpose(confStatus)
    
    On Error GoTo 0
    Exit Sub

'Vetva na error-handling
errHandler:

    Call errorHandler(SUB_NAME, PBL_worksheetName)

End Sub


'---------------------------------
'Konverzia dat z tabuliek typu REG
'---------------------------------
Sub REGdataConversion(startRange As String, endRange As String)

Const SUB_NAME = "REGdataConversion"

Dim i As Integer
Dim j As Integer
Dim firstRow As Integer
Dim lastRow As Integer
Dim leadingRowStart As Integer
Dim leadingRowEnd As Integer
Dim leadingColStart As Integer
Dim leadingColEnd As Integer
Dim leadingValueStart As Range
Dim leadingValueEnd As Range

Dim headerNames As Variant
Dim columnNames As Variant
    
Dim accountingEntry() As String
Dim sto() As String
Dim valuation() As String
Dim prices() As String
Dim unitMeasure() As String
Dim unitMult() As String
Dim transformation() As String
Dim activity() As String
Dim refArea() As String
Dim obsValue() As Variant
Dim obsStatus() As String
Dim confStatus() As String

Dim boolString As String

'Spustenie vetvy "errHandler" ak sa vyskytne chyba
    On Error GoTo errHandler

'Vymedzenie riadiacich prvkov cyklu
    Set leadingValueStart = Range(startRange)
    Set leadingValueEnd = Range(endRange)
    
    leadingRowStart = leadingValueStart.Row
    leadingRowEnd = leadingValueEnd.Row
    
    firstRow = leadingRowStart
    lastRow = leadingRowEnd
    
'Inicializacia riadiacich hodnot - zac. stlpec, kon. stlpec
    leadingColStart = leadingValueStart.Column
    leadingColEnd = leadingValueEnd.Column
    
'Vyplnenie hlavicky vystupnej tabulky
    headerNames = Array("FREQ", "REF_AREA", "COUNTERPART_AREA", "REF_SECTOR", "COUNTERPART_SECTOR", "ACCOUNTING_ENTRY", _
                        "STO", "ACTIVITY", "VALUATION", "PRICES", "UNIT_MEASURE", "TRANSFORMATION", "TIME_PERIOD", "OBS_VALUE", _
                        "OBS_STATUS", "CONF_STATUS", "COMMENT_OBS", "EMBARGO_DATE", "PRE_BREAK_VALUE", "COMMENT_DSET", _
                        "REF_PERIOD_DETAIL", "TIME_FORMAT", "TIME_PER_COLLECT", "DECIMALS", "TABLE_IDENTIFIER", "TITLE", _
                        "UNIT_MULT", "LAST_UPDATE", "COMPILING_ORG", "COMMENT_TS", "DATA_COMP", "CURRENCY", "DISS_ORG")
    columnNames = Array("A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", _
                        "S", "T", "U", "V", "W", "X", "Y", "Z", "AA", "AB", "AC", "AD", "AE", "AF", "AG")
    For i = 0 To 32
        PBL_outputWs.Range(columnNames(i) & "1").Value = headerNames(i)
    Next i

'Vypocet posledneho vyplneneho riadku vo vystupnej tabulke (i), od [i+1] sa zacnu kopirovat nove hodnoty
    i = PBL_outputWs.Cells(Rows.count, "N").End(xlUp).Row
    i = i + 1
    PBL_copyStart = i
    
'Pocitadlo hodnot OBS_VALUE pre spravne dimenzovanie poli
    j = 0
    
'Hlavny cyklus
    For PBL_rowStep = firstRow To lastRow
        
'Kontrola nacitania riadku
        boolString = PBL_inputWs.Cells(PBL_rowStep, leadingColStart - 2).Value
        If boolString = "1" Then
            For PBL_colStep = leadingColStart To leadingColEnd
                
'Kontrola nacitania stlpca
                boolString = PBL_inputWs.Cells(firstRow - 1, PBL_colStep).Value
                If boolString = "1" Then
                        
'Nacitanie dat do pomocnych premennych
                    ReDim Preserve accountingEntry(j)
                    ReDim Preserve sto(j)
                    ReDim Preserve valuation(j)
                    ReDim Preserve prices(j)
                    ReDim Preserve unitMeasure(j)
                    ReDim Preserve unitMult(j)
                    ReDim Preserve transformation(j)
                    ReDim Preserve activity(j)
                    ReDim Preserve refArea(j)
                    ReDim Preserve obsValue(j)
                    ReDim Preserve obsStatus(j)
                    ReDim Preserve confStatus(j)

                    accountingEntry(j) = PBL_inputWs.Cells(leadingRowStart - 9, PBL_colStep).Value
                    sto(j) = PBL_inputWs.Cells(leadingRowStart - 8, PBL_colStep).Value
                    valuation(j) = PBL_inputWs.Cells(leadingRowStart - 7, PBL_colStep).Value
                    prices(j) = PBL_inputWs.Cells(leadingRowStart - 6, PBL_colStep).Value
                    unitMeasure(j) = PBL_inputWs.Cells(leadingRowStart - 5, PBL_colStep).Value
                    unitMult(j) = PBL_inputWs.Cells(leadingRowStart - 4, PBL_colStep).Value
                    transformation(j) = PBL_inputWs.Cells(leadingRowStart - 3, PBL_colStep).Value
                    activity(j) = PBL_inputWs.Cells(leadingRowStart - 2, PBL_colStep).Value
                    refArea(j) = PBL_inputWs.Cells(PBL_rowStep, leadingColStart - 1).Value
                    obsValue(j) = PBL_inputWs.Cells(PBL_rowStep, PBL_colStep).Value
                    obsStatus(j) = PBL_inputWs.Cells(PBL_rowStep, PBL_colStep + 1).Value
                    confStatus(j) = PBL_inputWs.Cells(PBL_rowStep, PBL_colStep + 2).Value

                    j = j + 1
                End If
            Next PBL_colStep
        End If
    Next PBL_rowStep
    
'Ulozenie hodnot z pomocnych premennych do prislusnych stlpcov vystupneho harku na riadku "i"
    PBL_outputWs.Cells(i, 14).Resize(UBound(obsValue) + 1, 1).Value = PBL_xlNew.Transpose(obsValue)
    PBL_outputWs.Cells(i, 2).Resize(UBound(obsValue) + 1, 1).Value = PBL_xlNew.Transpose(refArea)
    PBL_outputWs.Cells(i, 6).Resize(UBound(obsValue) + 1, 1).Value = PBL_xlNew.Transpose(accountingEntry)
    PBL_outputWs.Cells(i, 7).Resize(UBound(obsValue) + 1, 1).Value = PBL_xlNew.Transpose(sto)
    PBL_outputWs.Cells(i, 8).Resize(UBound(obsValue) + 1, 1).Value = PBL_xlNew.Transpose(activity)
    PBL_outputWs.Cells(i, 9).Resize(UBound(obsValue) + 1, 1).Value = PBL_xlNew.Transpose(valuation)
    PBL_outputWs.Cells(i, 10).Resize(UBound(obsValue) + 1, 1).Value = PBL_xlNew.Transpose(prices)
    PBL_outputWs.Cells(i, 11).Resize(UBound(obsValue) + 1, 1).Value = PBL_xlNew.Transpose(unitMeasure)
    PBL_outputWs.Cells(i, 12).Resize(UBound(obsValue) + 1, 1).Value = PBL_xlNew.Transpose(transformation)
    PBL_outputWs.Cells(i, 15).Resize(UBound(obsValue) + 1, 1).Value = PBL_xlNew.Transpose(obsStatus)
    PBL_outputWs.Cells(i, 16).Resize(UBound(obsValue) + 1, 1).Value = PBL_xlNew.Transpose(confStatus)
    PBL_outputWs.Cells(i, 27).Resize(UBound(obsValue) + 1, 1).Value = PBL_xlNew.Transpose(unitMult)
  
            
    On Error GoTo 0
    Exit Sub
    
'Vetva na error-handling
errHandler:

    Call errorHandler(SUB_NAME, PBL_worksheetName)

End Sub


'----------------------------------
'Konverzia dat z tabuliek typu PENS
'----------------------------------
Sub PENSdataConversion(startRange As String, endRange As String)

Const SUB_NAME = "PENSdataConversion"

Dim i As Integer
Dim j As Integer
Dim firstRow As Integer
Dim lastRow As Integer
Dim leadingRowStart As Integer
Dim leadingRowEnd As Integer
Dim leadingColStart As Integer
Dim leadingColEnd As Integer
Dim leadingValueStart As Range
Dim leadingValueEnd As Range

Dim headerNames As Variant
Dim columnNames As Variant
    
Dim counterpartArea() As String
Dim refSector() As String
Dim pensionFundtype() As String
Dim sto() As String
Dim instrAsset() As String
Dim obsValue() As Variant
Dim obsStatus() As String
Dim confStatus() As String

Dim boolString As String

'Spustenie vetvy "errHandler" ak sa vyskytne chyba
    On Error GoTo errHandler

'Vymedzenie riadiacich prvkov cyklu
    Set leadingValueStart = Range(startRange)
    Set leadingValueEnd = Range(endRange)
    
    leadingRowStart = leadingValueStart.Row
    leadingRowEnd = leadingValueEnd.Row
    
    firstRow = leadingRowStart
    lastRow = leadingRowEnd
    
'Inicializacia riadiacich hodnot - zac. stlpec, kon. stlpec
    leadingColStart = leadingValueStart.Column
    leadingColEnd = leadingValueEnd.Column
    
'Vyplnenie hlavicky vystupnej tabulky
    headerNames = Array("FREQ", "REF_AREA", "COUNTERPART_AREA", "REF_SECTOR", "COUNTERPART_SECTOR", "STO", "INSTR_ASSET", "PENSION_FUNDTYPE", _
                        "ACCOUNTING_ENTRY", "UNIT_MEASURE", "TIME_PERIOD", "OBS_VALUE", "OBS_STATUS", "CONF_STATUS", "COMMENT_OBS", _
                        "EMBARGO_DATE", "PRE_BREAK_VALUE", "COMMENT_DSET", "REF_PERIOD_DETAIL", "TIME_FORMAT", "TIME_PER_COLLECT", _
                        "DECIMALS", "TABLE_IDENTIFIER", "TITLE", "UNIT_MULT", "LAST_UPDATE", "COMPILING_ORG", "COMMENT_TS", _
                        "DATA_COMP", "CURRENCY", "DISS_ORG")
    columnNames = Array("A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", _
                        "T", "U", "V", "W", "X", "Y", "Z", "AA", "AB", "AC", "AD", "AE")
    For i = 0 To 30
        PBL_outputWs.Range(columnNames(i) & "1").Value = headerNames(i)
    Next i
    
'Vypocet posledneho vyplneneho riadku vo vystupnej tabulke (i), od [i+1] sa zacnu kopirovat nove hodnoty
    i = PBL_outputWs.Cells(Rows.count, "L").End(xlUp).Row
    i = i + 1
    PBL_copyStart = i
        
'Pocitadlo hodnot OBS_VALUE pre spravne dimenzovanie poli
    j = 0
    
'Hlavny cyklus
    For PBL_rowStep = firstRow To lastRow
        
'Kontrola nacitania riadku
        boolString = PBL_inputWs.Cells(PBL_rowStep, leadingColStart - 3).Value
        If boolString = "1" Then
            For PBL_colStep = leadingColStart To leadingColEnd
                
'Kontrola nacitania stlpca
                boolString = PBL_inputWs.Cells(firstRow - 1, PBL_colStep).Value
                If boolString = "1" Then
                        
'Nacitanie dat do pomocnych premennych
                    ReDim Preserve counterpartArea(j)
                    ReDim Preserve refSector(j)
                    ReDim Preserve pensionFundtype(j)
                    ReDim Preserve sto(j)
                    ReDim Preserve instrAsset(j)
                    ReDim Preserve obsValue(j)
                    ReDim Preserve obsStatus(j)
                    ReDim Preserve confStatus(j)
                    
                    counterpartArea(j) = PBL_inputWs.Cells(leadingRowStart - 4, PBL_colStep).Value
                    refSector(j) = PBL_inputWs.Cells(leadingRowStart - 3, PBL_colStep).Value
                    pensionFundtype(j) = PBL_inputWs.Cells(leadingRowStart - 2, PBL_colStep).Value
                    sto(j) = PBL_inputWs.Cells(PBL_rowStep, leadingColStart - 2).Value
                    instrAsset(j) = PBL_inputWs.Cells(PBL_rowStep, leadingColStart - 1).Value
                    obsValue(j) = PBL_inputWs.Cells(PBL_rowStep, PBL_colStep).Value
                    obsStatus(j) = PBL_inputWs.Cells(PBL_rowStep, PBL_colStep + 1).Value
                    confStatus(j) = PBL_inputWs.Cells(PBL_rowStep, PBL_colStep + 2).Value

                    j = j + 1
                End If
            Next PBL_colStep
        End If
    Next PBL_rowStep
    
'Ulozenie hodnot z pomocnych premennych do prislusnych stlpcov vystupneho harku na riadku "i"
    PBL_outputWs.Cells(i, 12).Resize(UBound(obsValue) + 1, 1).Value = PBL_xlNew.Transpose(obsValue)
    PBL_outputWs.Cells(i, 3).Resize(UBound(obsValue) + 1, 1).Value = PBL_xlNew.Transpose(counterpartArea)
    PBL_outputWs.Cells(i, 4).Resize(UBound(obsValue) + 1, 1).Value = PBL_xlNew.Transpose(refSector)
    PBL_outputWs.Cells(i, 6).Resize(UBound(obsValue) + 1, 1).Value = PBL_xlNew.Transpose(sto)
    PBL_outputWs.Cells(i, 7).Resize(UBound(obsValue) + 1, 1).Value = PBL_xlNew.Transpose(instrAsset)
    PBL_outputWs.Cells(i, 8).Resize(UBound(obsValue) + 1, 1).Value = PBL_xlNew.Transpose(pensionFundtype)
    PBL_outputWs.Cells(i, 13).Resize(UBound(obsValue) + 1, 1).Value = PBL_xlNew.Transpose(obsStatus)
    PBL_outputWs.Cells(i, 14).Resize(UBound(obsValue) + 1, 1).Value = PBL_xlNew.Transpose(confStatus)
    
    On Error GoTo 0
    Exit Sub
    
'Vetva na error-handling
errHandler:

    Call errorHandler(SUB_NAME, PBL_worksheetName)

End Sub


'----------------------------------
'Konverzia dat z tabuliek typu MAIN
'----------------------------------
Sub MAINdataConversion(startRange As String, endRange As String)

Const SUB_NAME = "MAINdataConversion"

Dim i As Integer
Dim j As Integer
Dim firstRow As Integer
Dim lastRow As Integer
Dim leadingRowStart As Integer
Dim leadingRowEnd As Integer
Dim leadingColStart As Integer
Dim leadingColEnd As Integer
Dim leadingValueStart As Range
Dim leadingValueEnd As Range

Dim headerNames As Variant
Dim columnNames As Variant
    
Dim counterpartArea() As String
Dim refSector() As String
Dim accountingEntry() As String
Dim sto() As String
Dim instrAsset() As String
Dim unitMeasure() As String
Dim unitMult() As String
Dim prices() As String
Dim obsValue() As Variant
Dim obsStatus() As String
Dim confStatus() As String
Dim expenditure() As String
Dim activity() As String
Dim timePeriod() As String

Dim boolString As String

'Spustenie vetvy "errHandler" ak sa vyskytne chyba
    On Error GoTo errHandler

'Vymedzenie riadiacich prvkov cyklu
    Set leadingValueStart = Range(startRange)
    Set leadingValueEnd = Range(endRange)
    
    leadingRowStart = leadingValueStart.Row
    leadingRowEnd = leadingValueEnd.Row
    
    firstRow = leadingRowStart
    lastRow = leadingRowEnd
    
'Inicializacia riadiacich hodnot - zac. stlpec, kon. stlpec
    leadingColStart = leadingValueStart.Column
    leadingColEnd = leadingValueEnd.Column
    
'Vyplnenie hlavicky vystupnej tabulky
    headerNames = Array("FREQ", "ADJUSTMENT", "REF_AREA", "COUNTERPART_AREA", "REF_SECTOR", "COUNTERPART_SECTOR", "ACCOUNTING_ENTRY", _
                        "STO", "INSTR_ASSET", "ACTIVITY", "EXPENDITURE", "UNIT_MEASURE", "PRICES", "TRANSFORMATION", "TIME_PERIOD", _
                        "OBS_VALUE", "OBS_STATUS", "CONF_STATUS", "COMMENT_OBS", "EMBARGO_DATE", "PRE_BREAK_VALUE", "COMMENT_DSET", _
                        "REF_PERIOD_DETAIL", "TIME_FORMAT", "TIME_PER_COLLECT", "REF_YEAR_PRICE", "DECIMALS", "TABLE_IDENTIFIER", _
                        "TITLE", "UNIT_MULT", "LAST_UPDATE", "COMPILING_ORG", "COMMENT_TS", "DATA_COMP", "CURRENCY", "DISS_ORG")
    columnNames = Array("A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", _
                        "S", "T", "U", "V", "W", "X", "Y", "Z", "AA", "AB", "AC", "AD", "AE", "AF", "AG", "AH", _
                        "AI", "AJ")
    For i = 0 To 35
        PBL_outputWs.Range(columnNames(i) & "1").Value = headerNames(i)
    Next i
    
'Vypocet posledneho vyplneneho riadku vo vystupnej tabulke (i), od [i+1] sa zacnu kopirovat nove hodnoty
    i = PBL_outputWs.Cells(Rows.count, "P").End(xlUp).Row
    i = i + 1
    PBL_copyStart = i
    
'Pocitadlo hodnot OBS_VALUE pre spravne dimenzovanie poli
    j = 0
    
'Hlavny cyklus
    For PBL_rowStep = firstRow To lastRow
        
'Kontrola nacitania riadku
        boolString = PBL_inputWs.Cells(PBL_rowStep, leadingColStart - 1).Value
        If boolString = "1" Then
            For PBL_colStep = leadingColStart To leadingColEnd
                
'Kontrola nacitania stlpca
                boolString = PBL_inputWs.Cells(firstRow - 1, PBL_colStep).Value
                If boolString = "1" Then
                        
'Nacitanie dat do pomocnych premennych
                    ReDim Preserve counterpartArea(j)
                    ReDim Preserve refSector(j)
                    ReDim Preserve accountingEntry(j)
                    ReDim Preserve sto(j)
                    ReDim Preserve instrAsset(j)
                    ReDim Preserve unitMeasure(j)
                    ReDim Preserve unitMult(j)
                    ReDim Preserve prices(j)
                    ReDim Preserve obsValue(j)
                    ReDim Preserve obsStatus(j)
                    ReDim Preserve confStatus(j)
                    ReDim Preserve expenditure(j)
                    ReDim Preserve activity(j)
                    ReDim Preserve timePeriod(j)
                    
                    counterpartArea(j) = PBL_inputWs.Cells(leadingRowStart - 9, PBL_colStep).Value
                    refSector(j) = PBL_inputWs.Cells(leadingRowStart - 8, PBL_colStep).Value
                    accountingEntry(j) = PBL_inputWs.Cells(leadingRowStart - 7, PBL_colStep).Value
                    sto(j) = PBL_inputWs.Cells(leadingRowStart - 6, PBL_colStep).Value
                    instrAsset(j) = PBL_inputWs.Cells(leadingRowStart - 5, PBL_colStep).Value
                    unitMeasure(j) = PBL_inputWs.Cells(leadingRowStart - 4, PBL_colStep).Value
                    unitMult(j) = PBL_inputWs.Cells(leadingRowStart - 3, PBL_colStep).Value
                    prices(j) = PBL_inputWs.Cells(leadingRowStart - 2, PBL_colStep).Value
                    obsValue(j) = PBL_inputWs.Cells(PBL_rowStep, PBL_colStep).Value
                    obsStatus(j) = PBL_inputWs.Cells(PBL_rowStep, PBL_colStep + 1).Value
                    confStatus(j) = PBL_inputWs.Cells(PBL_rowStep, PBL_colStep + 2).Value
                    expenditure(j) = PBL_inputWs.Cells(PBL_rowStep, PBL_colStep + 3).Value
                    activity(j) = PBL_inputWs.Cells(PBL_rowStep, PBL_colStep + 4).Value
                    timePeriod(j) = PBL_inputWs.Cells(PBL_rowStep, PBL_colStep + 5).Value

                    j = j + 1
                End If
            Next PBL_colStep
        End If
    Next PBL_rowStep
    
'Ulozenie hodnot z pomocnych premennych do prislusnych stlpcov vystupneho harku na riadku "i"
    PBL_outputWs.Cells(i, 16).Resize(UBound(obsValue) + 1, 1).Value = PBL_xlNew.Transpose(obsValue)
    PBL_outputWs.Cells(i, 4).Resize(UBound(obsValue) + 1, 1).Value = PBL_xlNew.Transpose(counterpartArea)
    PBL_outputWs.Cells(i, 5).Resize(UBound(obsValue) + 1, 1).Value = PBL_xlNew.Transpose(refSector)
    PBL_outputWs.Cells(i, 7).Resize(UBound(obsValue) + 1, 1).Value = PBL_xlNew.Transpose(accountingEntry)
    PBL_outputWs.Cells(i, 8).Resize(UBound(obsValue) + 1, 1).Value = PBL_xlNew.Transpose(sto)
    PBL_outputWs.Cells(i, 9).Resize(UBound(obsValue) + 1, 1).Value = PBL_xlNew.Transpose(instrAsset)
    PBL_outputWs.Cells(i, 10).Resize(UBound(obsValue) + 1, 1).Value = PBL_xlNew.Transpose(activity)
    PBL_outputWs.Cells(i, 11).Resize(UBound(obsValue) + 1, 1).Value = PBL_xlNew.Transpose(expenditure)
    PBL_outputWs.Cells(i, 12).Resize(UBound(obsValue) + 1, 1).Value = PBL_xlNew.Transpose(unitMeasure)
    PBL_outputWs.Cells(i, 13).Resize(UBound(obsValue) + 1, 1).Value = PBL_xlNew.Transpose(prices)
    PBL_outputWs.Cells(i, 15).Resize(UBound(obsValue) + 1, 1).Value = PBL_xlNew.Transpose(timePeriod)
    PBL_outputWs.Cells(i, 17).Resize(UBound(obsValue) + 1, 1).Value = PBL_xlNew.Transpose(obsStatus)
    PBL_outputWs.Cells(i, 18).Resize(UBound(obsValue) + 1, 1).Value = PBL_xlNew.Transpose(confStatus)
    PBL_outputWs.Cells(i, 30).Resize(UBound(obsValue) + 1, 1).Value = PBL_xlNew.Transpose(unitMult)
    
    On Error GoTo 0
    Exit Sub
    
'Vetva na error-handling
errHandler:

    Call errorHandler(SUB_NAME, PBL_worksheetName)

End Sub

'--------------------------------
'Konverzia dat z tabuliek typu SU
'--------------------------------
Sub SUdataConversion(startRange As String, endRange As String)

Const SUB_NAME = "SUdataConversion"

Dim i As Integer
Dim j As Integer
Dim firstRow As Integer
Dim lastRow As Integer
Dim leadingRowStart As Integer
Dim leadingRowEnd As Integer
Dim leadingColStart As Integer
Dim leadingColEnd As Integer
Dim leadingValueStart As Range
Dim leadingValueEnd As Range

Dim headerNames As Variant
Dim columnNames As Variant
    
Dim custBreakdownLb() As String
Dim custBreakdown() As String
Dim sto() As String
Dim accountingEntry() As String
Dim refSector() As String
Dim counterpartArea() As String
Dim valuation() As String
Dim obsValue() As Variant
Dim activity() As String
Dim activityTo() As String
Dim product() As String
Dim productTo() As String
Dim timePeriod() As String

Dim boolString As String

'Spustenie vetvy "errHandler" ak sa vyskytne chyba
    On Error GoTo errHandler

'Vymedzenie riadiacich prvkov cyklu
    Set leadingValueStart = Range(startRange)
    Set leadingValueEnd = Range(endRange)
    
    leadingRowStart = leadingValueStart.Row
    leadingRowEnd = leadingValueEnd.Row
    
    firstRow = leadingRowStart
    lastRow = leadingRowEnd
    
'Inicializacia riadiacich hodnot - zac. stlpec, kon. stlpec
    leadingColStart = leadingValueStart.Column
    leadingColEnd = leadingValueEnd.Column
    
'Vyplnenie hlavicky vystupnej tabulky
    headerNames = Array("FREQ", "REF_AREA", "COUNTERPART_AREA", "REF_SECTOR", "COUNTERPART_SECTOR", "ACCOUNTING_ENTRY", "STO", "ACTIVITY", _
                        "ACTIVITY_TO", "PRODUCT", "PRODUCT_TO", "UNIT_MEASURE", "VALUATION", "PRICES", "CUST_BREAKDOWN", "TIME_PERIOD", _
                        "OBS_VALUE", "OBS_STATUS", "CONF_STATUS", "COMMENT_OBS", "EMBARGO_DATE", "PRE_BREAK_VALUE", "COMMENT_DSET", _
                        "REF_PERIOD_DETAIL", "TIME_FORMAT", "TIME_PER_COLLECT", "DECIMALS", "TABLE_IDENTIFIER", "TITLE", "UNIT_MULT", _
                        "LAST_UPDATE", "COMPILING_ORG", "CUST_BREAKDOWN_LB", "COMMENT_TS", "DATA_COMP", "CURRENCY", "DISS_ORG")
    columnNames = Array("A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", _
                        "U", "V", "W", "X", "Y", "Z", "AA", "AB", "AC", "AD", "AE", "AF", "AG", "AH", "AI", "AJ", "AK")
    For i = 0 To 36
        PBL_outputWs.Range(columnNames(i) & "1").Value = headerNames(i)
    Next i
    
'Vypocet posledneho vyplneneho riadku vo vystupnej tabulke (i), od [i+1] sa zacnu kopirovat nove hodnoty
    i = PBL_outputWs.Cells(Rows.count, "Q").End(xlUp).Row
    i = i + 1
    PBL_copyStart = i
    
'Pocitadlo hodnot OBS_VALUE pre spravne dimenzovanie poli
    j = 0
    
'Hlavny cyklus
    For PBL_rowStep = firstRow To lastRow
        
'Kontrola nacitania riadku
        boolString = PBL_inputWs.Cells(PBL_rowStep, leadingColStart - 1).Value
        If boolString = "1" Then
            For PBL_colStep = leadingColStart To leadingColEnd
                
'Kontrola nacitania stlpca
                boolString = PBL_inputWs.Cells(firstRow - 1, PBL_colStep).Value
                If boolString = "1" Then
                        
'Nacitanie dat do pomocnych premennych
                    ReDim Preserve custBreakdownLb(j)
                    ReDim Preserve custBreakdown(j)
                    ReDim Preserve sto(j)
                    ReDim Preserve accountingEntry(j)
                    ReDim Preserve refSector(j)
                    ReDim Preserve counterpartArea(j)
                    ReDim Preserve valuation(j)
                    ReDim Preserve obsValue(j)
                    ReDim Preserve activity(j)
                    ReDim Preserve activityTo(j)
                    ReDim Preserve product(j)
                    ReDim Preserve productTo(j)
                    ReDim Preserve timePeriod(j)
                    
                    custBreakdownLb(j) = PBL_inputWs.Cells(leadingRowStart - 8, PBL_colStep).Value
                    custBreakdown(j) = PBL_inputWs.Cells(leadingRowStart - 7, PBL_colStep).Value
                    sto(j) = PBL_inputWs.Cells(leadingRowStart - 6, PBL_colStep).Value
                    accountingEntry(j) = PBL_inputWs.Cells(leadingRowStart - 5, PBL_colStep).Value
                    refSector(j) = PBL_inputWs.Cells(leadingRowStart - 4, PBL_colStep).Value
                    counterpartArea(j) = PBL_inputWs.Cells(leadingRowStart - 3, PBL_colStep).Value
                    valuation(j) = PBL_inputWs.Cells(leadingRowStart - 2, PBL_colStep).Value
                    obsValue(j) = PBL_inputWs.Cells(PBL_rowStep, PBL_colStep).Value
                    activity(j) = PBL_inputWs.Cells(PBL_rowStep, PBL_colStep + 1).Value
                    activityTo(j) = PBL_inputWs.Cells(PBL_rowStep, PBL_colStep + 2).Value
                    product(j) = PBL_inputWs.Cells(PBL_rowStep, PBL_colStep + 3).Value
                    productTo(j) = PBL_inputWs.Cells(PBL_rowStep, PBL_colStep + 4).Value
                    timePeriod(j) = PBL_inputWs.Cells(PBL_rowStep, PBL_colStep + 5).Value

                    j = j + 1
                End If
            Next PBL_colStep
        End If
    Next PBL_rowStep
    
'Ulozenie hodnot z pomocnych premennych do prislusnych stlpcov vystupneho harku na riadku "i"
    PBL_outputWs.Cells(i, 17).Resize(UBound(obsValue) + 1, 1).Value = PBL_xlNew.Transpose(obsValue)
    PBL_outputWs.Cells(i, 7).Resize(UBound(obsValue) + 1, 1).Value = PBL_xlNew.Transpose(sto)
    PBL_outputWs.Cells(i, 6).Resize(UBound(obsValue) + 1, 1).Value = PBL_xlNew.Transpose(accountingEntry)
    PBL_outputWs.Cells(i, 4).Resize(UBound(obsValue) + 1, 1).Value = PBL_xlNew.Transpose(refSector)
    PBL_outputWs.Cells(i, 3).Resize(UBound(obsValue) + 1, 1).Value = PBL_xlNew.Transpose(counterpartArea)
    PBL_outputWs.Cells(i, 13).Resize(UBound(obsValue) + 1, 1).Value = PBL_xlNew.Transpose(valuation)
    PBL_outputWs.Cells(i, 33).Resize(UBound(obsValue) + 1, 1).Value = PBL_xlNew.Transpose(custBreakdownLb)
    PBL_outputWs.Cells(i, 15).Resize(UBound(obsValue) + 1, 1).Value = PBL_xlNew.Transpose(custBreakdown)
    PBL_outputWs.Cells(i, 8).Resize(UBound(obsValue) + 1, 1).Value = PBL_xlNew.Transpose(activity)
    PBL_outputWs.Cells(i, 9).Resize(UBound(obsValue) + 1, 1).Value = PBL_xlNew.Transpose(activityTo)
    PBL_outputWs.Cells(i, 10).Resize(UBound(obsValue) + 1, 1).Value = PBL_xlNew.Transpose(product)
    PBL_outputWs.Cells(i, 11).Resize(UBound(obsValue) + 1, 1).Value = PBL_xlNew.Transpose(productTo)
    PBL_outputWs.Cells(i, 16).Resize(UBound(obsValue) + 1, 1).Value = PBL_xlNew.Transpose(timePeriod)
    
    On Error GoTo 0
    Exit Sub
    
'Vetva na error-handling
errHandler:

    Call errorHandler(SUB_NAME, PBL_worksheetName)

End Sub
