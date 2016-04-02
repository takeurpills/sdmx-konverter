Attribute VB_Name = "M_conversions"
Option Explicit

'---------------------------------------------------------------
'Proced�ra na v�ber vstupn�ho s�boru cez OpenFile dial�gov� okno
'---------------------------------------------------------------
Sub OpenSourceFile()

Const SUB_NAME = "openSourceFile"

'Spustenie vetvy "errHandler" ak sa vyskytne chyba
    On Error GoTo errHandler

'Otvorenie dial�gov�ho okna na v�ber zdrojov�ho s�boru
    PBL_fileToOpen = Application.GetOpenFilename(FileFilter:="Excel,*.xls; *.xlsx; *.xlsm", Title:="Otvori� s�bor", MultiSelect:=False)
    
    If PBL_fileToOpen <> False Then
    
'Skontroluje �i be�� in�tancia "xlApp" do ktor�ho sa otv�ra zdrojov� s�bor
        If PBL_xlApp Is Nothing Then
        Else
            PBL_xlApp.Quit
            Set PBL_xlApp = Nothing
        End If

'Otvorenie zdrojov�ho s�boru do novej in�tancie v pozad�
        Set PBL_xlOld = GetObject(PBL_programName).Application
        Set PBL_xlApp = CreateObject("Excel.Application")
        Set PBL_inputWb = PBL_xlApp.Workbooks.Open(fileName:=PBL_fileToOpen, ReadOnly:=True)

'Vyp�sanie cesty zvolen�ho s�boru do formul�ru
        F_main.tbSourceFile.Value = PBL_inputWb.FullName
        F_main.tbSourceFile.SetFocus
    
        On Error GoTo 0
        
'Volanie proced�ry na vyplnenie listboxov
        Call InputItems
        
    End If
    Exit Sub
    
'Vetva na error-handling
errHandler:
    errorHandler (SUB_NAME)
    
End Sub


'----------------------------------------------------------------
' Proced�ra na naplnenie listboxov n�zvami vstupn�ch h�rkov v GUI
'----------------------------------------------------------------
Sub InputItems()

Const SUB_NAME = "inputItems"

Dim n As Integer  'Count
Dim w As Integer  'Width
Dim m As Integer  'Max
    
    F_main.lbLeft.Clear
    F_main.lbRight.Clear
    
    m = 0
    
    For n = 1 To PBL_xlApp.Workbooks(1).Sheets.Count
        F_main.lbLeft.AddItem PBL_xlApp.Workbooks(1).Sheets(n).Index
        F_main.lbLeft.List(n - 1, 1) = PBL_xlApp.Workbooks(1).Sheets(n).name
        F_main.labWidth.Caption = PBL_xlApp.Workbooks(1).Sheets(n).name
        w = F_main.labWidth.Width
        If w > m Then
            m = w
        End If
    Next n
    
    F_main.lbLeft.ColumnWidths = 18 & ";" & m + 20
    F_main.lbRight.ColumnWidths = 18 & ";" & m + 20
    
End Sub


'--------------------------
'Proced�ra reset formul�rov
'--------------------------
Sub UnloadForms()

Const SUB_NAME = "unloadForms"

    F_main.chbLeft.Value = False
    F_main.chbRight.Value = False
    F_main.tbSourceFile.Value = ""
    F_main.lbLeft.Clear
    F_main.lbRight.Clear
    F_main.optSEC.Value = False
    F_main.optREG.Value = False
    F_main.optPENS.Value = False
    F_main.optMAIN.Value = False
    F_main.optSU.Value = False

    Unload F_progress
    
    F_main.Show vbModeless

End Sub


'-----------------------------------------------------------
'Proced�ra na na��tanie metad�t z hlavi�ky do pomocn�ho po�a
'-----------------------------------------------------------
Sub ArrayPush(conversionType As Integer)

Const SUB_NAME = "arrayPush"

Dim i As Integer
Dim r As Integer
Dim s As Integer
Dim x As Integer
Dim paramValue As String
    
    Select Case conversionType
        Case PBL_SEC
            'Na��tanie fixn�ch popisn�ch d�t do pomocn�ho po�a z hlavi�ky h�rku - 6*5 = 30 hodn�t
            i = 0
        
            ReDim PBL_parameterFix(1 To 32)
        
            For s = 2 To 10 Step 2
                For r = 1 To 6
                    x = r + 6 * i
                    paramValue = PBL_inputWs.Cells(r, s)
                    PBL_parameterFix(x) = paramValue
                Next
                i = i + 1
            Next
        
            'Na��tanie zvy�n�ch dvoch hodn�t popisn�ch d�t - neboli zahrnut� do symetrick�ho cyklu
            PBL_parameterFix(31) = PBL_inputWs.Cells(1, 12)
            PBL_parameterFix(32) = PBL_inputWs.Cells(2, 12)

        Case PBL_REG
            'Na��tanie fixn�ch popisn�ch d�t do pomocn�ho po�a z hlavi�ky h�rku - 2*11 = 22 hodn�t
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
            'Na��tanie fixn�ch popisn�ch d�t do pomocn�ho po�a z hlavi�ky h�rku - 12+11 = 23 hodn�t
            
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
            'Na��tanie fixn�ch popisn�ch d�t do pomocn�ho po�a z hlavi�ky h�rku - 12+11 = 23 hodn�t
            
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
            
    End Select

End Sub


'----------------------------------------------------------------------
'Proced�ra na kop�rovanie metad�t z pomocn�ho po�a do v�stupnej tabu�ky
'----------------------------------------------------------------------
Sub ArrayFill(conversionType As Integer)

Const SUB_NAME = "arrayFill"

Dim i As Integer
Dim usedColumns As Variant
Dim copyStart As Integer
Dim copyEnd As Integer
    
    Select Case conversionType
        Case PBL_SEC
            ' Inicializ�cia premenn�ch - usedColumns = do ktor�ch st�pcov v�stupnej tabu�ky sa maj� nap��a� d�ta
            '                            copyEnd = spo��ta ko�ko riadkov je vyplnen�ch vo v�stupnej tabu�ke (to�ko riadkov bude naplnen�ch)
            usedColumns = Array(0, 1, 2, 3, 6, 7, 12, 13, 14, 15, 16, 17, 18, 19, 23, 24, 25, 26, 27, 28, 29, 30, 31, 32, 33, 34, 35, 36, 37, 38, 39, 40, 41)
            copyStart = PBL_outputWs.Cells(Rows.Count, "A").End(xlUp).Row + 1
            copyEnd = PBL_outputWs.Cells(Rows.Count, "T").End(xlUp).Row
        
            ' Nap��anie popisn�ch d�t vo v�stupnej tabu�ky od riadok = copyStart do riadok = copyEnd
            For PBL_rowStep = copyStart To copyEnd
                For i = 1 To 32
                    PBL_outputWs.Cells(PBL_rowStep, usedColumns(i)).Value = PBL_parameterFix(i)
                Next i
            Next PBL_rowStep
            
        Case PBL_REG
            ' Inicializ�cia premenn�ch - usedColumns = do ktor�ch st�pcov v�stupnej tabu�ky sa maj� nap��a� d�ta
            '                            copyEnd = spo��ta ko�ko riadkov je vyplnen�ch vo v�stupnej tabu�ke (to�ko riadkov bude naplnen�ch)
            usedColumns = Array(0, 1, 3, 4, 5, 12, 13, 17, 18, 19, 20, 21, 22, 23, 24, 25, 27, 28, 29, 30, 31, 32, 33)
            copyStart = PBL_outputWs.Cells(Rows.Count, "A").End(xlUp).Row + 1
            copyEnd = PBL_outputWs.Cells(Rows.Count, "N").End(xlUp).Row
            
            ' Nap��anie popisn�ch d�t vo v�stupnej tabu�ky od riadok = copyStart do riadok = copyEnd
            For PBL_rowStep = copyStart To copyEnd
                For i = 1 To 17
                    PBL_outputWs.Cells(PBL_rowStep, usedColumns(i)).Value = PBL_parameterFix(i)
                Next i
            Next PBL_rowStep
            
        Case PBL_PENS
            ' Inicializ�cia premenn�ch - usedColumns = do ktor�ch st�pcov v�stupnej tabu�ky sa maj� nap��a� d�ta
            '                            copyEnd = spo��ta ko�ko riadkov je vyplnen�ch vo v�stupnej tabu�ke (to�ko riadkov bude naplnen�ch)
            usedColumns = Array(0, 1, 2, 5, 6, 10, 11, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29, 30, 31)
            copyStart = PBL_outputWs.Cells(Rows.Count, "A").End(xlUp).Row + 1
            copyEnd = PBL_outputWs.Cells(Rows.Count, "L").End(xlUp).Row
            
            ' Nap��anie popisn�ch d�t vo v�stupnej tabu�ky od riadok = copyStart do riadok = copyEnd
            For PBL_rowStep = copyStart To copyEnd
                For i = 1 To 23
                    PBL_outputWs.Cells(PBL_rowStep, usedColumns(i)).Value = PBL_parameterFix(i)
                Next i
            Next PBL_rowStep
            
        Case PBL_MAIN
            ' Inicializ�cia premenn�ch - usedColumns = do ktor�ch st�pcov v�stupnej tabu�ky sa maj� nap��a� d�ta
            '                            copyEnd = spo��ta ko�ko riadkov je vyplnen�ch vo v�stupnej tabu�ke (to�ko riadkov bude naplnen�ch)
            usedColumns = Array(0, 1, 2, 3, 6, 13, 14, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 30, 31, 32, 33, 34, 35, 36)
            copyStart = PBL_outputWs.Cells(Rows.Count, "A").End(xlUp).Row + 1
            copyEnd = PBL_outputWs.Cells(Rows.Count, "P").End(xlUp).Row
            
            ' Nap��anie popisn�ch d�t vo v�stupnej tabu�ky od riadok = copyStart do riadok = copyEnd
            For PBL_rowStep = copyStart To copyEnd
                For i = 1 To 23
                    PBL_outputWs.Cells(PBL_rowStep, usedColumns(i)).Value = PBL_parameterFix(i)
                Next i
            Next PBL_rowStep
        
    End Select

End Sub


'----------------------------------------------------------------------
' Proced�ra na defin�ciu riadiacich prvkov jednotliv�ch typov konverzi�
'----------------------------------------------------------------------
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
            ' Inicializ�cia premenn�ch obsahuj�cich hodnoty riadiacich za�iatkov/koncov (I = instrumenty, B = bilan�n� polo�ky)
            startRangeInstrument = PBL_inputWs.Range("L3").Value
            endRangeInstrument = PBL_inputWs.Range("L4").Value
            startRangeBalance = PBL_inputWs.Range("L5").Value
            endRangeBalance = PBL_inputWs.Range("L6").Value
        
            ' Nastavenie riadiacich prvkov pre konverzn� cyklus a volanie konverzn�ho algoritmu pre polo�ky "I"
            typeFlag = "I"
            startRange = startRangeInstrument
            endRange = endRangeInstrument
        
                Call SECdataConversion(startRange, endRange, startRangeInstrument, endRangeInstrument, typeFlag)
            ' Nastavenie riadiacich prvkov pre konverzn� cyklus a volanie konverzn�ho algoritmu pre polo�ky "B"
            typeFlag = "B"
            startRange = startRangeBalance
            endRange = endRangeBalance
        
                Call SECdataConversion(startRange, endRange, startRangeInstrument, endRangeInstrument, typeFlag)
                
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
                
    End Select

End Sub


'----------------------------------------------
'Proced�ra na konverziu d�t z tabuliek typu SEC
'----------------------------------------------
Sub SECdataConversion(startRange As String, endRange As String, startRangeInstrument As String, endRangeInstrument As String, typeFlag As String)

Const SUB_NAME = "SECdataConversion"

Dim i As Integer
Dim firstRow As Integer
Dim lastRow As Integer
Dim leadingRowStart As Integer
Dim leadingRowEnd As Integer
Dim leadingColStart As Integer
Dim leadingColEnd As Integer
Dim leadingValueStart As Range
Dim leadingValueEnd As Range
Dim specAccEntry As Integer
Dim specBoolStr As Integer

Dim obsValue As Variant
Dim counterpartArea As String
Dim refSector As String
Dim accountingEntry As String
Dim STO As String
Dim instrAsset As String
Dim maturity As String
Dim obsStatus As String
Dim confStatus As String

Dim boolString As String
Dim boolSubString As String
Dim confSubString As String

    ' Spustenie "ErrHandler" ak sa vyskytne chyba
    On Error GoTo errHandler
    
    ' Vymedzenie �i sa jedn� o klasick� cyklus pre in�trumenty, alebo o �peci�lny cyklus pre bilan�n� polo�ky
    If typeFlag = "I" Then
        Set leadingValueStart = Range(startRange)
        Set leadingValueEnd = Range(endRange)
        
        leadingRowStart = leadingValueStart.Row
        leadingRowEnd = leadingValueEnd.Row
        
        firstRow = leadingRowStart
        lastRow = leadingRowEnd
        
        ' Vymedzenie miesta, kde sa nach�dzaj� variabiln� popisn� d�ta, ktor� sa l�ia v pr�pade polo�iek "I" alebo "B"
        specAccEntry = firstRow - 3
        specBoolStr = firstRow - 1
        
    ElseIf typeFlag = "B" Then
        Set leadingValueStart = Range(startRangeInstrument)
        Set leadingValueEnd = Range(endRangeInstrument)
    
        leadingRowStart = leadingValueStart.Row
        leadingRowEnd = leadingValueEnd.Row
    
        firstRow = Range(startRange).Row
        lastRow = Range(endRange).Row
        
        ' Vymedzenie miesta, kde sa nach�dzaj� variabiln� popisn� d�ta, ktor� sa l�ia v pr�pade polo�iek "I" alebo "B"
        specAccEntry = firstRow - 2
        specBoolStr = firstRow - 1
    End If
    
    ' Inicializ�cia riadiacich hodn�t - za�iato�n� st�pec, kone�n� st�pec
    leadingColStart = leadingValueStart.Column
    leadingColEnd = leadingValueEnd.Column
    

    ' V�po�et posledn�ho vyplnen�ho riadku vo v�stupnej tabu�ke (i), od [i+1] sa za�n� kop�rova� nov� hodnoty
    i = PBL_outputWs.Cells(Rows.Count, "T").End(xlUp).Row
    i = i + 1
    
    ' Hlavn� cyklus konverzie
    For PBL_rowStep = firstRow To lastRow
        
        ' Kontrola na na��tanie riadku
        If PBL_inputWs.Cells(PBL_rowStep, leadingColStart - 1).Value = 1 Then
        
            For PBL_colStep = leadingColStart To leadingColEnd
                
                ' Kontrola na na��tanie st�pca + kontrola na substring - confidential value
                boolSubString = ""
                boolString = PBL_inputWs.Cells(specBoolStr, PBL_colStep).Value
                
                If Len(boolString) > 0 Then
                
                    boolSubString = Left(boolString, 1)
                    confSubString = Right(boolString, 1)
                    If boolSubString = 1 Then
                        
                        ' Na��tanie d�t do pomocn�ch premenn�ch
                        obsValue = PBL_inputWs.Cells(PBL_rowStep, PBL_colStep).Value
                        counterpartArea = PBL_inputWs.Cells(leadingRowStart - 2, PBL_colStep).Value
                        refSector = PBL_inputWs.Cells(leadingRowStart - 4, PBL_colStep).Value
                        STO = PBL_inputWs.Cells(PBL_rowStep, leadingColStart - 3).Value
                        instrAsset = PBL_inputWs.Cells(PBL_rowStep, leadingColStart - 2).Value
                        maturity = PBL_inputWs.Cells(PBL_rowStep, leadingColStart - 4).Value
                        obsStatus = PBL_inputWs.Cells(PBL_rowStep, PBL_colStep + 1).Value
                        accountingEntry = PBL_inputWs.Cells(specAccEntry, PBL_colStep).Value
                        
                        ' Na��tanie hodnoty confidential value z pomocn�ho substringu
                        If confSubString = "1" Or "0" Then
                           confStatus = ""
                        Else: confStatus = confSubString
                        End If

                        ' Ulo�enie hodn�t z pomocn�ch premenn�ch do pr�slu�n�ch st�pcov v�stupn�ho h�rku riadku "i"
                        PBL_outputWs.Range("T" & i).Value = obsValue
                        PBL_outputWs.Range("D" & i).Value = counterpartArea
                        PBL_outputWs.Range("E" & i).Value = refSector
                        PBL_outputWs.Range("H" & i).Value = accountingEntry
                        PBL_outputWs.Range("I" & i).Value = STO
                        PBL_outputWs.Range("J" & i).Value = instrAsset
                        PBL_outputWs.Range("K" & i).Value = maturity
                        PBL_outputWs.Range("U" & i).Value = obsStatus
                        PBL_outputWs.Range("V" & i).Value = confStatus
                        
                        ' Inkrement�cia po��tadla riadkov
                        i = i + 1
                        
                    End If
                End If
            Next PBL_colStep
        End If
    Next PBL_rowStep
    
    On Error GoTo 0
    Exit Sub
    
errHandler:

    Call errorHandler(SUB_NAME, PBL_worksheetName)

End Sub


'----------------------------------------------
'Proced�ra na konverziu d�t z tabuliek typu REG
'----------------------------------------------
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
    
Dim refArea() As String
Dim accountingEntry() As String
Dim STO() As String
Dim instrAsset() As String
Dim unitMeasure() As String
Dim valuation() As String
Dim prices() As String
Dim obsValue() As Variant
Dim obsStatus() As String
Dim confStatus() As String
Dim unitMult() As String

Dim boolString As String

    ' Spustenie "ErrHandler" ak sa vyskytne chyba
    On Error GoTo errHandler

    ' Vymedzenie riadiacich prvkov cyklu

    Set leadingValueStart = Range(startRange)
    Set leadingValueEnd = Range(endRange)
    
    leadingRowStart = leadingValueStart.Row
    leadingRowEnd = leadingValueEnd.Row
    
    firstRow = leadingRowStart
    lastRow = leadingRowEnd
    
    ' Inicializ�cia riadiacich hodn�t - za�iato�n� st�pec, kone�n� st�pec
    leadingColStart = leadingValueStart.Column
    leadingColEnd = leadingValueEnd.Column
    
    ' Vyplnenie hlavi�ky dokumentu
    headerNames = Array("FREQ", "REF_AREA", "COUNTERPART_AREA", "REF_SECTOR", "COUNTERPART_SECTOR", "ACCOUNTING_ENTRY", _
                        "STO", "INSTR_ASSET", "UNIT_MEASURE", "VALUATION", "PRICES", "TRANSFORMATION", "TIME_PERIOD", _
                        "OBS_VALUE", "OBS_STATUS", "CONF_STATUS", "COMMENT_OBS", "PRE_BREAK_VALUE", "EMBARGO_DATE", _
                        "REF_PERIOD_DETAIL", "TIME_FORMAT", "TIME_PER_COLLECT", "DECIMALS", "TABLE_IDENTIFIER", _
                        "TITLE", "UNIT_MULT", "LAST_UPDATE", "COMPILING_ORG", "COMMENT_DSET", "COMMENT_TS", _
                        "DATA_COMP", "CURRENCY", "DISS_ORG")
    columnNames = Array("A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", _
                        "S", "T", "U", "V", "W", "X", "Y", "Z", "AA", "AB", "AC", "AD", "AE", "AF", "AG")
    For i = 0 To 32
        PBL_outputWs.Range(columnNames(i) & "1").Value = headerNames(i)
    Next i


    ' V�po�et posledn�ho vyplnen�ho riadku vo v�stupnej tabu�ke (i), od [i+1] sa za�n� kop�rova� nov� hodnoty
    i = PBL_outputWs.Cells(Rows.Count, "N").End(xlUp).Row
    i = i + 1
    
    j = 0
    
    ' Hlavn� cyklus konverzie
    For PBL_rowStep = firstRow To lastRow
        
        ' Kontrola na na��tanie riadku
        boolString = PBL_inputWs.Cells(PBL_rowStep, leadingColStart - 2).Value
        If boolString = "1" Then
        
            For PBL_colStep = leadingColStart To leadingColEnd
                
                ' Kontrola na na��tanie st�pca
                boolString = PBL_inputWs.Cells(firstRow - 1, PBL_colStep).Value
                If boolString = "1" Then
                        
                        ' Na��tanie d�t do pomocn�ch premenn�ch
                        ReDim Preserve refArea(j)
                        ReDim Preserve accountingEntry(j)
                        ReDim Preserve STO(j)
                        ReDim Preserve instrAsset(j)
                        ReDim Preserve unitMeasure(j)
                        ReDim Preserve valuation(j)
                        ReDim Preserve prices(j)
                        ReDim Preserve obsValue(j)
                        ReDim Preserve obsStatus(j)
                        ReDim Preserve confStatus(j)
                        ReDim Preserve unitMult(j)

                        refArea(j) = PBL_inputWs.Cells(PBL_rowStep, leadingColStart - 1).Value
                        accountingEntry(j) = PBL_inputWs.Cells(leadingRowStart - 8, PBL_colStep).Value
                        STO(j) = PBL_inputWs.Cells(leadingRowStart - 7, PBL_colStep).Value
                        instrAsset(j) = PBL_inputWs.Cells(leadingRowStart - 6, PBL_colStep).Value
                        unitMeasure(j) = PBL_inputWs.Cells(leadingRowStart - 3, PBL_colStep).Value
                        valuation(j) = PBL_inputWs.Cells(leadingRowStart - 5, PBL_colStep).Value
                        prices(j) = PBL_inputWs.Cells(leadingRowStart - 4, PBL_colStep).Value
                        obsValue(j) = PBL_inputWs.Cells(PBL_rowStep, PBL_colStep).Value
                        obsStatus(j) = PBL_inputWs.Cells(PBL_rowStep, PBL_colStep + 1).Value
                        confStatus(j) = PBL_inputWs.Cells(PBL_rowStep, PBL_colStep + 2).Value
                        unitMult(j) = PBL_inputWs.Cells(leadingRowStart - 2, PBL_colStep).Value

                        j = j + 1
                        
                End If
            Next PBL_colStep
        End If
    Next PBL_rowStep
    
            'Ulo�enie hodn�t z pomocn�ch premenn�ch do pr�slu�n�ch st�pcov v�stupn�ho h�rku riadku "i"
            PBL_outputWs.Cells(i, 14).Resize(UBound(obsValue) + 1, 1).Value = PBL_xlNew.Transpose(obsValue)
            PBL_outputWs.Cells(i, 2).Resize(UBound(obsValue) + 1, 1).Value = PBL_xlNew.Transpose(refArea)
            PBL_outputWs.Cells(i, 6).Resize(UBound(obsValue) + 1, 1).Value = PBL_xlNew.Transpose(accountingEntry)
            PBL_outputWs.Cells(i, 7).Resize(UBound(obsValue) + 1, 1).Value = PBL_xlNew.Transpose(STO)
            PBL_outputWs.Cells(i, 8).Resize(UBound(obsValue) + 1, 1).Value = PBL_xlNew.Transpose(instrAsset)
            PBL_outputWs.Cells(i, 9).Resize(UBound(obsValue) + 1, 1).Value = PBL_xlNew.Transpose(unitMeasure)
            PBL_outputWs.Cells(i, 10).Resize(UBound(obsValue) + 1, 1).Value = PBL_xlNew.Transpose(valuation)
            PBL_outputWs.Cells(i, 11).Resize(UBound(obsValue) + 1, 1).Value = PBL_xlNew.Transpose(prices)
            PBL_outputWs.Cells(i, 15).Resize(UBound(obsValue) + 1, 1).Value = PBL_xlNew.Transpose(obsStatus)
            PBL_outputWs.Cells(i, 16).Resize(UBound(obsValue) + 1, 1).Value = PBL_xlNew.Transpose(confStatus)
            PBL_outputWs.Cells(i, 26).Resize(UBound(obsValue) + 1, 1).Value = PBL_xlNew.Transpose(unitMult)
            
    On Error GoTo 0
    Exit Sub
    
errHandler:

    Call errorHandler(SUB_NAME, PBL_worksheetName)

End Sub


'-----------------------------------------------
'Proced�ra na konverziu d�t z tabuliek typu PENS
'-----------------------------------------------
Sub PENSdataConversion(startRange As String, endRange As String)

Const SUB_NAME = "PENSdataConversion"

Dim i As Integer
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
    
Dim counterpartArea As String
Dim refSector As String
Dim pensionFundtype As String
Dim STO As String
Dim instrAsset As String
Dim obsValue As Variant
Dim obsStatus As String
Dim confStatus As String

Dim boolString As String

    ' Spustenie "ErrHandler" ak sa vyskytne chyba
    On Error GoTo errHandler

    ' Vymedzenie riadiacich prvkov cyklu

    Set leadingValueStart = Range(startRange)
    Set leadingValueEnd = Range(endRange)
    
    leadingRowStart = leadingValueStart.Row
    leadingRowEnd = leadingValueEnd.Row
    
    firstRow = leadingRowStart
    lastRow = leadingRowEnd
    
    ' Inicializ�cia riadiacich hodn�t - za�iato�n� st�pec, kone�n� st�pec
    leadingColStart = leadingValueStart.Column
    leadingColEnd = leadingValueEnd.Column
    
    ' Vyplnenie hlavi�ky dokumentu
    headerNames = Array("FREQ", "REF_AREA", "COUNTERPART_AREA", "REF_SECTOR", "COUNTERPART_SECTOR", "ACCOUNTING_ENTRY", "STO", "INSTR_ASSET", "PENSION_FUNDTYPE", "UNIT_MEASURE", "TIME_PERIOD", "OBS_VALUE", "OBS_STATUS", "CONF_STATUS", "COMMENT_OBS", "PRE_BREAK_VALUE", "EMBARGO_DATE", "REF_PERIOD_DETAIL", "TIME_FORMAT", "TIME_PER_COLLECT", "DECIMALS", "TABLE_IDENTIFIER", "TITLE", "UNIT_MULT", "LAST_UPDATE", "COMPILING_ORG", "COMMENT_DSET", "COMMENT_TS", "DATA_COMP", "CURRENCY", "DISS_ORG")
    columnNames = Array("A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z", "AA", "AB", "AC", "AD", "AE")
    For i = 0 To 30
        PBL_outputWs.Range(columnNames(i) & "1").Value = headerNames(i)
    Next i
    
    ' V�po�et posledn�ho vyplnen�ho riadku vo v�stupnej tabu�ke (i), od [i+1] sa za�n� kop�rova� nov� hodnoty
    i = PBL_outputWs.Cells(Rows.Count, "L").End(xlUp).Row
    i = i + 1
    
    ' Hlavn� cyklus konverzie
    For PBL_rowStep = firstRow To lastRow
        
        ' Kontrola na na��tanie riadku
        If PBL_inputWs.Cells(PBL_rowStep, leadingColStart - 3).Value = 1 Then
        
            For PBL_colStep = leadingColStart To leadingColEnd
                
                ' Kontrola na na��tanie st�pca
                boolString = PBL_inputWs.Cells(firstRow - 1, PBL_colStep).Value
                
                    If boolString = "1" Then
                        
                        ' Na��tanie d�t do pomocn�ch premenn�ch
                        counterpartArea = PBL_inputWs.Cells(leadingRowStart - 4, PBL_colStep).Value
                        refSector = PBL_inputWs.Cells(leadingRowStart - 3, PBL_colStep).Value
                        pensionFundtype = PBL_inputWs.Cells(leadingRowStart - 2, PBL_colStep).Value
                        STO = PBL_inputWs.Cells(PBL_rowStep, leadingColStart - 2).Value
                        instrAsset = PBL_inputWs.Cells(PBL_rowStep, leadingColStart - 1).Value
                        obsValue = PBL_inputWs.Cells(PBL_rowStep, PBL_colStep).Value
                        obsStatus = PBL_inputWs.Cells(PBL_rowStep, PBL_colStep + 1).Value
                        confStatus = PBL_inputWs.Cells(PBL_rowStep, PBL_colStep + 2).Value

                        ' Ulo�enie hodn�t z pomocn�ch premenn�ch do pr�slu�n�ch st�pcov v�stupn�ho h�rku riadku "i"
                        PBL_outputWs.Range("L" & i).Value = obsValue
                        PBL_outputWs.Range("C" & i).Value = counterpartArea
                        PBL_outputWs.Range("D" & i).Value = refSector
                        PBL_outputWs.Range("G" & i).Value = STO
                        PBL_outputWs.Range("H" & i).Value = instrAsset
                        PBL_outputWs.Range("I" & i).Value = pensionFundtype
                        PBL_outputWs.Range("M" & i).Value = obsStatus
                        PBL_outputWs.Range("N" & i).Value = confStatus

                        ' Inkrement�cia po��tadla riadkov
                        i = i + 1
                        
                End If
            Next PBL_colStep
        End If
    Next PBL_rowStep
    
    On Error GoTo 0
    Exit Sub
    
errHandler:

    Call errorHandler(SUB_NAME, PBL_worksheetName)

End Sub


'-----------------------------------------------
'Proced�ra na konverziu d�t z tabuliek typu MAIN
'-----------------------------------------------
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
Dim STO() As String
Dim instrAsset() As String
Dim expenditure() As String
Dim unitMeasure() As String
Dim unitMult() As String
Dim obsValue() As Variant
Dim obsStatus() As String
Dim confStatus() As String
Dim activity() As String
Dim timePeriod() As String

Dim boolString As String

    ' Spustenie "ErrHandler" ak sa vyskytne chyba
    'On Error GoTo errHandler

    ' Vymedzenie riadiacich prvkov cyklu

    Set leadingValueStart = Range(startRange)
    Set leadingValueEnd = Range(endRange)
    
    leadingRowStart = leadingValueStart.Row
    leadingRowEnd = leadingValueEnd.Row
    
    firstRow = leadingRowStart
    lastRow = leadingRowEnd
    
    ' Inicializ�cia riadiacich hodn�t - za�iato�n� st�pec, kone�n� st�pec
    leadingColStart = leadingValueStart.Column
    leadingColEnd = leadingValueEnd.Column
    
    ' Vyplnenie hlavi�ky dokumentu
    headerNames = Array("FREQ", "ADJUSTMENT", "REF_AREA", "COUNTERPART_AREA", "REF_SECTOR", "COUNTERPART_SECTOR", "ACCOUNTING_ENTRY", "STO", "INSTR_ASSET", "ACTIVITY", "EXPENDITURE", "UNIT_MEASURE", "PRICES", "TRANSFORMATION", "TIME_PERIOD", "OBS_VALUE", "OBS_STATUS", "CONF_STATUS", "COMMENT_OBS", "PRE_BREAK_VALUE", "EMBARGO_DATE", "REF_PERIOD_DETAIL", "TIME_FORMAT", "TIME_PER_COLLECT", "REF_YEAR_PRICE", "DECIMALS", "TABLE_IDENTIFIER", "TITLE", "UNIT_MULT", "LAST_UPDATE", "COMPILING_ORG", "COMMENT_DSET", "COMMENT_TS", "DATA_COMP", "CURRENCY", "DISS_ORG")
    columnNames = Array("A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z", "AA", "AB", "AC", "AD", "AE", "AF", "AG", "AH", "AI", "AJ")
    For i = 0 To 35
        PBL_outputWs.Range(columnNames(i) & "1").Value = headerNames(i)
    Next i
    
    ' V�po�et posledn�ho vyplnen�ho riadku vo v�stupnej tabu�ke (i), od [i+1] sa za�n� kop�rova� nov� hodnoty
    i = PBL_outputWs.Cells(Rows.Count, "P").End(xlUp).Row
    i = i + 1
    
    j = 0
    
    ' Hlavn� cyklus konverzie
    For PBL_rowStep = firstRow To lastRow
        
        ' Kontrola na na��tanie riadku
        boolString = PBL_inputWs.Cells(PBL_rowStep, leadingColStart - 1).Value
        If boolString = "1" Then
        
            For PBL_colStep = leadingColStart To leadingColEnd
                
                ' Kontrola na na��tanie st�pca
                boolString = PBL_inputWs.Cells(firstRow - 1, PBL_colStep).Value
                    If boolString = "1" Then
                        
                        ' Na��tanie d�t do pomocn�ch premenn�ch
                        ReDim Preserve counterpartArea(j)
                        ReDim Preserve refSector(j)
                        ReDim Preserve accountingEntry(j)
                        ReDim Preserve STO(j)
                        ReDim Preserve instrAsset(j)
                        ReDim Preserve expenditure(j)
                        ReDim Preserve unitMeasure(j)
                        ReDim Preserve unitMult(j)
                        ReDim Preserve obsValue(j)
                        ReDim Preserve obsStatus(j)
                        ReDim Preserve confStatus(j)
                        ReDim Preserve activity(j)
                        ReDim Preserve timePeriod(j)
                        
                        counterpartArea(j) = PBL_inputWs.Cells(leadingRowStart - 9, PBL_colStep).Value
                        refSector(j) = PBL_inputWs.Cells(leadingRowStart - 8, PBL_colStep).Value
                        accountingEntry(j) = PBL_inputWs.Cells(leadingRowStart - 7, PBL_colStep).Value
                        STO(j) = PBL_inputWs.Cells(leadingRowStart - 6, PBL_colStep).Value
                        instrAsset(j) = PBL_inputWs.Cells(leadingRowStart - 5, PBL_colStep).Value
                        expenditure(j) = PBL_inputWs.Cells(leadingRowStart - 4, PBL_colStep).Value
                        unitMeasure(j) = PBL_inputWs.Cells(leadingRowStart - 3, PBL_colStep).Value
                        unitMult(j) = PBL_inputWs.Cells(leadingRowStart - 2, PBL_colStep).Value
                        obsValue(j) = PBL_inputWs.Cells(PBL_rowStep, PBL_colStep).Value
                        obsStatus(j) = PBL_inputWs.Cells(PBL_rowStep, PBL_colStep + 1).Value
                        confStatus(j) = PBL_inputWs.Cells(PBL_rowStep, PBL_colStep + 2).Value
                        activity(j) = PBL_inputWs.Cells(PBL_rowStep, PBL_colStep + 3).Value
                        timePeriod(j) = PBL_inputWs.Cells(PBL_rowStep, PBL_colStep + 4).Value

                        j = j + 1
                        
                End If
            Next PBL_colStep
        End If
    Next PBL_rowStep
    
            ' Ulo�enie hodn�t z pomocn�ch premenn�ch do pr�slu�n�ch st�pcov v�stupn�ho h�rku riadku "i"
            PBL_outputWs.Cells(i, 16).Resize(UBound(obsValue) + 1, 1).Value = PBL_xlNew.Transpose(obsValue)
            PBL_outputWs.Cells(i, 4).Resize(UBound(obsValue) + 1, 1).Value = PBL_xlNew.Transpose(counterpartArea)
            PBL_outputWs.Cells(i, 5).Resize(UBound(obsValue) + 1, 1).Value = PBL_xlNew.Transpose(refSector)
            PBL_outputWs.Cells(i, 7).Resize(UBound(obsValue) + 1, 1).Value = PBL_xlNew.Transpose(accountingEntry)
            PBL_outputWs.Cells(i, 8).Resize(UBound(obsValue) + 1, 1).Value = PBL_xlNew.Transpose(STO)
            PBL_outputWs.Cells(i, 9).Resize(UBound(obsValue) + 1, 1).Value = PBL_xlNew.Transpose(instrAsset)
            PBL_outputWs.Cells(i, 10).Resize(UBound(obsValue) + 1, 1).Value = PBL_xlNew.Transpose(activity)
            PBL_outputWs.Cells(i, 11).Resize(UBound(obsValue) + 1, 1).Value = PBL_xlNew.Transpose(expenditure)
            PBL_outputWs.Cells(i, 12).Resize(UBound(obsValue) + 1, 1).Value = PBL_xlNew.Transpose(unitMeasure)
            PBL_outputWs.Cells(i, 15).Resize(UBound(obsValue) + 1, 1).Value = PBL_xlNew.Transpose(timePeriod)
            PBL_outputWs.Cells(i, 17).Resize(UBound(obsValue) + 1, 1).Value = PBL_xlNew.Transpose(obsStatus)
            PBL_outputWs.Cells(i, 18).Resize(UBound(obsValue) + 1, 1).Value = PBL_xlNew.Transpose(confStatus)
            PBL_outputWs.Cells(i, 29).Resize(UBound(obsValue) + 1, 1).Value = PBL_xlNew.Transpose(unitMult)

    
    On Error GoTo 0
    Exit Sub
    
errHandler:

    Call errorHandler(SUB_NAME, PBL_worksheetName)

End Sub
