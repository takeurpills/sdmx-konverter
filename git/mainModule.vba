Attribute VB_Name = "mainModule"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                                                                                       '
'  N·zov:  Konvertor ESA2010                                                            '
'  Autor:  Martin TÛth - ätatistick˝ ˙rad SR                                            '
'                                                                                       '
'  Popis:                                                                               '
'                                                                                       '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Explicit

Public Const NA_SEC = 1
Public Const NA_REG = 2
Public Const NA_PENS = 3
Public Const NA_MAIN = 4
Public Const NA_SU = 5

Public xlApp As Object
Public xlNew As Object
Public xlOld As Object
Public versionId As String
Public instanceName As String
Public inputWsId() As Integer

Dim rowStep As Integer
Dim colStep As Integer

Dim inputWs As Worksheet
Dim outputWs As Worksheet
Dim inputWb As Workbook
       
Dim parameterFix() As Variant
Dim fileToOpen As Variant

'----------------------------------------------
' InicializaËn· proced˙ra pri spustenÌ programu
'----------------------------------------------
Sub programInit()
    
    versionId = "v0.3"
    instanceName = ActiveWorkbook.FullName
    
    mainForm.Show vbModeless
    
End Sub
    
'------------------------------------
' Hlavn· riadiaca proced˙ra konverzie
'------------------------------------
Sub mainSub(conversionType As Integer)
    
Dim i As Integer
Dim saveName As String
Dim folderPath As String
Dim timeStamp As String
Dim nameString As String
Dim typeString As String
Dim outputFile As String
Dim errorMsg As String
Dim successMsg As String
Dim conversionCheck As Boolean
Dim progIndicator As Integer
Dim conversionOk As Integer
Dim conversionFail As Integer

    progIndicator = 0
    conversionOk = 0
    conversionFail = 0
    
    ' Spustenie "ErrHandler" ak sa vyskytne chyba
    On Error GoTo Errhandler
    
    If fileToOpen <> False Then
        If (Not inputWsId) = True Then
            errorMsg = "Nie s˙ zvolenÈ pracovnÈ h·rky pre konverziu!"
            MsgBox errorMsg, vbExclamation, "InformatÌvna chyba"
        Else
            For i = 1 To UBound(inputWsId)
        
            Set xlNew = CreateObject("Excel.Application")
            xlNew.ScreenUpdating = False
            
            xlNew.Workbooks.Add (1)
            Set outputWs = xlNew.ActiveWorkbook.Worksheets(1)

            Set inputWs = inputWb.Worksheets(inputWsId(i))
            nameString = inputWs.name

            conversionCheck = False
            
            Select Case conversionType
                Case NA_SEC
                    If inputWs.Cells(1, 1).Value = "FREQ" And inputWs.Cells(6, 1).Value = "EXPENDITURE" Then
                    conversionCheck = True
                    End If
                Case NA_REG
                    If inputWs.Cells(1, 10).Value = "REG" Then
                    conversionCheck = True
                    End If
                Case NA_MAIN
                    If inputWs.Cells(12, 1).Value = "TIME_PER_COLLECT" And inputWs.Cells(1, 6).Value = "MAIN" Then
                    conversionCheck = True
                    End If
            End Select
            
            If conversionCheck = True Then
            
                ' Sp˙ötanie proced˙r
                If progIndicator = 0 Then
                    mainForm.Hide
                    progressForm.Show vbModeless
                    progIndicator = 1
                End If
                                
                Call arrayPush(conversionType)
                Call defineConversion(conversionType)
                Call arrayFill(conversionType)
                
                ' Uloûenie v˝stupu
                folderPath = xlApp.ActiveWorkbook.Path
                timeStamp = Format(CStr(Now), "yyyy_mm_dd_hhmmss")
                saveName = folderPath & "\" & nameString & "_" & timeStamp
    
                xlNew.Workbooks(1).SaveAs Filename:=saveName, FileFormat:=xlCSV, local:=True
                xlNew.ScreenUpdating = True
                xlNew.Workbooks(1).Close False
                xlNew.Quit
                
                conversionOk = conversionOk + 1
                
            Else
                errorMsg = "Zvolen˝ h·rok - """ & nameString & """ nem· spr·vny form·t alebo bol zvolen˝ nespr·vny typ konverzie!" & vbNewLine & vbNewLine
                errorMsg = errorMsg & "Konverzia h·rku sa nevykon·!"
                MsgBox errorMsg, vbCritical, "Kritick· chyba"
                xlNew.Quit
                
                conversionFail = conversionFail + 1
            End If
            Next
            
            Call unloadForms
            
            xlApp.Quit
            
            successMsg = "Konverzia h·rkov bola dokonËen·." & vbNewLine & vbNewLine
            successMsg = successMsg & "ï PoËet ˙speön˝ch konverziÌ: " & conversionOk & vbNewLine
            successMsg = successMsg & "ï PoËet ne˙speön˝ch konverziÌ: " & conversionFail
            
            MsgBox successMsg, vbInformation, "Inform·cia"
            
        End If
        
    Else
        errorMsg = "Nie je zvolen˝ s˙bor pre konverziu!"
        MsgBox errorMsg, vbExclamation, "InformatÌvna chyba"
    End If
    Exit Sub
    
Errhandler:

    On Error Resume Next

    errorMsg = "Nastala chyba (#1001). ProsÌm kontaktujte spr·vcu aplik·cie."

    MsgBox errorMsg, vbCritical, "Kritick· chyba"
    
    appClose
    
End Sub

'---------------------------
' Proced˙ra vstupnej tabuæky
'---------------------------
Sub openSourceFile()

    If xlApp Is Nothing Then
    Else
        xlApp.Quit
    End If

    fileToOpen = Application.GetOpenFilename(FileFilter:="Vstupn˝ s˙bor,*.xls; *.xlsx; *.xlsm", Title:="Otvoriù s˙bor", MultiSelect:=False)
    
    If fileToOpen <> False Then
        Set xlOld = GetObject(instanceName).Application
        Set xlApp = CreateObject("Excel.Application")
        Set inputWb = xlApp.Workbooks.Open(Filename:=fileToOpen, ReadOnly:=True)
      
        mainForm.tbSourceFile.Value = inputWb.FullName
        mainForm.tbSourceFile.SetFocus
    
        Call inputItems
      
    End If

End Sub

'---------------------------
' Proced˙ra vstupn˝ch h·rkov
'---------------------------
Sub inputItems()
Dim N As Long
Dim w As Integer  'Width
Dim m As Integer  'Max
    
    mainForm.lbLeft.Clear
    mainForm.lbRight.Clear
    
    m = 0
    
    For N = 1 To xlApp.Workbooks(1).Sheets.Count
        mainForm.lbLeft.AddItem xlApp.Workbooks(1).Sheets(N).Index
        mainForm.lbLeft.List(N - 1, 1) = xlApp.Workbooks(1).Sheets(N).name
        mainForm.labWidth.Caption = xlApp.Workbooks(1).Sheets(N).name
        w = mainForm.labWidth.Width
        If w > m Then
            m = w
        End If
    Next N
    
    mainForm.lbLeft.ColumnWidths = 18 & ";" & m + 20
    mainForm.lbRight.ColumnWidths = 18 & ";" & m + 20

End Sub


'---------------------------------------------------
'Proced˙ra naËÌtanie popisn˝ch d·t do pomocnÈho pola
'---------------------------------------------------
Sub arrayPush(conversionType As Integer)

Dim i As Integer
Dim r As Integer
Dim s As Integer
Dim x As Integer
Dim paramValue As String
    
    Select Case conversionType
        Case NA_SEC
            ' NaËÌtanie fixn˝ch popisn˝ch d·t do pomocnÈho pola z hlaviËky h·rku - 6*5 = 30 hodnÙt
            i = 0
        
            ReDim parameterFix(1 To 32)
        
            For s = 2 To 10 Step 2
                For r = 1 To 6
                    x = r + 6 * i
                    paramValue = inputWs.Cells(r, s)
                    parameterFix(x) = paramValue
                Next
                i = i + 1
            Next
        
            ' NaËÌtanie zvyön˝ch dvoch hodnÙt popisn˝ch d·t - neboli zahrnutÈ do symetrickÈho cyklu
            parameterFix(31) = inputWs.Cells(1, 12)
            parameterFix(32) = inputWs.Cells(2, 12)

        Case NA_REG
            ' NaËÌtanie fixn˝ch popisn˝ch d·t do pomocnÈho pola z hlaviËky h·rku - 6*5 = 30 hodnÙt
            i = 0
        
            ReDim parameterFix(1 To 17)
        
            For s = 2 To 6 Step 2
                For r = 1 To 5
                    x = r + 5 * i
                    paramValue = inputWs.Cells(r, s)
                    parameterFix(x) = paramValue
                Next
                i = i + 1
            Next
        
            ' NaËÌtanie zvyön˝ch dvoch hodnÙt popisn˝ch d·t - neboli zahrnutÈ do symetrickÈho cyklu
            parameterFix(16) = inputWs.Cells(1, 8)
            parameterFix(17) = inputWs.Cells(2, 8)
            
        Case NA_MAIN
            ' NaËÌtanie fixn˝ch popisn˝ch d·t do pomocnÈho pola z hlaviËky h·rku - 12+11 = 23 hodnÙt
            
            ReDim parameterFix(1 To 23)
            
            s = 2
            For r = 1 To 12
                paramValue = inputWs.Cells(r, s)
                parameterFix(r) = paramValue
            Next
            
            s = 4
            For r = 1 To 11
                x = r + 12
                paramValue = inputWs.Cells(r, s)
                parameterFix(x) = paramValue
            Next
            
    End Select

End Sub


'--------------------------------------------------------
'Proced˙ra kopÌrovanie popisn˝ch d·t do v˝stupnej tabuæky
'--------------------------------------------------------
Sub arrayFill(conversionType As Integer)

Dim i As Integer
Dim usedColumns As Variant
Dim rowCount As Integer
    
    Select Case conversionType
        Case NA_SEC
            ' Inicializ·cia premenn˝ch - usedColumns = do ktor˝ch stÂpcov v˝stupnej tabuæky sa maj˙ napÂÚaù d·ta
            '                            rowCount = spoËÌta koæko riadkov je vyplnen˝ch vo v˝stupnej tabuæke (toæko riadkov bude naplnen˝ch)
            usedColumns = Array(0, 1, 2, 3, 6, 7, 12, 13, 14, 15, 16, 17, 18, 19, 23, 24, 25, 26, 27, 28, 29, 30, 31, 32, 33, 34, 35, 36, 37, 38, 39, 40, 41)
            rowCount = outputWs.Cells(Rows.Count, "T").End(xlUp).Row
        
            ' NapÂÚanie popisn˝ch d·t vo v˝stupnej tabuæky od riadok = 1 do riadok = rowCount
            For rowStep = 2 To rowCount
                For i = 1 To 32
                    outputWs.Cells(rowStep, usedColumns(i)).Value = parameterFix(i)
                Next i
            Next rowStep
            
        Case NA_REG
            ' Inicializ·cia premenn˝ch - usedColumns = do ktor˝ch stÂpcov v˝stupnej tabuæky sa maj˙ napÂÚaù d·ta
            '                            rowCount = spoËÌta koæko riadkov je vyplnen˝ch vo v˝stupnej tabuæke (toæko riadkov bude naplnen˝ch)
            usedColumns = Array(0, 1, 3, 4, 5, 13, 17, 18, 19, 20, 21, 22, 23, 24, 26, 27, 28, 29)
            rowCount = outputWs.Cells(Rows.Count, "N").End(xlUp).Row
            
            ' NapÂÚanie popisn˝ch d·t vo v˝stupnej tabuæky od riadok = 1 do riadok = rowCount
            For rowStep = 2 To rowCount
                For i = 1 To 17
                    outputWs.Cells(rowStep, usedColumns(i)).Value = parameterFix(i)
                Next i
            Next rowStep
            
        Case NA_MAIN
            ' Inicializ·cia premenn˝ch - usedColumns = do ktor˝ch stÂpcov v˝stupnej tabuæky sa maj˙ napÂÚaù d·ta
            '                            rowCount = spoËÌta koæko riadkov je vyplnen˝ch vo v˝stupnej tabuæke (toæko riadkov bude naplnen˝ch)
            usedColumns = Array(0, 1, 2, 3, 6, 13, 14, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 30, 31, 32, 33, 34, 35, 36)
            rowCount = outputWs.Cells(Rows.Count, "P").End(xlUp).Row
            
            ' NapÂÚanie popisn˝ch d·t vo v˝stupnej tabuæky od riadok = 1 do riadok = rowCount
            For rowStep = 2 To rowCount
                For i = 1 To 23
                    outputWs.Cells(rowStep, usedColumns(i)).Value = parameterFix(i)
                Next i
            Next rowStep
        
    End Select
    
End Sub


'-------------------------------
' Proced˙ra vymedzenie konverzie
'-------------------------------
Sub defineConversion(conversionType As Integer)

Dim startRange As String
Dim endRange As String
Dim startRangeInstrument As String
Dim endRangeInstrument As String
Dim startRangeBalance As String
Dim endRangeBalance As String
Dim typeFlag As String

    Select Case conversionType
        Case NA_SEC
            ' Inicializ·cia premenn˝ch obsahuj˙cich hodnoty riadiacich zaËiatkov/koncov (I = instrumenty, B = bilanËnÈ poloûky)
            startRangeInstrument = inputWs.Range("L3").Value
            endRangeInstrument = inputWs.Range("L4").Value
            startRangeBalance = inputWs.Range("L5").Value
            endRangeBalance = inputWs.Range("L6").Value
        
            ' Nastavenie riadiacich prvkov pre konverzn˝ cyklus a volanie konverznÈho algoritmu pre poloûky "I"
            typeFlag = "I"
            startRange = startRangeInstrument
            endRange = endRangeInstrument
        
                Call SECdataConversion(startRange, endRange, startRangeInstrument, endRangeInstrument, typeFlag)
            ' Nastavenie riadiacich prvkov pre konverzn˝ cyklus a volanie konverznÈho algoritmu pre poloûky "B"
            typeFlag = "B"
            startRange = startRangeBalance
            endRange = endRangeBalance
        
                Call SECdataConversion(startRange, endRange, startRangeInstrument, endRangeInstrument, typeFlag)
                
        Case NA_REG
            startRangeInstrument = inputWs.Range("J2").Value
            endRangeInstrument = inputWs.Range("J3").Value
            
            startRange = startRangeInstrument
            endRange = endRangeInstrument
            
                Call REGdataConversion(startRange, endRange)
                
        Case NA_MAIN
            startRangeInstrument = inputWs.Range("F2").Value
            endRangeInstrument = inputWs.Range("F3").Value
            
            startRange = startRangeInstrument
            endRange = endRangeInstrument
            
                Call MAINdataConversion(startRange, endRange)
                
    End Select
    
End Sub
    
    
'---------------------------
'Proced˙ra konverzia d·t SEC
'---------------------------
Sub SECdataConversion(startRange As String, endRange As String, startRangeInstrument As String, endRangeInstrument As String, typeFlag As String)

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
    
    ' Vymedzenie Ëi sa jedn· o klasick˝ cyklus pre inötrumenty, alebo o öpeci·lny cyklus pre bilanËnÈ poloûky
    If typeFlag = "I" Then
        Set leadingValueStart = Range(startRange)
        Set leadingValueEnd = Range(endRange)
        
        leadingRowStart = leadingValueStart.Row
        leadingRowEnd = leadingValueEnd.Row
        
        firstRow = leadingRowStart
        lastRow = leadingRowEnd
        
        ' Vymedzenie miesta, kde sa nach·dzaj˙ variabilnÈ popisnÈ d·ta, ktorÈ sa lÌöia v prÌpade poloûiek "I" alebo "B"
        specAccEntry = firstRow - 3
        specBoolStr = firstRow - 1
        
    ElseIf typeFlag = "B" Then
        Set leadingValueStart = Range(startRangeInstrument)
        Set leadingValueEnd = Range(endRangeInstrument)
    
        leadingRowStart = leadingValueStart.Row
        leadingRowEnd = leadingValueEnd.Row
    
        firstRow = Range(startRange).Row
        lastRow = Range(endRange).Row
        
        ' Vymedzenie miesta, kde sa nach·dzaj˙ variabilnÈ popisnÈ d·ta, ktorÈ sa lÌöia v prÌpade poloûiek "I" alebo "B"
        specAccEntry = firstRow - 2
        specBoolStr = firstRow - 1
    End If
    
    ' Inicializ·cia riadiacich hodnÙt - zaËiatoËn˝ stÂpec, koneËn˝ stÂpec
    leadingColStart = leadingValueStart.Column
    leadingColEnd = leadingValueEnd.Column
    

    ' V˝poËet poslednÈho vyplnenÈho riadku vo v˝stupnej tabuæke (i), od [i+1] sa zaËn˙ kopÌrovaù novÈ hodnoty
    i = outputWs.Cells(Rows.Count, "T").End(xlUp).Row
    i = i + 1
    
    ' Hlavn˝ cyklus konverzie
    For rowStep = firstRow To lastRow
        
        ' Kontrola na naËÌtanie riadku
        If inputWs.Cells(rowStep, leadingColStart - 1).Value = 1 Then
        
            For colStep = leadingColStart To leadingColEnd
                
                ' Kontrola na naËÌtanie stÂpca + kontrola na substring - confidential value
                boolSubString = ""
                boolString = inputWs.Cells(specBoolStr, colStep).Value
                
                If Len(boolString) > 0 Then
                
                    boolSubString = Left(boolString, 1)
                    confSubString = Right(boolString, 1)
                    If boolSubString = 1 Then
                        
                        ' NaËÌtanie d·t do pomocn˝ch premenn˝ch
                        obsValue = inputWs.Cells(rowStep, colStep).Value
                        counterpartArea = inputWs.Cells(leadingRowStart - 2, colStep).Value
                        refSector = inputWs.Cells(leadingRowStart - 4, colStep).Value
                        STO = inputWs.Cells(rowStep, leadingColStart - 3).Value
                        instrAsset = inputWs.Cells(rowStep, leadingColStart - 2).Value
                        maturity = inputWs.Cells(rowStep, leadingColStart - 4).Value
                        obsStatus = inputWs.Cells(rowStep, colStep + 1).Value
                        accountingEntry = inputWs.Cells(specAccEntry, colStep).Value
                        
                        ' NaËÌtanie hodnoty confidential value z pomocnÈho substringu
                        If confSubString = "1" Or "0" Then
                           confStatus = ""
                        Else: confStatus = confSubString
                        End If

                        ' Uloûenie hodnÙt z pomocn˝ch premenn˝ch do prÌsluön˝ch stÂpcov v˝stupnÈho h·rku riadku "i"
                        outputWs.Range("T" & i).Value = obsValue
                        outputWs.Range("D" & i).Value = counterpartArea
                        outputWs.Range("E" & i).Value = refSector
                        outputWs.Range("H" & i).Value = accountingEntry
                        outputWs.Range("I" & i).Value = STO
                        outputWs.Range("J" & i).Value = instrAsset
                        outputWs.Range("K" & i).Value = maturity
                        outputWs.Range("U" & i).Value = obsStatus
                        outputWs.Range("V" & i).Value = confStatus
                        
                        ' Inkrement·cia poËÌtadla riadkov
                        i = i + 1
                        
                    End If
                End If
            Next colStep
        End If
    Next rowStep
End Sub

'---------------------------
'Proced˙ra konverzia d·t REG
'---------------------------
Sub REGdataConversion(startRange As String, endRange As String)

Dim i As Integer
Dim firstRow As Integer
Dim lastRow As Integer
Dim leadingRowStart As Integer
Dim leadingRowEnd As Integer
Dim leadingColStart As Integer
Dim leadingColEnd As Integer
Dim leadingValueStart As Range
Dim leadingValueEnd As Range
    
Dim obsValue As Variant
Dim accountingEntry As String
Dim STO As String
Dim obsStatus As String
Dim confStatus As String
Dim transformation As String
Dim prices As String
Dim valuation As String
Dim unitMeasure As String
Dim unitMult As String
Dim activity As String
Dim refArea As String

Dim boolString As String

    ' Vymedzenie riadiacich prvkov cyklu

    Set leadingValueStart = Range(startRange)
    Set leadingValueEnd = Range(endRange)
    
    leadingRowStart = leadingValueStart.Row
    leadingRowEnd = leadingValueEnd.Row
    
    firstRow = leadingRowStart
    lastRow = leadingRowEnd
    
    ' Inicializ·cia riadiacich hodnÙt - zaËiatoËn˝ stÂpec, koneËn˝ stÂpec
    leadingColStart = leadingValueStart.Column
    leadingColEnd = leadingValueEnd.Column
    

    ' V˝poËet poslednÈho vyplnenÈho riadku vo v˝stupnej tabuæke (i), od [i+1] sa zaËn˙ kopÌrovaù novÈ hodnoty
    i = outputWs.Cells(Rows.Count, "T").End(xlUp).Row
    i = i + 1
    
    ' Hlavn˝ cyklus konverzie
    For rowStep = firstRow To lastRow
        
        ' Kontrola na naËÌtanie riadku
        If inputWs.Cells(rowStep, leadingColStart - 1).Value = 1 Then
        
            For colStep = leadingColStart To leadingColEnd
                
                ' Kontrola na naËÌtanie stÂpca
                boolString = inputWs.Cells(firstRow - 1, colStep).Value
                
                    If boolString = "1" Then
                        
                        ' NaËÌtanie d·t do pomocn˝ch premenn˝ch
                        STO = inputWs.Cells(leadingRowStart - 9, colStep).Value
                        transformation = inputWs.Cells(leadingRowStart - 8, colStep).Value
                        accountingEntry = inputWs.Cells(leadingRowStart - 7, colStep).Value
                        prices = inputWs.Cells(leadingRowStart - 6, colStep).Value
                        valuation = inputWs.Cells(leadingRowStart - 5, colStep).Value
                        unitMeasure = inputWs.Cells(leadingRowStart - 4, colStep).Value
                        unitMult = inputWs.Cells(leadingRowStart - 3, colStep).Value
                        activity = inputWs.Cells(leadingRowStart - 2, colStep).Value
                        obsValue = inputWs.Cells(rowStep, colStep).Value
                        refArea = inputWs.Cells(rowStep, leadingColStart - 2).Value
                        obsStatus = inputWs.Cells(rowStep, colStep + 1).Value
                        confStatus = inputWs.Cells(rowStep, colStep + 2).Value

                        ' Uloûenie hodnÙt z pomocn˝ch premenn˝ch do prÌsluön˝ch stÂpcov v˝stupnÈho h·rku riadku "i"
                        outputWs.Range("N" & i).Value = obsValue
                        outputWs.Range("B" & i).Value = refArea
                        outputWs.Range("F" & i).Value = accountingEntry
                        outputWs.Range("G" & i).Value = STO
                        outputWs.Range("H" & i).Value = activity
                        outputWs.Range("I" & i).Value = valuation
                        outputWs.Range("J" & i).Value = prices
                        outputWs.Range("K" & i).Value = unitMeasure
                        outputWs.Range("L" & i).Value = transformation
                        outputWs.Range("O" & i).Value = obsStatus
                        outputWs.Range("P" & i).Value = confStatus
                        outputWs.Range("Y" & i).Value = unitMult

                        ' Inkrement·cia poËÌtadla riadkov
                        i = i + 1
                        
                End If
            Next colStep
        End If
    Next rowStep
End Sub

'---------------------------
'Proced˙ra konverzia d·t MAIN
'---------------------------
Sub MAINdataConversion(startRange As String, endRange As String)

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
Dim accountingEntry As String
Dim STO As String
Dim instrAsset As String
Dim expenditure As String
Dim unitMeasure As String
Dim unitMult As String
Dim obsValue As Variant
Dim obsStatus As String
Dim confStatus As String
Dim activity As String
Dim timePeriod As String

Dim boolString As String

    ' Vymedzenie riadiacich prvkov cyklu

    Set leadingValueStart = Range(startRange)
    Set leadingValueEnd = Range(endRange)
    
    leadingRowStart = leadingValueStart.Row
    leadingRowEnd = leadingValueEnd.Row
    
    firstRow = leadingRowStart
    lastRow = leadingRowEnd
    
    ' Inicializ·cia riadiacich hodnÙt - zaËiatoËn˝ stÂpec, koneËn˝ stÂpec
    leadingColStart = leadingValueStart.Column
    leadingColEnd = leadingValueEnd.Column
    
    ' Vyplnenie hlaviËky dokumentu
    headerNames = Array("FREQ", "ADJUSTMENT", "REF_AREA", "COUNTERPART_AREA", "REF_SECTOR", "COUNTERPART_SECTOR", "ACCOUNTING_ENTRY", "STO", "INSTR_ASSET", "ACTIVITY", "EXPENDITURE", "UNIT_MEASURE", "PRICES", "TRANSFORMATION", "TIME_PERIOD", "OBS_VALUE", "OBS_STATUS", "CONF_STATUS", "COMMENT_OBS", "PRE_BREAK_VALUE", "EMBARGO_DATE", "REF_PERIOD_DETAIL", "TIME_FORMAT", "TIME_PER_COLLECT", "REF_YEAR_PRICE", "DECIMALS", "TABLE_IDENTIFIER", "TITLE", "UNIT_MULT", "LAST_UPDATE", "COMPILING_ORG", "COMMENT_DSET", "COMMENT_TS", "DATA_COMP", "CURRENCY", "DISS_ORG")
    columnNames = Array("A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z", "AA", "AB", "AC", "AD", "AE", "AF", "AG", "AH", "AI", "AJ")
    For i = 0 To 35
        outputWs.Range(columnNames(i) & "1").Value = headerNames(i)
    Next i
    
    ' V˝poËet poslednÈho vyplnenÈho riadku vo v˝stupnej tabuæke (i), od [i+1] sa zaËn˙ kopÌrovaù novÈ hodnoty
    i = outputWs.Cells(Rows.Count, "A").End(xlUp).Row
    i = i + 1
    
    ' Hlavn˝ cyklus konverzie
    For rowStep = firstRow To lastRow
        
        ' Kontrola na naËÌtanie riadku
        If inputWs.Cells(rowStep, leadingColStart - 1).Value = 1 Then
        
            For colStep = leadingColStart To leadingColEnd
                
                ' Kontrola na naËÌtanie stÂpca
                boolString = inputWs.Cells(firstRow - 1, colStep).Value
                
                    If boolString = "1" Then
                        
                        ' NaËÌtanie d·t do pomocn˝ch premenn˝ch
                        counterpartArea = inputWs.Cells(leadingRowStart - 9, colStep).Value
                        refSector = inputWs.Cells(leadingRowStart - 8, colStep).Value
                        accountingEntry = inputWs.Cells(leadingRowStart - 7, colStep).Value
                        STO = inputWs.Cells(leadingRowStart - 6, colStep).Value
                        instrAsset = inputWs.Cells(leadingRowStart - 5, colStep).Value
                        expenditure = inputWs.Cells(leadingRowStart - 4, colStep).Value
                        unitMeasure = inputWs.Cells(leadingRowStart - 3, colStep).Value
                        unitMult = inputWs.Cells(leadingRowStart - 2, colStep).Value
                        obsValue = inputWs.Cells(rowStep, colStep).Value
                        obsStatus = inputWs.Cells(rowStep, colStep + 1).Value
                        confStatus = inputWs.Cells(rowStep, colStep + 2).Value
                        activity = inputWs.Cells(rowStep, colStep + 3).Value
                        timePeriod = inputWs.Cells(rowStep, colStep + 4).Value

                        ' Uloûenie hodnÙt z pomocn˝ch premenn˝ch do prÌsluön˝ch stÂpcov v˝stupnÈho h·rku riadku "i"
                        outputWs.Range("P" & i).Value = obsValue
                        outputWs.Range("D" & i).Value = counterpartArea
                        outputWs.Range("E" & i).Value = refSector
                        outputWs.Range("G" & i).Value = accountingEntry
                        outputWs.Range("H" & i).Value = STO
                        outputWs.Range("I" & i).Value = instrAsset
                        outputWs.Range("J" & i).Value = activity
                        outputWs.Range("K" & i).Value = expenditure
                        outputWs.Range("L" & i).Value = unitMeasure
                        outputWs.Range("O" & i).Value = timePeriod
                        outputWs.Range("Q" & i).Value = obsStatus
                        outputWs.Range("R" & i).Value = confStatus
                        outputWs.Range("AC" & i).Value = unitMult

                        ' Inkrement·cia poËÌtadla riadkov
                        i = i + 1
                        
                End If
            Next colStep
        End If
    Next rowStep
End Sub

'---------------------------
'Proced˙ra reset formul·rov
'---------------------------
Sub unloadForms()

    mainForm.chbLeft.Value = False
    mainForm.chbRight.Value = False
    mainForm.tbSourceFile.Value = ""
    mainForm.lbLeft.Clear
    mainForm.lbRight.Clear
    mainForm.optSEC.Value = False
    mainForm.optREG.Value = False
    mainForm.optPENS.Value = False
    mainForm.optMAIN.Value = False
    mainForm.optSU.Value = False

    Unload progressForm
    
    mainForm.Show vbModeless

End Sub

'---------------------------
'Proced˙ra vypnutia programu
'---------------------------
Sub appClose()

    Application.Visible = True

    Set xlOld = GetObject(instanceName).Application

    If xlOld.Workbooks.Count = 1 Then
        xlOld.ThisWorkbook.Saved = True
        If xlApp Is Nothing Then
        Else
            xlApp.Quit
        End If
            If xlNew Is Nothing Then
            Else
                xlNew.Quit
            End If
        xlOld.Quit
    ElseIf xlOld.Workbooks.Count > 1 Then
        xlOld.ThisWorkbook.Saved = True
        xlOld.Visible = True
         If xlApp Is Nothing Then
        Else
            xlApp.Quit
        End If
            If xlNew Is Nothing Then
            Else
                xlNew.Quit
            End If
        xlOld.ThisWorkbook.Close
    Else
        MsgBox "Nastala chyba. ProsÌm kontaktujte spr·vcu aplik·cie.", , "Chyba"
    End If

End Sub

