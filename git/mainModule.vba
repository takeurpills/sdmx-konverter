Attribute VB_Name = "mainModule"

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                                                                                       '
'  N·zov:  Konvertor ESA2010                                                            '
'  Autor:  Martin TÛth - ätatistick˝ ˙rad SR                                            '
'                                                                                       '
'  Popis:                                                                               '
'                                                                                       '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


'--------------------------------
'Deklar·cia glob·lnych premenn˝ch
'--------------------------------
    
    Option Explicit
    
    '--------------
    'Pohyby a skoky
    '--------------
    Dim rowStep As Integer
    Dim colStep As Integer
           
    '------------
    'PomocnÈ pole
    '------------
    Dim parameterFix() As Variant
    Public myArray() As Integer
    
    '-------------
    'PouûitÈ h·rky
    '-------------
    Dim inputWorksheet As Worksheet
    Dim outputWorksheet As Worksheet
    Dim inputWorkbook As Workbook
    
    '---------------
    'InÈ
    '---------------
    Public xlApp As Object
    Dim xlNew As Object
    Dim xlOld As Excel.Application
    
    Dim fileToOpen As Variant

    
'-----------------
' Hlavn· proced˙ra
'-----------------
Sub mainSub(conversionAlg As Integer)
    
Dim i As Integer
Dim saveName As String
Dim folderPath As String
Dim timeStamp As String
Dim nameString As String
Dim typeString As String
Dim outputFile As String
Dim errorMsg As String
Dim conversionCheck As Boolean
Dim progIndicator As Integer

    progIndicator = 0
    
    If fileToOpen <> False Then
        If (Not myArray) = True Then
            MsgBox "Nie s˙ zvolenÈ pracovnÈ h·rky pre konverziu!", , "Chyba"
        Else
            For i = 1 To UBound(myArray)
        
            Set xlNew = CreateObject("Excel.Application")
            xlNew.ScreenUpdating = False
            
            Select Case conversionAlg
                Case 1
                    outputFile = "esa2010-NA_SEC_accounts.xlsx"
                    typeString = "NA_SEC"
                Case 2
                    outputFile = "esa2010-NA_REG.xlsx"
                    typeString = "NA_REG"
                Case 3
                    outputFile = "esa2010-NA_PENS.xlsx"
                    typeString = "NA_PENS"
                Case 4
                    outputFile = "esa2010-NA_MAIN_accounts.xlsx"
                    typeString = "NA_MAIN"
                Case 5
                    outputFile = "esa2010-NA_SU.xlsx"
                    typeString = "NA_SU"
            End Select
                        
            ' Spustenie "ErrHandler" ak sa vyskytne chyba
            On Error GoTo Errhandler
            
            Set outputWorksheet = xlNew.Workbooks.Open(ThisWorkbook.Path & "\" & outputFile).Worksheets(1)
            
            ' Vypnutie "ErrHandler"
            On Error GoTo 0

            Set inputWorksheet = inputWorkbook.Worksheets(myArray(i))
            nameString = inputWorksheet.name

            conversionCheck = False
            
            Select Case conversionAlg
                Case 1
                    If inputWorksheet.Cells(1, 1).Value = "FREQ" And inputWorksheet.Cells(6, 1).Value = "EXPENDITURE" Then
                    conversionCheck = True
                    End If
                Case 2
                    If inputWorksheet.Cells(1, 10).Value = "REG" Then
                    conversionCheck = True
                    End If
                Case 4
                    If inputWorksheet.Cells(1, 1).Value = "FREQ" And inputWorksheet.Cells(1, 6).Value = "MAIN" Then
                    conversionCheck = True
                    End If
                Case Else
                    MsgBox "Nastala chyba. ProsÌm kontaktujte spr·vcu aplik·cie.", vbOKOnly, "Chyba"
            End Select
            
            If conversionCheck = True Then
            
                ' Sp˙ötanie proced˙r
                
                If progIndicator = 0 Then
                    mainForm.Hide
                    progressForm.Show vbModeless
                    progIndicator = 1
                End If
                                
                Call arrayPush(conversionAlg)
                Call defineConversion(conversionAlg)
                Call arrayFill(conversionAlg)
                
                ' Uloûenie v˝stupu
                folderPath = xlApp.ActiveWorkbook.Path
                timeStamp = Format(CStr(Now), "yyyy_mm_dd_hh_mm")
                saveName = folderPath & "\" & typeString & "_" & nameString & "_" & timeStamp
    
                xlNew.Workbooks(1).SaveAs Filename:=saveName, FileFormat:=51
                xlNew.ScreenUpdating = True
                xlNew.Quit
                
            Else
                MsgBox "Zvolen˝ h·rok - """ & nameString & """ nem· spr·vny form·t alebo bol zvolen˝ nespr·vny typ konverzie!", , "Chyba"
                xlNew.Quit
            End If
            Next
            
            Call unloadForms
            
            xlApp.Quit
        
'            If xlOld.Workbooks.Count = 1 Then
'               xlOld.ThisWorkbook.Saved = True
'               xlOld.Quit
'            ElseIf xlOld.Workbooks.Count > 1 Then
'               xlOld.ThisWorkbook.Saved = True
'               xlOld.Visible = True
'               xlOld.ThisWorkbook.Close
'            Else
'                MsgBox "Nastala chyba. ProsÌm kontaktujte spr·vcu aplik·cie.", vbOKOnly, "Chyba"
'            End If
            
        End If
        
    Else: MsgBox "Nie je zvolen˝ s˙bor pre konverziu!", , "Chyba"
    End If
    Exit Sub
    
Errhandler:

    errorMsg = "Nebol n·jden˝ s˙bor - """ & outputFile & """ !" & vbNewLine & vbNewLine
    errorMsg = errorMsg & "ï ProsÌm umiestnite tento s˙bor do prieËinka tejto aplik·cie!"

      MsgBox errorMsg, , "Chyba"
    
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
        Set xlOld = GetObject(, "Excel.Application")
        Set xlApp = CreateObject("Excel.Application")
        Set inputWorkbook = xlApp.Workbooks.Open(Filename:=fileToOpen, ReadOnly:=True)
      
        mainForm.tbSourceFile.Value = inputWorkbook.FullName
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
Sub arrayPush(conversionAlg As Integer)

Dim i As Integer
Dim r As Integer
Dim s As Integer
Dim x As Integer
Dim paramValue As String
    
    Select Case conversionAlg
        Case 1
            ' NaËÌtanie fixn˝ch popisn˝ch d·t do pomocnÈho pola z hlaviËky h·rku - 6*5 = 30 hodnÙt
            i = 0
        
            ReDim parameterFix(1 To 32)
        
            For s = 2 To 10 Step 2
                For r = 1 To 6
                    x = r + 6 * i
                    paramValue = inputWorksheet.Cells(r, s)
                    parameterFix(x) = paramValue
                Next
                i = i + 1
            Next
        
            ' NaËÌtanie zvyön˝ch dvoch hodnÙt popisn˝ch d·t - neboli zahrnutÈ do symetrickÈho cyklu
            parameterFix(31) = inputWorksheet.Cells(1, 12)
            parameterFix(32) = inputWorksheet.Cells(2, 12)

        Case 2
            ' NaËÌtanie fixn˝ch popisn˝ch d·t do pomocnÈho pola z hlaviËky h·rku - 6*5 = 30 hodnÙt
            i = 0
        
            ReDim parameterFix(1 To 17)
        
            For s = 2 To 6 Step 2
                For r = 1 To 5
                    x = r + 5 * i
                    paramValue = inputWorksheet.Cells(r, s)
                    parameterFix(x) = paramValue
                Next
                i = i + 1
            Next
        
            ' NaËÌtanie zvyön˝ch dvoch hodnÙt popisn˝ch d·t - neboli zahrnutÈ do symetrickÈho cyklu
            parameterFix(16) = inputWorksheet.Cells(1, 8)
            parameterFix(17) = inputWorksheet.Cells(2, 8)
            
        Case 4
            ' NaËÌtanie fixn˝ch popisn˝ch d·t do pomocnÈho pola z hlaviËky h·rku - 2*11 = 22 hodnÙt
            i = 0
            
            ReDim parameterFix(1 To 23)
            
            For s = 2 To 4 Step 2
                For r = 1 To 11
                    x = r + 11 * i
                    paramValue = inputWorksheet.Cells(r, s)
                    parameterFix(x) = paramValue
                Next
                i = i + 1
            Next
            ' NaËÌtanie zvyönej 23ej (pozÌcia [12][1]) hodnoty popisn˝ch d·t - nebola zahrnut· do symetrickÈho cyklu
            parameterFix(12) = inputWorksheet.Cells(12, 1)
            
    End Select

End Sub


'--------------------------------------------------------
'Proced˙ra kopÌrovanie popisn˝ch d·t do v˝stupnej tabuæky
'--------------------------------------------------------
Sub arrayFill(conversionAlg As Integer)

Dim i As Integer
Dim usedColumns As Variant
Dim rowCount As Integer
    
    Select Case conversionAlg
        Case 1
            ' Inicializ·cia premenn˝ch - usedColumns = do ktor˝ch stÂpcov v˝stupnej tabuæky sa maj˙ napÂÚaù d·ta
            '                            rowCount = spoËÌta koæko riadkov je vyplnen˝ch vo v˝stupnej tabuæke (toæko riadkov bude naplnen˝ch)
            usedColumns = Array(0, 1, 2, 3, 6, 7, 12, 13, 14, 15, 16, 17, 18, 19, 23, 24, 25, 26, 27, 28, 29, 30, 31, 32, 33, 34, 35, 36, 37, 38, 39, 40, 41)
            rowCount = outputWorksheet.Cells(Rows.Count, "T").End(xlUp).Row
        
            ' NapÂÚanie popisn˝ch d·t vo v˝stupnej tabuæky od riadok = 1 do riadok = rowCount
            For rowStep = 2 To rowCount
                For i = 1 To 32
                    outputWorksheet.Cells(rowStep, usedColumns(i)).Value = parameterFix(i)
                Next i
            Next rowStep
            
        Case 2
            ' Inicializ·cia premenn˝ch - usedColumns = do ktor˝ch stÂpcov v˝stupnej tabuæky sa maj˙ napÂÚaù d·ta
            '                            rowCount = spoËÌta koæko riadkov je vyplnen˝ch vo v˝stupnej tabuæke (toæko riadkov bude naplnen˝ch)
            usedColumns = Array(0, 1, 3, 4, 5, 13, 17, 18, 19, 20, 21, 22, 23, 24, 26, 27, 28, 29)
            rowCount = outputWorksheet.Cells(Rows.Count, "N").End(xlUp).Row
            
            ' NapÂÚanie popisn˝ch d·t vo v˝stupnej tabuæky od riadok = 1 do riadok = rowCount
            For rowStep = 2 To rowCount
                For i = 1 To 17
                    outputWorksheet.Cells(rowStep, usedColumns(i)).Value = parameterFix(i)
                Next i
            Next rowStep
            
        Case 4
            ' Inicializ·cia premenn˝ch - usedColumns = do ktor˝ch stÂpcov v˝stupnej tabuæky sa maj˙ napÂÚaù d·ta
            '                            rowCount = spoËÌta koæko riadkov je vyplnen˝ch vo v˝stupnej tabuæke (toæko riadkov bude naplnen˝ch)
            usedColumns = Array(0, 1, 2, 5, 12, 13, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 29, 30, 31, 32, 33, 34, 35)
            rowCount = outputWorksheet.Cells(Rows.Count, "N").End(xlUp).Row
            
            ' NapÂÚanie popisn˝ch d·t vo v˝stupnej tabuæky od riadok = 1 do riadok = rowCount
            For rowStep = 2 To rowCount
                For i = 1 To 23
                    outputWorksheet.Cells(rowStep, usedColumns(i)).Value = parameterFix(i)
                Next i
            Next rowStep
        
    End Select
    
End Sub


'-------------------------------
' Proced˙ra vymedzenie konverzie
'-------------------------------
Sub defineConversion(conversionAlg As Integer)

Dim startRange As String
Dim endRange As String
Dim startRangeInstrument As String
Dim endRangeInstrument As String
Dim startRangeBalance As String
Dim endRangeBalance As String
Dim typeFlag As String

    Select Case conversionAlg
        Case 1
            ' Inicializ·cia premenn˝ch obsahuj˙cich hodnoty riadiacich zaËiatkov/koncov (I = instrumenty, B = bilanËnÈ poloûky)
            startRangeInstrument = inputWorksheet.Range("L3").Value
            endRangeInstrument = inputWorksheet.Range("L4").Value
            startRangeBalance = inputWorksheet.Range("L5").Value
            endRangeBalance = inputWorksheet.Range("L6").Value
        
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
                
        Case 2
            startRangeInstrument = inputWorksheet.Range("J2").Value
            endRangeInstrument = inputWorksheet.Range("J3").Value
            
            startRange = startRangeInstrument
            endRange = endRangeInstrument
            
                Call REGdataConversion(startRange, endRange, startRangeInstrument, endRangeInstrument)
                
        Case 4
            startRangeInstrument = inputWorksheet.Range("F2").Value
            endRangeInstrument = inputWorksheet.Range("F3").Value
            
            startRange = startRangeInstrument
            endRange = endRangeInstrument
            
                Call MAINdataConversion(startRange, endRange, startRangeInstrument, endRangeInstrument)
                
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

Dim obsValue As Double
Dim counterPartArea As String
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
    i = outputWorksheet.Cells(Rows.Count, "T").End(xlUp).Row
    i = i + 1
    
    ' Hlavn˝ cyklus konverzie
    For rowStep = firstRow To lastRow
        
        ' Kontrola na naËÌtanie riadku
        If inputWorksheet.Cells(rowStep, leadingColStart - 1).Value = 1 Then
        
            For colStep = leadingColStart To leadingColEnd
                
                ' Kontrola na naËÌtanie stÂpca + kontrola na substring - confidential value
                boolSubString = ""
                boolString = inputWorksheet.Cells(specBoolStr, colStep).Value
                
                If Len(boolString) > 0 Then
                
                    boolSubString = Left(boolString, 1)
                    confSubString = Right(boolString, 1)
                    If boolSubString = 1 Then
                        
                        ' NaËÌtanie d·t do pomocn˝ch premenn˝ch
                        obsValue = inputWorksheet.Cells(rowStep, colStep).Value
                        counterPartArea = inputWorksheet.Cells(leadingRowStart - 2, colStep).Value
                        refSector = inputWorksheet.Cells(leadingRowStart - 4, colStep).Value
                        STO = inputWorksheet.Cells(rowStep, leadingColStart - 3).Value
                        instrAsset = inputWorksheet.Cells(rowStep, leadingColStart - 2).Value
                        maturity = inputWorksheet.Cells(rowStep, leadingColStart - 4).Value
                        obsStatus = inputWorksheet.Cells(rowStep, colStep + 1).Value
                        accountingEntry = inputWorksheet.Cells(specAccEntry, colStep).Value
                        
                        ' NaËÌtanie hodnoty confidential value z pomocnÈho substringu
                        If confSubString = "1" Or "0" Then
                           confStatus = ""
                        Else: confStatus = confSubString
                        End If

                        ' Uloûenie hodnÙt z pomocn˝ch premenn˝ch do prÌsluön˝ch stÂpcov v˝stupnÈho h·rku riadku "i"
                        outputWorksheet.Range("T" & i).Value = obsValue
                        outputWorksheet.Range("D" & i).Value = counterPartArea
                        outputWorksheet.Range("E" & i).Value = refSector
                        outputWorksheet.Range("H" & i).Value = accountingEntry
                        outputWorksheet.Range("I" & i).Value = STO
                        outputWorksheet.Range("J" & i).Value = instrAsset
                        outputWorksheet.Range("K" & i).Value = maturity
                        outputWorksheet.Range("U" & i).Value = obsStatus
                        outputWorksheet.Range("V" & i).Value = confStatus
                        
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
Sub REGdataConversion(startRange As String, endRange As String, startRangeInstrument As String, endRangeInstrument As String)

Dim i As Integer
Dim firstRow As Integer
Dim lastRow As Integer
Dim leadingRowStart As Integer
Dim leadingRowEnd As Integer
Dim leadingColStart As Integer
Dim leadingColEnd As Integer
Dim leadingValueStart As Range
Dim leadingValueEnd As Range
    
Dim obsValue As Double
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

    ' Vymedzenie Ëi sa jedn· o klasick˝ cyklus pre inötrumenty, alebo o öpeci·lny cyklus pre bilanËnÈ poloûky

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
    i = outputWorksheet.Cells(Rows.Count, "T").End(xlUp).Row
    i = i + 1
    
    ' Hlavn˝ cyklus konverzie
    For rowStep = firstRow To lastRow
        
        ' Kontrola na naËÌtanie riadku
        If inputWorksheet.Cells(rowStep, leadingColStart - 1).Value = 1 Then
        
            For colStep = leadingColStart To leadingColEnd
                
                ' Kontrola na naËÌtanie stÂpca
                boolString = inputWorksheet.Cells(firstRow - 1, colStep).Value
                
                    If boolString = "1" Then
                        
                        ' NaËÌtanie d·t do pomocn˝ch premenn˝ch
                        STO = inputWorksheet.Cells(leadingRowStart - 9, colStep).Value
                        transformation = inputWorksheet.Cells(leadingRowStart - 8, colStep).Value
                        accountingEntry = inputWorksheet.Cells(leadingRowStart - 7, colStep).Value
                        prices = inputWorksheet.Cells(leadingRowStart - 6, colStep).Value
                        valuation = inputWorksheet.Cells(leadingRowStart - 5, colStep).Value
                        unitMeasure = inputWorksheet.Cells(leadingRowStart - 4, colStep).Value
                        unitMult = inputWorksheet.Cells(leadingRowStart - 3, colStep).Value
                        activity = inputWorksheet.Cells(leadingRowStart - 2, colStep).Value
                        obsValue = inputWorksheet.Cells(rowStep, colStep).Value
                        refArea = inputWorksheet.Cells(rowStep, leadingColStart - 2).Value
                        obsStatus = inputWorksheet.Cells(rowStep, colStep + 1).Value
                        confStatus = inputWorksheet.Cells(rowStep, colStep + 2).Value

                        ' Uloûenie hodnÙt z pomocn˝ch premenn˝ch do prÌsluön˝ch stÂpcov v˝stupnÈho h·rku riadku "i"
                        outputWorksheet.Range("N" & i).Value = obsValue
                        outputWorksheet.Range("B" & i).Value = refArea
                        outputWorksheet.Range("F" & i).Value = accountingEntry
                        outputWorksheet.Range("G" & i).Value = STO
                        outputWorksheet.Range("H" & i).Value = activity
                        outputWorksheet.Range("I" & i).Value = valuation
                        outputWorksheet.Range("J" & i).Value = prices
                        outputWorksheet.Range("K" & i).Value = unitMeasure
                        outputWorksheet.Range("L" & i).Value = transformation
                        outputWorksheet.Range("O" & i).Value = obsStatus
                        outputWorksheet.Range("P" & i).Value = confStatus
                        outputWorksheet.Range("Y" & i).Value = unitMult

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

Dim box As OLEObject

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
    
'Dim MyActualForm As mainForm

'If TypeName(MyActualForm) = "Nothing" Then Set MyActualForm = New mainForm
'MyActualForm.Show vbModeless

End Sub
