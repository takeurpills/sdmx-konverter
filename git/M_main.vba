Attribute VB_Name = "M_main"
Option Explicit

'----------------------------------------------
'Inicializacna procedura pri spusteni programu
'----------------------------------------------
Sub ProgramInit()
    
    PBL_programVersion = "v1.02"
    PBL_programName = ActiveWorkbook.FullName
    
    F_main.Show vbModeless
    
End Sub
    
'------------------------------------
'Hlavna riadiaca procedura konverzie
'------------------------------------
Sub MainSub(conversionType As Integer)
    
Const SUB_NAME = "mainSub"
    
Dim i As Integer
Dim successCounter As Integer
Dim failCounter As Integer
Dim saveName As String
Dim folderPath As String
Dim timeStamp As String
Dim typeString As String
Dim outputFile As String
Dim infoMsg As String
Dim successMsg As String
Dim workbookName As String
Dim conversionCheck As Boolean
Dim progIndicator As Integer
Dim errorIndi As Integer
Dim checkedInputWsId() As Integer
Dim failInputWs() As String
Dim startVal As Range, endVal As Range

    progIndicator = 0
    PBL_conversionOk = 0
    PBL_conversionFail = 0

'Kontrola ci bol zvoleny vstupny subor a ci je aspon jeden harok zvoleny na konverziu
    If PBL_fileToOpen <> False Then
        If (Not PBL_inputWsId) = True Then
            infoMsg = "Nie sú zvolené pracovné hárky pre konverziu!"
            MsgBox infoMsg, vbExclamation, "Informatívna chyba"
        Else
            Set PBL_xlNew = CreateObject("Excel.Application")

            PBL_xlNew.Workbooks.Add (1)
            Set PBL_outputWs = PBL_xlNew.Workbooks(1).Worksheets(1)
            
            PBL_xlNew.ScreenUpdating = False
            PBL_xlNew.Calculation = xlCalculationManual

            successCounter = 1
            failCounter = 1

'Cyklus na test spravnosti formatu vybranych harkov
            For i = 1 To UBound(PBL_inputWsId)
            
                conversionCheck = False

                Set PBL_inputWs = PBL_inputWb.Worksheets(PBL_inputWsId(i))
                PBL_worksheetName = PBL_inputWs.name

                Select Case conversionType
                    Case PBL_SEC
                        Set startVal = PBL_inputWs.Range("F2")
                        Set endVal = PBL_inputWs.Range("F3")
                        If PBL_inputWs.Cells(13, 1).Value = "COLL_PERIOD" And PBL_inputWs.Cells(1, 6).Value = "NA_SEC" _
                        And cellValueRefTest(startVal, endVal) = True Then
                            conversionCheck = True
                        End If
                    Case PBL_T1100
                        Set startVal = PBL_inputWs.Range("F2")
                        Set endVal = PBL_inputWs.Range("F3")
                        If PBL_inputWs.Cells(13, 1).Value = "COLL_PERIOD" And PBL_inputWs.Cells(1, 6).Value = "NA_SEC_1100" _
                        And cellValueRefTest(startVal, endVal) = True Then
                            conversionCheck = True
                        End If
                    Case PBL_T9XX
                        Set startVal = PBL_inputWs.Range("K2")
                        Set endVal = PBL_inputWs.Range("K3")
                        If PBL_inputWs.Cells(19, 1).Value = "OBS_EDP_WBB" And PBL_inputWs.Cells(1, 11).Value = "NA_SEC_9XX" _
                        And cellValueRefTest(startVal, endVal) = True Then
                            conversionCheck = True
                        End If
                    Case PBL_T200
                        Set startVal = PBL_inputWs.Range("H2")
                        Set endVal = PBL_inputWs.Range("H3")
                        If PBL_inputWs.Cells(13, 1).Value = "TIME_PERIOD" And PBL_inputWs.Cells(1, 8).Value = "NA_SEC_200" _
                        And cellValueRefTest(startVal, endVal) = True Then
                            conversionCheck = True
                        End If
                    Case PBL_REG
                        Set startVal = PBL_inputWs.Range("F2")
                        Set endVal = PBL_inputWs.Range("F3")
                        If PBL_inputWs.Cells(11, 1).Value = "TIME_PERIOD" And PBL_inputWs.Cells(1, 6).Value = "NA_REG" _
                        And cellValueRefTest(startVal, endVal) = True Then
                            conversionCheck = True
                        End If
                    Case PBL_PENS
                        Set startVal = PBL_inputWs.Range("F2")
                        Set endVal = PBL_inputWs.Range("F3")
                        If PBL_inputWs.Cells(12, 1).Value = "UNIT_MULT" And PBL_inputWs.Cells(1, 6).Value = "NA_PENS" _
                        And cellValueRefTest(startVal, endVal) = True Then
                            conversionCheck = True
                        End If
                    Case PBL_MAIN
                        Set startVal = PBL_inputWs.Range("F2")
                        Set endVal = PBL_inputWs.Range("F3")
                        If PBL_inputWs.Cells(11, 1).Value = "TIME_PER_COLLECT" And PBL_inputWs.Cells(1, 6).Value = "NA_MAIN" _
                        And cellValueRefTest(startVal, endVal) = True Then
                            conversionCheck = True
                        End If
                    Case PBL_SU
                        Set startVal = PBL_inputWs.Range("F2")
                        Set endVal = PBL_inputWs.Range("F3")
                        If PBL_inputWs.Cells(12, 1).Value = "UNIT_MEASURE" And PBL_inputWs.Cells(1, 6).Value = "NA_SU" _
                        And cellValueRefTest(startVal, endVal) = True Then
                            conversionCheck = True
                        End If
                End Select

                If conversionCheck = True Then
                    ReDim Preserve checkedInputWsId(successCounter)
                    checkedInputWsId(successCounter) = PBL_inputWsId(i)
        
                    successCounter = successCounter + 1
                Else
                    ReDim Preserve failInputWs(failCounter)
                    failInputWs(failCounter) = PBL_worksheetName
                    
                    failCounter = failCounter + 1
                    PBL_conversionFail = IncrementConversions(PBL_FAIL)
                End If
            Next i
            
            If failCounter > 1 Then
                infoMsg = "Niektoré hárky nemajú správny formát pre zvolený typ konverzie:" & vbNewLine
                For i = 1 To UBound(failInputWs)
                    infoMsg = infoMsg & "   " & Chr(149) & failInputWs(i) & vbNewLine
                Next i
                infoMsg = infoMsg & vbNewLine & "Konverzia týchto hárkov sa nevykoná!"
                MsgBox infoMsg, vbExclamation, "Informatívna chyba"
            End If
            
            If successCounter > 1 Then
                For i = 1 To UBound(checkedInputWsId)
                
                Set PBL_inputWs = PBL_inputWb.Worksheets(checkedInputWsId(i))
                workbookName = Left(PBL_inputWb.name, (InStrRev(PBL_inputWb.name, ".", -1, vbTextCompare) - 1))
                PBL_worksheetName = PBL_inputWs.name

'Spustanie procedur konverzie
                    If progIndicator = 0 Then
                        F_main.Hide
                        F_progress.Show vbModeless
                        progIndicator = 1
                    End If
        
                    errorIndi = PBL_conversionFail
        
                    Call ArrayPush(conversionType)
                    Call DefineConversion(conversionType)
        
                    If errorIndi = PBL_conversionFail Then
                        Call ArrayFill(conversionType)
                        PBL_conversionOk = IncrementConversions(PBL_OK)
                    End If
                Next i
            End If

'Radenie hodnot, ulozenie vystupu a object-cleanup
            If PBL_conversionOk > 0 Then
            
                DataSort (conversionType)
            
                folderPath = PBL_xlApp.ActiveWorkbook.Path
                timeStamp = Format(CStr(Now), "yyyy_mm_dd_hhmmss")
                saveName = folderPath & "\" & workbookName & "_" & timeStamp
                With PBL_xlNew
                    .DecimalSeparator = "."
                    .UseSystemSeparators = False
                    .Workbooks(1).SaveAs fileName:=saveName, FileFormat:=xlCSV, local:=True
                End With
            End If
            
            With PBL_xlNew
                .Calculation = xlCalculationAutomatic
                .ScreenUpdating = True
                .DecimalSeparator = ","
                .UseSystemSeparators = True
                .Workbooks(1).Close False
                .Quit
            End With
            
'Vymaze posledny carriage return z .csv [classic excel bug = feature :) ]
            If PBL_conversionOk > 0 Then
                deleteLastLine (saveName)
            End If
            Call UnloadForms
             
            With PBL_xlApp
                If .Workbooks(1).ReadOnly = True Then
                    .Workbooks(1).Close False
                End If
                .Quit
            End With
            
            Set PBL_xlApp = Nothing
            Set PBL_xlNew = Nothing
            Set PBL_outputWs = Nothing
            Set PBL_inputWb = Nothing
            Set PBL_inputWs = Nothing
            
            successMsg = "Konverzia hárkov bola dokonèená." & vbNewLine & vbNewLine
            successMsg = successMsg & "• Poèet úspešných konverzií: " & PBL_conversionOk & vbNewLine
            successMsg = successMsg & "• Poèet neúspešných konverzií: " & PBL_conversionFail
            
            MsgBox successMsg, vbInformation, "Informácia"
        End If
        
    Else
        infoMsg = "Nie je zvolený súbor pre konverziu!"
        MsgBox infoMsg, vbExclamation, "Informatívna chyba"
    End If

End Sub


'---------------------------
'Vypnutie programu
'---------------------------
Sub AppClose()

    Set PBL_xlOld = GetObject(PBL_programName).Application
    
    PBL_xlOld.Visible = True

    If PBL_xlOld.Workbooks.count = 1 Then
        PBL_xlOld.ThisWorkbook.Saved = True
        If PBL_xlApp Is Nothing Then
        Else
            PBL_xlApp.Quit
            Set PBL_xlApp = Nothing
        End If
        If PBL_xlNew Is Nothing Then
        Else
            PBL_xlNew.Quit
            Set PBL_xlNew = Nothing
        End If
        PBL_xlOld.Quit
    ElseIf PBL_xlOld.Workbooks.count > 1 Then
        PBL_xlOld.ThisWorkbook.Saved = True
        PBL_xlOld.Visible = True
         If PBL_xlApp Is Nothing Then
        Else
            PBL_xlApp.Quit
            Set PBL_xlApp = Nothing
        End If
        If PBL_xlNew Is Nothing Then
        Else
            PBL_xlNew.Quit
            Set PBL_xlNew = Nothing
        End If
        PBL_xlOld.ThisWorkbook.Close
    Else
        MsgBox "Nastala chyba. Prosím kontaktujte správcu aplikácie.", , "Chyba"
    End If

End Sub
