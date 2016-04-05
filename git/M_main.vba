Attribute VB_Name = "M_main"

Option Explicit

'----------------------------------------------
' Inicializaèná procedúra pri spustení programu
'----------------------------------------------
Sub ProgramInit()
    
    PBL_programVersion = "v0.5"
    PBL_programName = ActiveWorkbook.FullName
    
    F_main.Show vbModeless
    
End Sub
    
'------------------------------------
' Hlavná riadiaca procedúra konverzie
'------------------------------------
Sub MainSub(conversionType As Integer)
    
Const SUB_NAME = "mainSub"
    
Dim i As Integer
Dim saveName As String
Dim folderPath As String
Dim timeStamp As String
Dim typeString As String
Dim outputFile As String
Dim errorMsg As String
Dim successMsg As String
Dim workbookName As String
Dim conversionCheck As Boolean
Dim progIndicator As Integer
Dim errorIndi As Integer

    progIndicator = 0
    PBL_conversionOk = 0
    PBL_conversionFail = 0

    If PBL_fileToOpen <> False Then
        If (Not PBL_inputWsId) = True Then
            errorMsg = "Nie sú zvolené pracovné hárky pre konverziu!"
            MsgBox errorMsg, vbExclamation, "Informatívna chyba"
        Else
            Set PBL_xlNew = CreateObject("Excel.Application")

            PBL_xlNew.Workbooks.Add (1)
            Set PBL_outputWs = PBL_xlNew.Workbooks(1).Worksheets(1)
            
            PBL_xlNew.ScreenUpdating = False
            PBL_xlNew.Calculation = xlCalculationManual

            For i = 1 To UBound(PBL_inputWsId)

                conversionCheck = False

                Set PBL_inputWs = PBL_inputWb.Worksheets(PBL_inputWsId(i))
                workbookName = Left(PBL_inputWb.name, (InStrRev(PBL_inputWb.name, ".", -1, vbTextCompare) - 1))
                PBL_worksheetName = PBL_inputWs.name

                Select Case conversionType
                    Case PBL_SEC
                        If PBL_inputWs.Cells(1, 1).Value = "FREQ" And PBL_inputWs.Cells(6, 1).Value = "SEC" Then
                        conversionCheck = True
                        End If
                    Case PBL_REG
                        If PBL_inputWs.Cells(11, 1).Value = "REF_SECTOR" And PBL_inputWs.Cells(1, 6).Value = "REG" Then
                        conversionCheck = True
                        End If
                    Case PBL_PENS
                        If PBL_inputWs.Cells(12, 1).Value = "UNIT_MULT" And PBL_inputWs.Cells(1, 6).Value = "PENS" Then
                        conversionCheck = True
                        End If
                    Case PBL_MAIN
                        If PBL_inputWs.Cells(12, 1).Value = "TIME_PER_COLLECT" And PBL_inputWs.Cells(1, 6).Value = "MAIN" Then
                        conversionCheck = True
                        End If
                End Select

                If conversionCheck = True Then

                    ' Spúštanie procedúr
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

                Else
                    Call errorHandler(SUB_NAME, PBL_worksheetName)
                End If

            Next i
                
            ' Uloženie výstupu
            If PBL_conversionOk > 0 Then
                folderPath = PBL_xlApp.ActiveWorkbook.Path
                timeStamp = Format(CStr(Now), "yyyy_mm_dd_hhmmss")
                saveName = folderPath & "\" & workbookName & "_" & timeStamp
    
                PBL_xlNew.Workbooks(1).SaveAs fileName:=saveName, FileFormat:=xlCSV, local:=True
            End If
            
            PBL_xlNew.Calculation = xlCalculationAutomatic
            PBL_xlNew.ScreenUpdating = True
            PBL_xlNew.Workbooks(1).Close False
            PBL_xlNew.Quit
            
            Call UnloadForms
             
            PBL_xlApp.Quit
            
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
        errorMsg = "Nie je zvolený súbor pre konverziu!"
        MsgBox errorMsg, vbExclamation, "Informatívna chyba"
    End If

End Sub


'---------------------------
'Procedúra vypnutia programu
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


