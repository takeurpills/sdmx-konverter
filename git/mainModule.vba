Attribute VB_Name = "mainModule"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                                                                                       '
'  Názov:  Konvertor ESA2010                                                            '
'  Autor:  Martin Tóth - Štatistický úrad SR                                            '
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

Public rowStep As Integer
Public colStep As Integer

Public inputWs As Worksheet
Public outputWs As Worksheet
Public inputWb As Workbook
       
Public parameterFix() As Variant
Public fileToOpen As Variant

'----------------------------------------------
' Inicializaèná procedúra pri spustení programu
'----------------------------------------------
Sub programInit()
    
    versionId = "v0.3"
    instanceName = ActiveWorkbook.FullName
    
    mainForm.Show vbModeless
    
End Sub
    
'------------------------------------
' Hlavná riadiaca procedúra konverzie
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
    


    If fileToOpen <> False Then
        If (Not inputWsId) = True Then
            errorMsg = "Nie sú zvolené pracovné hárky pre konverziu!"
            MsgBox errorMsg, vbExclamation, "Informatívna chyba"
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
                Case NA_PENS
                    If inputWs.Cells(12, 1).Value = "UNIT_MULT" And inputWs.Cells(1, 6).Value = "PENS" Then
                    conversionCheck = True
                    End If
                Case NA_MAIN
                    If inputWs.Cells(12, 1).Value = "TIME_PER_COLLECT" And inputWs.Cells(1, 6).Value = "MAIN" Then
                    conversionCheck = True
                    End If
            End Select
            
            If conversionCheck = True Then
            
                ' Spúštanie procedúr
                If progIndicator = 0 Then
                    mainForm.Hide
                    progressForm.Show vbModeless
                    progIndicator = 1
                End If
                                
                Call arrayPush(conversionType)
                Call defineConversion(conversionType)
                Call arrayFill(conversionType)
                
                ' Uloženie výstupu
                folderPath = xlApp.ActiveWorkbook.Path
                timeStamp = Format(CStr(Now), "yyyy_mm_dd_hhmmss")
                saveName = folderPath & "\" & nameString & "_" & timeStamp
    
                xlNew.Workbooks(1).SaveAs Filename:=saveName, FileFormat:=xlCSV, local:=True
                xlNew.ScreenUpdating = True
                xlNew.Workbooks(1).Close False
                xlNew.Quit
                
                conversionOk = conversionOk + 1
                
            Else
                errorMsg = "Zvolený hárok - """ & nameString & """ nemá správny formát alebo bol zvolený nesprávny typ konverzie!" & vbNewLine & vbNewLine
                errorMsg = errorMsg & "Konverzia hárku sa nevykoná!"
                MsgBox errorMsg, vbCritical, "Kritická chyba"
                xlNew.Quit
                
                conversionFail = conversionFail + 1
            End If
            Next
            
            Call unloadForms
            
            xlApp.Quit
            
            successMsg = "Konverzia hárkov bola dokonèená." & vbNewLine & vbNewLine
            successMsg = successMsg & "• Poèet úspešných konverzií: " & conversionOk & vbNewLine
            successMsg = successMsg & "• Poèet neúspešných konverzií: " & conversionFail
            
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
Sub appClose()

    Set xlOld = GetObject(instanceName).Application
    
    xlOld.Visible = True

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
        MsgBox "Nastala chyba. Prosím kontaktujte správcu aplikácie.", , "Chyba"
    End If

End Sub


