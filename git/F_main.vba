VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} F_main 
   Caption         =   "SDMX Konvertor"
   ClientHeight    =   8685
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10590
   OleObjectBlob   =   "F_main.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "F_main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'-------------------------
'Hlavny userform aplikacie
'-------------------------
Private Sub UserForm_Initialize()
        
'Minimalizuje instanciu vo fullscreen mode
    Application.WindowState = xlMaximized
    Application.WindowState = xlMinimized
    
'Nastavenie farieb elementov userformu
    Me.BackColor = RGB(204, 236, 255)
    chbLeft.BackColor = RGB(204, 236, 255)
    chbRight.BackColor = RGB(204, 236, 255)
    labIdLeft.BackColor = RGB(204, 236, 255)
    labIdRight.BackColor = RGB(204, 236, 255)
    labNameLeft.BackColor = RGB(204, 236, 255)
    labNameRight.BackColor = RGB(204, 236, 255)
    labSourceFile.BackColor = RGB(204, 236, 255)
    labConversionType.BackColor = RGB(204, 236, 255)
    labVersion.BackColor = RGB(204, 236, 255)
    optMAIN.BackColor = RGB(204, 236, 255)
    optPENS.BackColor = RGB(204, 236, 255)
    optREG.BackColor = RGB(204, 236, 255)
    optSEC.BackColor = RGB(204, 236, 255)
    optT9XX.BackColor = RGB(204, 236, 255)
    optT200.BackColor = RGB(204, 236, 255)
    optSU.BackColor = RGB(204, 236, 255)
    
'Nastavenie velkosti a formatu pismen elementov
    labConversionType.Font.Size = 9
    labIdLeft.Font.Size = 10
    labIdRight.Font.Size = 10
    labNameLeft.Font.Size = 10
    labNameRight.Font.Size = 10
    labSourceFile.Font.Size = 9
    chbLeft.Font.Size = 8
    chbRight.Font.Size = 8
    cbAdd.Font.Size = 11
    cbRemove.Font.Size = 11
    cbRun.Font.Size = 12
    cbBrowseFile.Font.Size = 10
    
    labConversionType.Font.Bold = True
    labSourceFile.Font.Bold = True
    cbBrowseFile.Font.Bold = True
    cbRun.Font.Bold = True
    
'Dotiahne informaciu o aktualnej verzii do UF
    labVersion.Caption = PBL_programVersion
    
'Umozni viacnasobny vyber prvkov z listboxov
    lbLeft.MultiSelect = 2
    lbRight.MultiSelect = 2
        
End Sub


Private Sub UserForm_Terminate()

'Vypne aplikaciu
    Call AppClose

End Sub


Private Sub cbBrowseFile_Click()

'Vyber vstupneho suboru cez OpenFile dialogove okno
    Call OpenSourceFile
    
End Sub


'----------------------------------------
'Pridanie prvkov do zoznamu pre konverziu
'----------------------------------------
Private Sub cbAdd_Click()

Dim i As Integer
Dim thisListBox As Object
Dim counter As Integer

    counter = 0
    
    For i = 0 To lbLeft.ListCount - 1
        If lbLeft.Selected(i - counter) Then
            lbRight.AddItem lbLeft.List(i - counter, 0)
            lbRight.List(lbRight.ListCount - 1, 1) = lbLeft.List(i - counter, 1)
            lbLeft.RemoveItem (i - counter)
            counter = counter + 1
        End If
    Next i
    
    chbLeft.Value = False
    
'Zoradenie prvkov podla ID
    Set thisListBox = Me.lbRight
    Call lbSort(thisListBox)

End Sub


'-----------------------------------------
'Odobratie prvkov zo zoznamu pre konverziu
'-----------------------------------------
Private Sub cbRemove_Click()

Dim i As Integer
Dim thisListBox As Object
Dim counter As Integer

    counter = 0
    
    For i = 0 To lbRight.ListCount - 1
        If lbRight.Selected(i - counter) Then
            lbLeft.AddItem lbRight.List(i - counter, 0)
            lbLeft.List(lbLeft.ListCount - 1, 1) = lbRight.List(i - counter, 1)
            lbRight.RemoveItem (i - counter)
            counter = counter + 1
        End If
    Next i
    
    chbRight.Value = False
    
'Zoradenie prvkov podla ID
    Set thisListBox = Me.lbLeft
    Call lbSort(thisListBox)

End Sub


'------------------------------------------------------------------
'Sort prvkov v listboxe podla ich ID (pricom ID = poradie v zosite)
'------------------------------------------------------------------
Sub lbSort(lbArgument As Object)

Dim i As Long
Dim j As Long
Dim x As Long
Dim temp As String
    
    With lbArgument
        For j = LBound(.List) To UBound(.List) - 1 Step 1
            For i = LBound(.List) To UBound(.List) - 1 Step 1
                If CInt(.List(i)) > CInt(.List(i + 1)) Then
                    For x = 0 To (.ColumnCount - 1) Step 1
                        temp = .List(i, x)
                        .List(i, x) = .List(i + 1, x)
                        .List(i + 1, x) = temp
                    Next x
                End If
            Next i
        Next j
    End With

End Sub


'---------------------------------------------
'Oznacenie/odznacenie kazdeho prvku v listboxe
'---------------------------------------------
Private Sub chbLeft_Click()

Dim i As Integer

    If chbLeft.Value = True Then
        For i = 0 To lbLeft.ListCount - 1
            lbLeft.Selected(i) = True
        Next i
    End If
    
    If chbLeft.Value = False Then
        For i = 0 To lbLeft.ListCount - 1
            lbLeft.Selected(i) = False
        Next i
    End If

End Sub


'---------------------------------------------
'Oznacenie/odznacenie kazdeho prvku v listboxe
'---------------------------------------------
Private Sub chbRight_Click()

Dim i As Integer

    If chbRight.Value = True Then
        For i = 0 To lbRight.ListCount - 1
            lbRight.Selected(i) = True
        Next i
    End If
    
    If chbRight.Value = False Then
        For i = 0 To lbRight.ListCount - 1
            lbRight.Selected(i) = False
        Next i
    End If

End Sub


'-------------------
'Spustenie konverzie
'-------------------
Private Sub cbRun_Click()

Dim i As Integer
Dim j As Integer
Dim counter As Integer
Dim conversionType As Integer
Dim infoMsg As String

    conversionType = 0
    counter = 0

'Overenie typu konverzie (typ template)
    If optSEC = True Then conversionType = PBL_SEC
    If optT9XX = True Then conversionType = PBL_T9XX
    If optT200 = True Then conversionType = PBL_T200
    If optREG = True Then conversionType = PBL_REG
    If optPENS = True Then conversionType = PBL_PENS
    If optMAIN = True Then conversionType = PBL_MAIN
    If optSU = True Then conversionType = PBL_SU
    
    If conversionType > 0 Then
        For i = 0 To lbRight.ListCount - 1
            lbRight.Selected(i) = True
        Next i
        
        Erase PBL_inputWsId()
        counter = lbRight.ListCount
        
        j = 1
        
'Push ID harkov ktore sa maju konvertovat do pola
        For i = 1 To counter
            If lbRight.Selected(i - 1) Then
                ReDim Preserve PBL_inputWsId(1 To j)
                PBL_inputWsId(j) = lbRight.List(i - 1, 0)
                j = j + 1
            End If
        Next
        
'Hlavna riadiaca procedura konverzie
        Call MainSub(conversionType)
        
    Else
        infoMsg = "Nie je zvolený typ výstupnej tabu¾ky!"
        MsgBox infoMsg, vbExclamation, "Informatívna chyba"
    End If

End Sub
