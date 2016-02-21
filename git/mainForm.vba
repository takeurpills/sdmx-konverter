VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} mainForm 
   Caption         =   "Konvertor"
   ClientHeight    =   8685
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10590
   OleObjectBlob   =   "mainForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "mainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim i As Integer

Private Sub UserForm_Initialize()
    
    Application.Visible = True
    
    lbLeft.MultiSelect = 2
    lbRight.MultiSelect = 2
    
    Me.BackColor = RGB(204, 236, 255)
    chbLeft.BackColor = RGB(204, 236, 255)
    chbRight.BackColor = RGB(204, 236, 255)
    labIdLeft.BackColor = RGB(204, 236, 255)
    labIdRight.BackColor = RGB(204, 236, 255)
    labNameLeft.BackColor = RGB(204, 236, 255)
    labNameRight.BackColor = RGB(204, 236, 255)
    labSourceFile.BackColor = RGB(204, 236, 255)
    labConversionAlg.BackColor = RGB(204, 236, 255)
    labVersion.BackColor = RGB(204, 236, 255)
    optMAIN.BackColor = RGB(204, 236, 255)
    optPENS.BackColor = RGB(204, 236, 255)
    optREG.BackColor = RGB(204, 236, 255)
    optSEC.BackColor = RGB(204, 236, 255)
    optSU.BackColor = RGB(204, 236, 255)
    
    labConversionAlg.Font.Size = 9
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
    
    labConversionAlg.Font.Bold = True
    labSourceFile.Font.Bold = True
    cbBrowseFile.Font.Bold = True
    cbRun.Font.Bold = True
    
    labVersion.Caption = versionId
        
End Sub

Private Sub UserForm_Terminate()

    appClose

End Sub

Private Sub cbBrowseFile_Click()

    Call openSourceFile
    
End Sub

Private Sub cbAdd_Click()

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
    
    Set thisListBox = Me.lbRight
    Call lbSort(thisListBox)

End Sub

Private Sub cbRemove_Click()

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
    
    Set thisListBox = Me.lbLeft
    Call lbSort(thisListBox)

End Sub

Private Sub chbLeft_Click()

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

Private Sub chbRight_Click()

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

Private Sub cbRun_Click()

Dim j As Integer
Dim myCount As Integer
Dim algType As Integer

algType = 0
myCount = 0

If optSEC = True Then algType = NA_SEC
If optREG = True Then algType = NA_REG
If optPENS = True Then algType = NA_PENS
If optMAIN = True Then algType = NA_MAIN
If optSU = True Then algType = NA_SU

If algType > 0 Then

    For i = 0 To lbRight.ListCount - 1
        lbRight.Selected(i) = True
    Next i
    
    Erase inputSheetsId()
    myCount = lbRight.ListCount
    
    j = 1
    
    For i = 1 To myCount
        If lbRight.Selected(i - 1) Then
            ReDim Preserve inputSheetsId(1 To j)
            inputSheetsId(j) = lbRight.List(i - 1, 0)
            j = j + 1
        End If
    Next
    
    Call mainSub(algType)

Else: MsgBox "Nie je zvolený typ výstupnej tabu¾ky!", , "Chyba"
End If

End Sub

Sub lbSort(lbArgument As Object)

Dim i As Long, j As Long, x As Long, temp As String
    
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

