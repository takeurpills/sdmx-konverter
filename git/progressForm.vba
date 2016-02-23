VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} progressForm 
   Caption         =   "Informácia"
   ClientHeight    =   1575
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5505
   OleObjectBlob   =   "progressForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "progressForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                                                                                       '
'  Názov:  Konvertor ESA2010                                                            '
'  Autor:  Martin Tóth - Štatistický úrad SR                                            '
'                                                                                       '
'  Popis:                                                                               '
'                                                                                       '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Explicit

Private Sub UserForm_Initialize()

    progressForm.BackColor = RGB(204, 236, 255)
    infoImage.BackColor = RGB(204, 236, 255)
    infoLab.BackColor = RGB(204, 236, 255)

    infoLab.Font.Size = 9
    infoLab.Font.Bold = True
    
End Sub
