VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} F_progress 
   Caption         =   "Inform�cia"
   ClientHeight    =   1575
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5505
   OleObjectBlob   =   "F_progress.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "F_progress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                                                                                       '
'  N�zov:  SDMX Konvertor ESA2010                                                       '
'  Autor:  Martin T�th - �tatistick� �rad SR                                            '
'                                                                                       '
'  Popis:  Aplik�cia sl��i na konverziu pracovn�ch v�stupn�ch tabuliek n�rodn�ch ��tov  '
'          vo form�te excel do tabuliek v zmysle SDMX �tandardu (pod�a definovan�ch     '
'          dom�nov�ch datab�zov�ch �trukt�r) v csv form�te                              '
'                                                                                       '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Explicit

Private Sub UserForm_Initialize()

    Me.BackColor = RGB(204, 236, 255)
    infoImage.BackColor = RGB(204, 236, 255)
    infoLab.BackColor = RGB(204, 236, 255)

    infoLab.Font.Size = 9
    infoLab.Font.Bold = True
    
End Sub
