Attribute VB_Name = "M_publicVar"
Option Explicit

'Deklaracia globalnych premennych projektu
Public Const PBL_SEC = 1
Public Const PBL_REG = 2
Public Const PBL_PENS = 3
Public Const PBL_MAIN = 4
Public Const PBL_SU = 5
Public Const PBL_NTL = 6

Public Const PBL_OK = "OK"
Public Const PBL_FAIL = "FAIL"

Public PBL_programVersion As String
Public PBL_programName As String

Public PBL_xlApp As Object
Public PBL_xlNew As Object
Public PBL_xlOld As Object
Public PBL_worksheetName As String
Public PBL_inputWsId() As Integer

Public PBL_conversionOk As Integer
Public PBL_conversionFail As Integer

Public PBL_rowStep As Integer
Public PBL_colStep As Integer

Public PBL_inputWs As Worksheet
Public PBL_outputWs As Worksheet
Public PBL_inputWb As Workbook
       
Public PBL_parameterFix() As Variant
Public PBL_fileToOpen As Variant

Public PBL_copyStart As Integer
