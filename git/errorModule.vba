Attribute VB_Name = "errorModule"
Sub errorHandler(errProcedure As String)

    errorMsg = "Nastala chyba (#1001). Pros�m kontaktujte spr�vcu aplik�cie." & errProcedure
    MsgBox errorMsg, vbCritical, "Kritick� chyba"
    'appClose
    
End Sub
