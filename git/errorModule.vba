Attribute VB_Name = "errorModule"
Sub errorHandler(errProcedure As String)

    errorMsg = "Nastala chyba (#1001). Prosím kontaktujte správcu aplikácie." & errProcedure
    MsgBox errorMsg, vbCritical, "Kritická chyba"
    'appClose
    
End Sub
