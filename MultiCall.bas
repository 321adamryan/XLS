Attribute VB_Name = "MultiCall"
Sub MultiCall()
    
    Call CATCheck
    Call JPGCheck
    Call PNGCheck
    Call PSDCheck
    Call R23PCheck
    Call RAWCheck
    Call ZIPCheck

    

    
    
     ThisWorkbook.Sheets("ImageReport").Activate

    MsgBox "iMAGE Report is Done!"
    
err:
    Application.DisplayAlerts = True


End Sub
