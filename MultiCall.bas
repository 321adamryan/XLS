Attribute VB_Name = "MultiCall"
Sub MultiCall()
    
    Call APFinder
    Call DPFinder
    Call FOFinder
    Call MILFinder
    Call OHPFinder
 
    

    
    
     ThisWorkbook.Sheets("Master Inventory List").Activate

    MsgBox "Report is Done!"
    
err:
    Application.DisplayAlerts = True


End Sub
