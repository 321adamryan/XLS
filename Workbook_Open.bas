Private Sub Workbook_Open()


'save file in ThisWorkbook


    Workbooks.Add.SaveAs Filename:="E:\Dropbox\WAREHOUSE\Master Inventory File_Open.xlsx"
    Workbooks("Master Inventory File_Open.xlsx").Close SaveChanges:=False
    
End Sub


Private Sub Workbook_BeforeClose(Cancel As Boolean)

Kill "E:\Dropbox\WAREHOUSE\Master Inventory File_Open.xlsx"
End Sub
