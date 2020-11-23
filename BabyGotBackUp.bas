Private Sub BabyGotBackUp()

Application.DisplayAlerts = False

    Dim BabyGotBackUpCSVFileName As String
    
    Dim tempWB As Workbook

    On Error GoTo err
       
'--------Backup-------------------

'yyyyMMdd
'MM_dd_yyyy

    BabyGotBackUpCSVFileName = ThisWorkbook.Path & "_BabyGotBackUp_" & VBA.Format(VBA.Now, "yyyyMMdd") & ".xlsx"
    ThisWorkbook.Sheets("Backup").Activate
    ActiveSheet.Copy
    Set tempWB = ActiveWorkbook
    With ActiveSheet.UsedRange
        .Cells.Copy
        .Cells.PasteSpecial xlPasteValues
        .Cells(1).Select
    End With
    Application.CutCopyMode = False
    With tempWB
    .SaveAs Filename:=BabyGotBackUpCSVFileName, FileFormat:=51, CreateBackup:=False
    .Close
    End With
       

    
'---------------Print------------
ThisWorkbook.Sheets("Backup").PrintOut

    
'--------CLOSE-------------------
    ThisWorkbook.Sheets("Backup").Activate

    MsgBox "Export is done!"
    
err:
    Application.DisplayAlerts = True
    
    

End Sub

