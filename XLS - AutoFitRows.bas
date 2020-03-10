You can use this code to auto-fit all the rows in a worksheet. When you run this code it will select all the cells in your worksheet and instantly auto-fit all the row.

Sub AutoFitRows()
' :::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
'
'       filename: AECrawlExpression.bas
'          coder: AdamRyan
'        program: Google Sheets
'    description: This script 
'      extention: BAS
'       licensce: OpenSource
'        website: adamryan.info
'
' :::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
Cells.Select
Cells.EntireRow.AutoFit
End Sub