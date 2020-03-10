This code will help you to remove text wrap from the entire worksheet with a single click.

It will first select all the columns and then remove text wrap and auto fit all the rows and columns.

Sub RemoveWrapText()
Cells.Select 
Selection.WrapText = False
Cells.EntireRow.AutoFit
Cells.EntireColumn.AutoFit
End Sub