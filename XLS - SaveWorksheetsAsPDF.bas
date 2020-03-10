Sub SaveWorkshetAsPDF()
Dimws As Worksheet
For Each ws In Worksheetsws.ExportAsFixedFormat xlTypePDF, “ENTER-FOLDER-NAME-HERE" & ws.Name & ".pdf" Nextws
End Sub