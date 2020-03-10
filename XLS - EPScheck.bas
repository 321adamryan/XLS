 ' AdamRyan
 'Filename=EPSCheck.bas
 'Description=Finds the existence of EPS files in a column and exports the results
 'Filetype=bas
 'License=OpenSource
Attribute VB_Name = "EPScheck"
Sub EPSCheck()
  Dim count&, lastRow&
  Dim folderPath, columnRead, columnWrite, imgExtion As String

  folderPath = "C:\Users\graphics2\Desktop\BN FTP\"
  columnRead = "J"
  columnResults = "Q"
  imgExtion = ".eps"
  

  lastRow = ThisWorkbook.Sheets(1).Range(columnRead & Rows.count).End(xlUp).row
  For count = 2 To lastRow
    Range(columnRead & count).Activate

    If Dir(folderPath & Range(columnRead & ActiveCell.row).Value & imgExtion) <> "" Then
        Range(columnResults & ActiveCell.row).Value = "File exists."
    Else
        Range(columnResults & ActiveCell.row).Value = "File doesn't exist."
    End If

  Next count
End Sub

