'AdamRyan
'Filename=PSDCheck.bas
'Description= Finds PSD files located in column A and prints their existence in column E
'Filetype=bas
'License=OpenSource

Attribute VB_Name = "PSDCheck"
Sub PSDCheck()
  Dim count&, lastRow&
  Dim folderPath, columnRead, columnWrite, imgExtion As String

  folderPath = "S:\00 Product Versions\HiRes\Ready4Droplet\"
  columnRead = "A"
  columnResults = "E"
  imgExtion = ".psd"
  

  lastRow = ThisWorkbook.Sheets(2).Range(columnRead & Rows.count).End(xlUp).row
  For count = 2 To lastRow
    Range(columnRead & count).Activate

    If Dir(folderPath & Range(columnRead & ActiveCell.row).Value & imgExtion) <> "" Then
        Range(columnResults & ActiveCell.row).Value = "PSD exists."
    Else
        Range(columnResults & ActiveCell.row).Value = "PSD doesn't exist."
    End If

  Next count
End Sub

