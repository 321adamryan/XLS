Attribute VB_Name = "PLTCheck"
Sub PLTCheck()
  Dim count&, lastRow&
  Dim folderPath, columnRead, columnWrite, imgExtion As String

  folderPath = "S:\00 Product Versions\HiRes\Ready4Droplet\"
  columnRead = "S"
  columnResults = "V"
  imgExtion = ".plt"
  

  lastRow = ThisWorkbook.Sheets(2).Range(columnRead & Rows.count).End(xlUp).row
  For count = 2 To lastRow
    Range(columnRead & count).Activate

    If Dir(folderPath & Range(columnRead & ActiveCell.row).Value & imgExtion) <> "" Then
        Range(columnResults & ActiveCell.row).Value = "PLT exists."
    Else
        Range(columnResults & ActiveCell.row).Value = "PLT doesn't exist."
    End If

  Next count
End Sub


