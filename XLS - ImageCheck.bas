Attribute VB_Name = "ImageCheck"
Sub ImageCheck()
  Dim count&, lastRow&
  Dim folderPath, columnRead, columnWrite, imgExtion As String

  folderPath = "S:\00 Product Versions\Staged\"
  columnRead = "c"
  columnResults = "g"
  imgExtion = ".jpg"
  

  lastRow = ThisWorkbook.Sheets(1).Range(columnRead & Rows.count).End(xlUp).Row
  For count = 2 To lastRow
    Range(columnRead & count).Activate

    If Dir(folderPath & Range(columnRead & ActiveCell.Row).Value & imgExtion) <> "" Then
        Range(columnResults & ActiveCell.Row).Value = "JPG exists."
    Else
        Range(columnResults & ActiveCell.Row).Value = "JPG doesn't exist."
    End If

  Next count
End Sub

