Attribute VB_Name = "Folder2Check"
Sub Folder2Check()
  Dim count&, lastRow&
  Dim folderPath, columnRead, columnWrite, imgExtion As String

  folderPath = "E:\Dropbox\_ChannelOnBoarding\Images\"
  columnRead = "A"
  columnResults = "B"
  imgExtion = "\"
  EXIST = "JPG Found"
  MIA = "JPG MIA"
  WLD = "*"

  
  lastRow = ThisWorkbook.Sheets(1).Range(columnRead & Rows.count).End(xlUp).Row
  For count = 2 To lastRow
    Range(columnRead & count).Activate

    If Dir(folderPath & Range(columnRead & ActiveCell.Row).Value & imgExtion) <> "" Then
        Range(columnResults & ActiveCell.Row).Value = EXIST
    Else
        Range(columnResults & ActiveCell.Row).Value = MIA
    End If
  Next count
End Sub


