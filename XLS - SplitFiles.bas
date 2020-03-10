'AdamRyan
'Filename=SplitFiles
'Description=Split excel file into multiple files based upon cell value
'Filetype=bas
'License=OpenSource

Public Sub SplitToFiles()


' Loops through a specified column, and split each distinct values into a separate file by making a copy and deleting rows below and above
'
' Note: Values in the column should be unique or sorted.
'
' The following cells are ignored when delimiting sections:
' - blank cells, or containing spaces only
' - same value repeated
' - cells containing "total"
'
' Files are saved in a "Split" subfolder from the location of the source workbook, and named after the section name.

Dim osh As Worksheet ' Original sheet
Dim iRow As Long ' Cursors
Dim iCol As Long
Dim iFirstRow As Long ' Constant
Dim iTotalRows As Long ' Constant
Dim iStartRow As Long ' Section delimiters
Dim iStopRow As Long
Dim sSectionName As String ' Section name (and filename)
Dim rCell As Range ' current cell
Dim owb As Workbook ' Original workbook
Dim sFilePath As String ' Constant
Dim iCount As Integer ' # of documents created

'enter the column# to process
iCol = 2
'change this value depending if their is a header row
iRow = 2
iFirstRow = iRow

Set osh = Application.ActiveSheet
Set owb = Application.ActiveWorkbook
iTotalRows = osh.UsedRange.Rows.count
sFilePath = Application.ActiveWorkbook.Path

If Dir(sFilePath + "\Split", vbDirectory) = "" Then
    MkDir sFilePath + "\Split"
End If

'Turn Off Screen Updating  Events
Application.EnableEvents = False
Application.ScreenUpdating = False

Do
    ' Get cell at cursor
    Set rCell = osh.Cells(iRow, iCol)
    sCell = Replace(rCell.Text, " ", "")

    If sCell = "" Or (rCell.Text = sSectionName And iStartRow <> 0) Or InStr(1, rCell.Text, "total", vbTextCompare) <> 0 Then
        ' Skip condition met
    Else
        ' Found new section
        If iStartRow = 0 Then
            ' StartRow delimiter not set, meaning beginning a new section
            sSectionName = rCell.Text
            iStartRow = iRow
        Else
            ' StartRow delimiter set, meaning we reached the end of a section
            iStopRow = iRow - 1

            ' Pass variables to a separate sub to create and save the new worksheet
            CopySheet osh, iFirstRow, iStartRow, iStopRow, iTotalRows, sFilePath, sSectionName, owb.fileFormat
            iCount = iCount + 1

            ' Reset section delimiters
            iStartRow = 0
            iStopRow = 0

            ' Ready to continue loop
            iRow = iRow - 1
        End If
    End If

    ' Continue until last row is reached
    If iRow < iTotalRows Then
            iRow = iRow + 1
    Else
        ' Finished. Save the last section
        iStopRow = iRow
        CopySheet osh, iFirstRow, iStartRow, iStopRow, iTotalRows, sFilePath, sSectionName, owb.fileFormat
        iCount = iCount + 1

        ' Exit
        Exit Do
    End If
Loop

'Turn On Screen Updating  Events
Application.ScreenUpdating = True
Application.EnableEvents = True

MsgBox Str(iCount) + " documents saved in " + sFilePath


End Sub

Public Sub DeleteRows(targetSheet As Worksheet, RowFrom As Long, RowTo As Long)

Dim rngRange As Range
Set rngRange = Range(targetSheet.Cells(RowFrom, 1), targetSheet.Cells(RowTo, 1)).EntireRow
rngRange.Select
rngRange.Delete

End Sub


Public Sub CopySheet(osh As Worksheet, iFirstRow As Long, iStartRow As Long, iStopRow As Long, iTotalRows As Long, sFilePath As String, sSectionName As String, fileFormat As XlFileFormat)
     Dim ash As Worksheet ' Copied sheet
     Dim awb As Workbook ' New workbook

     ' Copy book
     osh.Copy
     Set ash = Application.ActiveSheet

     ' Delete Rows after section
     If iTotalRows > iStopRow Then
         DeleteRows ash, iStopRow + 1, iTotalRows
     End If

     ' Delete Rows before section
     If iStartRow > iFirstRow Then
         DeleteRows ash, iFirstRow, iStartRow - 1
     End If

     ' Select left-topmost cell
     ash.Cells(1, 1).Select

     ' Clean up a few characters to prevent invalid filename
     sSectionName = Replace(sSectionName, "/", " ")
     sSectionName = Replace(sSectionName, "\", " ")
     sSectionName = Replace(sSectionName, ":", " ")
     sSectionName = Replace(sSectionName, "=", " ")
     sSectionName = Replace(sSectionName, "*", " ")
     sSectionName = Replace(sSectionName, ".", " ")
     sSectionName = Replace(sSectionName, "?", " ")

     ' Save in same format as original workbook
     ash.SaveAs sFilePath + "\Split\" + "9011757_" + sSectionName, fileFormat

     ' Close
     Set awb = ash.Parent
     awb.Close SaveChanges:=False
End Sub

