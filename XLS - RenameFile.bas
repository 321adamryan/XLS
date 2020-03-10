'AdamRyan
'Filename=RenameFile.bas
'Description=Pulls name from excel column 1 and replaces it with the corresponding cell in column 2
'Filetype=bas
'License=OpenSource

Attribute VB_Name = "RenameFile"
Sub RenameFile()
    Dim z As String
    Dim s As String
    Dim V As Integer
    Dim TotalRow As Integer
    
    TotalRow = ActiveSheet.UsedRange.Rows.Count
    
    For V = 1 To TotalRow
        
        ' Get value of each row in columns 1 start at row 2
        z = Cells(V + 1, 1).Value
        ' Get value of each row in columns 2 start at row 2
        s = Cells(V + 1, 2).Value
        
        Dim sOldPathName As String
        sOldPathName = z
        On Error Resume Next
        Name sOldPathName As s & ".jpg"
        
    Next V
    
    MsgBox "Congratulations! You have successfully renamed all the files"
    
End Sub

