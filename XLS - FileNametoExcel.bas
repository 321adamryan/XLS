 ' AdamRyan
 'Filename=FileNametoExcel.bas
 'Description=Pulls filenames from a folder and inserts them into an excel colum and preps the file for RenameFile.bas
 'Filetype=bas
 'License=OpenSource
Attribute VB_Name = "FileNametoExcel"
Sub FileNametoExcel()
    
    Dim fnam As Variant
    ' fnam is an array of files returned from GetOpenFileName
    ' note that fnam is of type boolean if no array is returned.
    ' That is, if the user clicks on cancel in the file open dialog box, fnam is set to FALSE
    
    Dim b As Integer 'counter for filname array
    Dim b1 As Integer 'counter for finding \ in filename
    Dim c As Integer 'extention marker
    
    ' format header
    Range("A1").Select
    ActiveCell.FormulaR1C1 = "Path and Filenames that had been selected to Rename"
    Range("A1").Select
    With Selection.Font
        .Name = "Arial"
        .FontStyle = "Bold"
        .Size = 10
    End With
    Columns("A:A").EntireColumn.AutoFit
    Range("B1").Select
    ActiveCell.FormulaR1C1 = "Input New Filenames Below"
    Range("B1").Select
    With Selection.Font
        .Name = "Arial"
        .FontStyle = "Bold"
        .Size = 10
    End With
    Columns("B:B").EntireColumn.AutoFit
    
    ' first open a blank sheet and go to top left  ActiveWorkbook.Worksheets.Add
    
    fnam = Application.GetOpenFilename("all files (*.*), *.*", 1, _
    "Select Files to Fill Range", "Get Data", True)
    
    If TypeName(fnam) = "Boolean" And Not (IsArray(fnam)) Then Exit Sub
    
    'if user hits cancel, then end
    
    For b = 1 To UBound(fnam)
        ' print out the filename (with path) into first column of new sheet
        ActiveSheet.Cells(b + 1, 1) = fnam(b)
    Next
    
    
End Sub

