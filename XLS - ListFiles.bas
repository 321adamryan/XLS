'AdamRyan
'Filename=ListFiles.bas
'Description=List files located within a folder
'Filetype=bas
'License=OpenSource

Attribute VB_Name = "ListFiles"
Option Explicit

'' ***************************************************************************

' Two versions of this macro are shown here.
' The first version is the modified version which will parse the directories into separate columns
' The second version is the original version.
Sub ListFiles()
    
    Dim vvRes                   ''Variant to collect result
    Dim viLoopCounter%          ''For loop counter
    Dim CRCol As Integer
    Dim i As Integer
    Dim nWS As Worksheet
    Application.ScreenUpdating = False
    On Error GoTo Endit ' this is messy. The macro will add a sheet anyway and then delete it without asking. It works though.
    Set nWS = Worksheets.Add
    nWS.Cells.Activate
    ''Set an error trap - gets around a "Cancel" situation
    
    
    ''Clear the target range, column "A"
    Cells.Columns(1).ClearContents
    
    ''Go to the top left cell
    Cells(1, 1).Select
    
    ''Show the file open box and get a result
    vvRes = Application.GetOpenFilename("The lot, *.*", MultiSelect:=True)
    
    ''Loop for each result in the "fileToOpen" result ..
    For viLoopCounter = LBound(vvRes) To UBound(vvRes)
        
        '' ..  and input to a cell
        Cells(viLoopCounter, 1) = vvRes(viLoopCounter)
        
    Next
    
    
    Columns("A:A").Select
    Selection.TextToColumns Destination:=Range("A1"), DataType:=xlDelimited, _
    TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=False, _
    Semicolon:=False, Comma:=False, Space:=False, other:=True, OtherChar _
    :="\", FieldInfo:=Array(Array(1, 1), Array(2, 1), Array(3, 1), Array(4, 1), Array(5, _
    1)), TrailingMinusNumbers:=True
    Cells.Select
    Selection.Columns.AutoFit
    Do While Range("A1").Value = ""
        Columns(1).EntireColumn.Delete
    Loop
    Range("A1").Select
    CRCol = Selection.CurrentRegion.Columns.Count
    For i = 1 To CRCol - 2
        
        Columns(1).EntireColumn.Delete
        
    Next i
    Range("A1").EntireRow.Insert
    Range("A1").Select
    With Selection
        .Value = "Folder"
        .Font.FontStyle = "Bold"
    End With
    Range("B1").Select
    With Selection
        .Value = "Filename"
        .Font.FontStyle = "Bold"
    End With
    On Error Resume Next
    nWS.Name = Range("A2").Value
    GoTo Finish ' the macro has run with no errors
Endit:     ' perhaps someone cancelled half-way, but a sheet has been added anyway. The section deletes that sheet.
    Application.DisplayAlerts = False
    ActiveSheet.Delete
    Application.DisplayAlerts = True
Finish:
End Sub

