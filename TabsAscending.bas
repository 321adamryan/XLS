
Sub TabsAscending()
 
For i = 1 To Application.Sheets.Count
    For j = 1 To Application.Sheets.Count - 1
        If UCase$(Application.Sheets(j).Name) > UCase$(Application.Sheets(j + 1).Name) Then
            Sheets(j).Move after:=Sheets(j + 1)
        End If
    Next
Next
MsgBox "The tabs have been sorted from A to Z."
 
End Sub