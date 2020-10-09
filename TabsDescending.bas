
Sub TabsDescending()
 
For i = 1 To Application.Sheets.Count
    For j = 1 To Application.Sheets.Count - 1
        If UCase$(Application.Sheets(j).Name) < UCase$(Application.Sheets(j + 1).Name) Then
            Application.Sheets(j).Move after:=Application.Sheets(j + 1)
        End If
    Next
Next
 
MsgBox "The tabs have been sorted from Z to A."
End Sub
