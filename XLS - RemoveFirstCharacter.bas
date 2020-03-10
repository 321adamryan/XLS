Attribute VB_Name = "RemoveFirstCharacter"
Sub RemoveFirstCharacter()

columnRead = "D"

For Each cell In Range(columnRead + "1", Range(columnRead + "65536").End(xlUp))
If Not IsEmpty(cell) Then
cell.Value = Right(cell, Len(cell) - 1)
End If
Next cell
End Sub
