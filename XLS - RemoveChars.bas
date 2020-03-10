Sub RemoveFirstThreeCharactersInEachCell()
For Each cell In Range("D1", Range("D65536").End(xlUp))
If Not IsEmpty(cell) Then
cell.Value = Right(cell, Len(cell) - 1)
End If
Next cell
End Sub