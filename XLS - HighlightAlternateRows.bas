By highlighting alternate rows you can make your data easily readable. And for this, you can use below VBA code. It will simply highlight every alternate row in selected range.


Sub highlightAlternateRows()
Dim rng As Range
For Each rng In Selection.Rows
If rng.RowMod 2 = 1 Then
rng.Style= "20% -Accent1"
rng.Value= rng^ (1 / 3)
Else
End If
Next rng
End Sub