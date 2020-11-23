Attribute VB_Name = "IsNumericRange"
Sub IsNumericRange()
Dim cell As Range
Dim bIsNumeric As Boolean

bIsNumeric = True
For Each cell In Range("b2:b670")
    If IsNumeric(cell) = False Then
        'Non-numeric value found. Exit loop
        bIsNumeric = False
        Exit For
    End If
Next cell

If bIsNumeric = True Then
    'All values in your range are numeric
    '**PLACE CODE HERE**
    MsgBox "Congrats! Values show REAL math!"
Else
    'There are non-numeric values in your range
    '**PLACE CODE HERE**
    MsgBox "Your FUZZY math has numerical errors, 2+Text=Error"
End If
End Sub

