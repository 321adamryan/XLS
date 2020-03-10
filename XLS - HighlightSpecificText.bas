Suppose you have a large dataset and you want to check for a particular value. For this, you can use this code. When you run it, you will get an input box to enter the value to search for.

Sub highlightValue()
Dim myStr As String
Dim myRg As Range
Dim myTxt As String
Dim myCell As Range
Dim myChar As String
Dim I As Long
Dim J As Long
On Error Resume Next
If ActiveWindow.RangeSelection.Count> 1 Then
myTxt= ActiveWindow.RangeSelection.AddressLocal
Else
myTxt= ActiveSheet.UsedRange.AddressLocal
End If
LInput: Set myRg= Application.InputBox("please select the data range:", "Selection Required", myTxt, , , , , 8)
If myRg Is Nothing Then
Exit Sub
If myRg.Areas.Count > 1 Then
MsgBox"not support multiple columns" GoToLInput
End If
If myRg.Columns.Count <> 2 Then
MsgBox"the selected range can only contain two columns "
GoTo LInput
End If
For I = 0 To myRg.Rows.Count-1
myStr= myRg.Range("B1").Offset(I, 0).Value
With myRg.Range("A1").Offset(I, 0)
.Font.ColorIndex= 1
For J = 1 To Len(.Text)
Mid(.Text, J, Len(myStr)) = myStrThen
.Characters(J, Len(myStr)).Font.ColorIndex= 3
Next
End With
Next I
End Sub