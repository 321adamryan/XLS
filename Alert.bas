Attribute VB_Name = "Alert"
Sub Alert()
Attribute Alert.VB_ProcData.VB_Invoke_Func = " \n14"


For x = 2 To 670

Set val3 = Worksheets("FinalAllocation").Cells(x, 2)
If val3.Value < 0 Then

MsgBox "Don't be so Negative!! Clean that up!"


End If
Next x


MsgBox "Positivity check complete!"
End Sub
