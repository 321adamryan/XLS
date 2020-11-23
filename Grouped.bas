Attribute VB_Name = "Grouped"
Sub Refresh()
Attribute Refresh.VB_ProcData.VB_Invoke_Func = " \n14"

' Refresh Macro
'
    ActiveWorkbook.RefreshAll
'
End Sub
Sub SaveNClose()
Attribute SaveNClose.VB_ProcData.VB_Invoke_Func = " \n14"
'
' SaveNClose Macro
'

'
    ActiveWorkbook.Save
    ActiveWorkbook.Close
End Sub
