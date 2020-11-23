Attribute VB_Name = "APFinder"
Function WorksheetExists2(WorksheetName As String, Optional wb As Workbook) As Boolean
    If wb Is Nothing Then Set wb = ThisWorkbook
    With wb
        On Error Resume Next
        WorksheetExists2 = (.Sheets(WorksheetName).Name = WorksheetName)
        On Error GoTo 0
    End With
End Function

Sub WorksheetFinder()

Dim ahickey As String

ahickey = "Allocated Product"


If WorksheetExists2(ahickey) Then
    MsgBox "Excelcior! Your sheet is fantastic!" & vbCrLf & vbCrLf & ahickey & " is in this workbook"
Else
    MsgBox "You just picked a fresh bouquet of Oopsie Daisies!" & vbCrLf & vbCrLf & ahickey & " does not exist!"
End If
End Sub
