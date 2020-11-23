Private Sub Workbook_Open()

Dim YellowBrickRoad As String
Dim ahickey As String


'save file in ThisWorkbook
ahickey = "Master Inventory File_IsOpen.xlsx"
YellowBrickRoad = ThisWorkbook.Path & "\" & ahickey


    Workbooks.Add.SaveAs Filename:=YellowBrickRoad
    Workbooks("Master Inventory File_IsOpen.xlsx").Close SaveChanges:=False
    
End Sub


Private Sub Workbook_BeforeClose(Cancel As Boolean)

Dim YellowBrickRoad As String
Dim ahickey As String

ahickey = "Master Inventory File_IsOpen.xlsx"
YellowBrickRoad = ThisWorkbook.Path & "\" & ahickey
Kill YellowBrickRoad
End Sub



