Sub ExtractHL()

'1-Windows folder copy dropbox link
'2-paste link in incognito
'3-select all
'4-copy
'5-paste in sheet
'6-run script



Dim HL As Hyperlink
For Each HL In ActiveSheet.Hyperlinks
HL.Range.Offset(0, 1).Value = HL.Address
Next
End Sub

