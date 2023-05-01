Sub SplitbookSCMA() 
 
Dim xWs As Worksheet 
Dim xPath As String 
xPath = Range("C41") 
Dim xDate As String 
xDate = Range("C42") 
Dim cleanName As Variant 
Dim lookupRange As Range 
Set lookupRange = Range("A1:B40") 
 
 
Application.ScreenUpdating = False 
Application.DisplayAlerts = False 
 
For Each xWs In ActiveWorkbook.Sheets 
 
    cleanName = Application.VLookup(xWs.name, lookupRange.Value, 2, True) 
    On Error GoTo SkipThisSheet: 
        xWs.Copy 
        Application.ActiveWorkbook.SaveAs fileName:=xPath & "\" & cleanName & " SC MA Member Report " & xDate & ".xlsx" 
SkipThisSheet: 
        Application.ActiveWorkbook.Close False 
Next 
 
 
Application.DisplayAlerts = True 
Application.ScreenUpdating = True 
MsgBox "All Done!" 
 
End Sub
