Sub SplitbookSCConnect() 
 
Dim xWs As Worksheet 
Dim xPath As String 
xPath = Range("C45") 
Dim xDate As String 
xDate = Range("C46") 
Dim cleanGrpName As Variant 
Dim cleanFileName As Variant 
Dim lookupRange As Range 
Set lookupRange = Range("A1:C40")
 
 
Application.ScreenUpdating = False 
Application.DisplayAlerts = False 
 
For Each xWs In ActiveWorkbook.Sheets 
 
    cleanGrpName = Application.VLookup(xWs.name, lookupRange.Value, 2, True) 
    cleanFileName = Application.VLookup(xWs.name, lookupRange.Value, 3, True) 
     
    On Error GoTo SkipThisSheet: 
        xWs.Copy 
        Application.ActiveWorkbook.SaveAs fileName:=xPath & "\" & cleanGrpName & " " & cleanFileName & " " & xDate & ".xlsx" 
SkipThisSheet: 
        Application.ActiveWorkbook.Close False 
Next 
 
 
Application.DisplayAlerts = True 
Application.ScreenUpdating = True 
MsgBox "All Done!" 
 
End Sub
