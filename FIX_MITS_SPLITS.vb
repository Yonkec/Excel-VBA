'Older version of split macro? 
 
 
Sub FIXMITSPLITSPersonal() 
 
  Application.ScreenUpdating = False 
  Application.EnableEvents = False 
 
Dim c, r, d As Integer 
Dim dataWidth, dataHeight As Long 
Dim bottomDeleteRange, topDeleteRange As String 
Dim summarySheet As Worksheet 
Dim filePath As String 
'Dim rowNum As Long 
Dim wb1, wb2 As Workbook 
Dim wsLkUp, ws2 As Worksheet 
Dim copyWidth, copyHeight As Long 
Dim sourceRange, destRange As Range 
'Dim col As Integer 
Dim fileCount As Integer, fileEpiType As String 
Dim cleanName As Variant 
Dim errorInFile As Boolean, errorInBatch As Boolean 
Dim episodeTracker As String 
 
''''''''''''''''''''''''''''''''''''''''' 
'Begin directory loading and preparation' 
''''''''''''''''''''''''''''''''''''''''' 
 
    Set wb1 = ActiveWorkbook 
    Set wsLkUp = wb1.Worksheets("Lookup") 
     
    'Request the Target Folder Path From User 
        Set pickFolder = Application.FileDialog(msoFileDialogFolderPicker) 
     
        With pickFolder 
            .Title = "Please select the folder containing the new MITS data" 
            .AllowMultiSelect = False 
               If .Show <> -1 Then GoTo NextCode 
               filePath = .SelectedItems(1) & "\" 
         End With 
 
    'If user cancels 
NextCode: 
    filePath = filePath 
    If filePath = "" Then GoTo ResetSettingsAndClose 
     
    'Initialize fileCount to track number of files being processed 
    fileCount = 0 
    episodeTracker = "Blank" 
 
    'Initialize Dir, pointing it to the user provided folder path 
    fileName = Dir(filePath & "*.xl*") 
 
'''''''''''''''''''''''''''''''''''''''''''''''' 
'Open original file and begin modifications' 
'''''''''''''''''''''''''''''''''''''''''''''''' 
 
Do While fileName <> "" 
 
    Set wb2 = Workbooks.Open(filePath & fileName) 
     
    'Ensure Workbook has opened before moving on to next line of code 
      DoEvents 
       
    Set ws2 = wb2.Worksheets(1) 
 
    'Determines the bottom of the data table, and which rows to delete above and below it 
    dataHeight = ws2.Range("C1048576").End(xlUp).Row 
    topDeleteRange = "1:6" 
    bottomDeleteRange = dataHeight - 5 & ":" & dataHeight + 10 
     
        ' Deletes the first 6 rows and any partial rows below the data, preparing it for import 
        ws2.Range("1:6").Select 
        Selection.Delete 
        ws2.Range(bottomDeleteRange).Select 
        Selection.Delete 
         
    'Determines the width of the resulting data table 
    dataWidth = ws2.Range("XFD1").End(xlToLeft).Column 
    dataHeight = ws2.Range("C1048576").End(xlUp).Row 
     
    'Loops through each column in the data 
    For c = 1 To dataWidth 
     
    'If a field header ends with " num" we assume it and the field immediately to the right of it 
    'relate to a split numerator/denominator quality metric and combine them accordingly 
    If Right(ws2.Cells(1, c), 4) = " num" Then 
        ws2.Columns(c).Select 
        Selection.Insert 
         
        ws2.Cells(1, c) = Left(ws2.Cells(1, c + 1), Len(ws2.Cells(1, c + 1)) - 3) 
         
        'Loops through each episode's line and inserts the combined metric value 
        For r = 2 To dataHeight 
         
        If ws2.Cells(r, c + 2) = 0 Or ws2.Cells(r, c + 2) = "null" Or ws2.Cells(r, c + 2) = "Null" Then 
            ws2.Cells(r, c) = 0 
        Else 
            ws2.Cells(r, c) = ws2.Cells(r, c + 1) / ws2.Cells(r, c + 2) 
        End If 
         
        Next 
         
        'Once the new combined field has been generated this removes the original num and den columns 
        For d = 1 To 2 
            ws2.Columns(c + 1).Select 
            Selection.Delete 
        Next d 
     
    End If 
    Next c 
 
''''''''''''''''''''''''''''''''''''''''''''' 
''Begin additional modifications of the file' 
''''''''''''''''''''''''''''''''''''''''''''' 
    Range("A:A").Select 
        Selection.Insert Shift:=xlToRight 
        Range("A1").Value = fileName 
    Range("B:H").Select 
        Selection.Insert Shift:=xlToRight 
    Range("A1").Select 
        Selection.TextToColumns Destination:=Range("A1"), DataType:=xlDelimited, Other:=True, _ 
        OtherChar:=".", FieldInfo:=Array(Array(1, 1), Array(2, 1), Array(3, 1), Array(4, 1), _ 
        Array(5, 1), Array(6, 3), Array(7, 1), Array(8, 1)) 
    Range("A2:H" & dataHeight).Select 
        Selection.FormulaR1C1 = "=R[-1]C" 
    Range("A1:H" & dataHeight).Select 
        Selection.Copy 
        Selection.PasteSpecial xlPasteValues 
    Range("F1:F" & dataHeight).Select 
        Selection.NumberFormat = "m/d/yyyy" 
    Range("A:B, H:H").Select 
        Selection.Delete 
         
    fileEpiType = Range("A1").Value 
         
    Range("A1").Value = "EPISODE" 
    Range("B1").Value = "rptBeginDate" 
    Range("C1").Value = "rptEndDate" 
    Range("D1").Value = "publishedDate" 
    Range("E1").Value = "MDC_ID" 
     
    Range("A:D").Select 
        Selection.Insert Shift:=xlToRight 
         
    Range("A1").Value = "Unique Record Key" 
    Range("B1").Value = "EOC Year" 
    Range("C1").Value = "EOC CODE" 
    Range("D1").Value = "Report Source" 
        
    'Estimates an EOC Year based on the reporting period dates. 
    'Periods that begin midyear are assumed to represent the following year's Episode period 
    'IE: 20150501 to 20160631 is assumed to represent a 2016 Episode 
    For r = 2 To dataHeight 
            Cells(r, 1) = Cells(r, 10) & Int(CDbl(Cells(r, 8))) & Cells(r, 5) 
             
            If Right(Cells(r, 6), 4) = "0101" Then 
                Cells(r, 2) = Left(Cells(r, 6), 4) 
                Cells(r, 3) = Cells(r, 5) & Left(Cells(r, 6), 4) 
                Cells(r, 4) = "MITS" 
            Else 
                Cells(r, 2) = Left(Cells(r, 6), 4) + 1 
                Cells(r, 3) = Cells(r, 5) & Left(Cells(r, 6), 4) + 1 
                Cells(r, 4) = "MITS" 
            End If 
    Next 
     
     
    'Iterates through finalized columns and standardizes the naming convention with a lookup 
    dataWidth = ws2.Range("XFD1").End(xlToLeft).Column 
     
    For c = 8 To dataWidth 
     
    cleanName = MetricLookups(ws2.Cells(2, 5).Value, ws2.Cells(1, c).Value, wsLkUp) 
     
'        cleanName = Application.Index(wsLkUp.Range("B:B"), _ 
'                    Application.Match(ws2.Cells(1, c), wsLkUp.Range("C:C"), 0)) 
         
        'Error Handling for bad lookups, if value not found then you have the option to quit the program early 
        'Selecting cancel quits the code but keeps the problematic file open for review 
        ' Continuing flags the file to be saved with a note indicating an issue was found within 
        If Not cleanName = "NotFound" Then 
            ws2.Cells(1, c).Value = cleanName 
        Else 
'            If MsgBox("File: " & vbNewLine & vbNewLine & fileName & vbNewLine & vbNewLine & _ 
'                      "At Column: " & c & vbNewLine & vbNewLine & ws2.Cells(1, c), 1, "Lookup Error") = 2 Then 
'                GoTo ResetSettingsAndCloseEarly 
'            Else 
                errorInFile = True 
                errorInBatch = True 
                ws2.Cells(1, c).Interior.ColorIndex = 27 
'            End If 
             
        End If 
         
    Next 
     
    'Iterates through the entire body of the file to remove any null values / prevent type mismatches in access 
    For c = 9 To dataWidth 
        For r = 2 To dataHeight 
            If ws2.Cells(r, c).Value = "null" Or ws2.Cells(r, c).Value = "Null" Then 
                ws2.Cells(r, c).ClearContents 
            End If 
        Next r 
    Next c 
     
    'Monitors the episode type and runs a counter for the number of files being processed per episode 
    If episodeTracker = ws2.Cells(2, 5).Value Or episodeTracker = "Blank" Then 
        fileCount = fileCount + 1 
    Else 
        fileCount = 1 
    End If 
 
    episodeTracker = ws2.Cells(2, 5).Value 
     
'''''''''''''''''''''' 
''Save close and exit' 
'''''''''''''''''''''' 
     
    'Save as new processed file and Close Workbook. If error was encountered, marks file for inspection 
    If errorInFile = True Then 
       wb2.SaveAs fileName:=filePath & "Processed" & "\" & fileEpiType & " - " & fileCount & " CHECK HEADERS", _ 
            FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False 
    Else 
        wb2.SaveAs fileName:=filePath & "Processed" & "\" & fileEpiType & " - " & fileCount, _ 
            FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False 
    End If 
            
        wb2.Close 
        errorInFile = False 
                       
    'Ensure Workbook has closed before moving on to next line of code 
      DoEvents 
 
    'Get next file name 
      fileName = Dir() 
Loop 
 
GoTo ResetSettingsAndClose: 
         
ResetSettingsAndCloseEarly: 
    MsgBox ("The problematic file has been left open to review." & vbNewLine & _ 
            "Do not save this, it is the original (partially modified) copy") 
 
ResetSettingsAndClose: 
    Application.EnableEvents = True 
    Application.ScreenUpdating = True 
    If errorInBatch = True Then 
        MsgBox ("Finished - Any errors are indicated in the processed file names") 
    Else 
    End If 
     
End Sub 
 
Public Function MetricLookups(episodeID As String, headerText As String, ByVal lookupWS As Worksheet) As String 
 
    Dim row_index As Long 
    Dim dataHeight As Long 
     
    dataHeight = lookupWS.Range("C1048576").End(xlUp).Row 
     
    For row_index = 2 To dataHeight 
 
        If (lookupWS.Cells(row_index, 1).Value = episodeID Or lookupWS.Cells(row_index, 1).Value = "GENERIC") _ 
        And (lookupWS.Cells(row_index, 3).Value = headerText) Then 
 
            MetricLookups = lookupWS.Cells(row_index, 2).Value 
            Exit Function 
 
        End If 
 
    Next row_index 
     
    MetricLookups = "NotFound" 
 
End Function 
        
 
 
 
 
 
 
 
 
 
------------------------------------------------------------------------------------------------------------------------------------- 
 
 
 
 
 
 
 
 
'Old version of split 
 
Sub FIXMITSPLITS() 
 
  Application.ScreenUpdating = False 
  Application.EnableEvents = False 
 
Dim c, r, d As Integer 
Dim dataWidth, dataHeight As Integer 
Dim bottomDeleteRange, topDeleteRange As String 
Dim summarySheet As Worksheet 
Dim filePath As String 
Dim rowNum As Long 
Dim wb As Workbook 
Dim ws As Worksheet 
Dim copyWidth, copyHeight As Integer 
Dim sourceRange, destRange As Range 
Dim col As Integer 
 
''''''''''''''''''''''''''''''''''''''''' 
'Begin directory loading and preparation' 
''''''''''''''''''''''''''''''''''''''''' 
 
    ' Create a new blank workbook and set a variable to the first sheet. 
    Set summarySheet = Workbooks.Add(xlWBATWorksheet).Worksheets(1) 
     
    'Request the Target Folder Path From User 
        Set pickFolder = Application.FileDialog(msoFileDialogFolderPicker) 
     
        With pickFolder 
            .Title = "Please select the folder containing the new MITS data" 
            .AllowMultiSelect = False 
               If .Show <> -1 Then GoTo NextCode 
               filePath = .SelectedItems(1) & "\" 
         End With 
 
    'If user cancels 
NextCode: 
    filePath = filePath 
    If filePath = "" Then GoTo ResetSettingsAndClose 
     
    'Initialize rowNum which keeps track of where to insert new rows in summary file, col to track number of Quality Metrics 
    rowNum = 1 
    col = 1 
 
    'Initialize Dir, pointing it to the user provided folder path 
    fileName = Dir(filePath & "*.xl*") 
 
'''''''''''''''''''''''''''''''''''''''''''''''' 
'Begin individual file modifications and import' 
'''''''''''''''''''''''''''''''''''''''''''''''' 
 
Do While fileName <> "" 
 
    Set wb = Workbooks.Open(filePath & fileName) 
     
    'Ensure Workbook has opened before moving on to next line of code 
      DoEvents 
       
    Set ws = wb.Worksheets(1) 
     
    summarySheet.Range("A" & rowNum).Value = fileName 
 
    'Determines the bottom of the data table, and which rows to delete above and below it 
    dataHeight = ws.Range("C1048576").End(xlUp).Row 
    topDeleteRange = "1:6" 
    bottomDeleteRange = dataHeight - 5 & ":" & dataHeight + 10 
     
        ' Deletes the first 6 rows and any partial rows below the data, preparing it for import 
        ws.Range("1:6").Select 
        Selection.Delete 
        ws.Range(bottomDeleteRange).Select 
        Selection.Delete 
         
    'Determines the width of the resulting data table 
    dataWidth = ws.Range("XFD1").End(xlToLeft).Column 
     
    'Loops through each column in the data 
    For c = 1 To dataWidth 
     
    'If a field header ends with " num" we assume it and the field immediately to the right of it 
    'relate to a split numerator/denominator quality metric and combine them accordingly 
    If Right(ws.Cells(1, c), 4) = " num" Then 
        ws.Columns(c).Select 
        Selection.Insert 
         
        ws.Cells(1, c) = Left(ws.Cells(1, c + 1), Len(ws.Cells(1, c + 1)) - 3) 
         
        'Loops through each episode's line and inserts the combined metric value 
        For r = 2 To dataHeight 
         
        If ws.Cells(r, c + 2) = 0 Or ws.Cells(r, c + 2) = "null" Or ws.Cells(r, c + 2) = "Null" Then 
            ws.Cells(r, c) = 0 
        Else 
            ws.Cells(r, c) = ws.Cells(r, c + 1) / ws.Cells(r, c + 2) 
        End If 
         
        Next 
         
        'Once the new combined field has been generated this removes the original num and den columns 
        For d = 1 To 2 
            ws.Columns(c + 1).Select 
            Selection.Delete 
        Next d 
     
    End If 
    Next c 
 
    'Redetermines the height and width of the processed data in the import file 
    copyHeight = ws.Range("C1048567").End(xlUp).Row 
    copyWidth = ws.Range("XFD1").End(xlToLeft).Column 
     
    'aligns the source and destination ranges and copies the file into the summary 
    Set sourceRange = ws.Range(Cells(1, 1), Cells(copyHeight, copyWidth)) 
     
    Set destRange = summarySheet.Range("B" & rowNum) 
    Set destRange = destRange.Resize(sourceRange.Rows.Count, sourceRange.Columns.Count) 
     
    destRange.Value = sourceRange.Value 
     
    'saves the new rowNum position, closes the data file without saving, and prepares to open the next file 
    rowNum = rowNum + destRange.Rows.Count 
    wb.Close savechanges:=False 
    fileName = Dir() 
     
Loop 
 
'''''''''''''''''''''''''''''''''''''''''''''''''' 
'Begin modification of the resulting summary file' 
'''''''''''''''''''''''''''''''''''''''''''''''''' 
 
    Range("B:H").Select 
        Selection.Insert Shift:=xlToRight 
    Range("A:A").Select 
        Selection.TextToColumns Destination:=Range("A1"), DataType:=xlDelimited, Other:=True, _ 
        OtherChar:=".", FieldInfo:=Array(Array(1, 1), Array(2, 1), Array(3, 1), Array(4, 1), _ 
        Array(5, 1), Array(6, 3), Array(7, 1), Array(8, 1)) 
    Range("A:H").Select 
        Selection.SpecialCells(xlCellTypeBlanks).Select 
        Selection.FormulaR1C1 = "=R[-1]C" 
    Range("A:H").Select 
        Selection.Copy 
        Selection.PasteSpecial xlPasteValues 
    Range("F:F").Select 
        Selection.NumberFormat = "m/d/yyyy" 
    Range("A:B, H:H").Select 
        Selection.Delete 
         
    Range("A1").Value = "EPISODE" 
    Range("B1").Value = "rptBeginDate" 
    Range("C1").Value = "rptEndDate" 
    Range("D1").Value = "publishedDate" 
    Range("E1").Value = "MDC_ID" 
     
'    For col = 1 To 10 
'        If col < 10 Then 
'            Cells(1, col + 29) = "QM0" & col 
'        Else 
'            Cells(1, col + 29) = "QM" & col 
'        End If 
'    Next 
' 
'     Range("A1:ZZ" & rowNum).RemoveDuplicates Columns:=Array(6, 7, 8, 9, 10), Header:=xlNo 
' 
    Range("A:C").Select 
        Selection.Insert Shift:=xlToRight 
         
    Range("A1").Value = "Unique Record Key" 
    Range("B1").Value = "EOC Year" 
    Range("C1").Value = "EOC CODE" 
         
    dataHeight = Range("D1048576").End(xlUp).Row 
     
    For r = 2 To dataHeight 
            Cells(r, 1) = Cells(r, 9) & Int(CDbl(Cells(r, 7))) & Cells(r, 4) 
             
            If Right(Cells(r, 5), 4) = "0101" Then 
                Cells(r, 2) = Left(Cells(r, 5), 4) 
                Cells(r, 3) = Cells(r, 4) & Left(Cells(r, 5), 4) 
            Else 
                Cells(r, 2) = Left(Cells(r, 5), 4) + 1 
                Cells(r, 3) = Cells(r, 4) & Left(Cells(r, 5), 4) + 1 
            End If 
    Next 
         
    Columns.AutoFit 
     
ResetSettingsAndClose: 
    Application.EnableEvents = True 
    Application.ScreenUpdating = True 
 
End Sub 
