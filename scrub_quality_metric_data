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
 
''''''''''''''''''''''''''''''''' 
'Initial loading and preparation' 
''''''''''''''''''''''''''''''''' 
 
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
 
    'Dynamically finds the bottom of the data table, and determines what rows to delete above and below it 
    dataHeight = ws.Range("C1048576").End(xlUp).Row 
    topDeleteRange = "1:6" 
    bottomDeleteRange = dataHeight - 5 & ":" & dataHeight + 10 
     
        ' Deletes the first 6 rows and any partial rows below the data, preparing it for import 
        ws.Range("1:6").Select 
        Selection.Delete 
        ws.Range(bottomDeleteRange).Select 
        Selection.Delete 
         
    'Dynamically counts the width of the resulting data table 
    dataWidth = ws.Range("XFD1").End(xlToLeft).Column 
     
    'Loops through each column in the data 
    For c = 1 To dataWidth 
     
    'If a field header ends with " num" we assume it and the field immediately to the right of it 
    'relate to a split numerator/denominator quality metric and combine them appropriately 
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
        OtherChar:=".", FieldInfo:=Array(Array(1, 1), Array(2, 1), Array(3, 2), Array(4, 1), _ 
        Array(5, 1), Array(6, 3), Array(7, 2), Array(8, 2)) 
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
     
    For col = 1 To 10 
        If col < 10 Then 
            Cells(1, col + 29) = "QM0" & col 
        Else 
            Cells(1, col + 29) = "QM" & col 
        End If 
    Next 
         
    Range("A1:ZZ" & rowNum).RemoveDuplicates Columns:=Array(6, 7, 8, 9, 10), Header:=xlNo 
         
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
