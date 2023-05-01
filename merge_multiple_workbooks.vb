Sub MergeAllWorkbooks() 
    Dim summarySheet As Worksheet 
    Dim FolderPath As String 
    Dim NRow As Long 
    Dim fileName As String 
    Dim WorkBk As Workbook 
    Dim sourceRange As Range 
    Dim destRange As Range 
    Dim LastRow As Long 
     
    ' Create a new workbook and set a variable to the first sheet. 
    Set summarySheet = Workbooks.Add(xlWBATWorksheet).Worksheets(1) 
     
    ' The folder containing the files to be merged 
    FolderPath = "C:\Users\wilsonjaso\Desktop\Test\" 
     
    ' NRow keeps track of where to insert new rows in the destination workbook. 
    NRow = 1 
     
    ' Call Dir the first time, pointing it to all Excel files in the folder path. 
    fileName = Dir(FolderPath & "*.xl*") 
     
    ' Loop until Dir returns an empty string. 
    Do While fileName <> "" 
        ' Open a workbook in the folder 
        Set WorkBk = Workbooks.Open(FolderPath & fileName) 
         
        ' Set the cell in column A to be the file name. 
        summarySheet.Range("A" & NRow).Value = fileName 
         
        ' Looks for the last row in the data and selects everything above it. 
        LastRow = WorkBk.Worksheets(1).Cells.Find(What:="*", _ 
                 after:=WorkBk.Worksheets(1).Cells.Range("A1"), _ 
                 SearchDirection:=xlPrevious, _ 
                 LookIn:=xlFormulas, _ 
                 SearchOrder:=xlByRows).Row 
        Set sourceRange = WorkBk.Worksheets(1).Range("A1:CZ" & LastRow) 
         
        ' Set the destination range to start at column B and 
        ' be the same size as the source range. 
        Set destRange = summarySheet.Range("B" & NRow) 
        Set destRange = destRange.Resize(sourceRange.Rows.Count, _ 
           sourceRange.Columns.Count) 
            
        ' Copy over the values from the source to the destination. 
        destRange.Value = sourceRange.Value 
         
        ' Increase NRow so that we know where to copy data next. 
        NRow = NRow + destRange.Rows.Count 
         
        ' Close the source workbook without saving changes. 
        WorkBk.Close savechanges:=False 
         
        ' Use Dir to get the next file name. 
        fileName = Dir() 
    Loop 
     
    'Breaks apart the file name using its periods as delimiters, turning the data into individual fields 
    Columns("B:H").Select 
    Selection.Insert Shift:=xlToRight 
    Columns("A:A").Select 
    Selection.TextToColumns Destination:=Range("A1"), DataType:=xlDelimited, _ 
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _ 
        Semicolon:=False, Comma:=False, Space:=False, Other:=True, OtherChar _ 
        :=".", FieldInfo:=Array(Array(1, 1), Array(2, 1), Array(3, 1), Array(4, 1), Array(5, _ 
        1), Array(6, 1), Array(7, 1), Array(8, 1)), TrailingMinusNumbers:=True 
    Range("A1:H1").Select 
    Range(Selection, Selection.End(xlDown)).Select 
    Columns("A:H").Select 
    Range("H1").Activate 
    Selection.SpecialCells(xlCellTypeBlanks).Select 
    Selection.FormulaR1C1 = "=R[-1]C" 
    Columns("A:H").Select 
    Columns("A:H").EntireColumn.AutoFit 
    Columns("F:F").Select 
    Selection.NumberFormat = "m/d/yyyy" 
    Range("A1").Select 
         
End Sub
