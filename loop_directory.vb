Sub LoopAllExcelFilesInFolder() 
 
Dim wb As Workbook 
Dim myPath As String 
Dim myFile As String 
Dim myExtension As String 
Dim FldrPicker As FileDialog 
 
  Application.ScreenUpdating = False 
  Application.EnableEvents = False 
 'Application.Calculation = xlCalculationManual 
 
'Request Target Folder Path From User 
  Set FldrPicker = Application.FileDialog(msoFileDialogFolderPicker) 
 
    With FldrPicker 
      .Title = "Select A Target Folder" 
      .AllowMultiSelect = False 
        If .Show <> -1 Then GoTo NextCode 
        myPath = .SelectedItems(1) & "\" 
    End With 
 
'If user cancels 
NextCode: 
  myPath = myPath 
  If myPath = "" Then GoTo ResetSettings 
 
'Target File Extension (must include wildcard "*") 
  myExtension = "*.xls*" 
 
'Target Path with Ending Extention 
  myFile = Dir(myPath & myExtension) 
 
'Loop through each Excel file in folder 
  Do While myFile <> "" 
    'Set variable equal to opened workbook 
      Set wb = Workbooks.Open(fileName:=myPath & myFile) 
     
    'Ensure Workbook has opened before moving on to next line of code 
      DoEvents 
     
    Range("A1").Select 
    Selection.AutoFilter 
    Range("B1").Select 
    ActiveSheet.Range("$A$1:$AD$27").AutoFilter Field:=2, Criteria1:= _ 
        "=Episode Included or Excluded", Operator:=xlOr, Criteria2:="=" 
    Rows("1:1").Select 
    Range(Selection, Selection.End(xlDown)).Select 
    Selection.Delete Shift:=xlUp 
    Range("A1").Select 
    Range(Selection, Selection.End(xlToRight)).Select 
    Range(Selection, Selection.End(xlDown)).Select 
    Selection.Copy 
     
    Windows("Book1.xlsx").Activate 
    Range("E13").Select 
    Selection.End(xlToLeft).Select 
    Selection.End(xlUp).Select 
    Selection.End(xlDown).Select 
    Range("A21").Select 
    ActiveSheet.Paste 
    Range("B14").Select 
     
    'Save and Close Workbook 
      wb.Close savechanges:=True 
       
    'Ensure Workbook has closed before moving on to next line of code 
      DoEvents 
 
    'Get next file name 
      myFile = Dir 
  Loop 
 
'Message Box when tasks are completed 
  MsgBox "Task Complete!" 
 
ResetSettings: 
    Application.EnableEvents = True 
   'Application.Calculation = xlCalculationAutomatic 
    Application.ScreenUpdating = True 
 
End Sub 
