Private Sub copyFiles_Click() 
 
Dim fso As Object 
Set fso = VBA.CreateObject("Scripting.FileSystemObject") 
 
Dim ws As Worksheet 
Dim folderStart As String 
Dim folderEnd As String 
Dim dataHeight As Integer 
Dim r As Integer 
 
  Application.ScreenUpdating = False 
  Application.EnableEvents = False 
   
''''''''''''''''''''''''''''''''''''''''' 
'Begin directory loading and preparation' 
''''''''''''''''''''''''''''''''''''''''' 
   
'Request Source Folder Path From User 
  Set FldrPicker = Application.FileDialog(msoFileDialogFolderPicker) 
 
    With FldrPicker 
      .Title = "Select A Source Directory" 
      .AllowMultiSelect = False 
        If .Show <> -1 Then GoTo NextCode 
        folderStart = .SelectedItems(1) & "\" 
    End With 
 
'Request Target Folder Path From User 
  Set FldrPicker = Application.FileDialog(msoFileDialogFolderPicker) 
 
    With FldrPicker 
      .Title = "Select A Target Directory" 
      .AllowMultiSelect = False 
        If .Show <> -1 Then GoTo NextCode 
        folderEnd = .SelectedItems(1) & "\" 
    End With 
 
'Validate that a directory was selected for both and it is not a duplicate 
NextCode: 
  If folderStart = "" Then GoTo ResetSettings 
  If folderEnd = "" Then GoTo ResetSettings 
  If folderEnd = folderStart Then GoTo ResetSettings 
   
'''''''''''''''''''''''''''''''''''''''''''''''' 
'Begin individual file modifications and import' 
'''''''''''''''''''''''''''''''''''''''''''''''' 
   
    'MsgBox "The name of the active sheet is " & ActiveSheet.Name 
     
    Set ws = Application.ActiveSheet 
     
    'Determines the bottom of the data table 
        dataHeight = ws.Range("A1048576").End(xlUp).Row 
         
    'Performs the actual file copying 
    For r = 2 To dataHeight 
     
        Filename = ws.Cells(r, 1) 
     
        Call fso.CopyFile(folderStart & Filename, folderEnd & Filename, True) 
         
    Next r 
 
MsgBox "A total of " & dataHeight & " files have been successfully copied" 
 
ResetSettings: 
    Application.EnableEvents = True 
    Application.ScreenUpdating = True 
 
End Sub 
