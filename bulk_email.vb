Private Sub CommandButton3_Click() 
 
    Dim OutApp As Object 
    Dim OutMail As Object 
    Dim sh As Worksheet 
    Dim cell As Range 
    Dim FileCell As Range 
    Dim rng As Range 
 
    With Application 
        .EnableEvents = False 
        .ScreenUpdating = True 
    End With 
 
    Set sh = Sheets("Lookup") 
 
    Set OutApp = CreateObject("Outlook.Application") 
 
    For Each cell In sh.Columns("B").Cells.SpecialCells(xlCellTypeConstants) 
 
        'Enter the path/file names in the C:Z column in each row 
        Set rng = sh.Cells(cell.Row, 1).Range("C1:Z1") 
 
        If cell.Value Like "?*@?*.?*" And _ 
           Application.WorksheetFunction.CountA(rng) > 0 Then 
            Set OutMail = OutApp.CreateItem(0) 
 
            With OutMail 
                .to = cell.Value 
                .Subject = "Testfile" 
                .Body = "Hi " & cell.Offset(0, -1).Value 
 
                For Each FileCell In rng.SpecialCells(xlCellTypeConstants) 
                    If Trim(FileCell) <> "" Then 
                        If Dir(FileCell.Value) <> "" Then 
                            .Attachments.Add FileCell.Value 
                        End If 
                    End If 
                Next FileCell 
 
                .Send  'Or use .Display 
            End With 
 
            Set OutMail = Nothing 
        End If 
    Next cell 
 
    Set OutApp = Nothing 
    With Application 
        .EnableEvents = True 
        .ScreenUpdating = True 
    End With 
End Sub 
