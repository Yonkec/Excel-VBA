Sub PrintOneLine() 
  
Dim rng As Range 
Dim WorkRng As Range 
Dim xWs As Worksheet 
  
On Error Resume Next 
xTitleId = "Print Each Line - Select only first column of cells IE A1:A10 not A1:D10" 
Set WorkRng = Application.Selection 
Set WorkRng = Application.InputBox("Range", xTitleId, WorkRng.Address, Type:=8) 
Set xWs = WorkRng.Parent 
      
For Each rng In WorkRng 
  xWs.PageSetup.PrintArea = rng.EntireRow.Address 
  xWs.PrintPreview 
Next 
    
End Sub
