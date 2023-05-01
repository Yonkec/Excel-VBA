Sub TelDel2() 
    Range("A1").Select 
    Selection.AutoFilter 
    Range("B1").Select 
    ActiveSheet.Range("$A$1:$AD$27").AutoFilter Field:=2, Criteria1:= _ 
        "=Episode Included or Excluded", Operator:=xlOr, Criteria2:="=" 
    Rows("1:1").Select 
    Range(Selection, Selection.End(xlDown)).Select 
    Selection.Delete Shift:=xlUp 
    Range("A1").Select 
End Sub 
 
 
 
 
'Split MITS file name apart as separate values 
 
Sub BreakFormatFileName() 
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
    ActiveWindow.ScrollColumn = 6 
    ActiveWindow.ScrollColumn = 5 
    ActiveWindow.ScrollColumn = 4 
    ActiveWindow.ScrollColumn = 2 
    ActiveWindow.ScrollColumn = 1 
    Range("A1").Select 
End Sub 
 
