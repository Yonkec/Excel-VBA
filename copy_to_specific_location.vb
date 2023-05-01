Sub CopyMove1() 
' 
' CopyMove1 Macro 
' 
' Keyboard Shortcut: Ctrl+Shift+B 
' 
    Range("A1").Select 
    Range(Selection, Selection.End(xlToRight)).Select 
    Range(Selection, Selection.End(xlDown)).Select 
    Selection.Copy 
    Windows("LINK - Acute PCI.xlsx").Activate 
    Range("E13").Select 
    Selection.End(xlToLeft).Select 
    Selection.End(xlUp).Select 
    Selection.End(xlDown).Select 
    Range("A21").Select 
    ActiveSheet.Paste 
    Range("B14").Select 
End Sub 
