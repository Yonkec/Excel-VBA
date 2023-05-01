Sub Highlight() 
' 
' Highlight Macro 
' Highlights selected cells Yellow 
' 
' Keyboard Shortcut: Ctrl+Shift+Q 
' 
    With Selection.Interior 
        .Pattern = xlSolid 
        .PatternColorIndex = xlAutomatic 
        .color = 65535 
        .TintAndShade = 0 
        .PatternTintAndShade = 0 
    End With 
End Sub 
Sub Clear_Highlighting() 
' 
' Clear_Highlighting Macro 
' Removes highlighting of selected cells 
' 
' Keyboard Shortcut: Ctrl+Shift+W 
' 
    With Selection.Interior 
        .Pattern = xlNone 
        .TintAndShade = 0 
        .PatternTintAndShade = 0 
    End With 
End Sub 
Sub outsideBorder() 
' 
' outsideBorder Macro 
' Adds an outer border to selected cells 
' 
' Keyboard Shortcut: Ctrl+Shift+C 
' 
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone 
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone 
    With Selection.Borders(xlEdgeLeft) 
        .LineStyle = xlContinuous 
        .ColorIndex = 0 
        .TintAndShade = 0 
        .Weight = xlThin 
    End With 
    With Selection.Borders(xlEdgeTop) 
        .LineStyle = xlContinuous 
        .ColorIndex = 0 
        .TintAndShade = 0 
        .Weight = xlThin 
    End With 
    With Selection.Borders(xlEdgeBottom) 
        .LineStyle = xlContinuous 
        .ColorIndex = 0 
        .TintAndShade = 0 
        .Weight = xlThin 
    End With 
    With Selection.Borders(xlEdgeRight) 
        .LineStyle = xlContinuous 
        .ColorIndex = 0 
        .TintAndShade = 0 
        .Weight = xlThin 
    End With 
    Selection.Borders(xlInsideVertical).LineStyle = xlNone 
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone 
End Sub 
Sub fullBorders() 
' 
' fullBorders Macro 
' Adds full borders to selected cells 
' 
' Keyboard Shortcut: Ctrl+Shift+X 
' 
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone 
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone 
    With Selection.Borders(xlEdgeLeft) 
        .LineStyle = xlContinuous 
        .ColorIndex = 0 
        .TintAndShade = 0 
        .Weight = xlThin 
    End With 
    With Selection.Borders(xlEdgeTop) 
        .LineStyle = xlContinuous 
        .ColorIndex = 0 
        .TintAndShade = 0 
        .Weight = xlThin 
    End With 
    With Selection.Borders(xlEdgeBottom) 
        .LineStyle = xlContinuous 
        .ColorIndex = 0 
        .TintAndShade = 0 
        .Weight = xlThin 
    End With 
    With Selection.Borders(xlEdgeRight) 
        .LineStyle = xlContinuous 
        .ColorIndex = 0 
        .TintAndShade = 0 
        .Weight = xlThin 
    End With 
    With Selection.Borders(xlInsideVertical) 
        .LineStyle = xlContinuous 
        .ColorIndex = 0 
        .TintAndShade = 0 
        .Weight = xlThin 
    End With 
    With Selection.Borders(xlInsideHorizontal) 
        .LineStyle = xlContinuous 
        .ColorIndex = 0 
        .TintAndShade = 0 
        .Weight = xlThin 
    End With 
End Sub 
Sub noBorders() 
' 
' noBorders Macro 
' Removes all borders on selected cells 
' 
' Keyboard Shortcut: Ctrl+Shift+Z 
' 
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone 
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone 
    Selection.Borders(xlEdgeLeft).LineStyle = xlNone 
    Selection.Borders(xlEdgeTop).LineStyle = xlNone 
    Selection.Borders(xlEdgeBottom).LineStyle = xlNone 
    Selection.Borders(xlEdgeRight).LineStyle = xlNone 
    Selection.Borders(xlInsideVertical).LineStyle = xlNone 
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone 
End Sub 
