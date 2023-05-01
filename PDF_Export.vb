'PDF EXPORT MACROS HUMANA MMO 
 
Private Sub Command57_Click() 
 
Dim rst As DAO.Recordset 
Dim db As DAO.Database 
Dim qdf As QueryDef 
 
Set db = CurrentDb 
Set qdf = db.QueryDefs("Data_Crosstab_Monthly") 
qdf.Parameters(0) = [Forms]![CCFEForm]![reportDate] 
Set rst = qdf.OpenRecordset 
 
Do While Not rst.EOF 
    strRptFilter = "[NHC_Group_Nm] = " & Chr(34) & rst![NHC_Group_Nm] & Chr(34) 
     
    DoCmd.OutputTo acOutputReport, "Monthly Group Breakdown", acFormatPDF, "C:\Users\wilsonjaso\Desktop\output" _ 
    & "\" & rst![NHC_Group_Nm] & " Humana CCF (Monthly Activity Total) " & format(Date, "yyyymmdd") & ".pdf" 
    DoEvents 
    rst.MoveNext 
Loop 
 
rst.Close 
Set rst = Nothing 
Set qdf = Nothing 
Set db = Nothing 
 
End Sub 
 
Private Sub Command58_Click() 
 
Dim rst As DAO.Recordset 
Dim db As DAO.Database 
Dim qdf As QueryDef 
 
Set db = CurrentDb 
Set qdf = db.QueryDefs("Data_Crosstab_PCPs") 
qdf.Parameters(0) = [Forms]![CCFEForm]![reportDate] 
Set rst = qdf.OpenRecordset 
 
Do While Not rst.EOF 
    strRptFilter = "[NHC_Group_Nm] = " & Chr(34) & rst![NHC_Group_Nm] & Chr(34) 
     
    DoCmd.OutputTo acOutputReport, "PCP Breakdown Report", acFormatPDF, "C:\Users\wilsonjaso\Desktop\output" _ 
    & "\" & rst![NHC_Group_Nm] & " Humana CCF (Monthly Activity by PCP) " & format(Date, "yyyymmdd") & ".pdf" 
    DoEvents 
    rst.MoveNext 
Loop 
 
rst.Close 
Set rst = Nothing 
Set qdf = Nothing 
Set db = Nothing 
 
End Sub 
 
Private Sub Command59_Click() 
 
Dim rst As DAO.Recordset 
Dim db As DAO.Database 
Dim qdf As QueryDef 
 
Set db = CurrentDb 
Set qdf = db.QueryDefs("Data_Crosstab_Retro") 
qdf.Parameters(0) = [Forms]![CCFEForm]![reportDate] 
Set rst = qdf.OpenRecordset 
 
Do While Not rst.EOF 
    strRptFilter = "[NHC_Group_Nm] = " & Chr(34) & rst![NHC_Group_Nm] & Chr(34) 
     
    DoCmd.OutputTo acOutputReport, "Retro Breakout Report", acFormatPDF, "C:\Users\wilsonjaso\Desktop\output" _ 
    & "\" & rst![NHC_Group_Nm] & " Humana CCF (Retroactivity by PCP) " & format(Date, "yyyymmdd") & ".pdf" 
    DoEvents 
    rst.MoveNext 
Loop 
 
rst.Close 
Set rst = Nothing 
Set qdf = Nothing 
Set db = Nothing 
 
End Sub 
 
Private Sub Command10_Click() 
 
Dim rst As DAO.Recordset 
Dim db As DAO.Database 
Dim qdf As QueryDef 
 
Set db = CurrentDb 
Set qdf = db.QueryDefs("invPayeeCrosstab_2") 
qdf.Parameters(0) = [Forms]![Invoicing Report Form]![formYear] 
qdf.Parameters(1) = [Forms]![Invoicing Report Form]![formPeriod] 
Set rst = qdf.OpenRecordset 
 
Do While Not rst.EOF 
    strRptFilter = "[Practice] = " & Chr(34) & rst![Practice] & Chr(34) 
     
    DoCmd.OutputTo acOutputReport, "invRptByPayee", acFormatPDF, "C:\Users\wilsonjaso\Desktop\output" _ 
    & "\" & rst![Practice] & " " & rst![Report Name] & " " & format(Date, "yyyymmdd") & ".pdf" 
    DoEvents 
    rst.MoveNext 
Loop 
 
rst.Close 
Set rst = Nothing 
Set qdf = Nothing 
Set db = Nothing 
 
End Sub 
