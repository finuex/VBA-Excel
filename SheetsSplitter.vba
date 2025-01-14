Sub SplitSheets()
    Dim ws As Worksheet
    Dim newWB As Workbook

    For Each ws In ThisWorkbook.Worksheets
        Set newWB = Workbooks.Add
        ws.Copy Before:=newWB.Sheets(1)
        newWB.SaveAs Filename:=ThisWorkbook.Path & "\" & ws.Name & ".xlsx"
        newWB.Close SaveChanges:=False
    Next ws
End Sub

