# Exlsheetvba

Sub SummaryOfSheets()
    Dim ws As Worksheet
    Dim summaryWs As Worksheet
    Dim lastRow As Long
    Dim completedRow As Long
    Dim pendingRow As Long
    
    ' Create or clear the summary sheet
    On Error Resume Next
    Set summaryWs = ThisWorkbook.Worksheets("Summary")
    If summaryWs Is Nothing Then
        Set summaryWs = ThisWorkbook.Worksheets.Add
        summaryWs.Name = "Summary"
    Else
        summaryWs.Cells.Clear
    End If
    On Error GoTo 0
    
    ' Headers for the summary sheet
    summaryWs.Range("A1").Value = "Sheet Name"
    summaryWs.Range("B1").Value = "Row Count"
    summaryWs.Range("C1").Value = "Status"
    
    completedRow = 2
    pendingRow = 2
    
    ' Loop through each sheet and gather information
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name <> "Summary" Then
            ' Get row count
            lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
            
            ' Add sheet name and row count to the summary sheet
            summaryWs.Cells(completedRow, 1).Value = ws.Name
            summaryWs.Cells(completedRow, 2).Value = lastRow - 1 ' Exclude header row
            
            ' Check the status column and categorize rows
            Dim i As Long
            For i = 2 To lastRow
                If ws.Cells(i, 3).Value = "Completed" Then
                    summaryWs.Cells(completedRow, 3).Value = "Completed"
                    completedRow = completedRow + 1
                ElseIf ws.Cells(i, 3).Value = "Pending" Then
                    summaryWs.Cells(pendingRow, 3).Value = "Pending"
                    pendingRow = pendingRow + 1
                End If
            Next i
        End If
    Next ws
    
    MsgBox "Summary Created Successfully!", vbInformation
End Sub
