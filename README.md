Sub SummaryOfSheets()
    Dim ws As Worksheet
    Dim summaryWs As Worksheet
    Dim lastRow As Long
    Dim totalRecords As Long
    Dim completedCount As Long
    Dim pendingCount As Long
    Dim summaryRow As Long

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
    summaryWs.Range("B1").Value = "Total Records"
    summaryWs.Range("C1").Value = "Completed"
    summaryWs.Range("D1").Value = "Pending"

    summaryRow = 2

    ' Loop through each sheet and gather information
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name <> "Summary" Then
            ' Get the last row in the sheet
            lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
            totalRecords = lastRow - 1 ' Exclude header row
            
            ' Initialize counts for each sheet
            completedCount = 0
            pendingCount = 0
            
            ' Count completed and pending based on the status column (assumed to be Column C)
            Dim i As Long
            For i = 2 To lastRow
                If ws.Cells(i, 3).Value = "Completed" Then
                    completedCount = completedCount + 1
                ElseIf ws.Cells(i, 3).Value = "Pending" Then
                    pendingCount = pendingCount + 1
                End If
            Next i
            
            ' Add data to the summary sheet
            summaryWs.Cells(summaryRow, 1).Value = ws.Name
            summaryWs.Cells(summaryRow, 2).Value = totalRecords
            summaryWs.Cells(summaryRow, 3).Value = completedCount
            summaryWs.Cells(summaryRow, 4).Value = pendingCount
            
            summaryRow = summaryRow + 1
        End If
    Next ws
    
    MsgBox "Summary Created Successfully!", vbInformation
End Sub
