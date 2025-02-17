Sub CompareSheetsByNameWithReport()
    Dim ws1 As Worksheet, ws2 As Worksheet, wsReport As Worksheet, wsSummary As Worksheet
    Dim dict1 As Object, dict2 As Object
    Dim lastRow1 As Long, lastRow2 As Long, lastCol1 As Long, lastCol2 As Long
    Dim i As Long, j As Long, reportRow As Long
    Dim key As Variant ' Ensure key is Variant
    Dim row1 As Range, row2 As Range
    Dim isModified As Boolean
    Dim oldValue As String, newValue As String
    Dim headers As Object
    Dim header As Variant
    Dim col As Long
    Dim modifiedCount As Long, addedCount As Long, removedCount As Long
    
    ' Set the sheets to compare (if sheet names change then names in "" need to match)
    Set ws1 = ThisWorkbook.Sheets("extract from SNow")
    Set ws2 = ThisWorkbook.Sheets("import from Ellipse")
    
    ' Create dictionaries to store row data keyed by "Name" column
    Set dict1 = CreateObject("Scripting.Dictionary")
    Set dict2 = CreateObject("Scripting.Dictionary")
    
    ' Create a dictionary for headers
    Set headers = CreateObject("Scripting.Dictionary")
    
    ' Add a new sheet for the report
    On Error Resume Next
    Set wsReport = ThisWorkbook.Sheets("Comparison Results")
    If wsReport Is Nothing Then
        Set wsReport = ThisWorkbook.Sheets.Add
        wsReport.Name = "Comparison Results"
    Else
        wsReport.Cells.Clear
    End If
    On Error GoTo 0
    
    ' Add headers to the report
    wsReport.Cells(1, 1).Value = "Name"
    wsReport.Cells(1, 2).Value = "Status"
    col = 3
    For j = 1 To ws1.Cells(1, ws1.Columns.Count).End(xlToLeft).Column
        If ws1.Cells(1, j).Value <> "Name" And ws1.Cells(1, j).Value <> "Updated" Then
            wsReport.Cells(1, col).Value = ws1.Cells(1, j).Value & " (Old)"
            wsReport.Cells(1, col + 1).Value = ws1.Cells(1, j).Value & " (New)"
            col = col + 2
        End If
    Next j
    wsReport.Rows(1).Font.Bold = True
    reportRow = 2
    
    ' Format the "Name" column as text to preserve leading zeros
    wsReport.Columns(1).NumberFormat = "@"
    
    ' Find the last rows and columns in each sheet
    lastRow1 = ws1.Cells(ws1.Rows.Count, 1).End(xlUp).Row
    lastRow2 = ws2.Cells(ws2.Rows.Count, 1).End(xlUp).Row
    lastCol1 = ws1.Cells(1, ws1.Columns.Count).End(xlToLeft).Column
    lastCol2 = ws2.Cells(1, ws2.Columns.Count).End(xlToLeft).Column
    
    ' Populate dictionary for headers
    For j = 1 To lastCol1
        headers(ws1.Cells(1, j).Value) = j
    Next j
    
    ' Populate dictionary for "extract"
    For i = 2 To lastRow1 ' Start at row 2 to skip headers
        dict1(CStr(ws1.Cells(i, headers("Name")).Value)) = i ' Use "Name" as key
    Next i
    
    ' Populate dictionary for "import"
    For i = 2 To lastRow2
        dict2(CStr(ws2.Cells(i, headers("Name")).Value)) = i ' Use "Name" as key
    Next i
    
    ' Initialize counters
    modifiedCount = 0
    addedCount = 0
    removedCount = 0
    
    ' Check for removed and modified rows
    For Each key In dict1.Keys
        If Not dict2.exists(key) Then
            ' Row in "extract" but not in "import" (removed row)
            ws1.Rows(dict1(key)).Interior.Color = RGB(255, 0, 0) ' Red
            wsReport.Cells(reportRow, 1).Value = key
            wsReport.Cells(reportRow, 2).Value = "Removed"
            removedCount = removedCount + 1
            reportRow = reportRow + 1
        Else
            ' Row exists in both; check for modifications
            Set row1 = ws1.Rows(dict1(key))
            Set row2 = ws2.Rows(dict2(key))
            isModified = False
            col = 3
            
            For Each header In headers.Keys
                If header <> "Name" And header <> "Updated" Then ' Exclude "Name" and "Updated"
                    If CStr(row1.Cells(1, headers(header)).Value) <> CStr(row2.Cells(1, headers(header)).Value) Then
                        isModified = True
                        oldValue = CStr(row1.Cells(1, headers(header)).Value)
                        newValue = CStr(row2.Cells(1, headers(header)).Value)
                        wsReport.Cells(reportRow, col).Value = oldValue
                        wsReport.Cells(reportRow, col + 1).Value = newValue
                    End If
                    col = col + 2
                End If
            Next header
            
            If isModified Then
                wsReport.Cells(reportRow, 1).Value = key
                wsReport.Cells(reportRow, 2).Value = "Modified"
                modifiedCount = modifiedCount + 1
                reportRow = reportRow + 1
            End If
        End If
    Next key
    
    ' Check for added rows
    For Each key In dict2.Keys
        If Not dict1.exists(key) Then
            ' Row in "import" but not in "extract" (added row)
            ws2.Rows(dict2(key)).Interior.Color = RGB(0, 255, 0) ' Green
            wsReport.Cells(reportRow, 1).Value = key
            wsReport.Cells(reportRow, 2).Value = "Added"
            addedCount = addedCount + 1
            reportRow = reportRow + 1
        End If
    Next key
    
    ' Format the report sheet
    wsReport.Columns("A:Z").AutoFit
    
    ' Add a new sheet for the summary
    On Error Resume Next
    Set wsSummary = ThisWorkbook.Sheets("Summary")
    If wsSummary Is Nothing Then
        Set wsSummary = ThisWorkbook.Sheets.Add
        wsSummary.Name = "Summary"
    Else
        wsSummary.Cells.Clear
    End If
    On Error GoTo 0
    
    ' Add headers and data to the summary sheet
    wsSummary.Cells(1, 1).Value = "Status"
    wsSummary.Cells(1, 2).Value = "Count"
    wsSummary.Rows(1).Font.Bold = True
    wsSummary.Cells(2, 1).Value = "Modified"
    wsSummary.Cells(2, 2).Value = modifiedCount
    wsSummary.Cells(3, 1).Value = "Added"
    wsSummary.Cells(3, 2).Value = addedCount
    wsSummary.Cells(4, 1).Value = "Removed"
    wsSummary.Cells(4, 2).Value = removedCount
    wsSummary.Columns("A:B").AutoFit
    
    MsgBox "Comparison Complete! Results are in 'Comparison Results' and 'Summary'.", vbInformation
End Sub
