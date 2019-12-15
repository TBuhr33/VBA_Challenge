Attribute VB_Name = "Module2"
Sub Final()

'modify in to loop through all worksheets in the workbook
Dim ws As Worksheet

'for loop through worksheets
For Each ws In Worksheets


    'set up headers for new table's and format % for needed columns
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    ws.Range("N2").Value = "Greatest % Increase"
    ws.Range("N3").Value = "Greatest % Decrease"
    ws.Range("N4").Value = "Greatest Total Volume"
    ws.Range("O1").Value = "Ticker"
    ws.Range("P1").Value = "Value"
    ws.Range("P2").NumberFormat = "0.00%"
    ws.Range("P3").NumberFormat = "0.00%"
    ws.Columns("K").NumberFormat = "0.00%"

    'establish variables
    Dim ticker As String
    Dim yearly_change As Double
    Dim percent_change As Double

    Dim total_volume As Double
    total_volume = 0

    Dim resultsTable As Integer
    resultsTable = 2

    'find last row of data
    Dim i As Long
    Dim LastRow As Long
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

    'run loop to place ticker values and corresponding yearly change,% change, and total  volumes
    For i = 2 To LastRow

        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    
            ticker = ws.Cells(i, 1).Value
            total_volume = total_volume + ws.Cells(i, 7).Value
        
            ws.Range("I" & resultsTable).Value = ticker
            ws.Range("L" & resultsTable).Value = total_volume
            total_volume = 0
        
            yearly_change = ws.Cells(i, 6).Value - ws.Cells(i, 3).Value
            ws.Range("J" & resultsTable).Value = yearly_change
            yearly_change = 0
        
            percent_change = ((ws.Cells(i, 6).Value - Cells(i, 3).Value) / 1) * 100
            ws.Range("K" & resultsTable).Value = percent_change
            percent_change = 0
        
            resultsTable = resultsTable + 1
        
        
        Else
        total_volume = total_volume + ws.Cells(i, 7).Value
        yearly_change = yearly_change + ws.Cells(i, 6).Value
        End If
    
    Next

    'create new loop iteration and find new last row in output table
    Dim x As Long
    Dim FormatLR As Long
    FormatLR = ws.Cells(Rows.Count, 10).End(xlUp).Row

    'run for loop to color corresponding positive or negative change
    For x = 2 To FormatLR
        If ws.Cells(x, 10).Value < 0 Then
            ws.Cells(x, 10).Interior.ColorIndex = 3
        
        ElseIf ws.Cells(x, 10).Value >= 0 Then
            ws.Cells(x, 10).Interior.ColorIndex = 4
        End If
    
    Next


    'now make 2nd output table, find max/minimum values in columns of first output table

    ws.Cells(2, 16).Value = WorksheetFunction.Max(Columns(11))

    ws.Cells(3, 16).Value = WorksheetFunction.Min(Columns(11))

    ws.Cells(4, 16).Value = WorksheetFunction.Max(Columns(12))

    'loop to find max/min and place corresponding ticker, repeat for each value
    For x = 2 To FormatLR
        If ws.Cells(x, 11).Value = ws.Cells(2, 16).Value Then
            ws.Cells(2, 15).Value = ws.Cells(x, 9).Value
        End If
    Next

    For x = 2 To FormatLR
        If ws.Cells(x, 11).Value = ws.Cells(3, 16).Value Then
            ws.Cells(3, 15).Value = ws.Cells(x, 9).Value
        End If
    Next

    For x = 2 To FormatLR
        If ws.Cells(x, 12).Value = ws.Cells(4, 16).Value Then
            ws.Cells(4, 15).Value = ws.Cells(x, 9).Value
        End If
    Next
 
Next ws


End Sub


