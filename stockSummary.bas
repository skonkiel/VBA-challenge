Attribute VB_Name = "Module1"
Option Explicit

Sub stockSummary()

Dim ws As Object

'Create a script that will loop through all the stocks for one year for each run and summarize the following information.

For Each ws In Worksheets

    Dim lastRow, vol, yearChange, pctChange, tableRow, counter As Double
    Dim ticker As String
    
    Dim openAmt, closeAmt As Double
    
    counter = 0
    
    ' Initialize Table Rows for the highlights table
    Dim highlightTableRow As Byte
    highlightTableRow = 2
    
       ' Find final row
    lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    ' Set labels for summary table headers
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly change"
    ws.Cells(1, 11).Value = "Percent change"
    ws.Cells(1, 12).Value = "Total stock volume"
    
    tableRow = 2

    vol = 0
    
    Dim i As Double
    
    For i = 2 To lastRow
        
        ' If the following cell does not match the current cell, write the Volume, etc values to the tablerow and increment the tableRow
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
       
            'Add to the total stock volume of the stock.
            vol = vol + ws.Range("G" & i)
            
            'Add the ticker symbol to the summary table
            ticker = ws.Cells(i, 1).Value
            ws.Cells(tableRow, 9).Value = ticker
            
            'Add the finalized volume
            ws.Cells(tableRow, 12).Value = vol
            
            vol = 0
            
            ' Get closing price
            closeAmt = ws.Cells(i, 6).Value
            ' Cells(tableRow, 14).Value = closeAmt
                  
            ' Get the opening price by doing (i - counter) to get opening price at start of ticker
            openAmt = ws.Cells((i - counter), 3).Value
            ' Cells(tableRow, 13).Value = openAmt
            
            'Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.
            yearChange = closeAmt - openAmt
            ws.Cells(tableRow, 10).Value = yearChange
            
            ' Highlight pos/neg
            If yearChange < 0 Then
                ws.Cells(tableRow, 10).Interior.ColorIndex = 3
            Else
                ws.Cells(tableRow, 10).Interior.ColorIndex = 4
            End If
            
             'The percent change from opening price at the beginning of a given year to the closing price at the end of that year.
            If openAmt <> "0" Then
                pctChange = yearChange / openAmt
            Else
                pctChange = "0"
            End If
        
            ws.Cells(tableRow, 11).Value = pctChange
            ws.Range("K" & tableRow).NumberFormat = "0.00%"
            
            ' Reset values
            counter = 0
            tableRow = tableRow + 1
            
        Else
        'If the cell matches the cell above it - add to the volume

            ' Add current cell value to the Total Volume
            vol = vol + ws.Range("G" & i)
            
            ' Increment the counter
            counter = counter + 1
                 
        End If
            
    Next i
    
    ' Return the stock with the "Greatest % increase", "Greatest % Decrease" and "Greatest total volume".
    ' Find row index for each value and use that to return ticker symbol
    Dim pctMax, pctMin, maxVol As Double ' variable that finds the max amt
    Dim pctMaxTicker, pctMinTicker, maxVolTicker As String ' stores location of max amt's ticker
    
    pctMax = 0
    pctMin = 0
    maxVol = 0
    
    ' Find final row of summary table
    Dim LastRowSum As Double
    LastRowSum = ws.Cells(Rows.Count, 9).End(xlUp).Row
    
    For i = 2 To LastRowSum
        ' find largest percentage change
        If (ws.Cells(i, 11).Value > pctMax) Then
            pctMax = ws.Cells(i, 11).Value
            pctMaxTicker = ws.Cells(i, 9).Value
        ' Find smallest percentage change
        ElseIf (ws.Cells(i, 11).Value < pctMin) Then
            pctMin = ws.Cells(i, 11).Value
            pctMinTicker = ws.Cells(i, 9).Value
        End If
    
        ' Find maximum stock volume
        If (ws.Cells(i, 12).Value > maxVol) Then
            maxVol = ws.Cells(i, 12).Value
            maxVolTicker = ws.Cells(i, 9).Value
        End If
    Next i
    
    ' Set highlight table values
    ws.Cells(highlightTableRow, 16).Value = pctMax
    ws.Cells(highlightTableRow, 15).Value = pctMaxTicker
    highlightTableRow = highlightTableRow + 1
    ws.Cells(highlightTableRow, 16).Value = pctMin
    ws.Cells(highlightTableRow, 15).Value = pctMinTicker
    highlightTableRow = highlightTableRow + 1
    ws.Cells(highlightTableRow, 16).Value = maxVol
    ws.Cells(highlightTableRow, 15).Value = maxVolTicker
    
    ws.Columns("A:Z").AutoFit
    
    ' Set highlight table - headers
    ws.Range("O1").Value = "Ticker"
    ws.Range("P1").Value = "Value"


        
    ' Format highlight table numbers
    ws.Range("P4").NumberFormat = "General"
    ws.Range("P2:P3").NumberFormat = "0.00%"
    
        ' Set column names for highlights table
    ws.Range("N2").Value = "Greatest % increase"
    ws.Range("N3").Value = "Greatest % Decrease"
    ws.Range("N4").Value = "Greatest total volume"

    Next ws
    
    MsgBox ("Analysis complete!")

End Sub
