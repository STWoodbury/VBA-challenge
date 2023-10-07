Attribute VB_Name = "Module1"
Sub stockAnalysis()

'Declare Variables
Dim openingPrice
Dim closingPrice
Dim stockTotal
Dim yearlyChange
Dim percentChange

'loop through the worksheets
For Each ws In Worksheets
    
    'Declare the data entry row
    Dim dataRow
    dataRow = 1
    
    'Add the headers to the data entry row
    ws.Cells(dataRow, 9) = "Ticker"
    ws.Cells(dataRow, 10) = "Yearly Change"
    ws.Cells(dataRow, 11) = "Percent Change"
    ws.Cells(dataRow, 12) = "Total Stock Volume"
    
    
    'Establish the last row
    lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
    'establish the initial opening price
    openingPrice = ws.Cells(2, 3).Value
    
    'Establish the initial stockTotal
    stockTotal = 0
    
    'Loop through the rows
    For i = 2 To lastRow
    
        'Conditional for unique tickers
        If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
            
            'Move down the data entry row
            dataRow = dataRow + 1
            
            'Insert the ticker
            ws.Cells(dataRow, 9) = ws.Cells(i, 1).Value
            
            'establish the closing price
            closingPrice = ws.Cells(i, 6).Value
            
            'Calculate and input the yearly change
            yearlyChange = closingPrice - openingPrice
            ws.Cells(dataRow, 10).Value = yearlyChange
                
                'Format the color of the yearly change
                If yearlyChange < 0 Then
                    ws.Cells(dataRow, 10).Interior.ColorIndex = 3
                Else
                    ws.Cells(dataRow, 10).Interior.ColorIndex = 4
                End If
            
            'Calculate and insert the percent change
            percentChange = yearlyChange / openingPrice
            ws.Cells(dataRow, 11).Value = percentChange
            
            'Reset the openingPrice to new ticker
            openingPrice = ws.Cells(i + 1, 3).Value
            
            'Format percentChange
            ws.Cells(dataRow, 11).NumberFormat = "0.00%"
            
            'Add the last row's total and insert data
            stockTotal = stockTotal + ws.Cells(i, 7).Value
            ws.Cells(dataRow, 12).Value = stockTotal
            
            'Reset stockTotal for next unique ticker
            stockTotal = 0
        Else
            'Add the stockTotal
            stockTotal = stockTotal + ws.Cells(i, 7).Value
        End If
    Next i
    
    'Declare Variables for comparisons
    Dim MaxTotalTicker
    Dim MaxTotal
    Dim MinPercentage
    Dim MinPercentageTicker
    Dim MaxPercentage
    Dim MaxPercentageTicker

    'Initialize Variables for first row of summary table
    MaxTotal = ws.Cells(2, 12).Value
    MaxTotalTicker = ws.Cells(2, 9).Value

    MinPercentage = ws.Cells(2, 11).Value
    MinPercentageTicker = ws.Cells(2, 9).Value

    MaxPercentage = ws.Cells(2, 11).Value
    MaxPercentageTicker = ws.Cells(2, 9).Value

    'Add headers and titles
    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(1, 17).Value = "Value"
    ws.Cells(2, 15).Value = "Greatest % Increase"
    ws.Cells(3, 15).Value = "Greatest % Decrease"
    ws.Cells(4, 15).Value = "Greatest Total Volume"

    'Find last row of the summary table
    lastRow = Range("I" & Rows.Count).End(xlUp).Row
    
    'Loop through the summary table to find max and min values and relative tickers
    For i = 2 To lastRow
        If ws.Cells(i, 12).Value > MaxTotal Then
            MaxTotal = ws.Cells(i, 12).Value
            MaxTotalTicker = ws.Cells(i, 9).Value
        End If
        If ws.Cells(i, 11).Value > MaxPercentage Then
            MaxPercentage = ws.Cells(i, 11).Value
            MaxPercentageTicker = ws.Cells(i, 9).Value
        End If
        If ws.Cells(i, 11).Value < MinPercentage Then
            MinPercentage = ws.Cells(i, 11).Value
            MinPercentageTicker = ws.Cells(i, 9).Value
        End If
    Next i

    'Add comparisons to table and format
    ws.Cells(2, 16).Value = MaxPercentageTicker
    ws.Cells(2, 17).Value = MaxPercentage
    ws.Cells(2, 17).NumberFormat = "0.00%"
    ws.Cells(3, 16).Value = MinPercentageTicker
    ws.Cells(3, 17).Value = MinPercentage
    ws.Cells(3, 17).NumberFormat = "0.00%"
    ws.Cells(4, 16).Value = MaxTotalTicker
    ws.Cells(4, 17).Value = MaxTotal

'Fix column widths
 ws.Cells.EntireColumn.AutoFit
    
'Advance to next worksheet
Next ws

End Sub

