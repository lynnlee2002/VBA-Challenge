# VBA-Challenge

Sub stockAnalysis()

'Create a variable for each worksheet
Dim ws As Worksheet

'Loop through and apply the following code to each worksheet
For Each ws In ThisWorkbook.Worksheets

    'Create a variable for the stock symbol
    Dim stockSymbol As String
    
    'Create a variable to hold the total stock volume
    Dim stockVolume As Double
    
    'Create a variable for the open price of the stock
    Dim openPrice As Double
    openPrice = ws.Cells(2, 3).Value
    
    'Create a variable for the close price of the stock
    Dim closePrice As Double
    
    'Create a variable for the yearly price change of the stock
    Dim yearlyChange As Double
    
    'Create a variable for the percent change in price of the stock
    Dim percentChange As Double
    
    'Create a variable for the row in the analysis area, starting from row 2
    Dim analysisRow As Integer
    analysisRow = 2
        
    'Obtain the number of rows based on column A
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    'Loop through each stock symbol, starting from the 2nd row
    'Create a variable for the loop
    Dim i As Long
    For i = 2 To lastRow
        
        'Check if the next symbol is different from the current one
        If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
            
            'If true, then the stock symbol is in the current cell and print to column I
            stockSymbol = ws.Cells(i, 1).Value
            ws.Range("I" & analysisRow).Value = stockSymbol
            
            'Obtain the close price of the stock
            closePrice = ws.Cells(i, 6).Value
            
            'If true, then calculate the yearly change of the stock and print to column J
            yearlyChange = closePrice - openPrice
            ws.Range("J" & analysisRow).Value = yearlyChange
            
            'If true, then calculate the percent change of the stock, check if any open price is 0.
            If openPrice = 0 Then
                percentChange = 0
            Else
                percentChange = yearlyChange / openPrice
            End If
            
            'Print the percent change of the stock to column K
            ws.Range("K" & analysisRow).Value = percentChange
            
            'Format the percent change column to percentage with 2 decimals
            ws.Range("K" & analysisRow).NumberFormat = "0.00%"
           
                            
            'If true, then sum up the trade volume up until that stock cell and print to column J
            stockVolume = stockVolume + ws.Cells(i, 7).Value
            ws.Range("L" & analysisRow).Value = stockVolume
            
            'Continue to the next row in the analysis area
            analysisRow = analysisRow + 1
            
            'Reset the trade volume to start calculating for the new stock
            stockVolume = 0
            
            'Reset the open price
            openPrice = ws.Cells(i + 1, 3).Value
            
        Else
            'If false, keep the open price and close price and only add the stock volume
            stockVolume = stockVolume + ws.Cells(i, 7).Value
        End If
    Next i
    
    'Print the analysis area's headers
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    
    'Fill cells with negative yearly changes in red and positive in green
    'Obtain the number of rows in the analysis area
    lastAnalysisRow = ws.Cells(ws.Rows.Count, 9).End(xlUp).Row
    'Loop through each row in the analysis area
    For i = 2 To lastAnalysisRow
        If ws.Cells(i, 10).Value < 0 Then
            ws.Cells(i, 10).Interior.Color = vbRed
        Else
            ws.Cells(i, 10).Interior.Color = vbGreen
        End If
        
    Next i
    
     
    'Create a variable for the greatest % increase
    Dim maxValue As Double
    maxValue = ws.Cells(2, 11).Value
    
    'Create a variable for the greatest % decrease
    Dim minValue As Double
    minValue = ws.Cells(2, 11).Value
    
    'Create a variable for the greatest total volume
    Dim maxVolume As Double
    maxVolume = ws.Cells(2, 12).Value
    
      
    For i = 2 To lastAnalysisRow
        'Look for the greatest % increase
        If ws.Cells(i, 11).Value >= maxValue Then
            maxValue = ws.Cells(i, 11).Value
            ws.Range("P2").Value = ws.Cells(i, 9).Value
            ws.Range("Q2").Value = ws.Cells(i, 11).Value
        End If
        
        'Look for the greatest % decrease
        If Cells(i, 11).Value <= minValue Then
            minValue = ws.Cells(i, 11).Value
            ws.Range("P3").Value = ws.Cells(i, 9).Value
            ws.Range("Q3").Value = ws.Cells(i, 11).Value
        End If
        
        'Look for the greatest total volume
        If ws.Cells(i, 12).Value >= maxVolume Then
            maxVolume = ws.Cells(i, 12).Value
            ws.Range("P4").Value = ws.Cells(i, 9).Value
            ws.Range("Q4").Value = ws.Cells(i, 12).Value
        End If
       
    Next i
    
    ws.Range("O2").Value = "Greatest % Increase"
    ws.Range("O3").Value = "Greatest % Decrease"
    ws.Range("O4").Value = "Greatest Total Volume"
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"
    
    ws.Range("Q2").NumberFormat = "0.00%"
    ws.Range("Q3").NumberFormat = "0.00%"
    
Next ws

End Sub
