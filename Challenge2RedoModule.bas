Attribute VB_Name = "Module2"
Sub StockAnalysis()

    ' Define variables
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim ticker As String
    Dim openingPrice As Double
    Dim closingPrice As Double
    Dim yearlyChange As Double
    Dim percentChange As Double
    Dim totalVolume As Double
    Dim summaryRow As Long
    
    ' Variables for tracking greatest increase, decrease, and volume
    Dim maxIncrease As Double
    Dim maxDecrease As Double
    Dim maxVolume As Double
    Dim maxIncreaseTicker As String
    Dim maxDecreaseTicker As String
    Dim maxVolumeTicker As String
    
    ' Loop through all worksheets
    For Each ws In Worksheets
        ' Find the last row of the worksheet
        lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        ' Set initial values for summary table
        summaryRow = 2
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        
        ' Loop through all rows in the worksheet
        For i = 2 To lastRow
            ' Check if the current row is the first row for a new ticker symbol
            If ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value Then
                ' Get the opening price for the new ticker symbol
                openingPrice = ws.Cells(i, 3).Value
            End If
            
            ' Accumulate the total volume for the ticker symbol
            totalVolume = totalVolume + ws.Cells(i, 7).Value
            
            ' Check if the current row is the last row for the ticker symbol
            If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
                ' Get the closing price for the ticker symbol
                closingPrice = ws.Cells(i, 6).Value
                
                ' Calculate the yearly change and percent change
                yearlyChange = closingPrice - openingPrice
                percentChange = yearlyChange / openingPrice
                
                ' Output the information to the summary table
                ws.Cells(summaryRow, 9).Value = ws.Cells(i, 1).Value
                ws.Cells(summaryRow, 10).Value = yearlyChange
                ws.Cells(summaryRow, 11).Value = percentChange
                ws.Cells(summaryRow, 12).Value = totalVolume
                
                ' Format the percent change as a percentage
                ws.Cells(summaryRow, 11).NumberFormat = "0.00%"
                
                ' Apply conditional formatting to highlight positive change in green and negative change in red
                If yearlyChange > 0 Then
                    ws.Cells(summaryRow, 10).Interior.Color = RGB(0, 255, 0) ' Green
                ElseIf yearlyChange < 0 Then
                    ws.Cells(summaryRow, 10).Interior.Color = RGB(255, 0, 0) ' Red
                End If
                
                ' Update the stock with the greatest increase
                If percentChange > maxIncrease Then
                    maxIncrease = percentChange
                    maxIncreaseTicker = ws.Cells(i, 1).Value
                End If
                
                ' Update the stock with the greatest decrease
                If percentChange < maxDecrease Then
                    maxDecrease = percentChange
                    maxDecreaseTicker = ws.Cells(i, 1).Value
                End If
                
                ' Update the stock with the greatest volume
                If totalVolume > maxVolume Then
                    maxVolume = totalVolume
                    maxVolumeTicker = ws.Cells(i, 1).Value
                End If
                
                ' Reset the variables for the next ticker symbol
                totalVolume = 0
                summaryRow = summaryRow + 1
            End If
        Next i
        
        ' Output the stocks with the greatest increase, decrease, and volume
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(2, 16).Value = maxIncreaseTicker
        ws.Cells(2, 17).Value = maxIncrease
        ws.Cells(2, 17).NumberFormat = "0.00%"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(3, 16).Value = maxDecreaseTicker
        ws.Cells(3, 17).Value = maxDecrease
        ws.Cells(3, 17).NumberFormat = "0.00%"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        ws.Cells(4, 16).Value = maxVolumeTicker
        ws.Cells(4, 17).Value = maxVolume
        
        ' Auto-fit the columns in the summary table
        ws.Columns("I:L").AutoFit
        
      
    Next ws
End Sub

