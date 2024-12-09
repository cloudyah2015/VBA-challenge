Sub AnalyzeQuarterlyStocks()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim currentRow As Long
    Dim ticker As String
    Dim openPrice As Double
    Dim closePrice As Double
    Dim quarterlyChange As Double
    Dim percentChange As Double
    Dim totalVolume As Double
    Dim outputRow As Long
    Dim maxIncrease As Double, maxDecrease As Double, maxVolume As Double
    Dim tickerMaxIncrease As String, tickerMaxDecrease As String, tickerMaxVolume As String
    
    ' Loop through all worksheets in the workbook
    For Each ws In ThisWorkbook.Worksheets
        With ws
            ' Find the last row of data
            lastRow = .Cells(.Rows.Count, 1).End(xlUp).Row
            currentRow = 2
            outputRow = 2
            
            ' Initialize variables to track maximums
            maxIncrease = -999999
            maxDecrease = 999999
            maxVolume = 0
            
            Do While currentRow <= lastRow
                ticker = .Cells(currentRow, 1).Value
                openPrice = .Cells(currentRow, 3).Value ' Opening price
                totalVolume = 0
                
                ' Loop through the rows for the same ticker
                Do While .Cells(currentRow, 1).Value = ticker
                    totalVolume = totalVolume + .Cells(currentRow, 7).Value ' Volume
                    closePrice = .Cells(currentRow, 6).Value ' Closing price
                    currentRow = currentRow + 1
                    If currentRow > lastRow Then Exit Do
                Loop
                
                ' Calculate the quarterly change and percentage change
                quarterlyChange = closePrice - openPrice
                If openPrice <> 0 Then
                    percentChange = (quarterlyChange / openPrice) * 100
                Else
                    percentChange = 0
                End If
                
                ' Output the results
                .Cells(outputRow, 9).Value = ticker
                .Cells(outputRow, 10).Value = quarterlyChange
                .Cells(outputRow, 11).Value = percentChange
                .Cells(outputRow, 12).Value = totalVolume
                
                ' Apply formatting
                If quarterlyChange > 0 Then
                    .Cells(outputRow, 10).Interior.Color = RGB(144, 238, 144) ' Light green
                Else
                    .Cells(outputRow, 10).Interior.Color = RGB(255, 182, 193) ' Light red
                End If
                
                ' Track maximum values
                If percentChange > maxIncrease Then
                    maxIncrease = percentChange
                    tickerMaxIncrease = ticker
                End If
                
                If percentChange < maxDecrease Then
                    maxDecrease = percentChange
                    tickerMaxDecrease = ticker
                End If
                
                If totalVolume > maxVolume Then
                    maxVolume = totalVolume
                    tickerMaxVolume = ticker
                End If
                
                outputRow = outputRow + 1
            Loop
            
            ' Output the maximums for the sheet
            .Cells(1, 15).Value = "Metric"
            .Cells(1, 16).Value = "Ticker"
            .Cells(1, 17).Value = "Value"
            
            .Cells(2, 15).Value = "Greatest % Increase"
            .Cells(2, 16).Value = tickerMaxIncrease
            .Cells(2, 17).Value = maxIncrease
            
            .Cells(3, 15).Value = "Greatest % Decrease"
            .Cells(3, 16).Value = tickerMaxDecrease
            .Cells(3, 17).Value = maxDecrease
            
            .Cells(4, 15).Value = "Greatest Total Volume"
            .Cells(4, 16).Value = tickerMaxVolume
            .Cells(4, 17).Value = maxVolume
        End With
    Next ws
    
    MsgBox "Quarterly Stock Analysis Completed! Yay!"
End Sub
