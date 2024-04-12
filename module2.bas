Attribute VB_Name = "Module1"
Sub StockAnalysis()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim ticker As String
    Dim openingPrice As Double
    Dim closingPrice As Double
    Dim yearlyChange As Double
    Dim percentageChange As Double
    Dim totalVolume As Double
    Dim outputRow As Long
    Dim maxPercentageIncrease As Double
    Dim maxPercentageDecrease As Double
    Dim maxTotalVolume As Double
    Dim maxPercentageIncreaseTicker As String
    Dim maxPercentageDecreaseTicker As String
    Dim maxTotalVolumeTicker As String
    
    For Each ws In Worksheets
        'This section finds the last row
        lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
        
        'The initial output row
        outputRow = 2
        
        'Header row for the output
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percentage Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        
        'Variables are inputed
        ticker = ws.Cells(2, 1).Value
        openingPrice = ws.Cells(2, 3).Value
        totalVolume = ws.Cells(2, 7).Value
        maxPercentageIncrease = 0
        maxPercentageDecrease = 0
        maxTotalVolume = 0
        maxPercentageIncreaseTicker = ""
        maxPercentageDecreaseTicker = ""
        maxTotalVolumeTicker = ""
        
        
        For i = 2 To lastRow
            'Check if it's a new ticker
            If ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value Then
                'Outputs the previous ticker data
                ws.Cells(outputRow, 9).Value = ticker
                ws.Cells(outputRow, 10).Value = yearlyChange
                If openingPrice <> 0 Then
                    ws.Cells(outputRow, 11).Value = percentageChange
                Else
                    ws.Cells(outputRow, 11).Value = "N/A"
                End If
                ws.Cells(outputRow, 12).Value = totalVolume
                
                'The greatest percentage increase
                If percentageChange > maxPercentageIncrease Then
                    maxPercentageIncrease = percentageChange
                    maxPercentageIncreaseTicker = ticker
                End If
                
                'If statement to calculate the greatest percentage decrease
                If percentageChange < maxPercentageDecrease Then
                    maxPercentageDecrease = percentageChange
                    maxPercentageDecreaseTicker = ticker
                End If
                
                'Next if statement to calculate the greatest total volume
                If totalVolume > maxTotalVolume Then
                    maxTotalVolume = totalVolume
                    maxTotalVolumeTicker = ticker
                End If
                
                '+1 moves to the next row for output
                outputRow = outputRow + 1
                
                'We are resetting variables for the next ticker
                ticker = ws.Cells(i, 1).Value
                openingPrice = ws.Cells(i, 3).Value
                totalVolume = ws.Cells(i, 7).Value
            End If
            
            'Updates closing price for each ticker
            closingPrice = ws.Cells(i, 6).Value
            
            'Updates total volume
            totalVolume = totalVolume + ws.Cells(i, 7).Value
            
            'Calculates yearly change
            yearlyChange = closingPrice - openingPrice
            
            'Calculates percentage change where openPrice is not equal to zero
            If openingPrice <> 0 Then
                percentageChange = (yearlyChange / openingPrice) * 100
            Else
                percentageChange = 0
            End If
            
           
        Next i
        
        'Output for the last ticker data
        ws.Cells(outputRow, 9).Value = ticker
        ws.Cells(outputRow, 10).Value = yearlyChange
        If openingPrice <> 0 Then
            ws.Cells(outputRow, 11).Value = percentageChange
        Else
            ws.Cells(outputRow, 11).Value = "N/A"
        End If
        ws.Cells(outputRow, 12).Value = totalVolume
        
        'Output for the stock with greatest % increase, greatest % decrease, and greatest total volume
        ws.Cells(2, 16).Value = maxPercentageIncreaseTicker
        ws.Cells(2, 17).Value = maxPercentageIncrease & "%"
        ws.Cells(3, 16).Value = maxPercentageDecreaseTicker
        ws.Cells(3, 17).Value = maxPercentageDecrease & "%"
        ws.Cells(4, 16).Value = maxTotalVolumeTicker
        ws.Cells(4, 17).Value = maxTotalVolume
        
        ' Conditional formatting
        With ws.Range(ws.Cells(2, 10), ws.Cells(lastRow, 10))
            .FormatConditions.Delete
            ' Add new formatting for positive and negative numbers
            .FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, Formula1:="0"
            .FormatConditions(1).Interior.Color = RGB(0, 255, 0) ' Green
            .FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, Formula1:="0"
            .FormatConditions(2).Interior.Color = RGB(255, 0, 0) ' Red
        End With
    Next ws
End Sub

