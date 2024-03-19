Sub Stock_Data_Part_II()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim ticker As String
    Dim yearlyChange As Double
    Dim percentageChange As Double
    Dim totalVolume As Double
    Dim maxPercentageIncrease As Double
    Dim minPercentageDecrease As Double
    Dim maxTotalVolume As Double
    Dim maxPercentageIncreaseTicker As String
    Dim minPercentageDecreaseTicker As String
    Dim maxTotalVolumeTicker As String
    Dim i As Long
    Dim summaryRow As Long
    Dim summaryColumn As Long
    
    ' Loop through each sheet
    For Each ws In ThisWorkbook.Sheets
        ' Find the last row in the worksheet
        lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
        
        ' Initialize variables to track maximum values for each sheet
        maxPercentageIncrease = 0
        minPercentageDecrease = 0
        maxTotalVolume = 0
        maxPercentageIncreaseTicker = ""
        minPercentageDecreaseTicker = ""
        maxTotalVolumeTicker = ""
        
        ' Find the next available column for summary
        summaryColumn = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column + 2
        
        ' Create headers for the summary table on the current worksheet
        ws.Cells(1, summaryColumn).Value = "Metric"
        ws.Cells(1, summaryColumn + 1).Value = "Ticker"
        ws.Cells(1, summaryColumn + 2).Value = "Value"
        summaryRow = 2
        
        ' Loop through each row in the worksheet
        For i = 2 To lastRow
            ' Assign the ticker symbol
            ticker = ws.Cells(i, 9).Value
                
            ' Assign the yearly change
            yearlyChange = ws.Cells(i, 10).Value
                
            ' Assign the percentage change
            percentageChange = ws.Cells(i, 11).Value
                
            ' Assign the total stock volume
            totalVolume = ws.Cells(i, 12).Value
                
            ' Check for maximum percentage increase
            If percentageChange > maxPercentageIncrease Then
                maxPercentageIncrease = percentageChange
                maxPercentageIncreaseTicker = ticker
            End If
                
            ' Check for minimum percentage decrease
            If minPercentageDecrease = 0 Or percentageChange < minPercentageDecrease Then
                minPercentageDecrease = percentageChange
                minPercentageDecreaseTicker = ticker
            End If
                
            ' Check for maximum total volume
            If totalVolume > maxTotalVolume Then
                maxTotalVolume = totalVolume
                maxTotalVolumeTicker = ticker
            End If
        Next i
        
        ' Populate the summary table on the current worksheet
        ws.Cells(summaryRow, summaryColumn).Value = "Greatest % Increase"
        ws.Cells(summaryRow + 1, summaryColumn).Value = "Greatest % Decrease"
        ws.Cells(summaryRow + 2, summaryColumn).Value = "Greatest Total Volume"
        
        ws.Cells(summaryRow, summaryColumn + 1).Value = maxPercentageIncreaseTicker
        ws.Cells(summaryRow + 1, summaryColumn + 1).Value = minPercentageDecreaseTicker
        ws.Cells(summaryRow + 2, summaryColumn + 1).Value = maxTotalVolumeTicker
        
        ws.Cells(summaryRow, summaryColumn + 2).Value = maxPercentageIncrease * 100 & "%" ' Append "%"
        ws.Cells(summaryRow + 1, summaryColumn + 2).Value = minPercentageDecrease * 100 & "%" ' Append "%"
        ws.Cells(summaryRow + 2, summaryColumn + 2).Value = maxTotalVolume * 100 & "%" ' Append "%"
    Next ws

End Sub

