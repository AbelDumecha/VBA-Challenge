# VBA-Challenge
Sub Stock_data_Part_I()
 
   Dim ws As Worksheet
    Dim lastRow As Long
    Dim ticker As String
    Dim openingPrice As Double
    Dim closingPrice As Double
    Dim yearlyChange As Double
    Dim percentageChange As Double
    Dim totalVolume As Double
    Dim summaryRow As Long
    Dim coloredYearlyChange As Range
    Dim cell As Range
    
    ' Loop through each sheet
    For Each ws In ThisWorkbook.Sheets
        ' Find the last row in the worksheet
        lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
        
        ' Initialize summary table headers
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percentage Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        
        ' Initialize summary row
        summaryRow = 2
        
        ' Loop through each row in the worksheet
        For i = 2 To lastRow
            ' Check if we are still within the same ticker
            If ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value Or i = 2 Then
                ' Assign the ticker symbol
                ticker = ws.Cells(i, 1).Value
                
                ' Assign the opening price
                openingPrice = ws.Cells(i, 3).Value
            End If
            
            ' Check if we are at the end of the ticker or at the last row
            If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Or i = lastRow Then
                ' Assign the closing price
                closingPrice = ws.Cells(i, 6).Value
                
                ' Calculate yearly change
                yearlyChange = closingPrice - openingPrice
                
                ' Calculate percentage change
                If openingPrice <> 0 Then
                    percentageChange = (yearlyChange / openingPrice) * 100
                Else
                    percentageChange = 0
                End If
                
                ' Calculate total stock volume
                totalVolume = Application.Sum(ws.Range(ws.Cells(i, 7), ws.Cells(summaryRow, 7)))
                
                ' Output results to summary table
                ws.Cells(summaryRow, 9).Value = ticker
                ws.Cells(summaryRow, 10).Value = yearlyChange
                ws.Cells(summaryRow, 11).Value = percentageChange
                ws.Cells(summaryRow, 12).Value = totalVolume
                
                ' Move to the next row in the summary table
                summaryRow = summaryRow + 1
                
            End If
        Next i
            ' Set range for yearly change column
        Set coloredYearlyChange = ws.Range(ws.Cells(2, 10), ws.Cells(lastRow, 10))
        
        ' Apply conditional formatting to highlight positive and negative changes
        For Each cell In coloredYearlyChange
            If cell.Value >= 0 Then
                cell.Interior.Color = RGB(0, 0, 255)
            Else
                cell.Interior.Color = RGB(255, 165, 0)
            End If
        Next cell
      
    Next ws
End Sub
