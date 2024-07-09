Sub StockMarketQuarterlyDataAnalysis()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim ticker As String
    Dim openingPrice As Double
    Dim closingPrice As Double
    Dim quarterlyChange As Double
    Dim percentChange As Double
    Dim greatestIncreaseTicker As String
    Dim greatestIncreaseValue As Double
    Dim greatestDecreaseTicker As String
    Dim greatestDecreaseValue As Double
    Dim greatestVolumeTicker As String
    Dim greatestVolumeValue As Double
    Dim Stock_Total As Double
    Dim Summary_Table_Row As Long
    
    ' Loop through each worksheet in the workbook
    For Each ws In ThisWorkbook.Worksheets
        ' Clear previous data in columns I:L and O:Q (if any)
        ws.Range("I:L").ClearContents
        ws.Range("O:Q").ClearContents
        
        ' Initialize variables for this worksheet
        Stock_Total = 0
        Summary_Table_Row = 2
        greatestIncreaseValue = 0
        greatestDecreaseValue = 0
        greatestVolumeValue = 0
        
        ' Reset conditional formatting
        ws.Cells.Interior.ColorIndex = xlNone
        
        ' Find the last row with data in column A
        lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        
        ' Loop through all rows with data in the worksheet
        For i = 2 To lastRow
            ' Check if we are still within the same ticker symbol
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                ' Set the ticker symbol
                ticker = ws.Cells(i, 1).Value
                
                ' Extract opening price (first row for the ticker)
                openingPrice = ws.Cells(Application.WorksheetFunction.Match(ticker, ws.Columns(1), 0), 3).Value
                
                ' Extract closing price (last row for the ticker)
                closingPrice = ws.Cells(i, 6).Value
                
                ' Add to the total stock volume for the ticker
                Stock_Total = Stock_Total + ws.Cells(i, 7).Value
                
                ' Print the ticker symbol in the summary table
                ws.Range("I" & Summary_Table_Row).Value = ticker
                
                ' Calculate the quarterly change and percent change
                quarterlyChange = closingPrice - openingPrice
                If openingPrice <> 0 Then
                    percentChange = ((closingPrice - openingPrice) / openingPrice)
                    ws.Range("J" & Summary_Table_Row).Value = quarterlyChange
                    ws.Range("K" & Summary_Table_Row).Value = Format(percentChange, "0.00%")
                Else
                    percentChange = 0 ' Handle division by zero
                End If
                
                ' Update greatest % increase, % decrease, and total volume
                If percentChange > greatestIncreaseValue Then
                    greatestIncreaseValue = percentChange
                    greatestIncreaseTicker = ticker
                End If
                
                If percentChange < greatestDecreaseValue Then
                    greatestDecreaseValue = percentChange
                    greatestDecreaseTicker = ticker
                End If
                
                If Stock_Total > greatestVolumeValue Then
                    greatestVolumeValue = Stock_Total
                    greatestVolumeTicker = ticker
                End If
                
                ' Print the total stock volume for the ticker in the summary table
                ws.Range("L" & Summary_Table_Row).Value = Stock_Total
                
                ' Apply conditional formatting to column J (Quarterly Change)
                If quarterlyChange < 0 Then
                    ws.Range("J" & Summary_Table_Row).Interior.Color = RGB(255, 0, 0) ' Red
                ElseIf quarterlyChange >= 0 Then
                    ws.Range("J" & Summary_Table_Row).Interior.Color = RGB(0, 255, 0) ' Green
                End If
                
                ' Apply conditional formatting to column K (Percent Change)
                If percentChange < 0 Then
                    ws.Range("K" & Summary_Table_Row).Interior.Color = RGB(255, 0, 0) ' Red
                ElseIf percentChange >= 0 Then
                    ws.Range("K" & Summary_Table_Row).Interior.Color = RGB(0, 255, 0) ' Green
                End If
                
                ' Move to the next row in the summary table
                Summary_Table_Row = Summary_Table_Row + 1
                
                ' Reset the total stock volume for the next ticker
                Stock_Total = 0
            Else
                ' Add to the total stock volume for the current ticker
                Stock_Total = Stock_Total + ws.Cells(i, 7).Value
            End If
        Next i
        
        ' Print headers for the summary table
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Quarterly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        
        ' Print the greatest % increase, % decrease, and total volume in the summary table
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(2, 16).Value = greatestIncreaseTicker
        ws.Cells(2, 17).Value = Format(greatestIncreaseValue, "0.00%")
        
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(3, 16).Value = greatestDecreaseTicker
        ws.Cells(3, 17).Value = Format(greatestDecreaseValue, "0.00%")
        
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        ws.Cells(4, 16).Value = greatestVolumeTicker
        ws.Cells(4, 17).Value = greatestVolumeValue
        
        ' Autofit columns for better visibility
        ws.Columns.AutoFit
    Next ws
End Sub
