Attribute VB_Name = "Module1"
Sub ProcessAllStocks()

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim rowNum As Long
    Dim ticker As String
    Dim openPrice As Double
    Dim closePrice As Double
    Dim yearlyChange As Double
    Dim percentageChange As Double
    Dim totalVolume As Double
    
    ' Initialize variables to track the maximum and minimum values
    Dim maxPercentageIncrease As Double: maxPercentageIncrease = 0
    Dim maxPercentageIncreaseTicker As String
    Dim minPercentageDecrease As Double: minPercentageDecrease = 0
    Dim minPercentageDecreaseTicker As String
    Dim maxTotalVolume As Double: maxTotalVolume = 0
    Dim maxTotalVolumeTicker As String
    Dim outputRow As Long: outputRow = 2
    
    ' Loop through each worksheet in the workbook
    For Each ws In ThisWorkbook.Sheets
        lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        totalVolume = 0
        openPrice = ws.Cells(2, 3).Value ' Initialize with the first open price of the sheet
        
        ' Loop through each row in the current worksheet
        For rowNum = 2 To lastRow
            ' Check if we are still within the same ticker
            If ws.Cells(rowNum, 1).Value <> ws.Cells(rowNum + 1, 1).Value Then
                ' Calculate the close price, yearly change, percentage change, and total volume
                ticker = ws.Cells(rowNum, 1).Value
                closePrice = ws.Cells(rowNum, 6).Value
                yearlyChange = closePrice - openPrice
                percentageChange = (yearlyChange / openPrice) * 100
                totalVolume = totalVolume + ws.Cells(rowNum, 7).Value
                
                ' Write the summary output for the current ticker
                ws.Cells(outputRow, 8).Value = yearlyChange
                ws.Cells(outputRow, 9).Value = percentageChange
                
                ' Apply conditional formatting to the yearly change
                If yearlyChange > 0 Then
                    ws.Cells(outputRow, 8).Interior.Color = RGB(0, 255, 0)
                ElseIf yearlyChange < 0 Then
                    ws.Cells(outputRow, 8).Interior.Color = RGB(255, 0, 0)
                End If
                
                outputRow = outputRow + 1
                
                ' Check for max and min values
                If percentageChange > maxPercentageIncrease Then
                    maxPercentageIncrease = percentageChange
                    maxPercentageIncreaseTicker = ticker
                End If
                
                If percentageChange < minPercentageDecrease Then
                    minPercentageDecrease = percentageChange
                    minPercentageDecreaseTicker = ticker
                End If
                
                If totalVolume > maxTotalVolume Then
                    maxTotalVolume = totalVolume
                    maxTotalVolumeTicker = ticker
                End If
                
                ' Reset the open price for the next ticker
                If rowNum + 1 <= lastRow Then
                    openPrice = ws.Cells(rowNum + 1, 3).Value
                End If
                
                ' Reset totalVolume for the next ticker
                totalVolume = 0
                
            Else
                ' Add to totalVolume if we are within the same ticker
                totalVolume = totalVolume + ws.Cells(rowNum, 7).Value
            End If
        Next rowNum
        
        ' Display the calculated max/min values in the current worksheet
        ws.Cells(1, 10).Value = "Greatest % Increase: " & maxPercentageIncreaseTicker & " (" & maxPercentageIncrease & "%)"
        ws.Cells(2, 10).Value = "Greatest % Decrease: " & minPercentageDecreaseTicker & " (" & minPercentageDecrease & "%)"
        ws.Cells(3, 10).Value = "Greatest Total Volume: " & maxTotalVolumeTicker & " (" & maxTotalVolume & ")"
        
        ' Reset outputRow for the next worksheet
        outputRow = 2
    Next ws

End Sub
