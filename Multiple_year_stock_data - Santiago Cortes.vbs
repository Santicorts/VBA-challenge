Sub multiple_year_stock_all_tabs()

    Dim ws As Worksheet
    
    ' Loop through each worksheet in the workbook
    For Each ws In ThisWorkbook.Worksheets
        ' Call the main function to process data for each worksheet
        ProcessData ws
    Next ws

End Sub

Sub ProcessData(ws As Worksheet)

    Dim i As Long
    Dim j As Long
    Dim ticker As String
    Dim openValue As Double ' Variable to store the open value for each ticker
    Dim closeValue As Double ' Variable to store the close value for each ticker
    Dim lastVolume As Double ' Variable to store the last volume value for each ticker
    Dim lastRow As Long
    Dim change As Double ' Variable to store the quarter change
    Dim percentChange As Double ' Variable to store the percent change
    Dim greatestWinner As String ' Variable to store the ticker of the greatest winner
    Dim greatestLoser As String ' Variable to store the ticker of the greatest loser
    Dim greatestMarketCap As String ' Variable to store the ticker of the greatest market cap
    Dim maxPercentChange As Double ' Variable to store the maximum percent change
    Dim minPercentChange As Double ' Variable to store the minimum percent change
    Dim maxVolume As Double ' Variable to store the maximum volume
    
    ' Find the last row in column A1
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    ' Initialize the row index for the summary table
    j = 2
    
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Quarter change"
    ws.Cells(1, 11).Value = "Percent change"
    ws.Cells(1, 12).Value = "Total market cap"
    
    ws.Cells(2, 15).Value = "Greatest winner"
    ws.Cells(3, 15).Value = "Greatest loser"
    ws.Cells(4, 15).Value = "Greatest market Cap"
    
    ' Loop through cells in Column A to find unique values
    For i = 2 To lastRow
        ' Check if the current ticker is different from the previous one
        If ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value Then
            ' If a new ticker is encountered, store the last close value, last total volume, and calculate quarter change and percent change for the previous ticker
            If j > 2 Then
                ' Print the last total volume (total volume for the previous ticker) in column L (12)
                ws.Cells(j - 1, 12).Value = lastVolume
                ' Calculate quarter change and percent change
                change = closeValue - openValue
                percentChange = (closeValue - openValue) / openValue
                ' Print quarter change and percent change in columns I and J (9 and 10)
                ws.Cells(j - 1, 10).Value = change
                ws.Cells(j - 1, 11).Value = percentChange
                
                ' Store the ticker with the greatest winner
                If percentChange > maxPercentChange Then
                    maxPercentChange = percentChange
                    greatestWinner = ticker
                End If
                
                ' Store the ticker with the greatest loser
                If percentChange < minPercentChange Then
                    minPercentChange = percentChange
                    greatestLoser = ticker
                End If
                
                ' Store the ticker with the greatest market cap
                If lastVolume > maxVolume Then
                    maxVolume = lastVolume
                    greatestMarketCap = ticker
                End If
            End If
            ' Print the new ticker in column I (9)
            ticker = ws.Cells(i, 1).Value
            ws.Cells(j, 9).Value = ticker
            ' Store the open value for the new ticker
            openValue = ws.Cells(i, 3).Value
            ' Reset the close value and last volume for the new ticker
            closeValue = ws.Cells(i, 6).Value
            lastVolume = ws.Cells(i, 7).Value
            ' Move to the next row in the summary table
            j = j + 1
        Else
            ' If the current ticker is the same as the previous one, update the close value and last volume
            closeValue = ws.Cells(i, 6).Value
            lastVolume = ws.Cells(i, 7).Value
        End If
    Next i
    
    ' Store the last total volume for the last ticker encountered
    ws.Cells(j - 1, 12).Value = lastVolume
    ' Calculate quarter change and percent change for the last ticker
    change = closeValue - openValue
    percentChange = (closeValue - openValue) / openValue
    ' Print quarter change and percent change for the last ticker in columns I and J (9 and 10)
    ws.Cells(j - 1, 10).Value = change
    ws.Cells(j - 1, 11).Value = percentChange
    
    ' Print the greatest winner, greatest loser, and greatest market cap in columns N and P (14 and 16)
    ws.Cells(2, 14).Value = greatestWinner
    ws.Cells(3, 14).Value = greatestLoser
    ws.Cells(4, 14).Value = greatestMarketCap
    ws.Cells(2, 16).Value = maxPercentChange
    ws.Cells(3, 16).Value = minPercentChange
    ws.Cells(4, 16).Value = maxVolume

End Sub

