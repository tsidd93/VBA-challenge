Sub CalculateQuarterlyData()

    Dim ws As Worksheet
    Dim lastrow As Long
    Dim openprice As Double
    Dim closeprice As Double
    Dim totalvolume As Double
    Dim i As Long
    Dim j As Long
    Dim ticker As String
    Dim prev_ticker As String
    Dim openpricecol As Long
    Dim closepricecol As Long
    Dim volumecol As Long
    Dim tickercol As Long
    
    ' Variables to track max values
    Dim maxIncrease As Double
    Dim maxDecrease As Double
    Dim maxVolume As Double
    Dim maxIncreaseTicker As String
    Dim maxDecreaseTicker As String
    Dim maxVolumeTicker As String
    
    For Each ws In Worksheets

        ' Define column locations
        tickercol = 1         ' Column A for tickers
        openpricecol = 3      ' Column C for open prices
        closepricecol = 6     ' Column F for close prices
        volumecol = 7         ' Column G for volume
        
        ' Find the last row in the worksheet
        lastrow = ws.Cells(ws.Rows.Count, tickercol).End(xlUp).Row
        
        ' Set headers for the output
        ws.Cells(1, 8).Value = "Ticker"
        ws.Cells(1, 9).Value = "Quarterly Change"
        ws.Cells(1, 10).Value = "Percent Change"
        ws.Cells(1, 11).Value = "Total Stock Volume"
        
        ' Set headers for greatest increase/decrease and total volume
        ws.Cells(2, 13).Value = "Greatest Percent Increase"
        ws.Cells(3, 13).Value = "Greatest Percent Decrease"
        ws.Cells(4, 13).Value = "Greatest Total Volume"
        
        ws.Cells(1, 14).Value = "Ticker"
        ws.Cells(1, 15).Value = "Value"
        
        ' Initialize variables for the first ticker
        j = 2 ' Start output from row 2
        prev_ticker = ws.Cells(2, tickercol).Value
        openprice = ws.Cells(2, openpricecol).Value
        totalvolume = 0
        
        ' Initialize max tracking variables
        maxIncrease = -99999
        maxDecrease = 99999
        maxVolume = 0
        
        ' Loop through the data rows
        For i = 2 To lastrow
            ticker = ws.Cells(i, tickercol).Value
            
            ' If a new ticker is found or the loop reaches the end of the data
            If ticker <> prev_ticker Then
                ' Calculate close price for the last occurrence of the previous ticker
                closeprice = ws.Cells(i - 1, closepricecol).Value
                
                ' Calculate quarterly and percent change
                Dim quarterlyChange As Double
                Dim percentChange As Double
                quarterlyChange = closeprice - openprice
                If openprice <> 0 Then
                    percentChange = ((closeprice - openprice) / openprice) * 100
                Else
                    percentChange = 0
                End If
                
                ' Output the result for the previous ticker
                ws.Cells(j, 8).Value = prev_ticker
                ws.Cells(j, 9).Value = quarterlyChange
                ws.Cells(j, 10).Value = percentChange
                ws.Cells(j, 11).Value = totalvolume
                
                ' Track max increase/decrease and volume
                If percentChange > maxIncrease Then
                    maxIncrease = percentChange
                    maxIncreaseTicker = prev_ticker
                End If
                If percentChange < maxDecrease Then
                    maxDecrease = percentChange
                    maxDecreaseTicker = prev_ticker
                End If
                If totalvolume > maxVolume Then
                    maxVolume = totalvolume
                    maxVolumeTicker = prev_ticker
                End If
                
                ' Reset for the new ticker
                j = j + 1
                prev_ticker = ticker
                openprice = ws.Cells(i, openpricecol).Value
                totalvolume = ws.Cells(i, volumecol).Value
            Else
                ' Accumulate volume for the same ticker
                totalvolume = totalvolume + ws.Cells(i, volumecol).Value
            End If
        Next i
        
        ' Handle the last ticker
        totalvolume = totalvolume + ws.Cells(lastrow, volumecol).Value ' Add the last row's volume
        closeprice = ws.Cells(lastrow, closepricecol).Value
        Dim lastQuarterlyChange As Double
        Dim lastPercentChange As Double
        lastQuarterlyChange = closeprice - openprice
        If openprice <> 0 Then
            lastPercentChange = ((closeprice - openprice) / openprice) * 100
        Else
            lastPercentChange = 0
        End If
        ws.Cells(j, 8).Value = prev_ticker
        ws.Cells(j, 9).Value = lastQuarterlyChange
        ws.Cells(j, 10).Value = lastPercentChange
        ws.Cells(j, 11).Value = totalvolume
        
        ' Final check for max values for the last ticker
        If lastPercentChange > maxIncrease Then
            maxIncrease = lastPercentChange
            maxIncreaseTicker = prev_ticker
        End If
        If lastPercentChange < maxDecrease Then
            maxDecrease = lastPercentChange
            maxDecreaseTicker = prev_ticker
        End If
        If totalvolume > maxVolume Then
            maxVolume = totalvolume
            maxVolumeTicker = prev_ticker
        End If
        
        ' Output the tracked max values in the new format
        ws.Cells(2, 14).Value = maxIncreaseTicker
        ws.Cells(2, 15).Value = maxIncrease
        ws.Cells(3, 14).Value = maxDecreaseTicker
        ws.Cells(3, 15).Value = maxDecrease
        ws.Cells(4, 14).Value = maxVolumeTicker
        ws.Cells(4, 15).Value = maxVolume
        
    Next ws
    
    MsgBox "Quarterly Calculation Complete"
End Sub

