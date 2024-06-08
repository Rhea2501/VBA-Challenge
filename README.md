# VBA-Challenge
Week 2 - VBA 
Used Stack Overflow for this part of the code: Xpert and Chat GPT 4.o for this part of the code: 

   
"    outputRow = 2

    ' Loop through each worksheet
    For Each ws In Worksheets
        If ws.Name <> "Summary_Data" Then
        
            ' Find the last row of data
            lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

            ' Start from 1st row with variables
            startRow = 2
            lastTicker = ws.Cells(startRow, 1).Value
            openPrice = ws.Cells(startRow, 3).Value
            totalStockVolume = 0

            ' Loop through each row of data
            For i = 2 To lastRow
                ticker = ws.Cells(i, 1).Value

                ' Check if we have a new ticker
                If ticker <> lastTicker Then
                    ' Calculate the values for the previous ticker
                    closePrice = ws.Cells(i - 1, 6).Value
                    quarterlyChange = closePrice - openPrice
                    If openPrice <> 0 Then
                        percentageChange = (quarterlyChange / openPrice) * 100
                    Else
                        percentageChange = 0
                    End If

                    ' Output the results to the summary sheet
                    wsSummary.Cells(outputRow, 1).Value = lastTicker
                    wsSummary.Cells(outputRow, 2).Value = quarterlyChange
                    wsSummary.Cells(outputRow, 3).Value = percentageChange
                    wsSummary.Cells(outputRow, 4).Value = totalStockVolume
                    outputRow = outputRow + 1

                    ' Reset for the new ticker
                    startRow = i
                    openPrice = ws.Cells(startRow, 3).Value
                    totalStockVolume = 0
                End If

                ' Add to the total volume
                totalStockVolume = totalStockVolume + ws.Cells(i, 7).Value

                ' Update last ticker
                lastTicker = ticker
            Next i

            ' Handle the last group
            closePrice = ws.Cells(lastRow, 6).Value
            quarterlyChange = closePrice - openPrice
            If openPrice <> 0 Then
                percentageChange = (quarterlyChange / openPrice) * 100
            Else
                percentageChange = 0
            End If

            wsSummary.Cells(outputRow, 1).Value = lastTicker
            wsSummary.Cells(outputRow, 2).Value = quarterlyChange
            wsSummary.Cells(outputRow, 3).Value = percentageChange
            wsSummary.Cells(outputRow, 4).Value = totalStockVolume
        End If
    Next ws
End Sub"
