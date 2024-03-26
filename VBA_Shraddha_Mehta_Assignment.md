Sub stocks()
    Dim ws As Worksheet
    Dim ticker As String
    Dim yearlychange As Double
    Dim percentchange As Double
    Dim lastrow As Long
    Dim openprice As Double
    Dim closeprice As Double
    Dim totalvolume As Double
    Dim outcome As Long
    Dim maxPercentIncrease As Double
    Dim maxPercentDecrease As Double
    Dim maxTotalVolume As Double
    Dim maxPercentIncreaseTicker As String
    Dim maxPercentDecreaseTicker As String
    Dim maxTotalVolumeTicker As String

    maxPercentIncrease = 0
    maxPercentDecrease = 0
    maxTotalVolume = 0

    For Each ws In ThisWorkbook.Worksheets
        ws.Range("i1").Value = "Ticker"
        ws.Range("j1").Value = "Yearly Change"
        ws.Range("k1").Value = "Percent Change"
        ws.Range("l1").Value = "Total Volume"

        lastrow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        outcome = 2
        ticker = ws.Cells(2, 1).Value
        totalvolume = 0

        For i = 2 To lastrow
            If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Or i = lastrow Then
                closeprice = ws.Cells(i, 6).Value
                openprice = ws.Cells(outcome, 3).Value
                yearlychange = closeprice - openprice

                If openprice <> 0 Then
                    percentchange = (yearlychange / openprice) * 100
                Else
                    percentchange = 0
                End If

                ' Output the values to the worksheet
                ws.Cells(outcome, 9).Value = ticker
                ws.Cells(outcome, 10).Value = yearlychange
                ws.Cells(outcome, 11).Value = percentchange
                ws.Cells(outcome, 12).Value = totalvolume

                ' Apply conditional formatting based on yearly change
                If yearlychange > 0 Then
                    ws.Cells(outcome, 10).Interior.Color = RGB(0, 255, 0) ' Green color
                Else
                    ws.Cells(outcome, 10).Interior.Color = RGB(255, 0, 0) ' Red color
                End If

                ' Check and update the greatest percentage increase, decrease, and total volume
                If percentchange > maxPercentIncrease Then
                    maxPercentIncrease = percentchange
                    maxPercentIncreaseTicker = ticker
                End If

                If percentchange < maxPercentDecrease Then
                    maxPercentDecrease = percentchange
                    maxPercentDecreaseTicker = ticker
                End If

                If totalvolume > maxTotalVolume Then
                    maxTotalVolume = totalvolume
                    maxTotalVolumeTicker = ticker
                End If

                ' Reset variables for next ticker
                outcome = outcome + 1
                ticker = ws.Cells(i + 1, 1).Value
                totalvolume = 0
            Else
                totalvolume = totalvolume + ws.Cells(i, 7).Value
            End If
        Next i
    Next ws

    ' Output the greatest percentage increase, decrease, and total volume along with the corresponding tickers
    For Each ws In ThisWorkbook.Worksheets
        ws.Cells(2, 16).Value = "Greatest % Increase"
        ws.Cells(2, 17).Value = maxPercentIncreaseTicker
        ws.Cells(2, 18).Value = maxPercentIncrease

        ws.Cells(3, 16).Value = "Greatest % Decrease"
        ws.Cells(3, 17).Value = maxPercentDecreaseTicker
        ws.Cells(3, 18).Value = maxPercentDecrease

        ws.Cells(4, 16).Value = "Greatest Total Volume"
        ws.Cells(4, 17).Value = maxTotalVolumeTicker
        ws.Cells(4, 18).Value = maxTotalVolume
    Next ws
End Sub
