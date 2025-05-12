Sub AnalyzeStockData():

    ' Set dimensions
    Dim stockVolume As Double
    Dim currentRow As Long
    Dim quarterlyChange As Double
    Dim summaryRow As Integer
    Dim tickerStartRow As Long
    Dim lastDataRow As Long
    Dim percentDifference As Double
    Dim tradingDays As Integer
    Dim dailyDifference As Double
    Dim averageDifference As Double

    ' Set title row
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Quarterly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Stock Volume"
    Range("P1").Value = "Ticker"
    Range("Q1").Value = "Value"
    Range("O2").Value = "Greatest % Increase"
    Range("O3").Value = "Greatest % Decrease"
    Range("O4").Value = "Greatest Total Volume"

    ' Set initial values
    summaryRow = 0
    stockVolume = 0
    quarterlyChange = 0
    tickerStartRow = 2

    ' get the row number of the last row with data
    lastDataRow = Cells(Rows.Count, "A").End(xlUp).Row

    For currentRow = 2 To lastDataRow

        ' If ticker changes then print results
        If Cells(currentRow + 1, 1).Value <> Cells(currentRow, 1).Value Then

            ' Stores results in variables
            stockVolume = stockVolume + Cells(currentRow, 7).Value

            ' Handle zero stockVolume volume
            If stockVolume = 0 Then
                ' print the results
                Range("I" & 2 + summaryRow).Value = Cells(currentRow, 1).Value
                Range("J" & 2 + summaryRow).Value = 0
                Range("K" & 2 + summaryRow).Value = "%" & 0
                Range("L" & 2 + summaryRow).Value = 0

            Else
                ' Find First non zero starting value
                If Cells(tickerStartRow, 3) = 0 Then
                    For find_value = tickerStartRow To currentRow
                        If Cells(find_value, 3).Value <> 0 Then
                            tickerStartRow = find_value
                            Exit For
                        End If
                     Next find_value
                End If

                ' Calculate Change
                quarterlyChange = (Cells(currentRow, 6) - Cells(tickerStartRow, 3))
                percentDifference = quarterlyChange / Cells(tickerStartRow, 3)

                ' tickerStartRow of the next stock ticker
                tickerStartRow = currentRow + 1

                ' print the results
                Range("I" & 2 + summaryRow).Value = Cells(currentRow, 1).Value
                Range("J" & 2 + summaryRow).Value = quarterlyChange
                Range("J" & 2 + summaryRow).NumberFormat = "0.00"
                Range("K" & 2 + summaryRow).Value = percentDifference
                Range("K" & 2 + summaryRow).NumberFormat = "0.00%"
                Range("L" & 2 + summaryRow).Value = stockVolume

                ' colors positives green and negatives red
                Select Case quarterlyChange
                    Case Is > 0
                        Range("J" & 2 + summaryRow).Interior.ColorIndex = 4
                    Case Is < 0
                        Range("J" & 2 + summaryRow).Interior.ColorIndex = 3
                    Case Else
                        Range("J" & 2 + summaryRow).Interior.ColorIndex = 0
                End Select

            End If

            ' reset variables for new stock ticker
            stockVolume = 0
            quarterlyChange = 0
            summaryRow = summaryRow + 1
            tradingDays = 0

        ' If ticker is still the same add results
        Else
            stockVolume = stockVolume + Cells(currentRow, 7).Value

        End If

    Next currentRow

    ' take the max and min and place them in a separate part in the worksheet
    Range("Q2") = "%" & WorksheetFunction.Max(Range("K2:K" & lastDataRow)) * 100
    Range("Q3") = "%" & WorksheetFunction.Min(Range("K2:K" & lastDataRow)) * 100
    Range("Q4") = WorksheetFunction.Max(Range("L2:L" & lastDataRow))

    ' returns one less because header row not a factor
    increase_number = WorksheetFunction.Match(WorksheetFunction.Max(Range("K2:K" & lastDataRow)), Range("K2:K" & lastDataRow), 0)
    decrease_number = WorksheetFunction.Match(WorksheetFunction.Min(Range("K2:K" & lastDataRow)), Range("K2:K" & lastDataRow), 0)
    volume_number = WorksheetFunction.Match(WorksheetFunction.Max(Range("L2:L" & lastDataRow)), Range("L2:L" & lastDataRow), 0)

    ' final ticker symbol for  stockVolume, greatest % of increase and decrease, and average
    Range("P2") = Cells(increase_number + 1, 9)
    Range("P3") = Cells(decrease_number + 1, 9)
    Range("P4") = Cells(volume_number + 1, 9)

End Sub
