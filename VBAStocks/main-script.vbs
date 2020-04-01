Sub loopTheYears()
    Dim rowLimit As Long
    Dim row As Long
    Dim rowIndex As Integer
    Dim openYear As Variant
    Dim closeYear As Variant
    Dim change As Variant
    Dim totalStock As Variant
    Dim greatestIncrease As Variant
    Dim greatestDecrease As Variant
    Dim greatestTotalVolume As Variant
    Dim tickers(3) As String
    rowLimit = Cells(Rows.Count, "I").End(xlUp).row
    greatestIncrease = 0
    greatestDecrease = 0
    greatestTotalVolume = 0
    rowLimit = Cells(Rows.Count, "A").End(xlUp).row
    openYear = Cells(2, 3).value
    totalStock = Cells(2, 7).value
    tickers(0) = Cells(2, 1).value
    tickers(1) = Cells(2, 1).value
    tickers(2) = Cells(2, 1).value
    Range("I1").value = "Ticker"
    Range("J1").value = "Yearly Changes"
    Range("K1").value = "Percentage Changes"
    Range("L1").value = "Total Stock Volume"
    Range("O2").value = "Greatest % Increase"
    Range("O3").value = "Greatest % Decrease"
    Range("O4").value = "Greatest Total Volume"
    Range("P1").value = "Ticker"
    Range("Q1").value = "Value"
    rowIndex = 2
    For row = 2 To rowLimit
        If Cells(row, 1).value = Cells(row + 1, 1).value Then
            totalStock = totalStock + Cells(row, 7)
            If openYear = 0 Then
                openYear = Cells(row, 3).value
            End If
        End If
        If Cells(row, 1).value <> Cells(row + 1, 1).value Then
            closeYear = Cells(row, 6).value
            Cells(rowIndex, 9).value = Cells(row, 1).value
            change = closeYear - openYear
            Cells(rowIndex, 10).value = change
            If change < 0 Then
                Cells(rowIndex, 10).Interior.ColorIndex = 3
            Else
                Cells(rowIndex, 10).Interior.ColorIndex = 4
            End If
            If change <> 0 Then
                If (change / openYear) * 100 > greatestIncrease Then
                    greatestIncrease = (change / openYear) * 100
                    tickers(0) = Cells(row, 1).value
                ElseIf (change / openYear) * 100 < greatestDecrease Then
                    greatestDecrease = (change / openYear) * 100
                    tickers(1) = Cells(row, 1).value
                End If
                 Cells(rowIndex, 11).value = (change / openYear) * 100 & "%"
            Else
                Cells(rowIndex, 11).value = 0
            End If
            totalStock = totalStock + Cells(row, 7)
            If totalStock > greatestTotalVolume Then
                greatestTotalVolume = totalStock
                tickers(2) = Cells(row, 1).value
            End If
            Cells(rowIndex, 12).value = totalStock
            rowIndex = rowIndex + 1
            openYear = 0
            closeYear = 0
            totalStock = 0
        End If
    Next row
        Range("P2").value = tickers(0)
        Range("Q2").value = greatestIncrease & "%"
        Range("P3").value = tickers(1)
        Range("Q3").value = greatestDecrease & "%"
        Range("P4").value = tickers(2)
        Range("Q4").value = greatestTotalVolume
End Sub


