Sub WorksheetLoop()

    ' Declare Current as a worksheet object variable.
    Dim Current As Worksheet

    ' Loop through all of the worksheets in the active workbook.
    For Each Current In Worksheets
        ' MsgBox ("Year " & Current.name)
        Current.Activate
        
        ' Ticker Summary Table
    
        'Insert column names
        Range("I1").Value = "Ticker"
        Range("J1").Value = "Yearly Change"
        Range("K1").Value = "Percent Change"
        Range("L1").Value = "Total Stock Volume"
    
        ' Find unique tickers
        Dim unique As Variant
        ' Declaring variables
        Dim o As Object, z As Variant, t As Long, lr As Long
        'setting values
        Set o = CreateObject("Scripting.Dictionary")
        lr = Cells(Rows.Count, 1).End(xlUp).Row
        z = Range("A2:A" & lr)
        'for next loop
        For t = 1 To UBound(z, 1)
            o(z(t, 1)) = 1
        Next t
        
        unique = Application.Transpose(o.keys)
        
        ' Find last row of sheet
        lastRow = Cells(Rows.Count, 1).End(xlUp).Row
        
        ' Loop through each ticker value
        For i = LBound(unique) To UBound(unique)
        ' For i = LBound(unique) To 10
            ticker = unique(i, 1)
            totalStockVolumn = 0
            maxDate = ""
            minDate = ""
            
            maxDateRow = 0
            minDateRow = 0
            
            For j = 2 To lastRow
                ' Calculate the total stock volumn
                rowTicker = Range("A" & j).Value
                If rowTicker = ticker Then
                    ' Find the max and min date
                    givenDate = Range("B" & j).Value
                    
                    If maxDate = "" Then
                    maxDate = givenDate
                    maxRow = j
                    End If
                    
                    If minDate = "" Then
                    minDate = givenDate
                    minRow = j
                    End If
                    
                    If givenDate > maxDate Then
                    maxDate = givenDate
                    maxRow = j
                    ElseIf givenDate < minDate Or minDate = "" Then
                    minDate = givenDate
                    minRow = j
                    End If
                    
                    ' Find the stock volumn
                    vol = Range("G" & j).Value
                    totalStockVolumn = totalStockVolumn + vol
                End If
            Next j
            
    
            ' Find open and close stock price
            openPrice = Range("C" & minRow)
            closePrice = Range("F" & maxRow)
            yearlyChange = closePrice - openPrice
            percentChange = ((closePrice - openPrice) / openPrice * 100) & "%"
            
            ' Write the ticker name and total stock volumn
            Range("I" & i + 1) = ticker
            Range("L" & i + 1) = totalStockVolumn
            Range("J" & i + 1) = yearlyChange
            Range("K" & i + 1) = percentChange
            
            If yearlyChange < 0 Then
                Range("J" & i + 1).Interior.ColorIndex = 3
            Else
                Range("J" & i + 1).Interior.ColorIndex = 4
            End If
            
        Next i
        
        ' Greatest table
        ' Insert column names
        Range("P1").Value = "Ticker"
        Range("Q1").Value = "Value"
        
        Range("O2").Value = "Greatest % Increase"
        Range("O3").Value = "Greatest % Decrease"
        Range("O4").Value = "Greatest Total Volume"
        
        ' Find last row of summary table
        lastRow = Cells(Rows.Count, 9).End(xlUp).Row
        
        greatestIncrease = -2
        greatestDecrease = 2
        greatestTotalVolume = 0
        
        greatestIncreaseTicker = ""
        greatestDecreaseTicker = ""
        greatestTotalVolumeTicker = ""
    
        
        For k = 2 To lastRow
            ' Calculate the total stock volume
            rowTicker = Range("I" & k).Value
            percentChange = Range("K" & k).Value
            totalVolume = Range("L" & k).Value
            
            If percentChange > greatestIncrease Then
                greatestIncrease = percentChange
                greatestIncreaseTicker = rowTicker
            End If
            If percentChange < greatestDecrease Then
                greatestDecrease = percentChange
                greatestDecreaseTicker = rowTicker
            End If
            If totalVolume > greatestTotalVolume Then
                greatestTotalVolume = totalVolume
                greatestTotalVolumeTicker = rowTicker
            End If
        Next k
        
        Range("P2").Value = greatestIncreaseTicker
        Range("P3").Value = greatestDecreaseTicker
        Range("P4").Value = greatestTotalVolumeTicker
        
        Range("Q2").Value = (greatestIncrease * 100) & "%"
        Range("Q3").Value = (greatestDecrease * 100) & "%"
        Range("Q4").Value = greatestTotalVolume
    
    Next

End Sub
