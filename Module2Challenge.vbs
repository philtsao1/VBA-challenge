Attribute VB_Name = "Module1"
Sub Module2Challenge()

Dim sheet As Worksheet

For Each sheet In ActiveWorkbook.Worksheets
    Dim ticker As String
    Dim volume As Long
    Dim lastRow As Double
    Dim openPrice As Double
    openPrice = sheet.Cells(2, 3).Value
    Dim closePrice As Double
    Dim yearlyChange As Double
    Dim percentChange As Double
    Dim totalStockVolume As Double
    Dim greatestPercentIncrease As Double
    Dim greatestPercentDecrease As Double
    Dim greatestTotalVolume As Double
    lastRow = sheet.Cells(Rows.Count, 1).End(xlUp).Row
    sheet.Range("I1").Value = "Ticker"
    sheet.Range("J1").Value = "Yearly Change"
    sheet.Range("K1").Value = "Percent Change"
    sheet.Range("L1").Value = "Total Stock Volume"
    sheet.Range("O2").Value = "Greatest % Increase"
    sheet.Range("O3").Value = "Greatest % Decrease"
    sheet.Range("O4").Value = "Greatest Total Volume"
    sheet.Range("P1").Value = "Ticker"
    sheet.Range("Q1").Value = "Value"
    greatestPercentIncrease = 0
    greatestPercentDecrease = 0
    Dim lastRowK As Double

    Dim tickerRowCount As Double
    tickerRowCount = 1
    For i = 2 To lastRow
        If sheet.Cells(i + 1, 1).Value <> sheet.Cells(i, 1).Value Then
            tickerRowCount = tickerRowCount + 1
            ticker = sheet.Cells(i, 1).Value
            sheet.Cells(tickerRowCount, "I").Value = ticker
            closePrice = sheet.Cells(i, 6).Value
            yearlyChange = (closePrice - openPrice)
            sheet.Cells(tickerRowCount, "J").Value = yearlyChange
        
            If yearlyChange < 0 Then
                sheet.Cells(tickerRowCount, "J").Interior.ColorIndex = 3
            ElseIf yearlyChange > 0 Then
                sheet.Cells(tickerRowCount, "J").Interior.ColorIndex = 4
            End If
             
          
        If yearlyChange = 0 Or openPrice = 0 Then
            sheet.Cells(tickerRowCount, "K").Value = 0
        Else
            sheet.Cells(tickerRowCount, "K").Value = FormatPercent((yearlyChange / openPrice), 2)
        End If
        
        With ActiveSheet
            lastRowK = .Cells(.Rows.Count, "K").End(xlUp).Row
        End With
        
        totalStockVolume = totalStockVolume + sheet.Cells(i, 7)
        

        openPrice = sheet.Cells(i + 1, 3).Value
        
        
        sheet.Cells(tickerRowCount, "L").Value = totalStockVolume
        
        totalStockVolume = 0
        
        
        'Loop through added columns and find the largest and lowest
        Dim j As Integer
        For j = 2 To lastRowK
            If greatestPercentIncrease < sheet.Cells(j, 11).Value Then
                greatestPercentIncrease = sheet.Cells(j, 11).Value
                sheet.Cells(2, 17).Value = FormatPercent(sheet.Cells(j, "K").Value, 2)
                sheet.Cells(2, 16).Value = sheet.Cells(j, "I").Value
            End If
            If greatestPercentDecrease > sheet.Cells(j, 11).Value Then
                greatestPercentDecrease = sheet.Cells(j, 11).Value
                sheet.Cells(3, 17).Value = FormatPercent(sheet.Cells(j, "K").Value, 2)
                sheet.Cells(3, 16).Value = sheet.Cells(j, "I").Value
            End If
            If greatestTotalVolume < sheet.Cells(j, 12).Value Then
                greatestTotalVolume = sheet.Cells(j, 12).Value
                sheet.Cells(4, 17).Value = sheet.Cells(j, "L").Value
                sheet.Cells(4, 16).Value = sheet.Cells(j, "I").Value
            End If
        Next j
        
        Else
            totalStockVolume = totalStockVolume + sheet.Cells(i, 7).Value
        End If
        
        
        Next i

Next sheet
End Sub

