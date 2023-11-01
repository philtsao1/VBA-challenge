# VBA-challenge

Instructions
Create a script that loops through all the stocks for one year and outputs the following information:

The ticker symbol

Yearly change from the opening price at the beginning of a given year to the closing price at the end of that year.

The percentage change from the opening price at the beginning of a given year to the closing price at the end of that year.

Add functionality to your script to return the stock with the "Greatest % increase", "Greatest % decrease", and "Greatest total volume". The solution should match the following image:

Make the appropriate adjustments to your VBA script to enable it to run on every worksheet (that is, every year) at once.

# 2018:
![image](https://github.com/philtsao1/VBA-challenge/assets/34148122/90845099-d0e8-45b1-9a9a-348572ac0cdc)

# 2019:
![image](https://github.com/philtsao1/VBA-challenge/assets/34148122/a57dda0b-5cbc-4667-880f-8d1c2a60e941)

# 2020:
![image](https://github.com/philtsao1/VBA-challenge/assets/34148122/aaaf0c0c-ad57-4b34-89d8-32435ad7576e)

```
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
```
