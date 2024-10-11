VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub AnalyzeStockData()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim ticker As String
    Dim openPrice As Double
    Dim closePrice As Double
    Dim volume As Double
    Dim quarterlyChange As Double
    Dim percentChange As Double
    Dim greatestIncrease As Double
    Dim greatestDecrease As Double
    Dim greatestVolume As Double
    Dim greatestIncreaseTicker As String
    Dim greatestDecreaseTicker As String
    Dim greatestVolumeTicker As String
    Dim outputRow As Long

    ' Loop through each worksheet (each quarter)
    For Each ws In ThisWorkbook.Worksheets
        lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
        
        ' Initialize variables
        greatestIncrease = -1
        greatestDecrease = 1
        greatestVolume = 0
        
        ' Loop through each row of stock data
        For i = 2 To lastRow
            ticker = ws.Cells(i, 1).Value
            openPrice = ws.Cells(i, 3).Value
            closePrice = ws.Cells(i, 6).Value
            volume = ws.Cells(i, 7).Value
            
            ' Calculate quarterly change and percentage change
            quarterlyChange = closePrice - openPrice
            If openPrice <> 0 Then
                percentChange = (quarterlyChange / openPrice) * 100
            Else
                percentChange = 0
            End If
            
            ' Output the results to the specified columns
            ws.Cells(i, 9).Value = ticker ' Ticker in Column I
            ws.Cells(i, 10).Value = quarterlyChange ' Quarterly Change in Column J
            ws.Cells(i, 11).Value = percentChange ' Percent Change in Column K
            ws.Cells(i, 12).Value = volume ' Total Stock Volume in Column L
            
            ' Check for greatest increase, decrease, and volume
            If percentChange > greatestIncrease Then
                greatestIncrease = percentChange
                greatestIncreaseTicker = ticker
            End If
            
            If percentChange < greatestDecrease Then
                greatestDecrease = percentChange
                greatestDecreaseTicker = ticker
            End If
            
            If volume > greatestVolume Then
                greatestVolume = volume
                greatestVolumeTicker = ticker
            End If
        Next i
        
        ' Output the results for greatest increase, decrease, and volume
        outputRow = 2 ' Starting output row for Tickers
        ws.Cells(outputRow, 17).Value = greatestIncreaseTicker ' Ticker for Greatest % Increase in Column Q
        ws.Cells(outputRow, 18).Value = greatestIncrease ' Value for Greatest % Increase in Column R
        
        outputRow = outputRow + 1
        ws.Cells(outputRow, 17).Value = greatestDecreaseTicker ' Ticker for Greatest % Decrease in Column Q
        ws.Cells(outputRow, 18).Value = greatestDecrease ' Value for Greatest % Decrease in Column R
        
        outputRow = outputRow + 1
        ws.Cells(outputRow, 17).Value = greatestVolumeTicker ' Ticker for Greatest Total Volume in Column Q
        ws.Cells(outputRow, 18).Value = greatestVolume ' Value for Greatest Total Volume in Column R
    Next ws
End Sub
