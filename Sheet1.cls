VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub CalculateQuarterlyChangesPerSheet()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim currentRow As Long
    Dim outputRow As Long
    Dim ticker As String
    Dim openPrice As Double
    Dim closePrice As Double
    Dim quarterlyChange As Double
    Dim percentChange As Double
    Dim totalVolume As Double
    Dim tickerStartRow As Long
    Dim tickerEndRow As Long
    Dim startPrice As Double
    Dim endPrice As Double
    
    Dim maxPercentIncrease As Double
    Dim minPercentIncrease As Double
    Dim maxTotalVolume As Double
    Dim maxTicker As String
    Dim minTicker As String
    Dim volTicker As String
    
    ' Loop through each sheet in the workbook
    For Each ws In ThisWorkbook.Sheets
        ' Skip the summary sheet if it exists
        If ws.Name <> "Quarterly Results" Then
            ' Initialize variables
            currentRow = 2
            outputRow = 2
            lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
            
            ' Initialize extreme values
            maxPercentIncrease = -1000000
            minPercentIncrease = 1000000
            maxTotalVolume = 0
            
            ' Add headers to the current sheet if not already present
            With ws
                If .Cells(1, 8).Value <> "Ticker" Then
                    .Cells(1, 8).Value = "Ticker"
                    .Cells(1, 9).Value = "Quarterly Change"
                    .Cells(1, 10).Value = "Percent Change"
                    .Cells(1, 11).Value = "Total Stock Volume"
                End If
            End With
            
            ' Loop through each row in the data sheet
            Do While currentRow <= lastRow
                ticker = ws.Cells(currentRow, 1).Value
                tickerStartRow = currentRow
                totalVolume = 0
                startPrice = ws.Cells(tickerStartRow, 3).Value
                
                ' Loop through each row for the current ticker
                Do While currentRow <= lastRow And ws.Cells(currentRow, 1).Value = ticker
                    totalVolume = totalVolume + ws.Cells(currentRow, 7).Value
                    currentRow = currentRow + 1
                Loop
                
                tickerEndRow = currentRow - 1
                endPrice = ws.Cells(tickerEndRow, 6).Value
                quarterlyChange = endPrice - startPrice
                percentChange = (quarterlyChange / startPrice) * 100
                
                ' Output results to the same sheet
                With ws
                    .Cells(outputRow, 8).Value = ticker
                    .Cells(outputRow, 9).Value = quarterlyChange
                    .Cells(outputRow, 10).Value = Format(percentChange, "0.00") & "%"
                    .Cells(outputRow, 11).Value = totalVolume
                    
                    ' Added Conditional formatting to Quarterly Change column
                    If quarterlyChange > 0 Then
                        .Cells(outputRow, 9).Interior.Color = RGB(0, 255, 0) ' Green
                    ElseIf quarterlyChange < 0 Then
                        .Cells(outputRow, 9).Interior.Color = RGB(255, 0, 0) ' Red
                    Else
                        .Cells(outputRow, 9).Interior.ColorIndex = xlNone ' No color
                    End If
                End With
                
                ' Check for greatest % increase, % decrease, and total volume
                If percentChange > maxPercentIncrease Then
                    maxPercentIncrease = percentChange
                    maxTicker = ticker
                End If
                
                If percentChange < minPercentIncrease Then
                    minPercentIncrease = percentChange
                    minTicker = ticker
                End If
                
                If totalVolume > maxTotalVolume Then
                    maxTotalVolume = totalVolume
                    volTicker = ticker
                End If
                
                outputRow = outputRow + 1
            Loop
            
            ' Output the greatest % increase, % decrease, and total volume to the current sheet
            With ws
                .Cells(outputRow, 8).Value = "Greatest % Increase"
                .Cells(outputRow, 9).Value = maxTicker
                .Cells(outputRow, 10).Value = Format(maxPercentIncrease, "0.00") & "%"
                
                .Cells(outputRow + 1, 8).Value = "Greatest % Decrease"
                .Cells(outputRow + 1, 9).Value = minTicker
                .Cells(outputRow + 1, 10).Value = Format(minPercentIncrease, "0.00") & "%"
                
                .Cells(outputRow + 2, 8).Value = "Greatest Total Volume"
                .Cells(outputRow + 2, 9).Value = volTicker
                .Cells(outputRow + 2, 10).Value = maxTotalVolume
            End With
            
        End If
    Next ws
    
End Sub

