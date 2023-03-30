Sub stocks()

Dim ws As Worksheet

For Each ws In ThisWorkbook.Worksheets

    Dim i, resultRow, lastRow, runningVol, stockOpen, greatestIncrease, greatestDecrease, greatestVolume As Double
    Dim increaseTicker, decreaseTicker, volumeTicker As String
        
    lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    stockOpen = 0
    runningVol = 0
    resultRow = 2
    greatestIncrease = 0
    greatestDecrease = 0
    greatestVolume = 0
    
    ' the below is just creating headers for where my data will go
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    
    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(1, 17).Value = "Value"
    
    ws.Cells(2, 15).Value = "Greatest % Increase"
    ws.Cells(3, 15).Value = "Greatest % Decrease"
    ws.Cells(4, 15).Value = "Greatest Total Volume"
    
    ws.Range("1:1").Font.Bold = True
    ws.Range("O2:O4").Font.Bold = True
    ws.Columns("I:O").AutoFit
    
    For i = 2 To lastRow
        If (stockOpen = 0) Then
            stockOpen = ws.Cells(i, 3).Value
        End If
        
        runningVol = runningVol + ws.Cells(i, 7).Value
        
        If (ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value) Then
            ws.Cells(resultRow, 9).Value = ws.Cells(i, 1).Value ' write ticker symbol
            
            ws.Cells(resultRow, 10).Value = ws.Cells(i, 6).Value - stockOpen ' write yearly change
            ws.Cells(resultRow, 10).NumberFormat = "$#,##0.00" ' format cell as currency
            
            If (ws.Cells(resultRow, 10).Value < 0) Then
                ws.Cells(resultRow, 10).Interior.ColorIndex = 3 ' color cell red if negative
            Else: ws.Cells(resultRow, 10).Interior.ColorIndex = 4 ' else cell color green
            End If
            
            ws.Cells(resultRow, 11).Value = ws.Cells(resultRow, 10).Value / stockOpen ' write percent change
            ws.Cells(resultRow, 11).NumberFormat = "0.00%" ' format cell as percentage with 2 decimal places
            
            If (ws.Cells(resultRow, 11).Value > greatestIncrease) Then
                greatestIncrease = ws.Cells(resultRow, 11).Value ' check greatest increase, overwrite if current result is bigger
                increaseTicker = ws.Cells(i, 1).Value
            End If
            
            If (ws.Cells(resultRow, 11).Value < greatestDecrease) Then
                greatestDecrease = ws.Cells(resultRow, 11).Value ' check greatest decrease, overwrite if current result is bigger
                decreaseTicker = ws.Cells(i, 1).Value
            End If
            
            ws.Cells(resultRow, 12).Value = runningVol ' write total stock volume
            
            If (runningVol > greatestVolume) Then
                greatestVolume = runningVol ' check greatest volume, overwrite if current result is bigger
                volumeTicker = ws.Cells(i, 1).Value
            End If
            
            stockOpen = 0 ' reset stock starting value
            runningVol = 0 ' reset stock volume counter
            resultRow = resultRow + 1 ' increment row of printed results
        End If
            
    Next i
    
    ws.Cells(2, 16).Value = increaseTicker
    ws.Cells(2, 17).Value = greatestIncrease
    ws.Cells(2, 17).NumberFormat = "0.00%" ' format cell as percentage with 2 decimal places
    
    ws.Cells(3, 16).Value = decreaseTicker
    ws.Cells(3, 17).Value = greatestDecrease
    ws.Cells(3, 17).NumberFormat = "0.00%" ' format cell as percentage with 2 decimal places
    
    ws.Cells(4, 16).Value = volumeTicker
    ws.Cells(4, 17).Value = greatestVolume
    
Next ws

End Sub

