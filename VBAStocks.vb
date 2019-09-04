Public Sub stockAnalysis()

    Dim annualChange As Double
    Dim closePrice As Double
    Dim lastRow As Double
    Dim openPrice As Double
    Dim stockTicker As String
    Dim totalChangesTracker As Double

    For Each ws In Worksheets
    
        stockVolume = 0
        totalChangesTracker = 2
        lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        ws.Range("K1").Value = "Percent Change"
        ws.Range("I1").Value = "Stock Ticker"
        ws.Range("L1").Value = "Total Stock Volume"
        ws.Range("J1").Value = "Yearly Change"
    
        For marketAnalysis = 2 To lastRow
            If ws.Cells(marketAnalysis + 1, 1).Value <> ws.Cells(marketAnalysis, 1).Value Then
                stockTicker = ws.Cells(marketAnalysis, 1).Value
                stockVolume = stockVolume + ws.Cells(marketAnalysis, 7).Value
                closePrice = ws.Cells(marketAnalysis, 6).Value
                annualChange = closePrice - openPrice
                If openPrice <= 0 Then
                    openPrice = 1
                End If
                percentChange = annualChange / openPrice
                ws.Range("I" & totalChangesTracker).Value = stockTicker
                ws.Range("J" & totalChangesTracker).Value = annualChange
                ws.Range("J" & totalChangesTracker).NumberFormat = "0.00"
                If annualChange > 0 Then
                    ws.Range("J" & totalChangesTracker).Interior.ColorIndex = 4
                    Else
                    ws.Range("J" & totalChangesTracker).Interior.ColorIndex = 3
                    End If
                ws.Range("K" & totalChangesTracker).Value = percentChange
                ws.Range("K" & totalChangesTracker).NumberFormat = "0.00%"
                ws.Range("L" & totalChangesTracker).Value = stockVolume
                totalChangesTracker = totalChangesTracker + 1
                stockVolume = 0
                openPrice = 0
            Else
                stockVolume = stockVolume + ws.Cells(marketAnalysis, 7).Value
                If openPrice = 0 Then
                    openPrice = ws.Cells(marketAnalysis, 3).Value
                End If
            End If
        Next marketAnalysis

    Next ws
    
MsgBox ("Calculations Complete")

End Sub
